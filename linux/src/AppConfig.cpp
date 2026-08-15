#include "AppConfig.h"

#include <QCoreApplication>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QMimeDatabase>
#include <QMimeType>
#include <QProcess>
#include <QProcessEnvironment>
#include <QRegularExpression>
#include <QStandardPaths>
#include <QtGlobal>

#include <initializer_list>
#include <utility>

namespace {

QString firstValue(const QHash<QString, QString> &values,
                   std::initializer_list<const char *> names,
                   const QString &fallback = QString()) {
    for (const char *name : names) {
        const auto it = values.constFind(normalizeConfigKey(QString::fromLatin1(name)));
        if (it != values.cend()) {
            return it.value();
        }
    }
    return fallback;
}

QStringList splitArguments(const QString &text, const QStringList &fallback) {
    if (text.trimmed().isEmpty()) {
        return fallback;
    }
#if QT_VERSION >= QT_VERSION_CHECK(5, 15, 0)
    return QProcess::splitCommand(text);
#else
    QStringList result;
    QString current;
    bool quoted = false;
    for (const QChar ch : text) {
        if (ch == QLatin1Char('"')) {
            quoted = !quoted;
        } else if (ch.isSpace() && !quoted) {
            if (!current.isEmpty()) {
                result.append(current);
                current.clear();
            }
        } else {
            current.append(ch);
        }
    }
    if (!current.isEmpty()) {
        result.append(current);
    }
    return result.isEmpty() ? fallback : result;
#endif
}

QString stripInlineComment(const QString &value) {
    bool singleQuoted = false;
    bool doubleQuoted = false;
    bool escaped = false;
    for (int index = 0; index < value.size(); ++index) {
        const QChar ch = value.at(index);
        if (escaped) {
            escaped = false;
            continue;
        }
        if (ch == QLatin1Char('\\') && (singleQuoted || doubleQuoted)) {
            escaped = true;
            continue;
        }
        if (ch == QLatin1Char('\'') && !doubleQuoted) {
            singleQuoted = !singleQuoted;
            continue;
        }
        if (ch == QLatin1Char('"') && !singleQuoted) {
            doubleQuoted = !doubleQuoted;
            continue;
        }
        if (ch == QLatin1Char(';') && !singleQuoted && !doubleQuoted &&
            (index == 0 || value.at(index - 1).isSpace())) {
            return value.left(index).trimmed();
        }
    }
    return value.trimmed();
}

} // namespace

QString normalizeConfigKey(QString key) {
    key.remove(QLatin1Char('_'));
    key.remove(QLatin1Char('-'));
    return key.trimmed().toLower();
}

QString expandPath(const QString &value) {
    QString expanded = value.trimmed();
    const QProcessEnvironment environment = QProcessEnvironment::systemEnvironment();
    static const QRegularExpression variable(
        QStringLiteral(R"(\$(?:\{([A-Za-z_][A-Za-z0-9_]*)\}|([A-Za-z_][A-Za-z0-9_]*)))"));
    qsizetype offset = 0;
    while (true) {
        const QRegularExpressionMatch match = variable.match(expanded, offset);
        if (!match.hasMatch()) {
            break;
        }
        const QString name = match.captured(1).isEmpty() ? match.captured(2) : match.captured(1);
        if (environment.contains(name)) {
            const QString replacement = environment.value(name);
            expanded.replace(match.capturedStart(), match.capturedLength(), replacement);
            offset = match.capturedStart() + replacement.size();
        } else {
            // An unknown variable is more useful left intact than silently erased.
            offset = match.capturedEnd();
        }
    }
    if (expanded == QStringLiteral("~")) {
        expanded = QDir::homePath();
    } else if (expanded.startsWith(QStringLiteral("~/"))) {
        expanded = QDir::home().filePath(expanded.mid(2));
    }
    return QDir::cleanPath(expanded);
}

bool parseConfigBool(const QString &value, bool fallback) {
    const QString normalized = value.trimmed().toLower();
    static const QSet<QString> trueValues{
        QStringLiteral("1"), QStringLiteral("true"), QStringLiteral("yes"),
        QStringLiteral("on"), QStringLiteral("y"), QStringLiteral("enabled")};
    static const QSet<QString> falseValues{
        QStringLiteral("0"), QStringLiteral("false"), QStringLiteral("no"),
        QStringLiteral("off"), QStringLiteral("n"), QStringLiteral("disabled")};
    if (trueValues.contains(normalized)) {
        return true;
    }
    if (falseValues.contains(normalized)) {
        return false;
    }
    return fallback;
}

bool AppConfig::load(const QString &explicitPath, AppConfig &config, QString &error) {
    config = AppConfig{};
    error.clear();

    QStringList candidates;
    if (!explicitPath.isNull()) {
        const QString path = expandPath(explicitPath);
        if (!QFileInfo(path).isFile()) {
            error = QStringLiteral("Configuration file does not exist or is not a file: %1").arg(path);
            return false;
        }
        candidates.append(path);
    } else if (qEnvironmentVariableIsSet("MEDIA_EXPLORER_CONFIG")) {
        const QString path = expandPath(qEnvironmentVariable("MEDIA_EXPLORER_CONFIG"));
        if (!QFileInfo(path).isFile()) {
            error = QStringLiteral("MEDIA_EXPLORER_CONFIG does not name an existing file: %1").arg(path);
            return false;
        }
        candidates.append(path);
    }
    candidates.append(QDir::home().filePath(QStringLiteral(".config/media-explorer/mediaexplorer.ini")));
    candidates.append(QDir(QCoreApplication::applicationDirPath()).filePath(QStringLiteral("mediaexplorer.ini")));

    QString selected;
    for (const QString &candidate : std::as_const(candidates)) {
        if (QFileInfo(candidate).isFile()) {
            selected = candidate;
            break;
        }
    }
    if (selected.isEmpty()) {
        config.ffprobeAvailable = !QStandardPaths::findExecutable(config.ffprobePath).isEmpty();
        config.ffmpegAvailable = !QStandardPaths::findExecutable(config.ffmpegPath).isEmpty();
        return true;
    }

    QFile file(selected);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        error = QStringLiteral("Cannot read configuration %1: %2").arg(selected, file.errorString());
        return false;
    }
    QString text = QString::fromUtf8(file.readAll());
    if (!text.isEmpty() && text.front().unicode() == 0xfeff) {
        text.remove(0, 1);
    }
    QHash<QString, QString> values;
    const QStringList lines = text.split(QRegularExpression(QStringLiteral("[\r\n]+")));
    for (QString line : lines) {
        line = line.trimmed();
        if (line.isEmpty() || line.startsWith(QLatin1Char('#')) ||
            line.startsWith(QLatin1Char(';')) || line.startsWith(QLatin1Char('['))) {
            continue;
        }
        const int equals = line.indexOf(QLatin1Char('='));
        if (equals <= 0) {
            continue;
        }
        values.insert(normalizeConfigKey(line.left(equals)),
                      stripInlineComment(line.mid(equals + 1)));
    }

    config.mappedRoot = expandPath(firstValue(
        values, {"mapped_root", "mapped_drive_root"}, config.mappedRoot));
    static const QRegularExpression mappingKey(
        QStringLiteral(R"(^(?:mapping|mappeddrive|mappedshare)([a-z0-9]+)$)"));
    for (auto it = values.cbegin(); it != values.cend(); ++it) {
        const QRegularExpressionMatch match = mappingKey.match(it.key());
        if (match.hasMatch() && !it.value().isEmpty()) {
            config.mappedShares.insert(match.captured(1).toLower(), it.value());
        }
    }
    const QString startPath = firstValue(values, {"start_path"});
    config.startPath = startPath.isEmpty() ? QString() : expandPath(startPath);
    config.ffprobePath = expandPath(firstValue(values, {"ffprobe_path"}, config.ffprobePath));
    config.ffmpegPath = expandPath(firstValue(values, {"ffmpeg_path"}, config.ffmpegPath));
    config.ffprobeAvailable = parseConfigBool(
        firstValue(values, {"ffprobe_available"}, config.ffprobeAvailable ? "true" : "false"),
        config.ffprobeAvailable);
    config.ffmpegAvailable = parseConfigBool(
        firstValue(values, {"ffmpeg_available"}, config.ffmpegAvailable ? "true" : "false"),
        config.ffmpegAvailable);
    config.videoCombineAvailable = parseConfigBool(
        firstValue(values, {"video_combine_available"},
                   config.videoCombineAvailable ? "true" : "false"),
        config.videoCombineAvailable);
    const QString upscaleDirectory = firstValue(values, {"upscale_directory", "upscaledirectory"});
    const QString topazQueue = firstValue(values, {"topaz_upscale_queue", "topazupscalequeue"});
    const QString loggingPath = firstValue(values, {"logging_path", "loggingpath"});
    config.upscaleDirectory = upscaleDirectory.isEmpty() ? QString() : expandPath(upscaleDirectory);
    config.topazUpscaleQueue = topazQueue.isEmpty() ? QString() : expandPath(topazQueue);
    config.loggingPath = loggingPath.isEmpty() ? QString() : expandPath(loggingPath);
    config.showHidden = parseConfigBool(firstValue(values, {"show_hidden"}), config.showHidden);
    config.followSymlinks = parseConfigBool(firstValue(values, {"follow_symlinks"}), config.followSymlinks);
    config.useTrash = parseConfigBool(firstValue(values, {"use_trash"}), config.useTrash);

    bool limitOk = false;
    const int limit = firstValue(values, {"metadata_prefetch_limit"}, QStringLiteral("500")).toInt(&limitOk);
    config.metadataPrefetchLimit = limitOk ? qBound(0, limit, 10000) : 500;
    config.ffprobeArgs = splitArguments(firstValue(values, {"ffprobe_args"}), {});
    config.ffmpegArgs = splitArguments(firstValue(values, {"ffmpeg_args"}), {});
    config.vlcArgs = splitArguments(firstValue(values, {"vlc_args"}), config.vlcArgs);

    const QString extensions = firstValue(values, {"video_extensions"});
    if (!extensions.isEmpty()) {
        QSet<QString> parsed;
        for (QString extension : extensions.split(
                 QRegularExpression(QStringLiteral("[,;\\s]+")), Qt::SkipEmptyParts)) {
            extension = extension.trimmed().toLower();
            if (!extension.startsWith(QLatin1Char('.'))) {
                extension.prepend(QLatin1Char('.'));
            }
            parsed.insert(extension);
        }
        if (!parsed.isEmpty()) {
            config.videoExtensions = parsed;
        }
    }
    config.configPath = selected;
    return true;
}

bool isVideoFile(const QString &path, const QSet<QString> &extensions) {
    const QString suffix = QFileInfo(path).suffix().toLower();
    return !suffix.isEmpty() && extensions.contains(QLatin1Char('.') + suffix);
}

QString formatBytes(qint64 bytes) {
    if (bytes < 0) {
        return {};
    }
    if (bytes < 1024) {
        return QStringLiteral("%1 B").arg(bytes);
    }
    double value = static_cast<double>(bytes);
    const QStringList units{QStringLiteral("KiB"), QStringLiteral("MiB"), QStringLiteral("GiB"),
                            QStringLiteral("TiB"), QStringLiteral("PiB")};
    for (const QString &unit : units) {
        value /= 1024.0;
        if (value < 1024.0 || unit == units.back()) {
            return QStringLiteral("%1 %2").arg(value, 0, 'f', 1).arg(unit);
        }
    }
    return {};
}

QString formatDuration(double seconds) {
    if (seconds < 0.0 || !qIsFinite(seconds)) {
        return {};
    }
    const qint64 total = qRound64(seconds);
    const qint64 hours = total / 3600;
    const qint64 minutes = (total / 60) % 60;
    const qint64 secs = total % 60;
    if (hours > 0) {
        return QStringLiteral("%1:%2:%3")
            .arg(hours).arg(minutes, 2, 10, QLatin1Char('0')).arg(secs, 2, 10, QLatin1Char('0'));
    }
    return QStringLiteral("%1:%2").arg(minutes).arg(secs, 2, 10, QLatin1Char('0'));
}

QString formatClock(qint64 milliseconds) {
    return formatDuration(qMax<qint64>(0, milliseconds) / 1000);
}

QString fileKind(const QString &path) {
    const QString suffix = QFileInfo(path).suffix().toUpper();
    const QMimeType mime = QMimeDatabase().mimeTypeForFile(path, QMimeDatabase::MatchExtension);
    if (mime.name().startsWith(QStringLiteral("video/"))) {
        return suffix.isEmpty() ? QStringLiteral("Video") : suffix + QStringLiteral(" video");
    }
    return suffix.isEmpty() ? QStringLiteral("File") : suffix + QStringLiteral(" file");
}

QString executableResolution(const QString &command) {
    if (command.contains(QLatin1Char('/'))) {
        const QFileInfo info(command);
        return info.isFile() && info.isExecutable() ? info.absoluteFilePath() : QString();
    }
    return QStandardPaths::findExecutable(command);
}
