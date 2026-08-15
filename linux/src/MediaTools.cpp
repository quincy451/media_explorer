#include "MediaTools.h"

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QLocale>
#include <QUuid>

#include <cerrno>
#include <cstring>
#include <exception>

#if defined(Q_OS_UNIX)
#include <fcntl.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <unistd.h>
#endif

#if defined(Q_OS_LINUX)
#include <sys/syscall.h>
#ifndef RENAME_NOREPLACE
#define RENAME_NOREPLACE (1U << 0)
#endif
#endif

namespace MediaTools {
namespace {

QString normalizedPathKey(const QString &path) {
    return QDir::cleanPath(QFileInfo(path).absoluteFilePath());
}

bool pathExistsWithoutFollowingFinalLink(const QString &path) {
#if defined(Q_OS_UNIX)
    struct stat status {};
    const QByteArray encoded = QFile::encodeName(path);
    return ::lstat(encoded.constData(), &status) == 0;
#else
    const QFileInfo info(path);
    return info.exists() || info.isSymLink();
#endif
}

struct SplitName {
    QString stem;
    QString extension;
};

SplitName splitLastExtension(const QString &fileName) {
    const int dot = fileName.lastIndexOf(QLatin1Char('.'));
    if (dot > 0) {
        return {fileName.left(dot), fileName.mid(dot)};
    }
    return {fileName, QString()};
}

QString sanitizedTag(QString tag) {
    tag.replace(QLatin1Char('/'), QLatin1Char('_'));
    tag.replace(QLatin1Char('\\'), QLatin1Char('_'));
    tag.replace(QChar(0), QLatin1Char('_'));
    return tag;
}

template <typename Collides>
QString chooseUniqueOutputPath(const QString &sourcePath, const QString &tag,
                               const QSet<QString> &reservedPaths, Collides collides) {
    const QFileInfo sourceInfo(sourcePath);
    const QString directory = sourceInfo.absolutePath();
    const SplitName name = splitLastExtension(sourceInfo.fileName());
    const QString base = name.stem + sanitizedTag(tag);

    QSet<QString> normalizedReserved;
    normalizedReserved.reserve(reservedPaths.size());
    for (const QString &path : reservedPaths) {
        normalizedReserved.insert(normalizedPathKey(path));
    }

    for (int suffix = 0; suffix < 10000; ++suffix) {
        const QString numbered = suffix == 0
                                     ? base
                                     : QStringLiteral("%1 (%2)").arg(base).arg(suffix);
        const QString candidate = QDir(directory).filePath(numbered + name.extension);
        if (!normalizedReserved.contains(normalizedPathKey(candidate)) && !collides(candidate)) {
            return candidate;
        }
    }
    return QString();
}

QString secondsArgument(qint64 milliseconds) {
    const double seconds = qMax<qint64>(0, milliseconds) / 1000.0;
    return QLocale::c().toString(seconds, 'f', 3);
}

QString ffconcatQuotedPath(const QString &path) {
    QString quoted;
    quoted.reserve(path.size() + 8);
    quoted += QLatin1Char('\'');
    for (const QChar ch : path) {
        if (ch == QLatin1Char('\'')) {
            // A quote cannot be escaped inside an ffconcat single-quoted token.
            // Close the token, append an escaped quote, then reopen it.
            quoted += QStringLiteral("'\\''");
        } else {
            // Backslashes are literal inside the quoted portion and must not be
            // doubled: doing so changes a valid Linux filename.
            quoted += ch;
        }
    }
    quoted += QLatin1Char('\'');
    return quoted;
}

QString utcTimestamp(const QDateTime &submittedUtc) {
    const QDateTime value = submittedUtc.isValid() ? submittedUtc : QDateTime::currentDateTimeUtc();
    return value.toUTC().toString(QStringLiteral("yyyy-MM-dd'T'HH:mm:ss.zzz'Z'"));
}

QByteArray jsonString(const QString &value) {
    const QByteArray utf8 = value.toUtf8();
    QByteArray escaped;
    escaped.reserve(utf8.size() + 8);
    escaped += '"';
    const char hex[] = "0123456789abcdef";
    for (const unsigned char byte : utf8) {
        switch (byte) {
        case '"': escaped += "\\\""; break;
        case '\\': escaped += "\\\\"; break;
        case '\b': escaped += "\\b"; break;
        case '\f': escaped += "\\f"; break;
        case '\n': escaped += "\\n"; break;
        case '\r': escaped += "\\r"; break;
        case '\t': escaped += "\\t"; break;
        default:
            if (byte < 0x20) {
                escaped += "\\u00";
                escaped += hex[(byte >> 4) & 0x0f];
                escaped += hex[byte & 0x0f];
            } else {
                escaped += static_cast<char>(byte);
            }
            break;
        }
    }
    escaped += '"';
    return escaped;
}

QString topazSchemaName(TopazProfile profile) {
    switch (profile) {
    case TopazProfile::General: return QStringLiteral("general");
    case TopazProfile::Repair: return QStringLiteral("repair");
    case TopazProfile::Stabilize: return QStringLiteral("stabilize");
    case TopazProfile::Deblur: return QStringLiteral("deblur");
    case TopazProfile::Denoise: return QStringLiteral("denoise");
    case TopazProfile::DeinterlaceRepair: return QStringLiteral("deinterlace_repair");
    case TopazProfile::Repair2Pass: return QStringLiteral("repair_2pass");
    case TopazProfile::GeneralGrain: return QStringLiteral("general_grain");
    case TopazProfile::RepairGrain: return QStringLiteral("repair_grain");
    }
    return QStringLiteral("general");
}

bool cancelled(const CancelCheck &isCancelled) {
    return isCancelled && isCancelled();
}

QString systemErrorString(int errorNumber) {
    return QString::fromLocal8Bit(std::strerror(errorNumber));
}

void retainOnce(QueueSubmitResult &result, const QString &path) {
    if (!path.isEmpty() && !result.retainedPaths.contains(path)) {
        result.retainedPaths.append(path);
    }
}

bool removeForCleanup(const QString &path, QueueSubmitResult &result) {
    if (path.isEmpty() || !pathExistsWithoutFollowingFinalLink(path)) {
        return true;
    }
    if (QFile::remove(path)) {
        return true;
    }
    retainOnce(result, path);
    return false;
}

#if defined(Q_OS_UNIX)
int openNewFile(const QString &path, mode_t mode) {
    int flags = O_WRONLY | O_CREAT | O_EXCL;
#ifdef O_CLOEXEC
    flags |= O_CLOEXEC;
#endif
    return ::open(QFile::encodeName(path).constData(), flags, mode);
}

bool regularFileNoFollow(const QString &path, QString &error) {
    struct stat status {};
    const QByteArray encoded = QFile::encodeName(path);
    if (::lstat(encoded.constData(), &status) != 0) {
        error = QStringLiteral("Could not inspect source %1: %2")
                    .arg(path, systemErrorString(errno));
        return false;
    }
    if (!S_ISREG(status.st_mode)) {
        error = QStringLiteral("Queue source must be a regular, non-symlink file: %1").arg(path);
        return false;
    }
    return true;
}

void bestEffortFsyncDirectory(const QString &directoryPath) {
    int flags = O_RDONLY;
#ifdef O_DIRECTORY
    flags |= O_DIRECTORY;
#endif
#ifdef O_CLOEXEC
    flags |= O_CLOEXEC;
#endif
    const int descriptor = ::open(QFile::encodeName(directoryPath).constData(), flags);
    if (descriptor < 0) {
        return;
    }
    (void)::fsync(descriptor);
    (void)::close(descriptor);
}
#endif

bool publishNoReplace(const QString &temporaryPath, const QString &finalPath,
                      bool &temporaryStillExists, QString &error) {
    temporaryStillExists = true;
#if defined(Q_OS_LINUX) && defined(SYS_renameat2)
    const QByteArray temporaryName = QFile::encodeName(temporaryPath);
    const QByteArray finalName = QFile::encodeName(finalPath);
    if (::syscall(SYS_renameat2, AT_FDCWD, temporaryName.constData(), AT_FDCWD,
                  finalName.constData(), RENAME_NOREPLACE) == 0) {
        temporaryStillExists = false;
        return true;
    }
    const int renameError = errno;
    if (renameError == EEXIST) {
        error = QStringLiteral("The queue job path appeared while publishing: %1").arg(finalPath);
        return false;
    }
    if (renameError != ENOSYS && renameError != EINVAL && renameError != EOPNOTSUPP
#ifdef ENOTSUP
        && renameError != ENOTSUP
#endif
    ) {
        error = QStringLiteral("Could not atomically publish %1: %2")
                    .arg(finalPath, systemErrorString(renameError));
        return false;
    }
#endif

#if defined(Q_OS_UNIX)
    const QByteArray linkTemporaryName = QFile::encodeName(temporaryPath);
    const QByteArray linkFinalName = QFile::encodeName(finalPath);
    // link(2) is an atomic create-if-absent operation. Both paths live in the
    // queue directory, so a successful link publishes the already-complete inode.
    if (::link(linkTemporaryName.constData(), linkFinalName.constData()) != 0) {
        const int linkError = errno;
        error = linkError == EEXIST
                    ? QStringLiteral("The queue job path appeared while publishing: %1").arg(finalPath)
                    : QStringLiteral("This filesystem cannot atomically publish %1 without replace: %2")
                          .arg(finalPath, systemErrorString(linkError));
        return false;
    }
    if (::unlink(linkTemporaryName.constData()) == 0) {
        temporaryStillExists = false;
    }
    return true;
#else
    Q_UNUSED(temporaryPath)
    Q_UNUSED(finalPath)
    error = QStringLiteral("Atomic no-replace queue publication is unsupported on this platform");
    return false;
#endif
}

void appendFailure(QStringList &failures, bool condition, const QString &message) {
    if (!condition) {
        failures.append(message);
    }
}

} // namespace

QString uniqueOutputPath(const QString &sourcePath, const QString &tag,
                         const QSet<QString> &reservedPaths) {
    if (sourcePath.isEmpty()) {
        return QString();
    }
    return chooseUniqueOutputPath(sourcePath, tag, reservedPaths, [](const QString &candidate) {
        return pathExistsWithoutFollowingFinalLink(candidate);
    });
}

QStringList trimFrontArguments(const QString &inputPath, const QString &outputPath,
                               qint64 positionMilliseconds) {
    return {QStringLiteral("-y"), QStringLiteral("-ss"), secondsArgument(positionMilliseconds),
            QStringLiteral("-i"), inputPath, QStringLiteral("-c"), QStringLiteral("copy"), outputPath};
}

QStringList trimEndArguments(const QString &inputPath, const QString &outputPath,
                             qint64 positionMilliseconds) {
    return {QStringLiteral("-y"), QStringLiteral("-i"), inputPath, QStringLiteral("-t"),
            secondsArgument(positionMilliseconds), QStringLiteral("-c"), QStringLiteral("copy"), outputPath};
}

QStringList horizontalFlipArguments(const QString &inputPath, const QString &outputPath) {
    return {QStringLiteral("-y"), QStringLiteral("-i"), inputPath, QStringLiteral("-vf"),
            QStringLiteral("hflip"), QStringLiteral("-c:a"), QStringLiteral("copy"), outputPath};
}

QString ffconcatContent(const QStringList &inputPaths, QString *error) {
    if (error) {
        error->clear();
    }
    if (inputPaths.isEmpty()) {
        if (error) *error = QStringLiteral("At least one input path is required");
        return QString();
    }

    QString content = QStringLiteral("ffconcat version 1.0\n");
    for (const QString &path : inputPaths) {
        if (path.isEmpty() || path.contains(QChar(0)) || path.contains(QLatin1Char('\n')) ||
            path.contains(QLatin1Char('\r'))) {
            if (error) *error = QStringLiteral("ffconcat paths must be non-empty single-line strings");
            return QString();
        }
        content += QStringLiteral("file ");
        content += ffconcatQuotedPath(path);
        content += QLatin1Char('\n');
    }
    return content;
}

QStringList combineStreamCopyArguments(const QString &listPath, const QString &outputPath) {
    return {QStringLiteral("-y"), QStringLiteral("-hide_banner"), QStringLiteral("-f"),
            QStringLiteral("concat"), QStringLiteral("-safe"), QStringLiteral("0"),
            QStringLiteral("-i"), listPath, QStringLiteral("-map"), QStringLiteral("0"),
            QStringLiteral("-c"), QStringLiteral("copy"), outputPath};
}

QStringList combineX264AacArguments(const QString &listPath, const QString &outputPath) {
    return {QStringLiteral("-y"), QStringLiteral("-hide_banner"), QStringLiteral("-f"),
            QStringLiteral("concat"), QStringLiteral("-safe"), QStringLiteral("0"),
            QStringLiteral("-i"), listPath, QStringLiteral("-map"), QStringLiteral("0:v:0"),
            QStringLiteral("-map"), QStringLiteral("0:a?"), QStringLiteral("-sn"),
            QStringLiteral("-dn"), QStringLiteral("-c:v"), QStringLiteral("libx264"),
            QStringLiteral("-preset"), QStringLiteral("fast"), QStringLiteral("-crf"),
            QStringLiteral("18"), QStringLiteral("-pix_fmt"), QStringLiteral("yuv420p"),
            QStringLiteral("-c:a"), QStringLiteral("aac"), QStringLiteral("-b:a"),
            QStringLiteral("192k"), outputPath};
}

QStringList ffprobeVideoArguments(const QString &mediaPath) {
    return {QStringLiteral("-v"), QStringLiteral("error"), QStringLiteral("-select_streams"),
            QStringLiteral("v:0"), QStringLiteral("-show_entries"),
            QStringLiteral("stream=codec_name,width,height"), QStringLiteral("-of"),
            QStringLiteral("default=noprint_wrappers=1"), mediaPath};
}

QStringList ffprobeAudioArguments(const QString &mediaPath) {
    return {QStringLiteral("-v"), QStringLiteral("error"), QStringLiteral("-select_streams"),
            QStringLiteral("a:0"), QStringLiteral("-show_entries"),
            QStringLiteral("stream=codec_name"), QStringLiteral("-of"),
            QStringLiteral("default=noprint_wrappers=1"), mediaPath};
}

QStringList ffprobeDurationArguments(const QString &mediaPath) {
    return {QStringLiteral("-v"), QStringLiteral("error"), QStringLiteral("-show_entries"),
            QStringLiteral("format=duration"), QStringLiteral("-of"),
            QStringLiteral("default=noprint_wrappers=1:nokey=1"), mediaPath};
}

QVector<TopazProfilePreset> topazProfilePresets() {
    return {
        {TopazProfile::General, QStringLiteral("General upscale"), QStringLiteral("general"), 0.0, 1},
        {TopazProfile::Repair, QStringLiteral("Repair (compression/noise/faces)"), QStringLiteral("repair"), 0.0, 1},
        {TopazProfile::Stabilize, QStringLiteral("Stabilize + upscale"), QStringLiteral("stabilize"), 0.0, 1},
        {TopazProfile::Deblur, QStringLiteral("Motion Deblur + upscale"), QStringLiteral("deblur"), 0.0, 1},
        {TopazProfile::Denoise, QStringLiteral("Denoise-heavy + upscale"), QStringLiteral("denoise"), 0.0, 1},
        {TopazProfile::DeinterlaceRepair, QStringLiteral("Deinterlace + Repair + upscale (rare)"),
         QStringLiteral("deinterlace_repair"), 0.0, 1},
        {TopazProfile::Repair2Pass, QStringLiteral("Repair+ (2-pass) + upscale (rare)"),
         QStringLiteral("repair_2pass"), 0.0, 1},
        {TopazProfile::GeneralGrain, QStringLiteral("General + Grain + upscale"),
         QStringLiteral("general_grain"), 0.01, 1},
        {TopazProfile::RepairGrain, QStringLiteral("Repair + Grain + upscale"),
         QStringLiteral("repair_grain"), 0.01, 1},
    };
}

TopazOptions topazOptionsForProfile(TopazTarget target, TopazProfile profile) {
    TopazOptions options;
    options.target = target;
    options.profile = profile;
    for (const TopazProfilePreset &preset : topazProfilePresets()) {
        if (preset.profile == profile) {
            options.grain = preset.grain;
            options.grainSize = preset.grainSize;
            break;
        }
    }
    return options;
}

Iw3Options iw3OptionsForPreset(Iw3Preset preset) {
    Iw3Options options;
    auto applyFast = [&options] {
        options.method = QStringLiteral("row_flow");
        options.depthModel = QStringLiteral("ZoeD_Any_N");
        options.emaNormalize = false;
        options.sceneDetect = false;
        options.tta = false;
        options.depthAA = false;
    };
    auto applyMaxQuality = [&options] {
        options.method = QStringLiteral("row_flow_v3_sym");
        options.depthModel = QStringLiteral("Any_V3_Mono");
        options.emaNormalize = true;
        options.sceneDetect = true;
        options.tta = true;
        options.depthAA = true;
    };

    switch (preset) {
    case Iw3Preset::FastPreview:
        options.divergence = 1.2;
        applyFast();
        break;
    case Iw3Preset::Recommended:
        options.divergence = 1.8;
        break;
    case Iw3Preset::Strong:
        options.divergence = 2.2;
        break;
    case Iw3Preset::MaxQuality:
        options.divergence = 1.8;
        applyMaxQuality();
        break;
    case Iw3Preset::MaxQualityStrong:
        options.divergence = 2.2;
        applyMaxQuality();
        break;
    case Iw3Preset::Wild:
        options.divergence = 2.6;
        applyMaxQuality();
        break;
    }
    return options;
}

QVector<Iw3PresetModel> iw3Presets() {
    return {
        {Iw3Preset::FastPreview, QStringLiteral("Fast preview (less stable depth)"),
         iw3OptionsForPreset(Iw3Preset::FastPreview)},
        {Iw3Preset::Recommended, QStringLiteral("Recommended (stable)"),
         iw3OptionsForPreset(Iw3Preset::Recommended)},
        {Iw3Preset::Strong, QStringLiteral("Strong (more pop)"),
         iw3OptionsForPreset(Iw3Preset::Strong)},
        {Iw3Preset::MaxQuality, QStringLiteral("Max quality (slow)"),
         iw3OptionsForPreset(Iw3Preset::MaxQuality)},
        {Iw3Preset::MaxQualityStrong, QStringLiteral("Max quality + strong (slow)"),
         iw3OptionsForPreset(Iw3Preset::MaxQualityStrong)},
        {Iw3Preset::Wild, QStringLiteral("Wild (artifacts possible)"),
         iw3OptionsForPreset(Iw3Preset::Wild)},
    };
}

QByteArray buildTopazJobJson(const QString &sourceOriginal, const QString &queuedFileName,
                             const TopazOptions &options, const QDateTime &submittedUtc) {
    const bool is8k = options.target == TopazTarget::K8;
    QByteArray json;
    json += "{\n";
    json += "  \"job_version\": 1,\n";
    json += "  \"submitted_utc\": " + jsonString(utcTimestamp(submittedUtc)) + ",\n";
    json += "  \"source_original\": " + jsonString(sourceOriginal) + ",\n";
    json += "  \"input_file\": " + jsonString(queuedFileName) + ",\n";
    json += "  \"target\": " + jsonString(is8k ? QStringLiteral("8k") : QStringLiteral("4k")) + ",\n";
    json += QByteArray("  \"target_w\": ") + (is8k ? "7680" : "3840") + ",\n";
    json += QByteArray("  \"target_h\": ") + (is8k ? "4320" : "2160") + ",\n";
    json += "  \"codec\": " + jsonString(is8k ? QStringLiteral("av1") : QStringLiteral("h264")) + ",\n";
    json += "  \"container\": \"mp4\",\n";
    json += "  \"profile\": " + jsonString(topazSchemaName(options.profile)) + ",\n";
    json += "  \"grain\": " + QByteArray::number(options.grain, 'f', 6) + ",\n";
    json += "  \"gsize\": " + QByteArray::number(options.grainSize) + "\n";
    json += "}\n";
    return json;
}

QByteArray buildIw3JobJson(const QString &sourceOriginal, const QString &queuedFileName,
                           const Iw3Options &options, const QDateTime &submittedUtc) {
    QByteArray json;
    json += "{\n";
    json += "  \"job_version\": 1,\n";
    json += "  \"job_type\": \"3d\",\n";
    json += "  \"submitted_utc\": " + jsonString(utcTimestamp(submittedUtc)) + ",\n";
    json += "  \"source_original\": " + jsonString(sourceOriginal) + ",\n";
    json += "  \"input_file\": " + jsonString(queuedFileName) + ",\n";
    json += "  \"output_mode\": \"full_sbs\",\n";
    json += "  \"divergence\": " + QByteArray::number(options.divergence, 'f', 4) + ",\n";
    json += "  \"convergence\": " + QByteArray::number(options.convergence, 'f', 4) + ",\n";
    json += "  \"method\": " + jsonString(options.method) + ",\n";
    json += "  \"depth_model\": " + jsonString(options.depthModel) + ",\n";
    json += "  \"synthetic_view\": " + jsonString(options.syntheticView) + ",\n";
    json += QByteArray("  \"preserve_screen_border\": ") + (options.preserveScreenBorder ? "1" : "0") + ",\n";
    json += QByteArray("  \"ema_normalize\": ") + (options.emaNormalize ? "1" : "0") + ",\n";
    json += QByteArray("  \"scene_detect\": ") + (options.sceneDetect ? "1" : "0") + ",\n";
    json += QByteArray("  \"tta\": ") + (options.tta ? "1" : "0") + ",\n";
    json += QByteArray("  \"depth_aa\": ") + (options.depthAA ? "1" : "0") + ",\n";
    json += QByteArray("  \"half_sbs\": ") + (options.halfSbs ? "1" : "0") + ",\n";
    json += "  \"gpu\": " + QByteArray::number(options.gpu) + ",\n";
    json += "  \"max_fps\": " + QByteArray::number(options.maxFps) + ",\n";
    json += "  \"video_format\": " + jsonString(options.videoFormat) + ",\n";
    json += "  \"video_codec\": " + jsonString(options.videoCodec) + ",\n";
    json += "  \"crf\": " + QByteArray::number(options.crf) + ",\n";
    json += "  \"preset\": " + jsonString(options.preset) + ",\n";
    json += "  \"pix_fmt\": " + jsonString(options.pixelFormat) + "\n";
    json += "}\n";
    return json;
}

QueueSubmitResult submitQueueJob(const QString &sourcePath, const QString &queueDirectory,
                                 const QueueJsonBuilder &jsonBuilder,
                                 const CancelCheck &isCancelled,
                                 const CopyProgress &progress) {
    QueueSubmitResult result;
    if (sourcePath.isEmpty() || queueDirectory.isEmpty() || !jsonBuilder) {
        result.error = QStringLiteral("Source, queue directory, and JSON builder are required");
        return result;
    }
    const QFileInfo queueInfo(queueDirectory);
    if (!queueInfo.exists() || !queueInfo.isDir()) {
        result.error = QStringLiteral("Queue directory does not exist: %1").arg(queueDirectory);
        return result;
    }

#if defined(Q_OS_UNIX)
    if (!regularFileNoFollow(sourcePath, result.error)) {
        return result;
    }
#endif

    QFile source(sourcePath);
    if (!source.open(QIODevice::ReadOnly)) {
        result.error = QStringLiteral("Could not open source %1: %2").arg(sourcePath, source.errorString());
        return result;
    }
    const qint64 expectedSize = source.size();
    const SplitName sourceName = splitLastExtension(QFileInfo(sourcePath).fileName());
    if (sourceName.stem.isEmpty()) {
        result.error = QStringLiteral("Source has no usable filename: %1").arg(sourcePath);
        return result;
    }

#if !defined(Q_OS_UNIX)
    result.error = QStringLiteral("Exclusive queue submission is supported only on Unix/Linux");
    return result;
#else
    int destinationDescriptor = -1;
    QString temporaryJsonPath;
    for (int suffix = 0; suffix < 10000; ++suffix) {
        const QString queuedStem = suffix == 0
                                       ? sourceName.stem
                                       : QStringLiteral("%1 (%2)").arg(sourceName.stem).arg(suffix);
        const QString videoPath = QDir(queueInfo.absoluteFilePath()).filePath(queuedStem + sourceName.extension);
        const QString jsonPath = QDir(queueInfo.absoluteFilePath()).filePath(queuedStem + QStringLiteral(".json"));
        if (normalizedPathKey(videoPath) == normalizedPathKey(jsonPath) ||
            pathExistsWithoutFollowingFinalLink(jsonPath)) {
            continue;
        }
        destinationDescriptor = openNewFile(videoPath, 0644);
        if (destinationDescriptor < 0) {
            if (errno == EEXIST || errno == ELOOP) {
                continue;
            }
            result.error = QStringLiteral("Could not reserve queue output %1: %2")
                               .arg(videoPath, systemErrorString(errno));
            return result;
        }
        result.videoPath = videoPath;
        result.jsonPath = jsonPath;
        break;
    }
    if (destinationDescriptor < 0) {
        result.error = QStringLiteral("Could not find a collision-free queue output name");
        return result;
    }

    QFile destination;
    if (!destination.open(destinationDescriptor, QIODevice::WriteOnly, QFileDevice::AutoCloseHandle)) {
        const QString openError = destination.errorString();
        ::close(destinationDescriptor);
        removeForCleanup(result.videoPath, result);
        result.error = QStringLiteral("Could not use reserved queue output %1: %2")
                           .arg(result.videoPath, openError);
        return result;
    }

    auto cancelAndClean = [&] {
        destination.close();
        removeForCleanup(temporaryJsonPath, result);
        removeForCleanup(result.videoPath, result);
        result.status = QueueSubmitStatus::Cancelled;
        result.error = QStringLiteral("Queue submission cancelled");
    };

    if (cancelled(isCancelled)) {
        cancelAndClean();
        return result;
    }

    constexpr qint64 kCopyBufferSize = 1024 * 1024;
    QByteArray buffer;
    buffer.resize(static_cast<int>(kCopyBufferSize));
    qint64 copied = 0;
    bool copyFailed = false;
    QString copyError;
    while (true) {
        if (cancelled(isCancelled)) {
            cancelAndClean();
            return result;
        }
        const qint64 bytesRead = source.read(buffer.data(), buffer.size());
        if (bytesRead < 0) {
            copyFailed = true;
            copyError = source.errorString();
            break;
        }
        if (bytesRead == 0) {
            break;
        }
        qint64 written = 0;
        while (written < bytesRead) {
            const qint64 amount = destination.write(buffer.constData() + written, bytesRead - written);
            if (amount <= 0) {
                copyFailed = true;
                copyError = destination.errorString();
                break;
            }
            written += amount;
            copied += amount;
            if (progress) progress(copied, expectedSize);
            if (cancelled(isCancelled)) {
                cancelAndClean();
                return result;
            }
        }
        if (copyFailed) break;
    }

    if (!copyFailed && expectedSize >= 0 && copied != expectedSize) {
        copyFailed = true;
        copyError = QStringLiteral("source size changed while copying (%1 expected, %2 copied)")
                        .arg(expectedSize).arg(copied);
    }
    if (!copyFailed && !destination.flush()) {
        copyFailed = true;
        copyError = destination.errorString();
    }
    if (!copyFailed && ::fsync(destination.handle()) != 0) {
        copyFailed = true;
        copyError = systemErrorString(errno);
    }
    destination.close();
    source.close();

    if (copyFailed) {
        removeForCleanup(result.videoPath, result);
        result.error = QStringLiteral("Could not fully copy %1: %2").arg(sourcePath, copyError);
        return result;
    }
    bestEffortFsyncDirectory(queueInfo.absoluteFilePath());
    if (cancelled(isCancelled)) {
        removeForCleanup(result.videoPath, result);
        result.status = QueueSubmitStatus::Cancelled;
        result.error = QStringLiteral("Queue submission cancelled");
        return result;
    }

    QByteArray json;
    try {
        json = jsonBuilder(sourcePath, QFileInfo(result.videoPath).fileName());
    } catch (const std::exception &exception) {
        retainOnce(result, result.videoPath);
        result.error = QStringLiteral("Queue JSON builder failed: %1")
                           .arg(QString::fromLocal8Bit(exception.what()));
        return result;
    } catch (...) {
        retainOnce(result, result.videoPath);
        result.error = QStringLiteral("Queue JSON builder failed with an unknown error");
        return result;
    }
    if (json.isEmpty()) {
        retainOnce(result, result.videoPath);
        result.error = QStringLiteral("Queue JSON builder returned no data");
        return result;
    }

    int jsonDescriptor = -1;
    for (int attempt = 0; attempt < 100; ++attempt) {
        const QString token = QUuid::createUuid().toString(QUuid::WithoutBraces);
        temporaryJsonPath = result.jsonPath + QStringLiteral(".%1._json").arg(token);
        jsonDescriptor = openNewFile(temporaryJsonPath, 0644);
        if (jsonDescriptor >= 0) break;
        if (errno != EEXIST && errno != ELOOP) break;
    }
    if (jsonDescriptor < 0) {
        retainOnce(result, result.videoPath);
        result.error = QStringLiteral("Could not create temporary queue JSON: %1")
                           .arg(systemErrorString(errno));
        return result;
    }

    QFile jsonFile;
    if (!jsonFile.open(jsonDescriptor, QIODevice::WriteOnly, QFileDevice::AutoCloseHandle)) {
        const QString openError = jsonFile.errorString();
        ::close(jsonDescriptor);
        removeForCleanup(temporaryJsonPath, result);
        retainOnce(result, result.videoPath);
        result.error = QStringLiteral("Could not use temporary queue JSON: %1").arg(openError);
        return result;
    }
    const qint64 jsonWritten = jsonFile.write(json);
    bool jsonComplete = jsonWritten == json.size() && jsonFile.flush();
    QString jsonError = jsonFile.errorString();
    if (jsonComplete && ::fsync(jsonFile.handle()) != 0) {
        jsonComplete = false;
        jsonError = systemErrorString(errno);
    }
    jsonFile.close();
    if (!jsonComplete) {
        removeForCleanup(temporaryJsonPath, result);
        retainOnce(result, result.videoPath);
        result.error = QStringLiteral("Could not fully write temporary queue JSON: %1").arg(jsonError);
        return result;
    }
    if (cancelled(isCancelled)) {
        removeForCleanup(temporaryJsonPath, result);
        removeForCleanup(result.videoPath, result);
        result.status = QueueSubmitStatus::Cancelled;
        result.error = QStringLiteral("Queue submission cancelled");
        return result;
    }

    bool temporaryStillExists = true;
    QString publishError;
    if (!publishNoReplace(temporaryJsonPath, result.jsonPath, temporaryStillExists, publishError)) {
        retainOnce(result, result.videoPath);
        retainOnce(result, temporaryJsonPath);
        result.error = publishError;
        return result;
    }
    if (temporaryStillExists) {
        retainOnce(result, temporaryJsonPath);
    }
    bestEffortFsyncDirectory(queueInfo.absoluteFilePath());
    result.status = QueueSubmitStatus::Published;
    return result;
#endif
}

QStringList selfTestFailures() {
    QStringList failures;

    const QString namingSource = QStringLiteral("/__media_tools_test__/clip.mp4");
    const QSet<QString> reserved = {
        QStringLiteral("/__media_tools_test__/clip_trimfront.mp4"),
        QStringLiteral("/__media_tools_test__/clip_trimfront (1).mp4")};
    const QString named = chooseUniqueOutputPath(
        namingSource, QStringLiteral("_trimfront"), reserved,
        [](const QString &) { return false; });
    appendFailure(failures, named == QStringLiteral("/__media_tools_test__/clip_trimfront (2).mp4"),
                  QStringLiteral("unique output naming snapshot differs"));

    appendFailure(failures,
                  trimFrontArguments(QStringLiteral("in.mp4"), QStringLiteral("out.mp4"), 1234) ==
                      QStringList({QStringLiteral("-y"), QStringLiteral("-ss"), QStringLiteral("1.234"),
                                   QStringLiteral("-i"), QStringLiteral("in.mp4"), QStringLiteral("-c"),
                                   QStringLiteral("copy"), QStringLiteral("out.mp4")}),
                  QStringLiteral("trim-front FFmpeg arguments differ"));
    appendFailure(failures,
                  trimEndArguments(QStringLiteral("in.mp4"), QStringLiteral("out.mp4"), 9876) ==
                      QStringList({QStringLiteral("-y"), QStringLiteral("-i"), QStringLiteral("in.mp4"),
                                   QStringLiteral("-t"), QStringLiteral("9.876"), QStringLiteral("-c"),
                                   QStringLiteral("copy"), QStringLiteral("out.mp4")}),
                  QStringLiteral("trim-end FFmpeg arguments differ"));
    appendFailure(failures,
                  horizontalFlipArguments(QStringLiteral("in.mp4"), QStringLiteral("out.mp4")) ==
                      QStringList({QStringLiteral("-y"), QStringLiteral("-i"), QStringLiteral("in.mp4"),
                                   QStringLiteral("-vf"), QStringLiteral("hflip"), QStringLiteral("-c:a"),
                                   QStringLiteral("copy"), QStringLiteral("out.mp4")}),
                  QStringLiteral("horizontal-flip FFmpeg arguments differ"));

    QString concatError;
    const QString concat = ffconcatContent(
        {QStringLiteral("/tmp/plain.mp4"), QStringLiteral("/tmp/O'Brien\\raw.mp4")}, &concatError);
    const QString expectedConcat = QStringLiteral(
        "ffconcat version 1.0\n"
        "file '/tmp/plain.mp4'\n"
        "file '/tmp/O'\\''Brien\\raw.mp4'\n");
    appendFailure(failures, concatError.isEmpty() && concat == expectedConcat,
                  QStringLiteral("ffconcat apostrophe/backslash escaping differs"));
    QString invalidConcatError;
    appendFailure(failures,
                  ffconcatContent({QStringLiteral("bad\npath.mp4")}, &invalidConcatError).isEmpty() &&
                      !invalidConcatError.isEmpty(),
                  QStringLiteral("ffconcat accepted a line-breaking path"));

    appendFailure(failures,
                  combineStreamCopyArguments(QStringLiteral("list.ffconcat"), QStringLiteral("out.mp4")) ==
                      QStringList({QStringLiteral("-y"), QStringLiteral("-hide_banner"), QStringLiteral("-f"),
                                   QStringLiteral("concat"), QStringLiteral("-safe"), QStringLiteral("0"),
                                   QStringLiteral("-i"), QStringLiteral("list.ffconcat"), QStringLiteral("-map"),
                                   QStringLiteral("0"), QStringLiteral("-c"), QStringLiteral("copy"),
                                   QStringLiteral("out.mp4")}),
                  QStringLiteral("stream-copy combine arguments differ"));
    const QStringList fallback = combineX264AacArguments(QStringLiteral("list.ffconcat"),
                                                          QStringLiteral("out.mp4"));
    appendFailure(failures,
                  fallback.mid(8, 8) == QStringList({QStringLiteral("-map"), QStringLiteral("0:v:0"),
                                                     QStringLiteral("-map"), QStringLiteral("0:a?"),
                                                     QStringLiteral("-sn"), QStringLiteral("-dn"),
                                                     QStringLiteral("-c:v"), QStringLiteral("libx264")}) &&
                      fallback.contains(QStringLiteral("192k")) && fallback.last() == QStringLiteral("out.mp4"),
                  QStringLiteral("x264/AAC fallback combine arguments differ"));

    const QVector<TopazProfilePreset> topaz = topazProfilePresets();
    appendFailure(failures, topaz.size() == 9 && topaz.at(0).schemaName == QStringLiteral("general") &&
                                topaz.at(5).schemaName == QStringLiteral("deinterlace_repair") &&
                                topaz.at(6).schemaName == QStringLiteral("repair_2pass") &&
                                topaz.at(7).grain == 0.01 && topaz.at(8).grain == 0.01,
                  QStringLiteral("Topaz profile presets differ"));

    const QVector<Iw3PresetModel> iw3 = iw3Presets();
    appendFailure(failures,
                  iw3.size() == 6 && iw3.at(0).options.divergence == 1.2 &&
                      iw3.at(0).options.method == QStringLiteral("row_flow") &&
                      !iw3.at(0).options.emaNormalize && iw3.at(1).options.divergence == 1.8 &&
                      iw3.at(2).options.divergence == 2.2 && iw3.at(3).options.tta &&
                      iw3.at(3).options.depthAA && iw3.at(3).options.divergence == 1.8 &&
                      iw3.at(4).options.divergence == 2.2 && iw3.at(5).options.divergence == 2.6 &&
                      iw3.at(5).options.depthModel == QStringLiteral("Any_V3_Mono"),
                  QStringLiteral("IW3 presets differ"));

    const QDateTime fixedUtc(QDate(2026, 8, 14), QTime(12, 34, 56, 789), Qt::UTC);
    TopazOptions topazOptions = topazOptionsForProfile(TopazTarget::K8, TopazProfile::GeneralGrain);
    const QByteArray topazJson = buildTopazJobJson(
        QStringLiteral("C:\\media\\clip.mp4"), QStringLiteral("clip (1).mp4"), topazOptions, fixedUtc);
    const QByteArray expectedTopazJson =
        "{\n"
        "  \"job_version\": 1,\n"
        "  \"submitted_utc\": \"2026-08-14T12:34:56.789Z\",\n"
        "  \"source_original\": \"C:\\\\media\\\\clip.mp4\",\n"
        "  \"input_file\": \"clip (1).mp4\",\n"
        "  \"target\": \"8k\",\n"
        "  \"target_w\": 7680,\n"
        "  \"target_h\": 4320,\n"
        "  \"codec\": \"av1\",\n"
        "  \"container\": \"mp4\",\n"
        "  \"profile\": \"general_grain\",\n"
        "  \"grain\": 0.010000,\n"
        "  \"gsize\": 1\n"
        "}\n";
    appendFailure(failures, topazJson == expectedTopazJson,
                  QStringLiteral("Topaz JSON schema snapshot differs"));

    const Iw3Options recommended = iw3OptionsForPreset(Iw3Preset::Recommended);
    const QByteArray iw3Json = buildIw3JobJson(
        QStringLiteral("/media/clip.mp4"), QStringLiteral("clip.mp4"), recommended, fixedUtc);
    const QByteArray expectedIw3Json =
        "{\n"
        "  \"job_version\": 1,\n"
        "  \"job_type\": \"3d\",\n"
        "  \"submitted_utc\": \"2026-08-14T12:34:56.789Z\",\n"
        "  \"source_original\": \"/media/clip.mp4\",\n"
        "  \"input_file\": \"clip.mp4\",\n"
        "  \"output_mode\": \"full_sbs\",\n"
        "  \"divergence\": 1.8000,\n"
        "  \"convergence\": 0.5000,\n"
        "  \"method\": \"row_flow_v3\",\n"
        "  \"depth_model\": \"Any_V2_N_B\",\n"
        "  \"synthetic_view\": \"right\",\n"
        "  \"preserve_screen_border\": 1,\n"
        "  \"ema_normalize\": 1,\n"
        "  \"scene_detect\": 1,\n"
        "  \"tta\": 0,\n"
        "  \"depth_aa\": 0,\n"
        "  \"half_sbs\": 0,\n"
        "  \"gpu\": 0,\n"
        "  \"max_fps\": 30,\n"
        "  \"video_format\": \"mp4\",\n"
        "  \"video_codec\": \"hevc_nvenc\",\n"
        "  \"crf\": 20,\n"
        "  \"preset\": \"p5\",\n"
        "  \"pix_fmt\": \"yuv420p\"\n"
        "}\n";
    appendFailure(failures, iw3Json == expectedIw3Json,
                  QStringLiteral("IW3 JSON schema snapshot differs"));

    return failures;
}

} // namespace MediaTools
