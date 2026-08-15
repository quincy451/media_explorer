#include "MountDiscovery.h"

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QProcess>
#include <QSet>
#include <QStandardPaths>
#include <QUrl>

#include <algorithm>
#include <dirent.h>
#include <unistd.h>
#include <utility>

namespace {

const QSet<QString> kPseudoFilesystems{
    QStringLiteral("autofs"), QStringLiteral("bpf"), QStringLiteral("cgroup"),
    QStringLiteral("cgroup2"), QStringLiteral("configfs"), QStringLiteral("debugfs"),
    QStringLiteral("devpts"), QStringLiteral("devtmpfs"), QStringLiteral("efivarfs"),
    QStringLiteral("fusectl"), QStringLiteral("hugetlbfs"), QStringLiteral("mqueue"),
    QStringLiteral("proc"), QStringLiteral("pstore"), QStringLiteral("securityfs"),
    QStringLiteral("sysfs"), QStringLiteral("tracefs")};

const QSet<QString> kNetworkFilesystems{
    QStringLiteral("cifs"), QStringLiteral("smb3"), QStringLiteral("nfs"),
    QStringLiteral("nfs4"), QStringLiteral("fuse.sshfs"), QStringLiteral("sshfs"),
    QStringLiteral("davfs"), QStringLiteral("davfs2")};

QString decodeMountField(const QString &field) {
    QString decoded;
    decoded.reserve(field.size());
    for (int index = 0; index < field.size(); ++index) {
        if (field.at(index) == QLatin1Char('\\') && index + 3 < field.size()) {
            bool ok = false;
            const ushort value = field.mid(index + 1, 3).toUShort(&ok, 8);
            if (ok) {
                decoded.append(QChar(value));
                index += 3;
                continue;
            }
        }
        decoded.append(field.at(index));
    }
    return decoded;
}

QStringList childNamesWithoutStat(const QString &directory) {
    QStringList names;
    const QByteArray encoded = QFile::encodeName(directory);
    DIR *handle = ::opendir(encoded.constData());
    if (!handle) {
        return names;
    }
    while (dirent *item = ::readdir(handle)) {
        const QByteArray raw(item->d_name);
        if (raw == "." || raw == ".." || raw.startsWith('.')) {
            continue;
        }
        names.append(QFile::decodeName(raw));
    }
    ::closedir(handle);
    return names;
}

int sourcePriority(EntrySource source) {
    switch (source) {
    case EntrySource::Home: return 0;
    case EntrySource::Filesystem: return 1;
    case EntrySource::MappingStable: return 2;
    case EntrySource::MappingGvfs: return 3;
    case EntrySource::MappingAutomount: return 4;
    case EntrySource::MappingDisconnected: return 5;
    case EntrySource::Mount: return 6;
    default: return 9;
    }
}

} // namespace

QVector<MountRecord> readMountRecords() {
    QVector<MountRecord> records;
    QFile file(QStringLiteral("/proc/self/mountinfo"));
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text)) {
        return records;
    }
    for (;;) {
        const QByteArray rawLine = file.readLine();
        if (rawLine.isEmpty()) {
            break;
        }
        const QString line = QString::fromUtf8(rawLine).trimmed();
        const QStringList fields = line.split(QLatin1Char(' '), Qt::SkipEmptyParts);
        const int separator = fields.indexOf(QStringLiteral("-"));
        if (separator < 0 || fields.size() < 7 || separator + 2 >= fields.size()) {
            continue;
        }
        records.append({QDir::cleanPath(decodeMountField(fields.at(4))), fields.at(separator + 1)});
    }
    return records;
}

MountState mappingMountState(const QString &path) {
    const QString normalized = QDir::cleanPath(QFileInfo(path).absoluteFilePath());
    bool automount = false;
    for (const MountRecord &record : readMountRecords()) {
        if (record.path != normalized) {
            continue;
        }
        if (!kPseudoFilesystems.contains(record.filesystem)) {
            return MountState::Mounted;
        }
        automount = automount || record.filesystem == QStringLiteral("autofs");
    }
    return automount ? MountState::Automount : MountState::Unmounted;
}

QString gvfsPathForUri(const QString &uri) {
    const QUrl url(uri);
    if (!url.isValid() || (url.scheme().compare(QStringLiteral("smb"), Qt::CaseInsensitive) != 0 &&
                           url.scheme().compare(QStringLiteral("cifs"), Qt::CaseInsensitive) != 0)) {
        return {};
    }
    const QString host = url.host().toLower();
    const QStringList encodedParts = url.path(QUrl::FullyEncoded).split(QLatin1Char('/'), Qt::SkipEmptyParts);
    if (host.isEmpty() || encodedParts.isEmpty()) {
        return {};
    }
    QStringList parts;
    for (const QString &encodedPart : encodedParts) {
        const QString part = QUrl::fromPercentEncoding(encodedPart.toUtf8());
        if (part.isEmpty() || part == QStringLiteral(".") || part == QStringLiteral("..") ||
            part.contains(QLatin1Char('/')) || part.contains(QLatin1Char('\\')) || part.contains(QChar(0))) {
            return {};
        }
        parts.append(part);
    }
    const QString share = parts.takeFirst().toLower();
    const QDir gvfsRoot(QStringLiteral("/run/user/%1/gvfs").arg(::getuid()));
    for (const QString &candidateName : gvfsRoot.entryList(QDir::Dirs | QDir::NoDotAndDotDot)) {
        const int colon = candidateName.indexOf(QLatin1Char(':'));
        if (colon <= 0) {
            continue;
        }
        const QString prefix = candidateName.left(colon).toLower();
        if (prefix != QStringLiteral("smb-share") && prefix != QStringLiteral("cifs-share")) {
            continue;
        }
        QHash<QString, QString> fields;
        for (const QString &component : candidateName.mid(colon + 1).split(QLatin1Char(','))) {
            const int equals = component.indexOf(QLatin1Char('='));
            if (equals > 0) {
                fields.insert(component.left(equals).toLower(),
                              QUrl::fromPercentEncoding(component.mid(equals + 1).toUtf8()).toLower());
            }
        }
        if (fields.value(QStringLiteral("server")) != host ||
            fields.value(QStringLiteral("share")) != share) {
            continue;
        }
        QString resolved = gvfsRoot.filePath(candidateName);
        for (const QString &part : std::as_const(parts)) {
            resolved = QDir(resolved).filePath(part);
        }
        // The mount directory came from a local GVFS directory listing.  Do
        // not stat its remote contents on the GUI thread; activation performs
        // a bounded worker probe.
        return QDir::cleanPath(resolved);
    }
    return {};
}

QVector<Entry> discoverMounts(const AppConfig &config) {
    QVector<Entry> result;
    QSet<QString> added;
    const auto add = [&result, &added](const QString &path, const QString &name, const QString &kind,
                                       EntrySource source, const QString &uri = QString(),
                                       bool requireDirectory = true) {
        const QString normalized = QDir::cleanPath(QFileInfo(path).absoluteFilePath());
        if (added.contains(normalized) || (requireDirectory && !QFileInfo(normalized).isDir())) {
            return;
        }
        added.insert(normalized);
        result.append({name, normalized, true, kind, 0, 0, {}, -1.0, source, uri});
    };

    add(QDir::homePath(), QStringLiteral("Home"), QStringLiteral("Home folder"), EntrySource::Home);
    add(QStringLiteral("/"), QStringLiteral("Filesystem"), QStringLiteral("Filesystem"),
        EntrySource::Filesystem);

    const QVector<MountRecord> records = readMountRecords();
    QSet<QString> mountedPaths;
    QSet<QString> automountPaths;
    for (const MountRecord &record : records) {
        if (!kPseudoFilesystems.contains(record.filesystem)) {
            mountedPaths.insert(record.path);
        }
        if (record.filesystem == QStringLiteral("autofs")) {
            automountPaths.insert(record.path);
        }
    }

    QSet<QString> mappingNames;
    for (auto it = config.mappedShares.cbegin(); it != config.mappedShares.cend(); ++it) {
        mappingNames.insert(it.key().toLower());
    }
    for (const QString &name : childNamesWithoutStat(config.mappedRoot)) {
        mappingNames.insert(name.toLower());
    }
    QStringList sortedMappings = mappingNames.values();
    std::sort(sortedMappings.begin(), sortedMappings.end(),
              [](const QString &left, const QString &right) {
                  return left.compare(right, Qt::CaseInsensitive) < 0;
              });
    for (const QString &mapping : std::as_const(sortedMappings)) {
        const QString mountpoint = QDir(config.mappedRoot).filePath(mapping);
        const QString normalized = QDir::cleanPath(QFileInfo(mountpoint).absoluteFilePath());
        const QString uri = config.mappedShares.value(mapping);
        const QString label = mapping.size() <= 3 ? mapping.toUpper() + QLatin1Char(':') : mapping;
        if (mountedPaths.contains(normalized)) {
            add(mountpoint, label, QStringLiteral("Mapped drive"),
                EntrySource::MappingStable, uri, false);
        } else if (automountPaths.contains(normalized)) {
            add(mountpoint, label, QStringLiteral("Mapped drive (on demand)"),
                EntrySource::MappingAutomount, uri, false);
        } else {
            const QString gvfs = uri.isEmpty() ? QString() : gvfsPathForUri(uri);
            if (!gvfs.isEmpty()) {
                add(gvfs, label, QStringLiteral("Mapped drive (GVFS)"),
                    EntrySource::MappingGvfs, uri, false);
            } else {
                add(mountpoint, label + QStringLiteral(" (disconnected)"),
                    QStringLiteral("Mapped drive (disconnected)"),
                    EntrySource::MappingDisconnected, uri, false);
            }
        }
    }

    for (const MountRecord &record : records) {
        if (record.path == QStringLiteral("/") || kPseudoFilesystems.contains(record.filesystem)) {
            continue;
        }
        const bool interesting = record.path.startsWith(QStringLiteral("/mnt/")) ||
                                 record.path.startsWith(QStringLiteral("/media/")) ||
                                 record.path.startsWith(QStringLiteral("/run/media/"));
        if (!interesting && !kNetworkFilesystems.contains(record.filesystem)) {
            continue;
        }
        const QString name = QFileInfo(record.path).fileName().isEmpty()
                                 ? record.path
                                 : QFileInfo(record.path).fileName();
        const QString kind = kNetworkFilesystems.contains(record.filesystem)
                                 ? QStringLiteral("Network mount")
                                 : QStringLiteral("Mounted %1").arg(record.filesystem);
        add(record.path, name, kind, EntrySource::Mount, {}, false);
    }

    std::sort(result.begin(), result.end(), [](const Entry &left, const Entry &right) {
        const int leftPriority = sourcePriority(left.source);
        const int rightPriority = sourcePriority(right.source);
        return leftPriority == rightPriority
                   ? left.name.compare(right.name, Qt::CaseInsensitive) < 0
                   : leftPriority < rightPriority;
    });
    return result;
}

MountProbeResult probeStableMount(const QString &path, int timeoutMs) {
    MountProbeResult result{QDir::cleanPath(QFileInfo(path).absoluteFilePath()),
                            mappingMountState(path), false, {}};
    const QString timeout = QStandardPaths::findExecutable(QStringLiteral("timeout"));
    const QString stat = QStandardPaths::findExecutable(QStringLiteral("stat"));
    if (timeout.isEmpty() || stat.isEmpty()) {
        result.error = QStringLiteral("timeout/stat utilities are unavailable; stable path was not probed");
        return result;
    }
    QProcess process;
    process.setProcessChannelMode(QProcess::MergedChannels);
    const int seconds = qMax(1, (timeoutMs + 999) / 1000);
    process.start(timeout,
                  {QStringLiteral("--signal=TERM"), QStringLiteral("--kill-after=2s"),
                   QStringLiteral("%1s").arg(seconds), stat, QStringLiteral("--"),
                   QDir(result.path).filePath(QStringLiteral("."))});
    if (!process.waitForStarted(2000)) {
        result.error = process.errorString();
        return result;
    }
    if (!process.waitForFinished(timeoutMs + 5000)) {
        process.kill();
        process.waitForFinished(2000);
        result.error = QStringLiteral("stable-path probe exceeded %1 seconds").arg(seconds);
    } else if (process.exitCode() != 0) {
        result.error = QString::fromUtf8(process.readAll()).trimmed();
        if (result.error.isEmpty()) {
            result.error = QStringLiteral("stable-path probe exited %1").arg(process.exitCode());
        }
    } else {
        result.accessible = true;
    }
    result.state = mappingMountState(result.path);
    if (!result.accessible && result.error.isEmpty()) {
        result.error = QStringLiteral("the path did not become accessible");
    }
    return result;
}

QString entrySourceName(EntrySource source) {
    switch (source) {
    case EntrySource::Folder: return QStringLiteral("folder");
    case EntrySource::Search: return QStringLiteral("search");
    case EntrySource::Home: return QStringLiteral("home");
    case EntrySource::Filesystem: return QStringLiteral("filesystem");
    case EntrySource::MappingStable: return QStringLiteral("mapping-stable");
    case EntrySource::MappingGvfs: return QStringLiteral("mapped-gvfs");
    case EntrySource::MappingAutomount: return QStringLiteral("mapping-automount");
    case EntrySource::MappingDisconnected: return QStringLiteral("mapping-disconnected");
    case EntrySource::Mount: return QStringLiteral("mount");
    }
    return QStringLiteral("unknown");
}
