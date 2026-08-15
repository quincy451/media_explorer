#include "CombineJobState.h"

#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonParseError>
#include <QProcessEnvironment>
#include <QRegularExpression>
#include <QSaveFile>
#include <QSet>
#include <QUuid>

#include <cmath>
#include <limits>
#include <utility>

#if defined(Q_OS_UNIX)
#include <cerrno>
#include <cstring>
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

namespace CombineJobState {
namespace {

constexpr qint64 kMaximumManifestBytes = 8 * 1024 * 1024;
constexpr int kMaximumSources = 4096;
constexpr int kMaximumPathCharacters = 32768;
constexpr int kMaximumTitleCharacters = 1024;
constexpr int kMaximumDimension = 131072;
constexpr qint64 kMaximumDurationMs = 10000000000000LL;
constexpr int kMaximumAttempt = 1000000;

const QString kSchema = QStringLiteral("org.media-explorer.combine-job");

void assignError(QString *error, const QString &message) {
    if (error) {
        *error = message;
    }
}

QString systemErrorString(int errorNumber) {
#if defined(Q_OS_UNIX)
    return QString::fromLocal8Bit(std::strerror(errorNumber));
#else
    Q_UNUSED(errorNumber);
    return QStringLiteral("operating-system error");
#endif
}

QString newIdentifier() {
    return QUuid::createUuid().toString(QUuid::WithoutBraces)
        .remove(QLatin1Char('-')).toLower();
}

bool validIdentifier(const QString &id) {
    static const QRegularExpression expression(QStringLiteral("^[0-9a-f]{32}$"));
    return expression.match(id).hasMatch();
}

bool pathExistsNoFollow(const QString &path) {
#if defined(Q_OS_UNIX)
    struct stat status {};
    return ::lstat(QFile::encodeName(path).constData(), &status) == 0;
#else
    const QFileInfo info(path);
    return info.exists() || info.isSymLink();
#endif
}

bool isRegularFileNoFollow(const QString &path) {
#if defined(Q_OS_UNIX)
    struct stat status {};
    if (::lstat(QFile::encodeName(path).constData(), &status) != 0) {
        return false;
    }
    return S_ISREG(status.st_mode);
#else
    const QFileInfo info(path);
    return !info.isSymLink() && info.isFile();
#endif
}

bool isOwnedByCurrentUser(const QFileInfo &info) {
#if defined(Q_OS_UNIX)
    return info.ownerId() == static_cast<uint>(::geteuid());
#else
    Q_UNUSED(info);
    return true;
#endif
}

bool setPrivateDirectoryPermissions(const QString &path, QString *error) {
    const QFile::Permissions permissions = QFileDevice::ReadOwner |
                                           QFileDevice::WriteOwner |
                                           QFileDevice::ExeOwner;
    if (!QFile::setPermissions(path, permissions)) {
        assignError(error, QStringLiteral("Could not set private permissions on directory: %1")
                               .arg(path));
        return false;
    }
    return true;
}

bool setPrivateFilePermissions(const QString &path, QString *error) {
    const QFile::Permissions permissions = QFileDevice::ReadOwner |
                                           QFileDevice::WriteOwner;
    if (!QFile::setPermissions(path, permissions)) {
        assignError(error, QStringLiteral("Could not set private permissions on manifest: %1")
                               .arg(path));
        return false;
    }
    return true;
}

bool ensurePrivateDirectory(const QString &path, QString *error) {
    if (!QDir().mkpath(path)) {
        assignError(error, QStringLiteral("Could not create state directory: %1").arg(path));
        return false;
    }
    const QFileInfo info(path);
    if (info.isSymLink() || !info.isDir()) {
        assignError(error, QStringLiteral("State path is not a real directory: %1").arg(path));
        return false;
    }
    if (!isOwnedByCurrentUser(info)) {
        assignError(error, QStringLiteral("State directory is not owned by the current user: %1")
                               .arg(path));
        return false;
    }
    return setPrivateDirectoryPermissions(path, error);
}

void bestEffortFsyncDirectory(const QString &path) {
#if defined(Q_OS_UNIX)
    int flags = O_RDONLY;
#ifdef O_DIRECTORY
    flags |= O_DIRECTORY;
#endif
#ifdef O_CLOEXEC
    flags |= O_CLOEXEC;
#endif
    const int descriptor = ::open(QFile::encodeName(path).constData(), flags);
    if (descriptor >= 0) {
        (void)::fsync(descriptor);
        (void)::close(descriptor);
    }
#else
    Q_UNUSED(path);
#endif
}

bool validPath(const QString &path, const QString &field, QString *error) {
    if (path.isEmpty() || path.size() > kMaximumPathCharacters ||
        path.contains(QChar(0)) || path.contains(QLatin1Char('\n')) ||
        path.contains(QLatin1Char('\r'))) {
        assignError(error, QStringLiteral("%1 is empty, too long, or contains an unsafe character")
                               .arg(field));
        return false;
    }
    if (!QDir::isAbsolutePath(path) || QDir::cleanPath(path) != path ||
        path == QDir::rootPath()) {
        assignError(error, QStringLiteral("%1 must be a clean, absolute, non-root path")
                               .arg(field));
        return false;
    }
    return true;
}

bool pathIsBelow(const QString &path, const QString &directory) {
    const QString relative = QDir(directory).relativeFilePath(path);
    return !relative.isEmpty() && relative != QStringLiteral(".") &&
           relative != QStringLiteral("..") &&
           !relative.startsWith(QStringLiteral("../")) &&
           !QDir::isAbsolutePath(relative);
}

bool hasOnlyKeys(const QJsonObject &object, const QSet<QString> &required,
                 const QString &where, QString *error) {
    for (const QString &key : required) {
        if (!object.contains(key)) {
            assignError(error, QStringLiteral("%1 is missing required field '%2'").arg(where, key));
            return false;
        }
    }
    for (auto it = object.constBegin(); it != object.constEnd(); ++it) {
        if (!required.contains(it.key())) {
            assignError(error, QStringLiteral("%1 contains unknown field '%2'")
                                   .arg(where, it.key()));
            return false;
        }
    }
    return true;
}

bool jsonString(const QJsonObject &object, const QString &key, QString &value,
                QString *error) {
    const QJsonValue jsonValue = object.value(key);
    if (!jsonValue.isString()) {
        assignError(error, QStringLiteral("Field '%1' must be a string").arg(key));
        return false;
    }
    value = jsonValue.toString();
    return true;
}

bool jsonInteger(const QJsonObject &object, const QString &key, qint64 minimum,
                 qint64 maximum, qint64 &value, QString *error) {
    const QJsonValue jsonValue = object.value(key);
    if (!jsonValue.isDouble()) {
        assignError(error, QStringLiteral("Field '%1' must be an integer").arg(key));
        return false;
    }
    const double number = jsonValue.toDouble(std::numeric_limits<double>::quiet_NaN());
    if (!std::isfinite(number) || std::floor(number) != number ||
        number < static_cast<double>(minimum) || number > static_cast<double>(maximum)) {
        assignError(error, QStringLiteral("Field '%1' is outside its allowed integer range")
                               .arg(key));
        return false;
    }
    value = static_cast<qint64>(number);
    return true;
}

QString utcText(const QDateTime &dateTime) {
    return dateTime.toUTC().toString(QStringLiteral("yyyy-MM-dd'T'HH:mm:ss.zzz'Z'"));
}

bool parseUtc(const QString &text, QDateTime &dateTime) {
    if (!text.endsWith(QLatin1Char('Z'))) {
        return false;
    }
    dateTime = QDateTime::fromString(text, QStringLiteral("yyyy-MM-dd'T'HH:mm:ss.zzz'Z'"));
    if (!dateTime.isValid()) {
        return false;
    }
    dateTime.setTimeSpec(Qt::UTC);
    return utcText(dateTime) == text;
}

QByteArray serialize(const Job &job) {
    QJsonArray sources;
    for (const QString &source : job.sources) {
        sources.append(source);
    }

    QJsonObject expected;
    expected.insert(QStringLiteral("width"), job.expectedWidth);
    expected.insert(QStringLiteral("height"), job.expectedHeight);
    expected.insert(QStringLiteral("duration_ms"), static_cast<double>(job.expectedDurationMs));

    QJsonObject root;
    root.insert(QStringLiteral("schema"), kSchema);
    root.insert(QStringLiteral("version"), kManifestVersion);
    root.insert(QStringLiteral("job_id"), job.id);
    root.insert(QStringLiteral("title"), job.title);
    root.insert(QStringLiteral("output_path"), job.outputPath);
    root.insert(QStringLiteral("working_directory"), job.workingDirectory);
    root.insert(QStringLiteral("list_path"), job.listPath);
    root.insert(QStringLiteral("encoded_path"), job.encodedPath);
    root.insert(QStringLiteral("publish_path"), job.publishPath);
    root.insert(QStringLiteral("sources"), sources);
    root.insert(QStringLiteral("expected"), expected);
    root.insert(QStringLiteral("stage"), stageName(job.stage));
    root.insert(QStringLiteral("attempt"), job.attempt);
    root.insert(QStringLiteral("created_utc"), utcText(job.createdUtc));
    root.insert(QStringLiteral("updated_utc"), utcText(job.updatedUtc));
    return QJsonDocument(root).toJson(QJsonDocument::Indented);
}

bool deserialize(const QByteArray &bytes, Job &job, QString *error) {
    QJsonParseError parseError;
    const QJsonDocument document = QJsonDocument::fromJson(bytes, &parseError);
    if (parseError.error != QJsonParseError::NoError || !document.isObject()) {
        assignError(error, QStringLiteral("Invalid JSON: %1").arg(parseError.errorString()));
        return false;
    }
    const QJsonObject root = document.object();
    static const QSet<QString> rootKeys{
        QStringLiteral("schema"), QStringLiteral("version"), QStringLiteral("job_id"),
        QStringLiteral("title"), QStringLiteral("output_path"),
        QStringLiteral("working_directory"), QStringLiteral("list_path"),
        QStringLiteral("encoded_path"), QStringLiteral("publish_path"),
        QStringLiteral("sources"), QStringLiteral("expected"), QStringLiteral("stage"),
        QStringLiteral("attempt"), QStringLiteral("created_utc"),
        QStringLiteral("updated_utc")};
    if (!hasOnlyKeys(root, rootKeys, QStringLiteral("manifest"), error)) {
        return false;
    }

    QString schema;
    if (!jsonString(root, QStringLiteral("schema"), schema, error) || schema != kSchema) {
        if (schema != kSchema) {
            assignError(error, QStringLiteral("Unsupported combine manifest schema"));
        }
        return false;
    }
    qint64 version = 0;
    if (!jsonInteger(root, QStringLiteral("version"), kManifestVersion,
                     kManifestVersion, version, error)) {
        return false;
    }
    Q_UNUSED(version);

    Job parsed;
    QString stage;
    QString created;
    QString updated;
    if (!jsonString(root, QStringLiteral("job_id"), parsed.id, error) ||
        !jsonString(root, QStringLiteral("title"), parsed.title, error) ||
        !jsonString(root, QStringLiteral("output_path"), parsed.outputPath, error) ||
        !jsonString(root, QStringLiteral("working_directory"), parsed.workingDirectory, error) ||
        !jsonString(root, QStringLiteral("list_path"), parsed.listPath, error) ||
        !jsonString(root, QStringLiteral("encoded_path"), parsed.encodedPath, error) ||
        !jsonString(root, QStringLiteral("publish_path"), parsed.publishPath, error) ||
        !jsonString(root, QStringLiteral("stage"), stage, error) ||
        !jsonString(root, QStringLiteral("created_utc"), created, error) ||
        !jsonString(root, QStringLiteral("updated_utc"), updated, error)) {
        return false;
    }
    if (!stageFromName(stage, parsed.stage)) {
        assignError(error, QStringLiteral("Unknown combine stage: %1").arg(stage));
        return false;
    }
    if (!parseUtc(created, parsed.createdUtc) || !parseUtc(updated, parsed.updatedUtc)) {
        assignError(error, QStringLiteral("Manifest timestamps must be canonical UTC values"));
        return false;
    }

    const QJsonValue sourceValue = root.value(QStringLiteral("sources"));
    if (!sourceValue.isArray()) {
        assignError(error, QStringLiteral("Field 'sources' must be an array"));
        return false;
    }
    const QJsonArray sourceArray = sourceValue.toArray();
    if (sourceArray.size() < 2 || sourceArray.size() > kMaximumSources) {
        assignError(error, QStringLiteral("Field 'sources' must contain between 2 and %1 paths")
                               .arg(kMaximumSources));
        return false;
    }
    for (const QJsonValue &value : sourceArray) {
        if (!value.isString()) {
            assignError(error, QStringLiteral("Every source must be a string path"));
            return false;
        }
        parsed.sources.append(value.toString());
    }

    const QJsonValue expectedValue = root.value(QStringLiteral("expected"));
    if (!expectedValue.isObject()) {
        assignError(error, QStringLiteral("Field 'expected' must be an object"));
        return false;
    }
    const QJsonObject expected = expectedValue.toObject();
    static const QSet<QString> expectedKeys{
        QStringLiteral("width"), QStringLiteral("height"), QStringLiteral("duration_ms")};
    if (!hasOnlyKeys(expected, expectedKeys, QStringLiteral("expected"), error)) {
        return false;
    }
    qint64 width = 0;
    qint64 height = 0;
    qint64 duration = 0;
    qint64 attempt = 0;
    if (!jsonInteger(expected, QStringLiteral("width"), 0, kMaximumDimension, width, error) ||
        !jsonInteger(expected, QStringLiteral("height"), 0, kMaximumDimension, height, error) ||
        !jsonInteger(expected, QStringLiteral("duration_ms"), 0, kMaximumDurationMs,
                     duration, error) ||
        !jsonInteger(root, QStringLiteral("attempt"), 0, kMaximumAttempt, attempt, error)) {
        return false;
    }
    parsed.expectedWidth = static_cast<int>(width);
    parsed.expectedHeight = static_cast<int>(height);
    parsed.expectedDurationMs = duration;
    parsed.attempt = static_cast<int>(attempt);

    if (!validateJob(parsed, error)) {
        return false;
    }
    job = std::move(parsed);
    return true;
}

bool readManifest(const QString &path, Job &job, QString *error) {
    if (!isRegularFileNoFollow(path)) {
        assignError(error, QStringLiteral("Manifest is not a regular, non-symlink file: %1")
                               .arg(path));
        return false;
    }
    const QFileInfo info(path);
    if (!isOwnedByCurrentUser(info)) {
        assignError(error, QStringLiteral("Manifest is not owned by the current user: %1")
                               .arg(path));
        return false;
    }
    if (info.size() <= 0 || info.size() > kMaximumManifestBytes) {
        assignError(error, QStringLiteral("Manifest has an invalid size: %1").arg(path));
        return false;
    }
    if (!setPrivateFilePermissions(path, error)) {
        return false;
    }
    QFile file(path);
    if (!file.open(QIODevice::ReadOnly)) {
        assignError(error, QStringLiteral("Could not read manifest %1: %2")
                               .arg(path, file.errorString()));
        return false;
    }
    const QByteArray bytes = file.read(kMaximumManifestBytes + 1);
    if (bytes.size() > kMaximumManifestBytes || !file.atEnd()) {
        assignError(error, QStringLiteral("Manifest is too large: %1").arg(path));
        return false;
    }
    return deserialize(bytes, job, error);
}

bool sameIdentity(const Job &left, const Job &right) {
    return left.id == right.id && left.title == right.title &&
           left.outputPath == right.outputPath &&
           left.workingDirectory == right.workingDirectory &&
           left.listPath == right.listPath && left.encodedPath == right.encodedPath &&
           left.sources == right.sources &&
           left.createdUtc == right.createdUtc;
}

bool writeWithSaveFile(const QString &path, const QByteArray &bytes, QString *error) {
    QSaveFile file(path);
    file.setDirectWriteFallback(false);
    if (!file.open(QIODevice::WriteOnly)) {
        assignError(error, QStringLiteral("Could not open atomic manifest file %1: %2")
                               .arg(path, file.errorString()));
        return false;
    }
    if (!file.setPermissions(QFileDevice::ReadOwner | QFileDevice::WriteOwner)) {
        file.cancelWriting();
        assignError(error, QStringLiteral("Could not make manifest private: %1").arg(path));
        return false;
    }
    if (file.write(bytes) != bytes.size()) {
        const QString detail = file.errorString();
        file.cancelWriting();
        assignError(error, QStringLiteral("Could not write manifest %1: %2").arg(path, detail));
        return false;
    }
    if (!file.commit()) {
        assignError(error, QStringLiteral("Could not commit manifest %1: %2")
                               .arg(path, file.errorString()));
        return false;
    }
    return true;
}

bool publishNoReplace(const QString &stagingPath, const QString &finalPath, QString *error) {
#if defined(Q_OS_LINUX) && defined(SYS_renameat2)
    if (::syscall(SYS_renameat2, AT_FDCWD, QFile::encodeName(stagingPath).constData(),
                  AT_FDCWD, QFile::encodeName(finalPath).constData(),
                  RENAME_NOREPLACE) == 0) {
        return true;
    }
    const int renameError = errno;
    if (renameError != ENOSYS && renameError != EINVAL && renameError != EOPNOTSUPP) {
        assignError(error, QStringLiteral("Could not publish manifest without replacement: %1")
                               .arg(systemErrorString(renameError)));
        return false;
    }
#endif
#if defined(Q_OS_UNIX)
    const QByteArray oldName = QFile::encodeName(stagingPath);
    const QByteArray newName = QFile::encodeName(finalPath);
    if (::link(oldName.constData(), newName.constData()) != 0) {
        assignError(error, QStringLiteral("Could not publish manifest without replacement: %1")
                               .arg(systemErrorString(errno)));
        return false;
    }
    if (::unlink(oldName.constData()) != 0) {
        const int unlinkError = errno;
        (void)::unlink(newName.constData());
        assignError(error, QStringLiteral("Could not retire manifest staging path: %1")
                               .arg(systemErrorString(unlinkError)));
        return false;
    }
    return true;
#else
    if (pathExistsNoFollow(finalPath)) {
        assignError(error, QStringLiteral("Manifest path already exists: %1").arg(finalPath));
        return false;
    }
    if (!QFile::rename(stagingPath, finalPath)) {
        assignError(error, QStringLiteral("Could not publish manifest: %1").arg(finalPath));
        return false;
    }
    return true;
#endif
}

QString exactPath(const QString &pendingDirectory, const QString &id) {
    return QDir(pendingDirectory).filePath(id + QStringLiteral(".json"));
}

bool validateOwnedPath(const QString &pendingDirectory, const Job &job, QString *error) {
    if (!validIdentifier(job.id)) {
        assignError(error, QStringLiteral("Job identifier is invalid"));
        return false;
    }
    const QString expected = exactPath(pendingDirectory, job.id);
    if (job.manifestPath != expected || QDir::cleanPath(job.manifestPath) != expected) {
        assignError(error, QStringLiteral("Job does not identify its exact owned manifest"));
        return false;
    }
    return true;
}

QString quarantineManifest(const QString &sourcePath, const QString &quarantineDirectory,
                           QString *error) {
    const QFileInfo sourceInfo(sourcePath);
    const QString timestamp = QDateTime::currentDateTimeUtc()
                                  .toString(QStringLiteral("yyyyMMdd'T'HHmmsszzz'Z'"));
    for (int attempt = 0; attempt < 64; ++attempt) {
        const QString targetName = QStringLiteral("%1.invalid-%2-%3")
                                       .arg(sourceInfo.fileName(), timestamp, newIdentifier());
        const QString targetPath = QDir(quarantineDirectory).filePath(targetName);
        if (pathExistsNoFollow(targetPath)) {
            continue;
        }
        QString moveError;
        if (publishNoReplace(sourcePath, targetPath, &moveError)) {
            // Never chmod through a quarantined symbolic link. The containing
            // directory is private, and a regular file can safely be mode 0600.
            if (isRegularFileNoFollow(targetPath)) {
                (void)setPrivateFilePermissions(targetPath, nullptr);
            }
            bestEffortFsyncDirectory(sourceInfo.absolutePath());
            bestEffortFsyncDirectory(quarantineDirectory);
            return targetPath;
        }
        assignError(error, QStringLiteral("Could not quarantine malformed manifest %1: %2")
                               .arg(sourcePath, moveError));
        return QString();
    }
    assignError(error, QStringLiteral("Could not choose a non-conflicting quarantine name for %1")
                           .arg(sourcePath));
    return QString();
}

} // namespace

QString stageName(Stage stage) {
    switch (stage) {
    case Stage::Prepared: return QStringLiteral("prepared");
    case Stage::InputsValidated: return QStringLiteral("inputs_validated");
    case Stage::StreamCopyRunning: return QStringLiteral("stream_copy_running");
    case Stage::TranscodeRunning: return QStringLiteral("transcode_running");
    case Stage::Publishing: return QStringLiteral("publishing");
    case Stage::Failed: return QStringLiteral("failed");
    case Stage::Completed: return QStringLiteral("completed");
    }
    return QString();
}

bool stageFromName(const QString &name, Stage &stage) {
    if (name == QStringLiteral("prepared")) {
        stage = Stage::Prepared;
    } else if (name == QStringLiteral("inputs_validated")) {
        stage = Stage::InputsValidated;
    } else if (name == QStringLiteral("stream_copy_running")) {
        stage = Stage::StreamCopyRunning;
    } else if (name == QStringLiteral("transcode_running")) {
        stage = Stage::TranscodeRunning;
    } else if (name == QStringLiteral("publishing")) {
        stage = Stage::Publishing;
    } else if (name == QStringLiteral("failed")) {
        stage = Stage::Failed;
    } else if (name == QStringLiteral("completed")) {
        stage = Stage::Completed;
    } else {
        return false;
    }
    return true;
}

bool validateJob(const Job &job, QString *error) {
    if (!validIdentifier(job.id)) {
        assignError(error, QStringLiteral("Job identifier must be 32 lowercase hexadecimal digits"));
        return false;
    }
    if (job.title.trimmed().isEmpty() || job.title.size() > kMaximumTitleCharacters ||
        job.title.contains(QChar(0))) {
        assignError(error, QStringLiteral("Job title is empty, too long, or contains NUL"));
        return false;
    }
    if (!validPath(job.outputPath, QStringLiteral("output_path"), error) ||
        !validPath(job.workingDirectory, QStringLiteral("working_directory"), error) ||
        !validPath(job.listPath, QStringLiteral("list_path"), error) ||
        !validPath(job.encodedPath, QStringLiteral("encoded_path"), error)) {
        return false;
    }
    if (!pathIsBelow(job.listPath, job.workingDirectory)) {
        assignError(error, QStringLiteral("list_path must be below working_directory"));
        return false;
    }
    const QFileInfo outputInfo(job.outputPath);
    const QFileInfo workingInfo(job.workingDirectory);
    const QFileInfo listInfo(job.listPath);
    const QFileInfo encodedInfo(job.encodedPath);
    static const QRegularExpression workName(
        QStringLiteral("^\\.media-explorer-combine-[A-Za-z0-9._-]{1,160}$"));
    if (workingInfo.absolutePath() != outputInfo.absolutePath() ||
        !workName.match(workingInfo.fileName()).hasMatch()) {
        assignError(error,
                    QStringLiteral("working_directory must be a .media-explorer-combine-* "
                                   "child of the output directory"));
        return false;
    }
    if (listInfo.absolutePath() != job.workingDirectory || listInfo.fileName().isEmpty()) {
        assignError(error, QStringLiteral("list_path must be a direct child of working_directory"));
        return false;
    }
    if (encodedInfo.absolutePath() != job.workingDirectory || encodedInfo.fileName().isEmpty() ||
        job.encodedPath == job.listPath) {
        assignError(error, QStringLiteral("encoded_path must be a distinct direct child of working_directory"));
        return false;
    }
    if (!job.publishPath.isEmpty()) {
        if (!validPath(job.publishPath, QStringLiteral("publish_path"), error) ||
            QFileInfo(job.publishPath).absolutePath() != outputInfo.absolutePath()) {
            assignError(error, QStringLiteral("publish_path must be in the output directory"));
            return false;
        }
    }
    if (job.outputPath == job.listPath || job.outputPath == job.workingDirectory ||
        job.outputPath == job.encodedPath || job.publishPath == job.workingDirectory ||
        job.publishPath == job.listPath || job.publishPath == job.encodedPath) {
        assignError(error, QStringLiteral("Output path must be distinct from working paths"));
        return false;
    }
    if (job.sources.size() < 2 || job.sources.size() > kMaximumSources) {
        assignError(error, QStringLiteral("A combine job must contain between 2 and %1 sources")
                               .arg(kMaximumSources));
        return false;
    }
    QSet<QString> uniqueSources;
    for (qsizetype index = 0; index < job.sources.size(); ++index) {
        const QString &source = job.sources.at(index);
        if (!validPath(source, QStringLiteral("sources[%1]").arg(index), error)) {
            return false;
        }
        if (source == job.outputPath || source == job.listPath || source == job.encodedPath ||
            source == job.publishPath || source == job.workingDirectory) {
            assignError(error, QStringLiteral("A source path collides with an output or working path"));
            return false;
        }
        if (uniqueSources.contains(source)) {
            assignError(error, QStringLiteral("Source paths must be unique"));
            return false;
        }
        uniqueSources.insert(source);
    }
    if (job.expectedWidth < 0 || job.expectedWidth > kMaximumDimension ||
        job.expectedHeight < 0 || job.expectedHeight > kMaximumDimension ||
        ((job.expectedWidth == 0) != (job.expectedHeight == 0))) {
        assignError(error, QStringLiteral("Expected dimensions must both be zero or both be positive"));
        return false;
    }
    if (job.expectedDurationMs < 0 || job.expectedDurationMs > kMaximumDurationMs) {
        assignError(error, QStringLiteral("Expected duration is outside its allowed range"));
        return false;
    }
    if (job.attempt < 0 || job.attempt > kMaximumAttempt) {
        assignError(error, QStringLiteral("Attempt is outside its allowed range"));
        return false;
    }
    if ((job.stage == Stage::InputsValidated ||
         job.stage == Stage::StreamCopyRunning ||
         job.stage == Stage::TranscodeRunning ||
         job.stage == Stage::Publishing ||
         job.stage == Stage::Completed) &&
        (job.expectedWidth == 0 || job.expectedHeight == 0)) {
        assignError(error, QStringLiteral("This stage requires validated positive dimensions"));
        return false;
    }
    if ((job.stage == Stage::StreamCopyRunning ||
         job.stage == Stage::TranscodeRunning ||
         job.stage == Stage::Publishing ||
         job.stage == Stage::Completed) && job.attempt == 0) {
        assignError(error, QStringLiteral("A running or completed job requires a positive attempt"));
        return false;
    }
    if ((job.stage == Stage::Publishing || job.stage == Stage::Completed) &&
        job.publishPath.isEmpty()) {
        assignError(error, QStringLiteral("Publishing and completed stages require publish_path"));
        return false;
    }
    if (!job.createdUtc.isValid() || !job.updatedUtc.isValid() ||
        job.createdUtc.toUTC() > job.updatedUtc.toUTC()) {
        assignError(error, QStringLiteral("Job timestamps are invalid or out of order"));
        return false;
    }
    return true;
}

Store::Store(const QString &stateHomeOverride) {
    if (!stateHomeOverride.isEmpty()) {
        stateHome_ = QDir::cleanPath(stateHomeOverride);
        return;
    }
    const QString xdgState = QProcessEnvironment::systemEnvironment()
                                 .value(QStringLiteral("XDG_STATE_HOME"));
    if (!xdgState.isEmpty() && QDir::isAbsolutePath(xdgState)) {
        stateHome_ = QDir::cleanPath(xdgState);
    } else {
        stateHome_ = QDir::cleanPath(QDir::home().filePath(QStringLiteral(".local/state")));
    }
}

QString Store::applicationStateDirectory() const {
    return QDir(stateHome_).filePath(QStringLiteral("media-explorer"));
}

QString Store::pendingDirectory() const {
    return QDir(applicationStateDirectory()).filePath(QStringLiteral("pending"));
}

QString Store::quarantineDirectory() const {
    return QDir(applicationStateDirectory()).filePath(QStringLiteral("quarantine"));
}

bool Store::initialize(QString *error) const {
    if (error) {
        error->clear();
    }
    if (stateHome_.isEmpty() || !QDir::isAbsolutePath(stateHome_) ||
        stateHome_ == QDir::rootPath()) {
        assignError(error, QStringLiteral("State home must be an absolute, non-root path"));
        return false;
    }
    if (!ensurePrivateDirectory(applicationStateDirectory(), error) ||
        !ensurePrivateDirectory(pendingDirectory(), error) ||
        !ensurePrivateDirectory(quarantineDirectory(), error)) {
        return false;
    }
    return true;
}

bool Store::create(Job &job, QString *error) const {
    if (error) {
        error->clear();
    }
    if (!job.id.isEmpty() || !job.manifestPath.isEmpty()) {
        assignError(error, QStringLiteral("create requires an unassigned job id and manifest path"));
        return false;
    }
    if (!initialize(error)) {
        return false;
    }

    const Job original = job;
    for (int collisionAttempt = 0; collisionAttempt < 64; ++collisionAttempt) {
        Job candidate = original;
        candidate.id = newIdentifier();
        const QDateTime now = QDateTime::currentDateTimeUtc();
        candidate.createdUtc = now;
        candidate.updatedUtc = now;
        candidate.manifestPath = exactPath(pendingDirectory(), candidate.id);
        if (!validateJob(candidate, error)) {
            return false;
        }
        if (pathExistsNoFollow(candidate.manifestPath)) {
            continue;
        }

        const QString stagingPath = QDir(pendingDirectory()).filePath(
            QStringLiteral(".%1-%2.staged").arg(candidate.id, newIdentifier()));
        if (pathExistsNoFollow(stagingPath)) {
            continue;
        }
        if (!writeWithSaveFile(stagingPath, serialize(candidate), error)) {
            return false;
        }
        QString publishError;
        if (!publishNoReplace(stagingPath, candidate.manifestPath, &publishError)) {
            if (!pathExistsNoFollow(candidate.manifestPath)) {
                assignError(error, publishError);
                return false;
            }
            // A UUID collision (or adversarial pre-creation) never replaces the
            // existing entry. Remove only our uniquely named staging file.
            (void)QFile::remove(stagingPath);
            continue;
        }
        // From this point onward the manifest is committed and recoverable,
        // even if a best-effort permission tightening step reports failure.
        job = candidate;
        if (!setPrivateFilePermissions(candidate.manifestPath, error)) {
            // Preserve the successfully published manifest for recovery.
            return false;
        }
        bestEffortFsyncDirectory(pendingDirectory());
        job = std::move(candidate);
        return true;
    }
    assignError(error, QStringLiteral("Could not allocate a unique combine manifest name"));
    return false;
}

bool Store::save(Job &job, QString *error) const {
    if (error) {
        error->clear();
    }
    if (!initialize(error) || !validateOwnedPath(pendingDirectory(), job, error)) {
        return false;
    }
    Job existing;
    if (!readManifest(job.manifestPath, existing, error)) {
        return false;
    }
    existing.manifestPath = job.manifestPath;
    if (!sameIdentity(existing, job)) {
        assignError(error, QStringLiteral("Refusing to change immutable combine job identity"));
        return false;
    }
    if (job.attempt < existing.attempt) {
        assignError(error, QStringLiteral("Refusing to decrease the combine attempt counter"));
        return false;
    }
    if (!existing.publishPath.isEmpty() && job.publishPath.isEmpty()) {
        assignError(error, QStringLiteral("Refusing to clear the recorded publication path"));
        return false;
    }

    Job candidate = job;
    candidate.updatedUtc = QDateTime::currentDateTimeUtc();
    if (candidate.updatedUtc < candidate.createdUtc) {
        candidate.updatedUtc = candidate.createdUtc;
    }
    if (!validateJob(candidate, error)) {
        return false;
    }
    if (!writeWithSaveFile(candidate.manifestPath, serialize(candidate), error)) {
        return false;
    }
    if (!setPrivateFilePermissions(candidate.manifestPath, error)) {
        return false;
    }
    bestEffortFsyncDirectory(pendingDirectory());
    job = std::move(candidate);
    return true;
}

bool Store::setStage(Job &job, Stage stage, int attempt, QString *error) const {
    const Stage previousStage = job.stage;
    const int previousAttempt = job.attempt;
    job.stage = stage;
    job.attempt = attempt;
    if (save(job, error)) {
        return true;
    }
    job.stage = previousStage;
    job.attempt = previousAttempt;
    return false;
}

bool Store::setPublishPath(Job &job, const QString &publishPath, QString *error) const {
    const QString previous = job.publishPath;
    job.publishPath = QDir::cleanPath(publishPath);
    if (save(job, error)) {
        return true;
    }
    job.publishPath = previous;
    return false;
}

ScanResult Store::loadPending() const {
    ScanResult result;
    QString initializationError;
    if (!initialize(&initializationError)) {
        result.errors.append(initializationError);
        return result;
    }

    QDir directory(pendingDirectory());
    const QFileInfoList entries = directory.entryInfoList(
        QStringList{QStringLiteral("*.json")},
        QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot,
        QDir::Name | QDir::IgnoreCase);
    for (const QFileInfo &entry : entries) {
        const QString path = entry.absoluteFilePath();
        Job job;
        QString loadError;
        bool valid = readManifest(path, job, &loadError);
        if (valid && entry.fileName() != job.id + QStringLiteral(".json")) {
            valid = false;
            loadError = QStringLiteral("Manifest filename does not match its job identifier");
        }
        if (valid) {
            job.manifestPath = exactPath(pendingDirectory(), job.id);
            if (job.manifestPath != path) {
                valid = false;
                loadError = QStringLiteral("Manifest path is not the exact owned pending path");
            }
        }
        if (valid) {
            result.jobs.append(std::move(job));
            continue;
        }

        QString quarantineError;
        const QString quarantined = quarantineManifest(path, quarantineDirectory(),
                                                       &quarantineError);
        if (!quarantined.isEmpty()) {
            result.quarantinedPaths.append(quarantined);
            result.errors.append(QStringLiteral("Quarantined %1: %2").arg(path, loadError));
        } else {
            result.errors.append(QStringLiteral("Malformed manifest %1: %2; %3")
                                     .arg(path, loadError, quarantineError));
        }
    }
    return result;
}

bool Store::remove(const Job &job, QString *error) const {
    if (error) {
        error->clear();
    }
    if (!initialize(error) || !validateOwnedPath(pendingDirectory(), job, error)) {
        return false;
    }
    Job existing;
    if (!readManifest(job.manifestPath, existing, error)) {
        return false;
    }
    existing.manifestPath = job.manifestPath;
    if (!sameIdentity(existing, job)) {
        assignError(error, QStringLiteral("Refusing to remove a different combine manifest"));
        return false;
    }
    if (!QFile::remove(job.manifestPath)) {
        assignError(error, QStringLiteral("Could not remove combine manifest: %1")
                               .arg(job.manifestPath));
        return false;
    }
    bestEffortFsyncDirectory(pendingDirectory());
    return true;
}

} // namespace CombineJobState
