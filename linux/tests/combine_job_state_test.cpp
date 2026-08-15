#include "CombineJobState.h"

#include <QCoreApplication>
#include <QDir>
#include <QFile>
#include <QFileInfo>
#include <QTemporaryDir>
#include <QTextStream>

#if defined(Q_OS_UNIX)
#include <sys/stat.h>
#endif

namespace {

bool writeBytes(const QString &path, const QByteArray &bytes) {
    QFile file(path);
    return file.open(QIODevice::WriteOnly | QIODevice::Truncate) &&
           file.write(bytes) == bytes.size() && file.flush();
}

void expect(bool condition, const QString &message, QStringList &failures) {
    if (!condition) {
        failures.append(message);
    }
}

#if defined(Q_OS_UNIX)
int modeBits(const QString &path) {
    struct stat status {};
    if (::stat(QFile::encodeName(path).constData(), &status) != 0) {
        return -1;
    }
    return status.st_mode & 0777;
}
#endif

} // namespace

int main(int argc, char **argv) {
    QCoreApplication application(argc, argv);
    QTemporaryDir temporary;
    QStringList failures;
    expect(temporary.isValid(), QStringLiteral("temporary directory creation failed"), failures);
    if (!temporary.isValid()) {
        return 1;
    }

    const QString stateHome = QDir(temporary.path()).filePath(QStringLiteral("state"));
    const QString work = QDir(temporary.path()).filePath(
        QStringLiteral(".media-explorer-combine-test"));
    expect(QDir().mkpath(work), QStringLiteral("work directory creation failed"), failures);

    CombineJobState::Store store(stateHome);
    QString error;
    expect(store.initialize(&error), QStringLiteral("initialize failed: %1").arg(error), failures);

    CombineJobState::Job job;
    job.title = QString::fromUtf8("Combine — résumé 日本語");
    job.outputPath = QDir(temporary.path()).filePath(QString::fromUtf8("résultat.mp4"));
    job.workingDirectory = work;
    job.listPath = QDir(work).filePath(QStringLiteral("inputs.ffconcat"));
    job.encodedPath = QDir(work).filePath(QStringLiteral("encoded-output.mp4"));
    job.sources = QStringList{
        QDir(temporary.path()).filePath(QString::fromUtf8("entrée 1.mp4")),
        QDir(temporary.path()).filePath(QString::fromUtf8("entrée 2.mp4"))};
    job.expectedWidth = 1920;
    job.expectedHeight = 1080;
    job.expectedDurationMs = 123456;

    expect(store.create(job, &error), QStringLiteral("create failed: %1").arg(error), failures);
    expect(QFileInfo(job.manifestPath).isFile(), QStringLiteral("manifest was not created"), failures);
    expect(job.id.size() == 32, QStringLiteral("fresh id has wrong length"), failures);
    QFile manifest(job.manifestPath);
    expect(manifest.open(QIODevice::ReadOnly), QStringLiteral("manifest cannot be read"), failures);
    const QByteArray manifestBytes = manifest.readAll();
    expect(manifestBytes.contains(QString::fromUtf8("résultat").toUtf8()),
           QStringLiteral("manifest is not UTF-8"), failures);

#if defined(Q_OS_UNIX)
    expect(modeBits(store.applicationStateDirectory()) == 0700,
           QStringLiteral("application state directory is not mode 0700"), failures);
    expect(modeBits(store.pendingDirectory()) == 0700,
           QStringLiteral("pending directory is not mode 0700"), failures);
    expect(modeBits(store.quarantineDirectory()) == 0700,
           QStringLiteral("quarantine directory is not mode 0700"), failures);
    expect(modeBits(job.manifestPath) == 0600,
           QStringLiteral("manifest is not mode 0600"), failures);
#endif

    CombineJobState::ScanResult scan = store.loadPending();
    expect(scan.jobs.size() == 1, QStringLiteral("loadPending did not return one job"), failures);
    if (scan.jobs.size() == 1) {
        expect(scan.jobs.front().title == job.title, QStringLiteral("Unicode title did not round-trip"), failures);
        expect(scan.jobs.front().sources == job.sources, QStringLiteral("sources did not round-trip"), failures);
    }

    expect(store.setStage(job, CombineJobState::Stage::InputsValidated, 0, &error),
           QStringLiteral("validated-stage save failed: %1").arg(error), failures);
    expect(store.setStage(job, CombineJobState::Stage::StreamCopyRunning, 1, &error),
           QStringLiteral("running-stage save failed: %1").arg(error), failures);
    const QString publishPath = QDir(temporary.path()).filePath(QStringLiteral("published.mp4"));
    expect(store.setPublishPath(job, publishPath, &error),
           QStringLiteral("publish-path save failed: %1").arg(error), failures);
    expect(store.setStage(job, CombineJobState::Stage::Publishing, 1, &error),
           QStringLiteral("publishing-stage save failed: %1").arg(error), failures);
    scan = store.loadPending();
    expect(scan.jobs.size() == 1, QStringLiteral("publishing manifest did not reload"), failures);
    if (scan.jobs.size() == 1) {
        expect(scan.jobs.front().encodedPath == job.encodedPath,
               QStringLiteral("encoded path did not round-trip"), failures);
        expect(scan.jobs.front().publishPath == publishPath,
               QStringLiteral("publish path did not round-trip"), failures);
        expect(scan.jobs.front().stage == CombineJobState::Stage::Publishing,
               QStringLiteral("publishing stage did not round-trip"), failures);
    }

    CombineJobState::Job stale = job;
    stale.attempt = 0;
    expect(!store.save(stale, &error), QStringLiteral("attempt counter was allowed to decrease"), failures);
    expect(QFileInfo(job.manifestPath).isFile(), QStringLiteral("failed save removed manifest"), failures);

    CombineJobState::Job forged = job;
    forged.manifestPath = QDir(temporary.path()).filePath(QStringLiteral("unrelated.json"));
    expect(!store.remove(forged, &error), QStringLiteral("forged manifest path was accepted"), failures);
    expect(QFileInfo(job.manifestPath).isFile(), QStringLiteral("forged removal touched real manifest"), failures);

    const QString malformed = QDir(store.pendingDirectory()).filePath(QStringLiteral("bad.json"));
    expect(writeBytes(malformed, QByteArrayLiteral("{\"version\":1}")),
           QStringLiteral("could not create malformed fixture"), failures);
    scan = store.loadPending();
    expect(scan.jobs.size() == 1, QStringLiteral("valid manifest disappeared during quarantine"), failures);
    expect(scan.quarantinedPaths.size() == 1,
           QStringLiteral("malformed manifest was not quarantined"), failures);
    expect(!QFileInfo::exists(malformed), QStringLiteral("malformed original remains pending"), failures);
    if (!scan.quarantinedPaths.isEmpty()) {
        expect(QFileInfo(scan.quarantinedPaths.front()).isFile(),
               QStringLiteral("quarantine did not preserve malformed bytes"), failures);
    }

    const QString unrelated = QDir(store.pendingDirectory()).filePath(QStringLiteral("keep.me"));
    expect(writeBytes(unrelated, QByteArrayLiteral("keep")),
           QStringLiteral("could not create unrelated fixture"), failures);
    expect(store.remove(job, &error), QStringLiteral("remove failed: %1").arg(error), failures);
    expect(!QFileInfo::exists(job.manifestPath), QStringLiteral("manifest remains after remove"), failures);
    expect(QFileInfo::exists(unrelated), QStringLiteral("remove touched unrelated file"), failures);

    QTextStream output(stdout);
    if (failures.isEmpty()) {
        output << "combine-job-state-test: PASS\n";
        return 0;
    }
    for (const QString &failure : failures) {
        output << "FAIL: " << failure << '\n';
    }
    return 1;
}
