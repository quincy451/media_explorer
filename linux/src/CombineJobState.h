#pragma once

#include <QDateTime>
#include <QString>
#include <QStringList>
#include <QVector>

namespace CombineJobState {

constexpr int kManifestVersion = 2;

enum class Stage {
    Prepared,
    InputsValidated,
    StreamCopyRunning,
    TranscodeRunning,
    Publishing,
    Failed,
    Completed
};

QString stageName(Stage stage);
bool stageFromName(const QString &name, Stage &stage);

// Paths are stored as clean absolute paths.  The list file must live below the
// working directory; no command line or shell text is persisted.
struct Job {
    QString id;
    QString title;
    QString outputPath;
    QString workingDirectory;
    QString listPath;
    QString encodedPath;
    QString publishPath;
    QStringList sources;
    int expectedWidth = 0;
    int expectedHeight = 0;
    qint64 expectedDurationMs = 0;
    Stage stage = Stage::Prepared;
    int attempt = 0;
    QDateTime createdUtc;
    QDateTime updatedUtc;

    // Assigned by Store::create/loadPending. It is never serialized and must
    // be exactly <pendingDirectory>/<id>.json for update or removal.
    QString manifestPath;
};

struct ScanResult {
    QVector<Job> jobs;
    QStringList quarantinedPaths;
    QStringList errors;
};

// The default location is
// $XDG_STATE_HOME/media-explorer/pending, falling back to
// ~/.local/state/media-explorer/pending.  A state-home override is provided for
// tests and portable callers; it must itself be an absolute path.
class Store {
public:
    explicit Store(const QString &stateHomeOverride = QString());

    QString applicationStateDirectory() const;
    QString pendingDirectory() const;
    QString quarantineDirectory() const;

    bool initialize(QString *error = nullptr) const;

    // create always assigns a fresh identifier and refuses a pre-populated id
    // or manifestPath. It never replaces an existing path.
    bool create(Job &job, QString *error = nullptr) const;

    // save atomically updates the exact manifest created/loaded by this Store.
    // Job identity (title, planned output, work/list/encoded paths, and sources)
    // is immutable. The publication candidate is mutable but cannot be cleared;
    // attempt cannot move backwards. The updated timestamp is assigned on success.
    bool save(Job &job, QString *error = nullptr) const;

    // Convenience integration point for validation and process stage changes.
    bool setStage(Job &job, Stage stage, int attempt,
                  QString *error = nullptr) const;

    // Records the exact no-replace publication candidate before the rename.
    // Recovery can then distinguish an encoded file awaiting publication from
    // an already-published output after a crash.
    bool setPublishPath(Job &job, const QString &publishPath,
                        QString *error = nullptr) const;

    // Every malformed *.json entry is moved intact to quarantine when
    // possible. It is never overwritten or silently deleted.
    ScanResult loadPending() const;

    // Removes only the exact, regular, owned manifest represented by job.
    // Outputs, sources, working files, and quarantined manifests are untouched.
    bool remove(const Job &job, QString *error = nullptr) const;

private:
    QString stateHome_;
};

// Validates the data/schema invariants without reading or writing the store.
bool validateJob(const Job &job, QString *error = nullptr);

} // namespace CombineJobState
