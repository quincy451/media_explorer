#pragma once

#include <QByteArray>
#include <QDateTime>
#include <QSet>
#include <QString>
#include <QStringList>
#include <QVector>

#include <functional>

namespace MediaTools {

// Output names use the Windows convention: name.ext, name (1).ext, ... .
// Existing files, including dangling symbolic links, and reservedPaths collide.
QString uniqueOutputPath(const QString &sourcePath, const QString &tag,
                         const QSet<QString> &reservedPaths = {});

QStringList trimFrontArguments(const QString &inputPath, const QString &outputPath,
                               qint64 positionMilliseconds);
QStringList trimEndArguments(const QString &inputPath, const QString &outputPath,
                             qint64 positionMilliseconds);
QStringList horizontalFlipArguments(const QString &inputPath, const QString &outputPath);

// Returns an empty string and sets error for an empty/invalid path list. Paths
// are quoted according to ffconcat's token grammar, including literal apostrophes
// and backslashes.
QString ffconcatContent(const QStringList &inputPaths, QString *error = nullptr);
QStringList combineStreamCopyArguments(const QString &listPath, const QString &outputPath);
QStringList combineX264AacArguments(const QString &listPath, const QString &outputPath);

QStringList ffprobeVideoArguments(const QString &mediaPath);
QStringList ffprobeAudioArguments(const QString &mediaPath);
QStringList ffprobeDurationArguments(const QString &mediaPath);

enum class TopazTarget { K4, K8 };
enum class TopazProfile {
    General,
    Repair,
    Stabilize,
    Deblur,
    Denoise,
    DeinterlaceRepair,
    Repair2Pass,
    GeneralGrain,
    RepairGrain
};

struct TopazOptions {
    TopazTarget target = TopazTarget::K4;
    TopazProfile profile = TopazProfile::General;
    double grain = 0.0;
    int grainSize = 1;
};

struct TopazProfilePreset {
    TopazProfile profile;
    QString displayName;
    QString schemaName;
    double grain;
    int grainSize;
};

QVector<TopazProfilePreset> topazProfilePresets();
TopazOptions topazOptionsForProfile(TopazTarget target, TopazProfile profile);

enum class Iw3Preset {
    FastPreview,
    Recommended,
    Strong,
    MaxQuality,
    MaxQualityStrong,
    Wild
};

struct Iw3Options {
    double divergence = 1.8;
    double convergence = 0.5;
    QString method = QStringLiteral("row_flow_v3");
    QString depthModel = QStringLiteral("Any_V2_N_B");
    QString syntheticView = QStringLiteral("right");
    bool preserveScreenBorder = true;
    bool emaNormalize = true;
    bool sceneDetect = true;
    bool tta = false;
    bool depthAA = false;
    int gpu = 0;
    int maxFps = 30;
    QString videoFormat = QStringLiteral("mp4");
    QString videoCodec = QStringLiteral("hevc_nvenc");
    int crf = 20;
    QString preset = QStringLiteral("p5");
    QString pixelFormat = QStringLiteral("yuv420p");
    bool halfSbs = false;
};

struct Iw3PresetModel {
    Iw3Preset preset;
    QString displayName;
    Iw3Options options;
};

QVector<Iw3PresetModel> iw3Presets();
Iw3Options iw3OptionsForPreset(Iw3Preset preset);

// An invalid submittedUtc means "now". Output matches the Windows queue JSON
// schema and formatting, including its numeric 0/1 option fields.
QByteArray buildTopazJobJson(const QString &sourceOriginal,
                             const QString &queuedFileName,
                             const TopazOptions &options,
                             const QDateTime &submittedUtc = QDateTime());
QByteArray buildIw3JobJson(const QString &sourceOriginal,
                           const QString &queuedFileName,
                           const Iw3Options &options,
                           const QDateTime &submittedUtc = QDateTime());

enum class QueueSubmitStatus { Published, Cancelled, Failed };

struct QueueSubmitResult {
    QueueSubmitStatus status = QueueSubmitStatus::Failed;
    QString videoPath;
    QString jsonPath;
    QString error;

    // Complete recovery artifacts or partial files that could not be cleaned.
    // A normal success or clean cancellation leaves this empty.
    QStringList retainedPaths;

    bool succeeded() const { return status == QueueSubmitStatus::Published; }
};

using QueueJsonBuilder =
    std::function<QByteArray(const QString &sourceOriginal, const QString &queuedFileName)>;
using CancelCheck = std::function<bool()>;
using CopyProgress = std::function<void(qint64 bytesCopied, qint64 totalBytes)>;

// Accepts only a regular, non-symlink source. Copies to an exclusively-created
// video name, then writes a non-triggering temporary JSON file and publishes the
// final .json atomically without replace. A failed publish preserves the complete
// video/temp JSON and reports both. Directory fsync is best-effort because some
// network filesystems do not implement it.
QueueSubmitResult submitQueueJob(const QString &sourcePath,
                                 const QString &queueDirectory,
                                 const QueueJsonBuilder &jsonBuilder,
                                 const CancelCheck &isCancelled = {},
                                 const CopyProgress &progress = {});

// Pure deterministic checks; no files or external programs are created/run.
QStringList selfTestFailures();

} // namespace MediaTools
