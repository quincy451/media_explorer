#pragma once

#include "AppConfig.h"
#include "CombineJobState.h"
#include "MountDiscovery.h"

#include <QHash>
#include <QKeySequence>
#include <QMainWindow>
#include <QPoint>
#include <QSet>
#include <QSize>
#include <QStringList>
#include <QThreadPool>
#include <QVector>
#include <QVariant>

#include <atomic>
#include <functional>
#include <memory>

class QAction;
class QCloseEvent;
class QEvent;
class QLabel;
class QLineEdit;
class QListWidget;
class QListWidgetItem;
class QLockFile;
class QKeyEvent;
class QMenu;
class QMouseEvent;
class QProgressBar;
class QPushButton;
class QSlider;
class QResizeEvent;
class QStackedWidget;
class QTableWidget;
class QTextBrowser;
class QTimer;
class QWidget;

struct libvlc_instance_t;
struct libvlc_media_player_t;
struct libvlc_media_t;

class CallbackWidget final : public QWidget {
public:
    explicit CallbackWidget(QWidget *parent = nullptr);
    std::function<void()> doubleClickCallback;
    std::function<void()> resizeCallback;

protected:
    void mouseDoubleClickEvent(QMouseEvent *event) override;
    void resizeEvent(QResizeEvent *event) override;
};

class MediaExplorerWindow final : public QMainWindow {
public:
    enum class KeyCommand {
        None, Help, Activate, NavigateUp, CancelFileOperation, DeleteSelected,
        RenameSelected, Refresh, NewFolder, SelectAllVideos, PlaySelected, Search,
        CopySelected, CutSelected, Paste, SubmitTopaz, SubmitIw3, ReorderUp,
        ReorderDown, Combine, EditLocation, ShowRoots, ToggleFullscreen, Pause,
        Resume, ExitPlayback, PanLeft, PanRight, PanUp, PanDown, CenterPan,
        Playlist, VideoProperties, VideoTools, RenameDeferred, CopyDeferred,
        ZoomIn, ZoomOut, ToggleLooping, Previous, Next, RemoveDeferred, VolumeUp, VolumeDown,
        SeekBack10, SeekForward10, SeekBack60, SeekForward60
    };

    explicit MediaExplorerWindow(AppConfig config, QWidget *parent = nullptr);
    ~MediaExplorerWindow() override;

    static QStringList browserShortcutMatrix();
    static QStringList playbackShortcutMatrix();
    static KeyCommand browserCommandFor(int key, Qt::KeyboardModifiers modifiers);
    static KeyCommand playbackCommandFor(int key, Qt::KeyboardModifiers modifiers);

    // Narrow live-display acceptance hooks; they exercise the same production
    // playlist/libVLC wiring without exposing libVLC handles.
    bool startPlaybackForTest(const QStringList &paths);
    int playerStateForTest() const;
    qint64 playerTimeForTest() const;
    qint64 playerLengthForTest() const;
    quint64 playbackGenerationForTest() const;
    bool setPlayerTimeForTest(qint64 milliseconds);
    QSize videoSizeForTest() const;
    void stopPlaybackForTest();

protected:
    bool eventFilter(QObject *watched, QEvent *event) override;
    void closeEvent(QCloseEvent *event) override;

private:
    enum class ViewMode { Roots, Folder, Searching, Search };
    enum class ClipboardMode { Copy, Move };
    enum class FfmpegOperation { TrimFront, TrimEnd, HorizontalFlip };
    enum class DeferredKind { Delete, Rename, Copy };

    struct Metadata {
        qint64 modifiedMs = 0;
        qint64 size = 0;
        int width = 0;
        int height = 0;
        double duration = -1.0;
        QString codec;
        QString error;
    };

    struct DeferredAction {
        DeferredKind kind;
        QString source;
        QString destination;
    };

    struct QueueOptions {
        bool iw3 = false;
        QString target = QStringLiteral("4k");
        QString profile = QStringLiteral("general");
        double grain = 0.0;
        int grainSize = 1;
        double divergence = 1.8;
        QString method = QStringLiteral("row_flow_v3");
        QString depthModel = QStringLiteral("Any_V2_N_B");
        bool tta = false;
        bool depthAa = false;
    };

    struct PendingMediaOutput {
        QString title;
        QString suggestedPath;
        QString encodedPath;
        QString workingDirectory;
        QString extraCleanupPath;
        QString details;
    };

    void buildUi();
    void buildMenus();
    QAction *addAction(QMenu *menu, const QString &text, const std::function<void()> &callback,
                       const QKeySequence &shortcut = {});
    bool dispatchBrowserKey(QKeyEvent *event);
    bool dispatchPlaybackKey(QKeyEvent *event);
    bool editableOrModalContext(QObject *watched) const;

    void setBusy(const QString &owner, const QString &message,
                 const std::shared_ptr<std::atomic_bool> &cancel = {}, int maximum = 0);
    void finishBusy(const QString &owner, const QString &message = {});
    void cancelBusy();
    void invalidateScan();
    void invalidateSearch();
    void invalidateMappingProbe();
    void clearMetadataQueue();

    void showRoots();
    void navigate(const QString &path, const QString &selectPath = {});
    void scanFinished(quint64 generation, const QString &path, QVector<Entry> entries,
                      const QString &error, bool cancelled, const QString &selectPath);
    void refresh();
    void goToLocation();
    void goUp();
    void activateSelection();
    void activateRow(int row);
    void activateMapping(const Entry &entry);
    void mappingProbeFinished(quint64 token, const Entry &entry, const MountProbeResult &result);
    void openMappingUri(const Entry &entry, const QString &stableError);

    QVector<Entry> selectedEntries() const;
    QStringList selectedVideoPaths() const;
    void selectAllVideos();
    void selectPaths(const QStringList &paths);
    void moveSelectedRow(int direction);
    void sortAndFill(bool keepSelection = false);
    void fillTable();
    QVariant sortValue(const Entry &entry, int column) const;

    void queueMetadata();
    void metadataFinished(quint64 generation, const QString &path, const Metadata &metadata);
    static Metadata probeMetadata(const Entry &entry, const AppConfig &config,
                                  const std::shared_ptr<std::atomic_bool> &cancel = {});
    void showVideoProperties();
    void showPropertiesDialog(const QString &path, const Metadata &metadata);

    void promptSearch();
    void startSearch(const QStringList &scopes, const QStringList &terms, bool refine);
    void searchFinished(quint64 token, QVector<Entry> entries, int directories, int files,
                        QStringList errors, bool cancelled);

    bool fileChangesAllowed(const QString &action) const;
    void copySelected();
    void cutSelected();
    void paste();
    void deleteSelected();
    void newFolder();
    void renameSelected();
    void openExternal();
    void runFileOperation(const QString &operation, const QStringList &sources,
                          const QString &destination = {},
                          const QVector<DeferredAction> &deferred = {});
    void fileOperationFinished(const QString &operation, int completed, int total,
                               const QStringList &errors, bool cancelled);

    bool ensurePlayer();
    void playSelected();
    void playEntries(const QVector<Entry> &entries);
    void playIndex(int index, bool resetView = true);
    void embedVideo();
    void pausePlayback(bool userRequested = true);
    void resumePlayback(bool userRequested = true);
    bool beginPlaybackPauseHold();
    void endPlaybackPauseHold();
    void updatePlaybackTitle();
    void toggleLooping();
    static int playbackIndexAfterEnd(int currentIndex, int playlistSize, bool looping);
    void processPendingPlaybackEnd();
    void previousVideo();
    void nextVideo();
    void seekBy(qint64 milliseconds);
    void setVolume(int value);
    void pollPlayer();
    void toggleFullscreen();
    void stopPlayback();
    void removeCurrentDeferred();
    void renameCurrentDeferred();
    void copyCurrentDeferred();
    void adjustZoom(bool zoomIn);
    void panVideo(int dx, int dy);
    void centerVideoPan();
    void updateVideoGeometry();
    void showPlaylistChooser();

    void showVideoTools();
    void startFfmpegTool(FfmpegOperation operation);
    void combineSelected();
    void beginCombineJob(CombineJobState::Job job);
    bool ensureCombineLock(bool showDialog);
    void loadPendingCombineJobs();
    void resumeNextCombineJob();
    void finishActiveCombineJob(bool published, bool cancelled, const QString &error = {});
    bool publishActiveCombineOutput(QString &publishedPath, QString &error);
    void startMediaProcess(const QString &title, const QString &program,
                           const QStringList &arguments, const QString &outputPath,
                           double expectedSeconds = -1.0,
                           const QString &fallbackProgram = {},
                           const QStringList &fallbackArguments = {},
                           const QString &temporaryPath = {},
                           bool deferPublication = false,
                           const QString &encodedPathOverride = {});
    void mediaProcessFinished(const QString &title, const QString &outputPath,
                              const QString &encodedPath, const QString &workingDirectory,
                              const QString &temporaryPath, bool deferPublication,
                              bool success, bool cancelled, const QString &details);
    bool publishMediaOutput(PendingMediaOutput output, QString &publishedPath, QString &error);
    void publishPendingMediaOutputs();

    void submitTopazQueue();
    void submitIw3Queue();
    bool promptTopazOptions(QueueOptions &options);
    bool promptIw3Options(QueueOptions &options);
    void startQueueSubmit(const QStringList &sources, const QueueOptions &options);

    void showHelp();
    void showAbout();
    void showError(const QString &title, const QString &message) const;

    AppConfig config_;
    QThreadPool workerPool_;
    QThreadPool metadataPool_;
    QThreadPool mappingPool_;
    QVector<Entry> entries_;
    QHash<QString, int> rowByPath_;
    ViewMode viewMode_ = ViewMode::Roots;
    QString currentDirectory_;
    QStringList searchTerms_;
    QStringList searchScopes_;
    QString searchReturnDirectory_;
    bool searchReturnRoots_ = true;
    int sortColumn_ = 0;
    bool sortAscending_ = true;
    quint64 scanGeneration_ = 0;
    quint64 searchToken_ = 0;
    quint64 mappingToken_ = 0;
    quint64 metadataGeneration_ = 0;
    std::shared_ptr<std::atomic_bool> scanCancel_;
    std::shared_ptr<std::atomic_bool> searchCancel_;
    std::shared_ptr<std::atomic_bool> busyCancel_;
    std::shared_ptr<std::atomic_bool> fileOperationCancel_;
    std::shared_ptr<std::atomic_bool> mediaProcessCancel_;
    QString busyOwner_;
    QHash<QString, Metadata> metadataCache_;
    QSet<QString> metadataPending_;
    QStringList clipboardPaths_;
    ClipboardMode clipboardMode_ = ClipboardMode::Copy;
    QVector<DeferredAction> deferredActions_;
    QVector<PendingMediaOutput> pendingMediaOutputs_;
    QString activeMediaOutput_;
    CombineJobState::Store combineStore_;
    std::unique_ptr<QLockFile> combineLock_;
    CombineJobState::Job activeCombineJob_;
    QVector<CombineJobState::Job> combineResumeQueue_;
    bool hasActiveCombineJob_ = false;

    QStackedWidget *stack_ = nullptr;
    QWidget *browserPage_ = nullptr;
    QWidget *playerPage_ = nullptr;
    QLineEdit *location_ = nullptr;
    QTableWidget *table_ = nullptr;
    QLabel *playerTitle_ = nullptr;
    CallbackWidget *videoViewport_ = nullptr;
    CallbackWidget *videoFrame_ = nullptr;
    QListWidget *playlistWidget_ = nullptr;
    QSlider *seekSlider_ = nullptr;
    QLabel *timeLabel_ = nullptr;
    QSlider *volumeSlider_ = nullptr;
    QPushButton *pauseButton_ = nullptr;
    QWidget *playerControls_ = nullptr;
    QProgressBar *progress_ = nullptr;
    QPushButton *cancelButton_ = nullptr;
    QTimer *playerTimer_ = nullptr;

    libvlc_instance_t *vlcInstance_ = nullptr;
    libvlc_media_player_t *vlcPlayer_ = nullptr;
    libvlc_media_t *currentMedia_ = nullptr;
    QStringList playlist_;
    int playlistIndex_ = -1;
    bool seeking_ = false;
    bool playbackErrorShown_ = false;
    int lastVlcState_ = -1;
    bool fullscreen_ = false;
    bool wasMaximizedBeforeFullscreen_ = false;
    bool exitingPlayback_ = false;
    QSet<QObject *> playbackDialogs_;
    int playbackPauseHolds_ = 0;
    bool resumeAfterPlaybackPause_ = false;
    bool playbackPauseRestorePending_ = false;
    quint64 playbackPauseRestoreToken_ = 0;
    bool looping_ = false;
    bool playbackEndPending_ = false;
    quint64 playbackEndGeneration_ = 0;
    quint64 playbackMediaGeneration_ = 0;
    double zoom_ = 1.0;
    QPoint panOffset_;
};
