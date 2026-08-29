#include "../src/AppConfig.h"
#include "../src/MountDiscovery.h"

#include <QAction>
#include <QApplication>
#include <QDialog>
#include <QEvent>
#include <QKeyEvent>
#include <QLineEdit>
#include <QMenu>
#include <QPlainTextEdit>
#include <QPushButton>
#include <QShortcut>
#include <QSlider>
#include <QStackedWidget>
#include <QTableWidget>
#include <QTextEdit>
#include <QTimer>
#include <QWidget>

#include <atomic>
#include <functional>
#include <memory>

// This translation unit intentionally reaches into the window's state.  The
// production header and binary remain unchanged; the access is used only to
// put the UI into browser/player modes and verify observable routing effects.
#define private public
#define protected public
#include "../src/MediaExplorerWindow.h"
#undef protected
#undef private

#include <iostream>

namespace {

class TestContext {
public:
    void expect(bool condition, const QString &message) {
        if (condition) return;
        ++failures;
        std::cerr << "FAIL: " << message.toStdString() << '\n';
    }

    int failures = 0;
};

class FallthroughProbe final : public QObject {
public:
    bool eventFilter(QObject *, QEvent *event) override {
        if (armed && event->type() == QEvent::KeyPress) ++keyPresses;
        return false;
    }

    void reset() { keyPresses = 0; }

    bool armed = false;
    int keyPresses = 0;
};

int sendKey(QApplication &application, FallthroughProbe &probe, QWidget *target,
            int key, Qt::KeyboardModifiers modifiers = Qt::NoModifier) {
    target->setFocus(Qt::OtherFocusReason);
    application.processEvents();
    probe.reset();
    // The MediaExplorerWindow filter is application-wide.  A filter on the
    // concrete target runs only when that global dispatcher lets the event
    // through, which makes this a reliable consumed-vs-native measurement on
    // both Qt's xcb and offscreen platform plugins.
    target->installEventFilter(&probe);
    probe.armed = true;
    QKeyEvent event(QEvent::KeyPress, key, modifiers);
    QApplication::sendEvent(target, &event);
    probe.armed = false;
    target->removeEventFilter(&probe);
    return probe.keyPresses;
}

struct Route {
    int key;
    Qt::KeyboardModifiers modifiers;
    MediaExplorerWindow::KeyCommand expected;
    const char *description;
};

void checkBrowserClassifier(TestContext &test) {
    using Command = MediaExplorerWindow::KeyCommand;
    const Qt::KeyboardModifiers control = Qt::ControlModifier;
    const Qt::KeyboardModifiers shift = Qt::ShiftModifier;
    const Qt::KeyboardModifiers alt = Qt::AltModifier;
    const Qt::KeyboardModifiers keypad = Qt::KeypadModifier;

    const Route routes[] = {
        {Qt::Key_F1, {}, Command::Help, "browser F1"},
        {Qt::Key_Return, {}, Command::Activate, "browser Return"},
        {Qt::Key_Enter, keypad, Command::Activate, "browser keypad Enter"},
        {Qt::Key_Left, {}, Command::NavigateUp, "browser Left"},
        {Qt::Key_Backspace, {}, Command::NavigateUp, "browser Backspace"},
        {Qt::Key_Left, alt, Command::NavigateUp, "browser Alt+Left"},
        {Qt::Key_Escape, {}, Command::CancelFileOperation, "browser Escape"},
        {Qt::Key_Delete, {}, Command::DeleteSelected, "browser Delete"},
        {Qt::Key_F2, {}, Command::RenameSelected, "browser F2"},
        {Qt::Key_F5, {}, Command::Refresh, "browser F5"},
        {Qt::Key_N, control | shift, Command::NewFolder, "browser Ctrl+Shift+N"},
        {Qt::Key_A, control, Command::SelectAllVideos, "browser Ctrl+A"},
        {Qt::Key_P, control, Command::PlaySelected, "browser Ctrl+P"},
        {Qt::Key_F, control, Command::Search, "browser Ctrl+F"},
        {Qt::Key_C, control, Command::CopySelected, "browser Ctrl+C"},
        {Qt::Key_X, control, Command::CutSelected, "browser Ctrl+X"},
        {Qt::Key_V, control, Command::Paste, "browser Ctrl+V"},
        {Qt::Key_U, control, Command::SubmitTopaz, "browser Ctrl+U"},
        {Qt::Key_3, control, Command::SubmitIw3, "browser Ctrl+3"},
        {Qt::Key_3, control | keypad, Command::SubmitIw3, "browser keypad Ctrl+3"},
        {Qt::Key_Up, control, Command::ReorderUp, "browser Ctrl+Up"},
        {Qt::Key_Down, control, Command::ReorderDown, "browser Ctrl+Down"},
        {Qt::Key_Plus, control, Command::Combine, "browser Ctrl+Plus"},
        {Qt::Key_Plus, control | keypad, Command::Combine, "browser keypad Ctrl+Plus"},
        {Qt::Key_Equal, control, Command::Combine, "browser Ctrl+Equal"},
        {Qt::Key_Equal, control | shift, Command::Combine, "browser Ctrl+Shift+Equal"},
        {Qt::Key_L, control, Command::EditLocation, "browser Ctrl+L"},
        {Qt::Key_Home, control, Command::ShowRoots, "browser Ctrl+Home"},
    };
    for (const Route &route : routes) {
        const Command actual = MediaExplorerWindow::browserCommandFor(route.key, route.modifiers);
        test.expect(actual == route.expected, QString::fromLatin1(route.description));
    }

    const Route negatives[] = {
        {Qt::Key_Up, {}, Command::None, "browser ordinary Up remains native"},
        {Qt::Key_Down, {}, Command::None, "browser ordinary Down remains native"},
        {Qt::Key_Right, {}, Command::None, "browser ordinary Right remains native"},
        {Qt::Key_A, {}, Command::None, "browser plain A"},
        {Qt::Key_Plus, {}, Command::None, "browser plain Plus"},
        {Qt::Key_Home, {}, Command::None, "browser plain Home"},
        {Qt::Key_Right, control, Command::None, "browser Ctrl+Right"},
        {Qt::Key_Right, alt, Command::None, "browser Alt+Right"},
    };
    for (const Route &route : negatives) {
        const Command actual = MediaExplorerWindow::browserCommandFor(route.key, route.modifiers);
        test.expect(actual == route.expected, QString::fromLatin1(route.description));
    }
}

void checkPlaybackClassifier(TestContext &test) {
    using Command = MediaExplorerWindow::KeyCommand;
    const Qt::KeyboardModifiers control = Qt::ControlModifier;
    const Qt::KeyboardModifiers shift = Qt::ShiftModifier;
    const Qt::KeyboardModifiers alt = Qt::AltModifier;
    const Qt::KeyboardModifiers keypad = Qt::KeypadModifier;

    const Route routes[] = {
        {Qt::Key_F1, {}, Command::Help, "playback F1"},
        {Qt::Key_Return, {}, Command::ToggleFullscreen, "playback Return"},
        {Qt::Key_Enter, keypad, Command::ToggleFullscreen, "playback keypad Enter"},
        {Qt::Key_Space, {}, Command::Pause, "playback Space forces pause"},
        {Qt::Key_Tab, {}, Command::Resume, "playback Tab forces resume"},
        {Qt::Key_Escape, {}, Command::ExitPlayback, "playback Escape"},
        {Qt::Key_Left, alt, Command::PanLeft, "playback Alt+Left"},
        {Qt::Key_Right, alt, Command::PanRight, "playback Alt+Right"},
        {Qt::Key_Up, alt, Command::PanUp, "playback Alt+Up"},
        {Qt::Key_Down, alt, Command::PanDown, "playback Alt+Down"},
        {Qt::Key_Home, alt, Command::CenterPan, "playback Alt+Home"},
        {Qt::Key_G, control, Command::Playlist, "playback Ctrl+G"},
        {Qt::Key_P, control, Command::VideoProperties, "playback Ctrl+P"},
        {Qt::Key_V, control, Command::VideoTools, "playback Ctrl+V"},
        {Qt::Key_R, control, Command::RenameDeferred, "playback Ctrl+R"},
        {Qt::Key_C, control, Command::CopyDeferred, "playback Ctrl+C"},
        {Qt::Key_L, control, Command::ToggleLooping, "playback Ctrl+L"},
        {Qt::Key_Z, control, Command::ZoomIn, "playback Ctrl+Z"},
        {Qt::Key_X, control, Command::ZoomOut, "playback Ctrl+X"},
        {Qt::Key_Left, control, Command::Previous, "playback Ctrl+Left"},
        {Qt::Key_Right, control, Command::Next, "playback Ctrl+Right"},
        {Qt::Key_Delete, {}, Command::RemoveDeferred, "playback Delete"},
        {Qt::Key_Up, {}, Command::VolumeUp, "playback Up"},
        {Qt::Key_Down, {}, Command::VolumeDown, "playback Down"},
        {Qt::Key_Left, {}, Command::SeekBack10, "playback Left"},
        {Qt::Key_Right, {}, Command::SeekForward10, "playback Right"},
        {Qt::Key_Left, shift, Command::SeekBack60, "playback Shift+Left"},
        {Qt::Key_Right, shift, Command::SeekForward60, "playback Shift+Right"},

        // Modifier precedence is a release requirement: Alt > Ctrl > Shift.
        {Qt::Key_Left, alt | control | shift, Command::PanLeft,
         "playback Alt wins for Left"},
        {Qt::Key_Right, alt | control | shift, Command::PanRight,
         "playback Alt wins for Right"},
        {Qt::Key_Up, alt | control, Command::PanUp, "playback Alt wins for Up"},
        {Qt::Key_Down, alt | control, Command::PanDown, "playback Alt wins for Down"},
        {Qt::Key_Home, alt | control, Command::CenterPan, "playback Alt wins for Home"},
        {Qt::Key_Left, control | shift, Command::Previous,
         "playback Ctrl wins over Shift for Left"},
        {Qt::Key_Right, control | shift, Command::Next,
         "playback Ctrl wins over Shift for Right"},
        {Qt::Key_Up, control, Command::VolumeUp, "playback Ctrl+Up is volume"},
        {Qt::Key_Down, control, Command::VolumeDown, "playback Ctrl+Down is volume"},
    };
    for (const Route &route : routes) {
        const Command actual = MediaExplorerWindow::playbackCommandFor(route.key, route.modifiers);
        test.expect(actual == route.expected, QString::fromLatin1(route.description));
    }

    const Route negatives[] = {
        {Qt::Key_A, {}, Command::None, "playback plain A"},
        {Qt::Key_Home, {}, Command::None, "playback plain Home"},
        {Qt::Key_3, control, Command::None, "playback Ctrl+3 remains browser-only"},
        {Qt::Key_U, control, Command::None, "playback Ctrl+U remains browser-only"},
        {Qt::Key_Plus, control, Command::None, "playback Ctrl+Plus remains browser-only"},
    };
    for (const Route &route : negatives) {
        const Command actual = MediaExplorerWindow::playbackCommandFor(route.key, route.modifiers);
        test.expect(actual == route.expected, QString::fromLatin1(route.description));
    }

    test.expect(MediaExplorerWindow::playbackIndexAfterEnd(1, 3, true) == 1,
                QStringLiteral("looping replays the current playlist item"));
    test.expect(MediaExplorerWindow::playbackIndexAfterEnd(1, 3, false) == 2,
                QStringLiteral("non-looping playback advances to the next playlist item"));
    test.expect(MediaExplorerWindow::playbackIndexAfterEnd(2, 3, false) == -1,
                QStringLiteral("non-looping playback stops after the final playlist item"));
    test.expect(MediaExplorerWindow::playbackIndexAfterEnd(-1, 3, true) == -1,
                QStringLiteral("invalid playback indexes are never replayed"));
}

void checkAdvertisedMatrix(TestContext &test) {
    const QStringList expectedBrowser{
        QStringLiteral("Enter=open/play"), QStringLiteral("Left|Backspace|Alt+Left=up"),
        QStringLiteral("Ctrl+A=select-videos"), QStringLiteral("Ctrl+P=play"),
        QStringLiteral("Ctrl+F=search/refine"), QStringLiteral("Ctrl+Up|Ctrl+Down=reorder"),
        QStringLiteral("Ctrl++=combine"), QStringLiteral("Ctrl+U=topaz"),
        QStringLiteral("Ctrl+3=iw3"), QStringLiteral("Ctrl+C|Ctrl+X|Ctrl+V=clipboard"),
        QStringLiteral("Delete=delete"), QStringLiteral("Escape=cancel"),
        QStringLiteral("F2=rename"), QStringLiteral("Ctrl+Shift+N=new-folder"),
        QStringLiteral("F5=refresh"), QStringLiteral("Ctrl+L=location"),
        QStringLiteral("Ctrl+Home=mounts"), QStringLiteral("F1=help")};
    const QStringList expectedPlayback{
        QStringLiteral("F1=help"), QStringLiteral("Enter=fullscreen"),
        QStringLiteral("Space=pause"), QStringLiteral("Tab=resume"),
        QStringLiteral("Escape=exit"), QStringLiteral("Ctrl+G=playlist"),
        QStringLiteral("Ctrl+P=properties"), QStringLiteral("Ctrl+V=tools"),
        QStringLiteral("Ctrl+R=rename-on-exit"), QStringLiteral("Ctrl+C=copy-on-exit"),
        QStringLiteral("Ctrl+L=loop-current"), QStringLiteral("Ctrl+Z|Ctrl+X=zoom"),
        QStringLiteral("Delete=delete-on-exit"),
        QStringLiteral("Alt+Arrows|Alt+Home=pan"), QStringLiteral("Up|Down=volume"),
        QStringLiteral("Left|Right=seek"), QStringLiteral("Shift+Left|Shift+Right=seek-60"),
        QStringLiteral("Ctrl+Left|Ctrl+Right=previous/next")};
    test.expect(MediaExplorerWindow::browserShortcutMatrix() == expectedBrowser,
                QStringLiteral("browser help/self-test matrix is complete"));
    test.expect(MediaExplorerWindow::playbackShortcutMatrix() == expectedPlayback,
                QStringLiteral("playback help/self-test matrix is complete"));
}

void checkEventDelivery(QApplication &application, TestContext &test,
                        FallthroughProbe &probe, MediaExplorerWindow &window) {
    window.show();
    window.activateWindow();
    application.processEvents();

    test.expect(window.findChild<QTableWidget *>(QStringLiteral("fileTable")) == window.table_,
                QStringLiteral("file table has a stable test identity"));
    test.expect(window.findChild<QWidget *>(QStringLiteral("videoFrame")) == window.videoFrame_,
                QStringLiteral("video frame has a stable test identity"));
    test.expect(window.findChild<QSlider *>(QStringLiteral("seekSlider")) == window.seekSlider_,
                QStringLiteral("seek slider has a stable test identity"));
    test.expect(window.videoFrame_->focusPolicy() != Qt::NoFocus,
                QStringLiteral("native video surface accepts keyboard focus"));

    // Menu entries display their keys in labels only.  A real QAction shortcut
    // would bypass the mode dispatcher and could execute a command twice.
    for (QAction *action : window.findChildren<QAction *>()) {
        test.expect(action->shortcut().isEmpty(),
                    QStringLiteral("no duplicate QAction shortcut: %1").arg(action->text()));
    }

    window.stack_->setCurrentWidget(window.browserPage_);
    window.viewMode_ = MediaExplorerWindow::ViewMode::Roots;
    test.expect(sendKey(application, probe, window.table_, Qt::Key_Up) == 1,
                QStringLiteral("ordinary browser Up falls through to the table"));
    test.expect(sendKey(application, probe, window.table_, Qt::Key_Up,
                        Qt::ControlModifier) == 0,
                QStringLiteral("browser Ctrl+Up is consumed exactly once"));
    window.location_->setReadOnly(true);
    test.expect(sendKey(application, probe, window.table_, Qt::Key_L,
                        Qt::ControlModifier) == 0,
                QStringLiteral("browser Ctrl+L is consumed"));
    test.expect(!window.location_->isReadOnly() && window.location_->hasFocus(),
                QStringLiteral("browser Ctrl+L still edits the location"));

    // Browser Escape must target only the active file mutation.  It must not
    // cancel search/scan/media work or leave Search view.
    auto fileCancel = std::make_shared<std::atomic_bool>(false);
    auto scanCancel = std::make_shared<std::atomic_bool>(false);
    auto searchCancel = std::make_shared<std::atomic_bool>(false);
    auto busyCancel = std::make_shared<std::atomic_bool>(false);
    auto mediaCancel = std::make_shared<std::atomic_bool>(false);
    window.fileOperationCancel_ = fileCancel;
    window.scanCancel_ = scanCancel;
    window.searchCancel_ = searchCancel;
    window.busyCancel_ = busyCancel;
    window.mediaProcessCancel_ = mediaCancel;
    test.expect(sendKey(application, probe, window.table_, Qt::Key_Escape) == 0,
                QStringLiteral("browser Escape is consumed"));
    test.expect(fileCancel->load(), QStringLiteral("browser Escape cancels the file operation"));
    test.expect(!scanCancel->load() && !searchCancel->load() && !busyCancel->load() &&
                    !mediaCancel->load(),
                QStringLiteral("browser Escape cancels no unrelated worker"));
    window.fileOperationCancel_.reset();
    window.scanCancel_.reset();
    window.searchCancel_.reset();
    window.busyCancel_.reset();
    window.mediaProcessCancel_.reset();

    window.viewMode_ = MediaExplorerWindow::ViewMode::Search;
    searchCancel = std::make_shared<std::atomic_bool>(false);
    window.searchCancel_ = searchCancel;
    sendKey(application, probe, window.table_, Qt::Key_Escape);
    test.expect(window.viewMode_ == MediaExplorerWindow::ViewMode::Search && !searchCancel->load(),
                QStringLiteral("browser Escape is a no-op in Search without a file operation"));
    window.searchCancel_.reset();

    // Line and multiline editors retain their keys instead of invoking the
    // browser dispatcher.
    window.stack_->setCurrentWidget(window.browserPage_);
    window.location_->setReadOnly(false);
    test.expect(sendKey(application, probe, window.location_, Qt::Key_F1) == 1,
                QStringLiteral("QLineEdit keys bypass the global dispatcher"));

    QPlainTextEdit plainEditor(window.browserPage_);
    plainEditor.setGeometry(20, 80, 240, 100);
    plainEditor.show();
    test.expect(sendKey(application, probe, &plainEditor, Qt::Key_F1) == 1,
                QStringLiteral("QPlainTextEdit keys bypass the global dispatcher"));
    plainEditor.hide();

    QTextEdit richEditor(window.browserPage_);
    richEditor.setGeometry(20, 80, 240, 100);
    richEditor.show();
    test.expect(sendKey(application, probe, &richEditor, Qt::Key_F1) == 1,
                QStringLiteral("QTextEdit keys bypass the global dispatcher"));
    richEditor.hide();

    QDialog modal(&window);
    modal.setModal(true);
    QPushButton modalButton(QStringLiteral("modal target"), &modal);
    modal.show();
    application.processEvents();
    test.expect(QApplication::activeModalWidget() == &modal,
                QStringLiteral("modal test window became active"));
    test.expect(sendKey(application, probe, &modalButton, Qt::Key_F1) == 1,
                QStringLiteral("modal dialog keys bypass the global dispatcher"));
    modal.close();
    application.processEvents();

    QMenu popup(&window);
    popup.addAction(QStringLiteral("popup target"));
    popup.popup(window.mapToGlobal(QPoint(20, 20)));
    application.processEvents();
    test.expect(sendKey(application, probe, &popup, Qt::Key_F1) == 1,
                QStringLiteral("popup menu keys bypass the global dispatcher"));
    popup.close();
    application.processEvents();

    // The same player commands must be intercepted from both the native video
    // surface and the seek slider.  Volume changes by exactly five provide an
    // observable once-only assertion.
    window.stack_->setCurrentWidget(window.playerPage_);
    window.exitingPlayback_ = false;
    window.videoViewport_->resize(640, 360);
    application.processEvents();
    window.volumeSlider_->setValue(100);
    test.expect(sendKey(application, probe, window.videoFrame_, Qt::Key_Up) == 0,
                QStringLiteral("playback Up is consumed from videoFrame"));
    test.expect(window.videoFrame_->hasFocus(), QStringLiteral("videoFrame really held focus"));
    test.expect(window.volumeSlider_->value() == 105,
                QStringLiteral("videoFrame Up executes volume exactly once"));
    test.expect(sendKey(application, probe, window.videoFrame_, Qt::Key_Space) == 0,
                QStringLiteral("Space is consumed from videoFrame"));
    test.expect(sendKey(application, probe, window.videoFrame_, Qt::Key_Tab) == 0,
                QStringLiteral("Tab is consumed from videoFrame"));
    test.expect(!window.looping_, QStringLiteral("playback looping starts off"));
    test.expect(sendKey(application, probe, window.videoFrame_, Qt::Key_L,
                        Qt::ControlModifier) == 0,
                QStringLiteral("playback Ctrl+L is consumed from videoFrame"));
    test.expect(window.looping_, QStringLiteral("playback Ctrl+L enables looping"));
    sendKey(application, probe, window.videoFrame_, Qt::Key_L, Qt::ControlModifier);
    test.expect(!window.looping_, QStringLiteral("second playback Ctrl+L disables looping"));

    window.zoom_ = 2.0;
    window.panOffset_ = {};
    window.updateVideoGeometry();
    window.volumeSlider_->setValue(100);
    test.expect(sendKey(application, probe, window.videoFrame_, Qt::Key_Up,
                        Qt::AltModifier | Qt::ControlModifier) == 0,
                QStringLiteral("Alt+Ctrl+Up is consumed from videoFrame"));
    test.expect(window.volumeSlider_->value() == 100 && window.panOffset_.y() == -40,
                QStringLiteral("Alt pan wins over Ctrl/volume and runs once"));
    sendKey(application, probe, window.videoFrame_, Qt::Key_Home, Qt::AltModifier);
    test.expect(window.panOffset_.isNull(), QStringLiteral("Alt+Home recenters once"));

    window.zoom_ = 1.0;
    sendKey(application, probe, window.videoFrame_, Qt::Key_Z, Qt::ControlModifier);
    test.expect(qFuzzyCompare(window.zoom_, 1.25),
                QStringLiteral("Ctrl+Z zooms in exactly once"));
    sendKey(application, probe, window.videoFrame_, Qt::Key_X, Qt::ControlModifier);
    test.expect(qFuzzyCompare(window.zoom_, 1.0),
                QStringLiteral("Ctrl+X zooms toward fit exactly once"));

    window.volumeSlider_->setValue(100);
    test.expect(sendKey(application, probe, window.seekSlider_, Qt::Key_Up) == 0,
                QStringLiteral("playback Up is consumed from seekSlider"));
    test.expect(window.seekSlider_->hasFocus(), QStringLiteral("seekSlider really held focus"));
    test.expect(window.volumeSlider_->value() == 105,
                QStringLiteral("seekSlider Up executes volume exactly once"));
    test.expect(sendKey(application, probe, window.seekSlider_, Qt::Key_L,
                        Qt::ControlModifier) == 0,
                QStringLiteral("playback Ctrl+L is consumed from seekSlider"));
    test.expect(window.looping_, QStringLiteral("seekSlider Ctrl+L enables looping"));
    sendKey(application, probe, window.seekSlider_, Qt::Key_L, Qt::ControlModifier);
    test.expect(!window.looping_, QStringLiteral("second seekSlider Ctrl+L disables looping"));
    window.seekSlider_->setValue(500);
    test.expect(sendKey(application, probe, window.seekSlider_, Qt::Key_Left) == 0,
                QStringLiteral("Left seek is consumed from seekSlider"));
    test.expect(window.seekSlider_->value() == 500,
                QStringLiteral("native slider does not also process routed Left"));
    test.expect(sendKey(application, probe, window.seekSlider_, Qt::Key_Space) == 0,
                QStringLiteral("Space is consumed from seekSlider"));
    test.expect(sendKey(application, probe, window.seekSlider_, Qt::Key_Tab) == 0,
                QStringLiteral("Tab is consumed from seekSlider"));
    test.expect(sendKey(application, probe, window.seekSlider_, Qt::Key_A) == 1,
                QStringLiteral("unbound player key falls through"));

    // Win32's Ctrl+V dialog accepts digits directly.  Inspect the live modal
    // while both optional tool families are enabled, then reject it before any
    // media operation can start.
    window.config_.upscaleDirectory = QStringLiteral("/tmp/media-explorer-key-test");
    window.config_.ffmpegAvailable = true;
    window.playlist_ = QStringList{
        QStringLiteral("/tmp/keyboard-acceptance-placeholder.mp4")};
    window.playlistIndex_ = 0;
    QStringList videoToolKeys;
    QTimer::singleShot(0, &window, [&test, &videoToolKeys] {
        auto *dialog = qobject_cast<QDialog *>(QApplication::activeModalWidget());
        test.expect(dialog && dialog->windowTitle() == QStringLiteral("Video tools"),
                    QStringLiteral("Ctrl+V opens the Video tools modal"));
        if (!dialog) return;
        for (QShortcut *shortcut : dialog->findChildren<QShortcut *>())
            videoToolKeys.append(shortcut->key().toString(QKeySequence::PortableText));
        dialog->reject();
    });
    QTimer::singleShot(500, &window, [] {
        if (auto *dialog = qobject_cast<QDialog *>(QApplication::activeModalWidget()))
            dialog->reject();
    });
    window.showVideoTools();
    videoToolKeys.sort();
    test.expect(videoToolKeys == QStringList({QStringLiteral("1"), QStringLiteral("2"),
                                              QStringLiteral("3"), QStringLiteral("4")}),
                QStringLiteral("Video tools modal binds direct keys 1, 2, 3, and 4"));
    window.playlist_.clear();
    window.playlistIndex_ = -1;

    window.viewMode_ = MediaExplorerWindow::ViewMode::Roots;
    window.looping_ = true;
    window.stopPlayback();
    test.expect(!window.looping_, QStringLiteral("looping resets when playback exits"));
    window.exitingPlayback_ = false;
    window.close();
    application.processEvents();
}

} // namespace

int main(int argc, char **argv) {
    QApplication application(argc, argv);
    application.setQuitOnLastWindowClosed(false);
    TestContext test;
    FallthroughProbe probe;

    checkBrowserClassifier(test);
    checkPlaybackClassifier(test);
    checkAdvertisedMatrix(test);

    AppConfig config;
    config.startPath.clear();
    config.ffmpegAvailable = false;
    config.ffprobeAvailable = false;
    config.videoCombineAvailable = false;
    config.useTrash = false;
    MediaExplorerWindow window(config);
    checkEventDelivery(application, test, probe, window);

    if (test.failures == 0) {
        std::cout << "keyboard-acceptance: PASS\n";
        return 0;
    }
    std::cerr << "keyboard-acceptance: " << test.failures << " failure(s)\n";
    return 1;
}
