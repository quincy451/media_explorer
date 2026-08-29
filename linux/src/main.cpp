#include "AppConfig.h"
#include "MediaExplorerWindow.h"
#include "MediaTools.h"
#include "MountDiscovery.h"

#include <QApplication>
#include <QDir>
#include <QJsonArray>
#include <QJsonDocument>
#include <QJsonObject>
#include <QTextStream>

#include <vlc/vlc.h>

#include <iterator>
#include <utility>

namespace {

struct Options {
    QString configPath;
    bool selfTest = false;
    bool shortcutTest = false;
    bool help = false;
    bool version = false;
    QString error;
};

Options parseOptions(int argc, char **argv) {
    Options options;
    for (int index = 1; index < argc; ++index) {
        const QString argument = QString::fromLocal8Bit(argv[index]);
        if (argument == QStringLiteral("--self-test") || argument == QStringLiteral("--smoke-test"))
            options.selfTest = true;
        else if (argument == QStringLiteral("--self-test-shortcuts"))
            options.shortcutTest = true;
        else if (argument == QStringLiteral("--help") || argument == QStringLiteral("-h"))
            options.help = true;
        else if (argument == QStringLiteral("--version"))
            options.version = true;
        else if (argument == QStringLiteral("--config")) {
            if (++index >= argc) { options.error = QStringLiteral("--config requires a path"); break; }
            options.configPath = QString::fromLocal8Bit(argv[index]);
            if (options.configPath.isEmpty()) { options.error = QStringLiteral("--config requires a path"); break; }
        } else if (argument.startsWith(QStringLiteral("--config="))) {
            options.configPath = argument.mid(QStringLiteral("--config=").size());
            if (options.configPath.isEmpty()) options.error = QStringLiteral("--config requires a path");
        } else {
            options.error = QStringLiteral("Unknown option: %1").arg(argument);
            break;
        }
    }
    return options;
}

void printUsage(QTextStream &stream) {
    stream << "Media Explorer — native Linux Qt/libVLC media browser\n\n"
              "Usage: media-explorer [options]\n"
              "  --config PATH          Use this INI file (highest precedence)\n"
              "  --self-test            Run noninteractive full smoke test\n"
              "  --smoke-test           Alias for --self-test\n"
              "  --self-test-shortcuts  Validate the complete keyboard routing matrix\n"
              "  --version              Print version\n"
              "  --help                 Show this help\n";
}

struct TestReporter {
    void check(const QString &name, bool success, const QString &detail = {}) {
        QJsonObject item{{QStringLiteral("name"), name}, {QStringLiteral("ok"), success}};
        if (!detail.isEmpty()) item.insert(QStringLiteral("detail"), detail);
        checks.append(item);
        if (!success) ++failures;
    }
    QJsonArray checks;
    int failures = 0;
};

void testShortcuts(TestReporter &reporter) {
    using Command = MediaExplorerWindow::KeyCommand;
    struct Route { int key; Qt::KeyboardModifiers modifiers; Command command; };
    const auto ctrl = Qt::ControlModifier;
    const auto shift = Qt::ShiftModifier;
    const auto alt = Qt::AltModifier;
    const Route browser[] = {
        {Qt::Key_F1, {}, Command::Help}, {Qt::Key_Return, {}, Command::Activate},
        {Qt::Key_Enter, Qt::KeypadModifier, Command::Activate},
        {Qt::Key_Left, {}, Command::NavigateUp}, {Qt::Key_Backspace, {}, Command::NavigateUp},
        {Qt::Key_Left, alt, Command::NavigateUp}, {Qt::Key_Escape, {}, Command::CancelFileOperation},
        {Qt::Key_Delete, {}, Command::DeleteSelected}, {Qt::Key_F2, {}, Command::RenameSelected},
        {Qt::Key_F5, {}, Command::Refresh}, {Qt::Key_N, ctrl | shift, Command::NewFolder},
        {Qt::Key_A, ctrl, Command::SelectAllVideos}, {Qt::Key_P, ctrl, Command::PlaySelected},
        {Qt::Key_F, ctrl, Command::Search}, {Qt::Key_C, ctrl, Command::CopySelected},
        {Qt::Key_X, ctrl, Command::CutSelected}, {Qt::Key_V, ctrl, Command::Paste},
        {Qt::Key_U, ctrl, Command::SubmitTopaz}, {Qt::Key_3, ctrl, Command::SubmitIw3},
        {Qt::Key_Up, ctrl, Command::ReorderUp}, {Qt::Key_Down, ctrl, Command::ReorderDown},
        {Qt::Key_Plus, ctrl, Command::Combine}, {Qt::Key_Equal, ctrl | shift, Command::Combine},
        {Qt::Key_L, ctrl, Command::EditLocation}, {Qt::Key_Home, ctrl, Command::ShowRoots}
    };
    bool browserOk = true;
    for (const Route &route : browser)
        browserOk = browserOk &&
                    MediaExplorerWindow::browserCommandFor(route.key, route.modifiers) == route.command;
    browserOk = browserOk &&
                MediaExplorerWindow::browserCommandFor(Qt::Key_Up, {}) == Command::None &&
                MediaExplorerWindow::browserCommandFor(Qt::Key_Right, ctrl) == Command::None;
    reporter.check(QStringLiteral("browser-shortcut-matrix"), browserOk,
                   QStringLiteral("%1 routed commands").arg(std::size(browser)));

    const Route playback[] = {
        {Qt::Key_F1, {}, Command::Help}, {Qt::Key_Return, {}, Command::ToggleFullscreen},
        {Qt::Key_Space, {}, Command::Pause}, {Qt::Key_Tab, {}, Command::Resume},
        {Qt::Key_Escape, {}, Command::ExitPlayback}, {Qt::Key_G, ctrl, Command::Playlist},
        {Qt::Key_P, ctrl, Command::VideoProperties}, {Qt::Key_V, ctrl, Command::VideoTools},
        {Qt::Key_R, ctrl, Command::RenameDeferred}, {Qt::Key_C, ctrl, Command::CopyDeferred},
        {Qt::Key_L, ctrl, Command::ToggleLooping},
        {Qt::Key_Z, ctrl, Command::ZoomIn}, {Qt::Key_X, ctrl, Command::ZoomOut},
        {Qt::Key_Delete, {}, Command::RemoveDeferred}, {Qt::Key_Up, {}, Command::VolumeUp},
        {Qt::Key_Down, {}, Command::VolumeDown}, {Qt::Key_Left, {}, Command::SeekBack10},
        {Qt::Key_Right, {}, Command::SeekForward10}, {Qt::Key_Left, shift, Command::SeekBack60},
        {Qt::Key_Right, shift, Command::SeekForward60}, {Qt::Key_Left, ctrl, Command::Previous},
        {Qt::Key_Right, ctrl, Command::Next}, {Qt::Key_Left, alt | ctrl | shift, Command::PanLeft},
        {Qt::Key_Right, alt | ctrl | shift, Command::PanRight},
        {Qt::Key_Up, alt | ctrl, Command::PanUp}, {Qt::Key_Down, alt | ctrl, Command::PanDown},
        {Qt::Key_Home, alt | ctrl, Command::CenterPan}
    };
    bool playbackOk = true;
    for (const Route &route : playback)
        playbackOk = playbackOk &&
                     MediaExplorerWindow::playbackCommandFor(route.key, route.modifiers) == route.command;
    playbackOk = playbackOk &&
                 MediaExplorerWindow::playbackCommandFor(Qt::Key_A, {}) == Command::None &&
                 MediaExplorerWindow::playbackCommandFor(Qt::Key_3, ctrl) == Command::None;
    reporter.check(QStringLiteral("playback-shortcut-matrix"), playbackOk,
                   QStringLiteral("%1 routed commands").arg(std::size(playback)));
    reporter.check(QStringLiteral("advertised-shortcut-matrices"),
                   MediaExplorerWindow::browserShortcutMatrix().size() == 18 &&
                       MediaExplorerWindow::playbackShortcutMatrix().size() == 18);
}

int runSelfTest(QApplication &application, const AppConfig &config, bool shortcutsOnly) {
    TestReporter reporter;
    reporter.check(QStringLiteral("qapplication"), QApplication::instance() != nullptr,
                   QApplication::platformName());
    testShortcuts(reporter);
    if (!shortcutsOnly) {
        reporter.check(QStringLiteral("config-mapped-root"), QDir::isAbsolutePath(config.mappedRoot),
                       config.mappedRoot);
        reporter.check(QStringLiteral("video-extension-recognition"),
                       isVideoFile(QStringLiteral("sample.MP4"), config.videoExtensions) &&
                           !isVideoFile(QStringLiteral("sample.txt"), config.videoExtensions));
        const QVector<MountRecord> mountRecords = readMountRecords();
        bool rootMountFound = false;
        for (const MountRecord &record : mountRecords) {
            rootMountFound = rootMountFound || record.path == QStringLiteral("/");
        }
        reporter.check(QStringLiteral("mountinfo-read"), !mountRecords.isEmpty() && rootMountFound,
                       QStringLiteral("%1 mount records").arg(mountRecords.size()));
        const QVector<Entry> mounts = discoverMounts(config);
        bool mountsValid = !mounts.isEmpty();
        for (const Entry &entry : mounts) mountsValid = mountsValid && entry.directory && !entry.path.isEmpty();
        reporter.check(QStringLiteral("mount-discovery"), mountsValid,
                       QStringLiteral("%1 roots discovered").arg(mounts.size()));

        const QString ffmpeg = executableResolution(config.ffmpegPath);
        const QString ffprobe = executableResolution(config.ffprobePath);
        reporter.check(QStringLiteral("ffmpeg-lookup"), !config.ffmpegAvailable || !ffmpeg.isEmpty(),
                       config.ffmpegAvailable ? ffmpeg : QStringLiteral("disabled"));
        reporter.check(QStringLiteral("ffprobe-lookup"), !config.ffprobeAvailable || !ffprobe.isEmpty(),
                       config.ffprobeAvailable ? ffprobe : QStringLiteral("disabled"));
        const QStringList mediaToolFailures = MediaTools::selfTestFailures();
        reporter.check(QStringLiteral("media-tools"), mediaToolFailures.isEmpty(),
                       mediaToolFailures.join(QStringLiteral("; ")));

        QVector<QByteArray> storage;
        for (const QString &argument : config.vlcArgs) storage.append(argument.toUtf8());
        QVector<const char *> arguments;
        for (const QByteArray &argument : std::as_const(storage)) arguments.append(argument.constData());
        libvlc_instance_t *instance = libvlc_new(arguments.size(),
                                                  arguments.isEmpty() ? nullptr : arguments.constData());
        bool vlcOk = instance != nullptr;
        if (instance) {
            libvlc_media_player_t *player = libvlc_media_player_new(instance);
            vlcOk = player != nullptr;
            if (player) libvlc_media_player_release(player);
            libvlc_release(instance);
        }
        reporter.check(QStringLiteral("libvlc-init-release"), vlcOk,
                       QString::fromUtf8(libvlc_get_version()));
    }

    QJsonObject result{{QStringLiteral("application"), QStringLiteral("Media Explorer")},
                       {QStringLiteral("platform"), application.platformName()},
                       {QStringLiteral("ok"), reporter.failures == 0},
                       {QStringLiteral("failures"), reporter.failures},
                       {QStringLiteral("checks"), reporter.checks}};
    QTextStream(stdout) << QJsonDocument(result).toJson(QJsonDocument::Indented);
    return reporter.failures == 0 ? 0 : 1;
}

} // namespace

int main(int argc, char **argv) {
    const Options options = parseOptions(argc, argv);
    QTextStream standardOutput(stdout);
    QTextStream standardError(stderr);
    if (!options.error.isEmpty()) {
        standardError << options.error << "\n\n";
        printUsage(standardError);
        return 2;
    }
    if (options.help) { printUsage(standardOutput); return 0; }
    if (options.version) { standardOutput << "Media Explorer 2.0.0-linux-native\n"; return 0; }

    const bool testing = options.selfTest || options.shortcutTest;
    qputenv("QT_QPA_PLATFORM", testing ? QByteArray("offscreen") : QByteArray("xcb"));
    QApplication::setAttribute(Qt::AA_UseHighDpiPixmaps);
    QApplication application(argc, argv);
    application.setApplicationName(QStringLiteral("Media Explorer"));
    application.setApplicationVersion(QStringLiteral("2.0.0-linux-native"));
    application.setOrganizationName(QStringLiteral("Media Explorer"));

    AppConfig config;
    QString configError;
    if (!AppConfig::load(options.configPath.isEmpty() ? QString() : options.configPath,
                         config, configError)) {
        standardError << configError << '\n';
        return 2;
    }
    if (testing) return runSelfTest(application, config, options.shortcutTest && !options.selfTest);
    if (QApplication::platformName() != QStringLiteral("xcb")) {
        standardError << "Embedded VLC playback requires Qt platform xcb; got "
                      << QApplication::platformName() << '\n';
        return 3;
    }
    MediaExplorerWindow window(std::move(config));
    window.show();
    return application.exec();
}
