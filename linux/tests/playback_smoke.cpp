#include "../src/AppConfig.h"
#include "../src/MediaExplorerWindow.h"

#include <QApplication>
#include <QElapsedTimer>
#include <QEventLoop>
#include <QFile>
#include <QFileInfo>
#include <QMessageBox>
#include <QSize>
#include <QTemporaryDir>
#include <QTextStream>
#include <QThread>
#include <QTimer>

#include <vlc/vlc.h>

namespace {

struct Options {
    QString clipPath;
    int cycles = 1;
    int timeoutMs = 15000;
    bool help = false;
    QString error;
};

bool parsePositiveInteger(const QString &text, int maximum, int &value) {
    bool ok = false;
    const int parsed = text.toInt(&ok);
    if (!ok || parsed < 1 || parsed > maximum) return false;
    value = parsed;
    return true;
}

Options parseOptions(int argc, char **argv) {
    Options options;
    bool positionalClipSeen = false;

    const QString cyclesEnvironment = qEnvironmentVariable("MEDIA_EXPLORER_PLAYBACK_CYCLES");
    if (!cyclesEnvironment.isEmpty() &&
        !parsePositiveInteger(cyclesEnvironment, 1000, options.cycles)) {
        options.error = QStringLiteral(
            "MEDIA_EXPLORER_PLAYBACK_CYCLES must be an integer from 1 through 1000");
        return options;
    }
    const QString timeoutEnvironment = qEnvironmentVariable("MEDIA_EXPLORER_PLAYBACK_TIMEOUT_MS");
    if (!timeoutEnvironment.isEmpty() &&
        !parsePositiveInteger(timeoutEnvironment, 120000, options.timeoutMs)) {
        options.error = QStringLiteral(
            "MEDIA_EXPLORER_PLAYBACK_TIMEOUT_MS must be an integer from 1 through 120000");
        return options;
    }
    options.clipPath = qEnvironmentVariable("MEDIA_EXPLORER_TEST_CLIP");

    for (int index = 1; index < argc; ++index) {
        const QString argument = QString::fromLocal8Bit(argv[index]);
        if (argument == QStringLiteral("--help") || argument == QStringLiteral("-h")) {
            options.help = true;
        } else if (argument == QStringLiteral("--cycles")) {
            if (++index >= argc ||
                !parsePositiveInteger(QString::fromLocal8Bit(argv[index]), 1000, options.cycles)) {
                options.error = QStringLiteral("--cycles requires an integer from 1 through 1000");
                return options;
            }
        } else if (argument.startsWith(QStringLiteral("--cycles="))) {
            if (!parsePositiveInteger(argument.mid(QStringLiteral("--cycles=").size()),
                                      1000, options.cycles)) {
                options.error = QStringLiteral("--cycles requires an integer from 1 through 1000");
                return options;
            }
        } else if (argument == QStringLiteral("--timeout-ms")) {
            if (++index >= argc ||
                !parsePositiveInteger(QString::fromLocal8Bit(argv[index]),
                                      120000, options.timeoutMs)) {
                options.error = QStringLiteral(
                    "--timeout-ms requires an integer from 1 through 120000");
                return options;
            }
        } else if (argument.startsWith(QStringLiteral("--timeout-ms="))) {
            if (!parsePositiveInteger(argument.mid(QStringLiteral("--timeout-ms=").size()),
                                      120000, options.timeoutMs)) {
                options.error = QStringLiteral(
                    "--timeout-ms requires an integer from 1 through 120000");
                return options;
            }
        } else if (argument.startsWith(QLatin1Char('-'))) {
            options.error = QStringLiteral("Unknown option: %1").arg(argument);
            return options;
        } else if (positionalClipSeen) {
            options.error = QStringLiteral("Only one video clip may be specified");
            return options;
        } else {
            options.clipPath = argument;
            positionalClipSeen = true;
        }
    }
    return options;
}

void printUsage(QTextStream &stream) {
    stream << "Usage: playback-smoke [--cycles N] [--timeout-ms N] VIDEO\n"
              "\n"
              "Runs real embedded libVLC playback through MediaExplorerWindow.\n"
              "VIDEO may instead be supplied with MEDIA_EXPLORER_TEST_CLIP.\n"
              "The cycle count defaults to 1; use --cycles 50 for a soak run.\n";
}

QString stateName(int state) {
    switch (static_cast<libvlc_state_t>(state)) {
    case libvlc_NothingSpecial: return QStringLiteral("NothingSpecial");
    case libvlc_Opening: return QStringLiteral("Opening");
    case libvlc_Buffering: return QStringLiteral("Buffering");
    case libvlc_Playing: return QStringLiteral("Playing");
    case libvlc_Paused: return QStringLiteral("Paused");
    case libvlc_Stopped: return QStringLiteral("Stopped");
    case libvlc_Ended: return QStringLiteral("Ended");
    case libvlc_Error: return QStringLiteral("Error");
    }
    return QStringLiteral("Unavailable(%1)").arg(state);
}

struct PlaybackResult {
    bool playing = false;
    bool videoSizeKnown = false;
    int lastState = -1;
    QSize size;
};

PlaybackResult waitForPlayback(QApplication &application, MediaExplorerWindow &window,
                               int timeoutMs) {
    PlaybackResult result;
    QElapsedTimer elapsed;
    elapsed.start();
    while (elapsed.elapsed() < timeoutMs) {
        application.processEvents(QEventLoop::AllEvents, 50);
        result.lastState = window.playerStateForTest();
        if (result.lastState == static_cast<int>(libvlc_Playing)) result.playing = true;
        const QSize currentSize = window.videoSizeForTest();
        if (currentSize.width() > 0 && currentSize.height() > 0) {
            result.videoSizeKnown = true;
            result.size = currentSize;
        }
        if (result.playing && result.videoSizeKnown) break;
        if (result.lastState == static_cast<int>(libvlc_Error)) break;
        QThread::msleep(20);
    }
    return result;
}

void drainEvents(QApplication &application, int milliseconds) {
    QElapsedTimer elapsed;
    elapsed.start();
    while (elapsed.elapsed() < milliseconds) {
        application.processEvents(QEventLoop::AllEvents, 25);
        QThread::msleep(10);
    }
}

bool waitForStopped(QApplication &application, MediaExplorerWindow &window, int timeoutMs,
                    int &lastState) {
    QElapsedTimer elapsed;
    elapsed.start();
    do {
        application.processEvents(QEventLoop::AllEvents, 25);
        lastState = window.playerStateForTest();
        if (lastState == static_cast<int>(libvlc_Stopped)) return true;
        QThread::msleep(10);
    } while (elapsed.elapsed() < timeoutMs);
    return false;
}

} // namespace

int main(int argc, char **argv) {
    QTextStream standardOutput(stdout);
    QTextStream standardError(stderr);
    const Options options = parseOptions(argc, argv);
    if (!options.error.isEmpty()) {
        standardError << options.error << "\n\n";
        printUsage(standardError);
        return 2;
    }
    if (options.help) {
        printUsage(standardOutput);
        return 0;
    }
    if (options.clipPath.isEmpty()) {
        standardError << "A video clip is required.\n\n";
        printUsage(standardError);
        return 2;
    }
    if (qEnvironmentVariable("DISPLAY").isEmpty()) {
        standardError << "playback-smoke: DISPLAY is required for real xcb video embedding\n";
        return 2;
    }

    const QFileInfo clipInfo(options.clipPath);
    const QString clipPath = clipInfo.canonicalFilePath();
    if (clipPath.isEmpty() || !clipInfo.isFile() || !clipInfo.isReadable()) {
        standardError << "playback-smoke: clip is not a readable regular file: "
                      << options.clipPath << '\n';
        return 2;
    }

    QApplication application(argc, argv);
    application.setQuitOnLastWindowClosed(false);
    application.setApplicationName(QStringLiteral("Media Explorer Playback Smoke"));
    if (QApplication::platformName() != QStringLiteral("xcb")) {
        standardError << "playback-smoke: Qt platform xcb is required; got "
                      << QApplication::platformName() << '\n';
        return 2;
    }

    // Never inspect or resume a user's interrupted combine job during a test.
    QTemporaryDir isolatedState;
    if (!isolatedState.isValid()) {
        standardError << "playback-smoke: could not create isolated application state\n";
        return 2;
    }
    qputenv("XDG_STATE_HOME", QFile::encodeName(isolatedState.path()));

    AppConfig config;
    config.startPath.clear();
    config.ffprobeAvailable = false;
    config.ffmpegAvailable = false;
    config.videoCombineAvailable = false;
    config.useTrash = false;

    int failures = 0;
    {
        MediaExplorerWindow window(config);
        window.show();
        window.activateWindow();
        application.processEvents();

        QString unexpectedDialog;
        QTimer dialogGuard;
        dialogGuard.setInterval(50);
        QObject::connect(&dialogGuard, &QTimer::timeout, [&unexpectedDialog] {
            QWidget *modal = QApplication::activeModalWidget();
            if (!modal) return;
            unexpectedDialog = modal->windowTitle();
            if (auto *message = qobject_cast<QMessageBox *>(modal)) {
                const QString detail = message->text();
                if (!detail.isEmpty()) unexpectedDialog += QStringLiteral(": ") + detail;
            }
            modal->close();
        });
        dialogGuard.start();

        for (int cycle = 1; cycle <= options.cycles; ++cycle) {
            unexpectedDialog.clear();
            if (!window.startPlaybackForTest(QStringList{clipPath})) {
                standardError << "playback-smoke: cycle " << cycle
                              << " could not start the production playback path\n";
                ++failures;
                break;
            }

            const PlaybackResult result = waitForPlayback(application, window, options.timeoutMs);
            if (!result.playing || !result.videoSizeKnown || !unexpectedDialog.isEmpty()) {
                standardError << "playback-smoke: cycle " << cycle << " failed: state="
                              << stateName(result.lastState) << ", size="
                              << result.size.width() << 'x' << result.size.height();
                if (!unexpectedDialog.isEmpty())
                    standardError << ", dialog=" << unexpectedDialog;
                standardError << '\n';
                ++failures;
            } else {
                standardOutput << "playback-smoke: cycle " << cycle << '/' << options.cycles
                               << " Playing " << result.size.width() << 'x'
                               << result.size.height() << '\n';
                standardOutput.flush();
            }

            window.stopPlaybackForTest();
            int stoppedState = -1;
            if (!waitForStopped(application, window, 2000, stoppedState)) {
                standardError << "playback-smoke: cycle " << cycle
                              << " did not stop cleanly: state=" << stateName(stoppedState) << '\n';
                ++failures;
            }
            if (failures != 0) break;
        }

        dialogGuard.stop();
        window.stopPlaybackForTest();
        window.close();
        drainEvents(application, 100);
        if (window.isVisible()) {
            standardError << "playback-smoke: window did not close cleanly\n";
            ++failures;
        }
    }

    if (failures == 0) {
        standardOutput << "playback-smoke: PASS (" << options.cycles << " cycle"
                       << (options.cycles == 1 ? "" : "s") << ")\n";
        return 0;
    }
    standardError << "playback-smoke: FAIL\n";
    return 1;
}
