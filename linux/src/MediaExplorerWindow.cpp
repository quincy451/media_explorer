#include "MediaExplorerWindow.h"
#include "MediaTools.h"

#include <QAbstractItemView>
#include <QAbstractSpinBox>
#include <QAction>
#include <QApplication>
#include <QCloseEvent>
#include <QComboBox>
#include <QDateTime>
#include <QDesktopServices>
#include <QDialog>
#include <QDialogButtonBox>
#include <QDir>
#include <QEvent>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QFormLayout>
#include <QGuiApplication>
#include <QHBoxLayout>
#include <QHeaderView>
#include <QInputDialog>
#include <QItemSelectionModel>
#include <QJsonDocument>
#include <QJsonObject>
#include <QJsonArray>
#include <QKeyEvent>
#include <QLabel>
#include <QLineEdit>
#include <QListWidget>
#include <QLocale>
#include <QLockFile>
#include <QMainWindow>
#include <QMenu>
#include <QMenuBar>
#include <QMessageBox>
#include <QMetaObject>
#include <QMouseEvent>
#include <QPointer>
#include <QPlainTextEdit>
#include <QProcess>
#include <QProgressBar>
#include <QPushButton>
#include <QRunnable>
#include <QResizeEvent>
#include <QRegularExpression>
#include <QSaveFile>
#include <QShortcut>
#include <QSlider>
#include <QSplitter>
#include <QStackedWidget>
#include <QStandardPaths>
#include <QStatusBar>
#include <QTableWidget>
#include <QTemporaryFile>
#include <QTemporaryDir>
#include <QTextBrowser>
#include <QTextEdit>
#include <QThread>
#include <QTimer>
#include <QUrl>
#include <QUuid>
#include <QVBoxLayout>

#include <vlc/vlc.h>

#include <algorithm>
#include <cerrno>
#include <cmath>
#include <cstring>
#include <functional>
#include <memory>
#include <fcntl.h>
#include <sys/stat.h>
#include <sys/syscall.h>
#include <unistd.h>
#include <utility>

namespace {

constexpr const char *kApplicationName = "Media Explorer";
constexpr const char *kVersion = "2.0.0-linux-native";

class LambdaRunnable final : public QRunnable {
public:
    explicit LambdaRunnable(std::function<void()> callback) : callback_(std::move(callback)) {
        setAutoDelete(true);
    }
    void run() override { callback_(); }

private:
    std::function<void()> callback_;
};

bool pathLexists(const QString &path) {
    struct stat status {};
    return ::lstat(QFile::encodeName(path).constData(), &status) == 0;
}

int renameNoReplace(const QString &source, const QString &destination) {
#if defined(SYS_renameat2)
#ifndef RENAME_NOREPLACE
#define RENAME_NOREPLACE (1U << 0)
#endif
    return static_cast<int>(::syscall(SYS_renameat2, AT_FDCWD,
                                      QFile::encodeName(source).constData(), AT_FDCWD,
                                      QFile::encodeName(destination).constData(), RENAME_NOREPLACE));
#else
    Q_UNUSED(source)
    Q_UNUSED(destination)
    errno = ENOSYS;
    return -1;
#endif
}

QString availableDestination(const QString &destination) {
    if (!pathLexists(destination)) {
        return destination;
    }
    const QFileInfo info(destination);
    const QString parent = info.absolutePath();
    const QString suffix = info.completeSuffix().isEmpty() ? QString() : QLatin1Char('.') + info.completeSuffix();
    QString stem = info.fileName();
    if (!suffix.isEmpty()) {
        stem.chop(suffix.size());
    }
    QString candidate = QDir(parent).filePath(stem + QStringLiteral(" (copy)") + suffix);
    int counter = 2;
    while (pathLexists(candidate)) {
        candidate = QDir(parent).filePath(
            QStringLiteral("%1 (%2)%3").arg(stem).arg(counter++).arg(suffix));
    }
    return candidate;
}

bool isWithin(const QString &child, const QString &parent) {
    const QString childCanonical = QFileInfo(child).canonicalFilePath();
    const QString parentCanonical = QFileInfo(parent).canonicalFilePath();
    if (childCanonical.isEmpty() || parentCanonical.isEmpty()) {
        return false;
    }
    return childCanonical == parentCanonical || childCanonical.startsWith(parentCanonical + QLatin1Char('/'));
}

bool cleanupCreatedPath(const QString &path) {
    struct stat status {};
    if (::lstat(QFile::encodeName(path).constData(), &status) != 0) return errno == ENOENT;
    if (S_ISDIR(status.st_mode) && !S_ISLNK(status.st_mode)) {
        QDir directory(path);
        for (const QString &name : directory.entryList(QDir::AllEntries | QDir::Hidden | QDir::System |
                                                        QDir::NoDotAndDotDot)) {
            if (!cleanupCreatedPath(directory.filePath(name))) return false;
        }
        return QDir().rmdir(path);
    }
    return QFile::remove(path);
}

QString nestedMountPoint(const QString &root);

bool cleanupOwnedWorkDirectory(const QString &path, const QString &expectedParent) {
    if (path.isEmpty()) return true;
    const QFileInfo info(path);
    const QString name = info.fileName();
    const QString actualParent = QFileInfo(info.absolutePath()).canonicalFilePath();
    const QString requiredParent = QFileInfo(expectedParent).canonicalFilePath();
    if (actualParent.isEmpty() || requiredParent.isEmpty() || actualParent != requiredParent ||
        (!name.startsWith(QStringLiteral(".media-explorer-process-")) &&
         !name.startsWith(QStringLiteral(".media-explorer-combine-")))) return false;
    struct stat status {};
    if (::lstat(QFile::encodeName(info.absoluteFilePath()).constData(), &status) != 0)
        return errno == ENOENT;
    if (!S_ISDIR(status.st_mode) || S_ISLNK(status.st_mode) ||
        status.st_uid != ::geteuid()) return false;
    if (!nestedMountPoint(info.absoluteFilePath()).isEmpty()) return false;
    return cleanupCreatedPath(info.absoluteFilePath());
}

bool combineWorkspaceIsOwned(const CombineJobState::Job &job, QString &error) {
    struct stat status {};
    if (::lstat(QFile::encodeName(job.workingDirectory).constData(), &status) != 0 ||
        !S_ISDIR(status.st_mode) || S_ISLNK(status.st_mode) ||
        status.st_uid != ::geteuid()) {
        error = QStringLiteral("The combine working directory is missing, unsafe, or not owned by this user: %1")
                    .arg(job.workingDirectory);
        return false;
    }
    if (::lstat(QFile::encodeName(job.listPath).constData(), &status) == 0) {
        if (!S_ISREG(status.st_mode) || S_ISLNK(status.st_mode) ||
            status.st_uid != ::geteuid()) {
            error = QStringLiteral("The combine list is not an owned regular file: %1")
                        .arg(job.listPath);
            return false;
        }
    } else if (errno != ENOENT) {
        error = QStringLiteral("Could not inspect the combine list: %1")
                    .arg(QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    return true;
}

bool removeOwnedCombineEncodedFile(const CombineJobState::Job &job, QString &error) {
    struct stat status {};
    if (::lstat(QFile::encodeName(job.encodedPath).constData(), &status) != 0) {
        if (errno == ENOENT) return true;
        error = QStringLiteral("Could not inspect the prior combine output: %1")
                    .arg(QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    if (!S_ISREG(status.st_mode) || S_ISLNK(status.st_mode) ||
        status.st_uid != ::geteuid()) {
        error = QStringLiteral("Refusing to replace an unsafe combine output path: %1")
                    .arg(job.encodedPath);
        return false;
    }
    if (!QFile::remove(job.encodedPath)) {
        error = QStringLiteral("Could not remove the prior incomplete combine output: %1")
                    .arg(job.encodedPath);
        return false;
    }
    return true;
}

bool rebuildCombineList(const CombineJobState::Job &job, QString &error) {
    if (!combineWorkspaceIsOwned(job, error)) return false;
    const QString contents = MediaTools::ffconcatContent(job.sources, &error);
    if (contents.isEmpty()) return false;
    QSaveFile file(job.listPath);
    file.setDirectWriteFallback(false);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text) ||
        file.write(contents.toUtf8()) != contents.toUtf8().size() || !file.commit()) {
        error = QStringLiteral("Could not atomically rebuild the combine input list: %1")
                    .arg(file.errorString());
        return false;
    }
    if (!QFile::setPermissions(job.listPath, QFileDevice::ReadOwner | QFileDevice::WriteOwner)) {
        error = QStringLiteral("Could not apply owner-writable permissions to the combine input list: %1")
                    .arg(job.listPath);
        return false;
    }
    return true;
}

bool ownedRegularFileState(const QString &path, bool &exists, QString &error) {
    exists = false;
    struct stat status {};
    if (::lstat(QFile::encodeName(path).constData(), &status) != 0) {
        if (errno == ENOENT) return true;
        error = QStringLiteral("Could not inspect %1: %2")
                    .arg(path, QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    exists = true;
    if (!S_ISREG(status.st_mode) || S_ISLNK(status.st_mode) ||
        status.st_uid != ::geteuid() || status.st_size <= 0) {
        error = QStringLiteral("Recovery path is not a non-empty owned regular file: %1").arg(path);
        return false;
    }
    return true;
}

bool copyRegularFile(const QString &source, const QString &destination, mode_t mode,
                     const std::shared_ptr<std::atomic_bool> &cancel, QString &error) {
    QFile input(source);
    if (!input.open(QIODevice::ReadOnly)) {
        error = QStringLiteral("Could not open %1: %2").arg(source, input.errorString());
        return false;
    }
    int flags = O_WRONLY | O_CREAT | O_EXCL;
#ifdef O_CLOEXEC
    flags |= O_CLOEXEC;
#endif
    const int descriptor = ::open(QFile::encodeName(destination).constData(), flags, mode & 0777);
    if (descriptor < 0) {
        error = QStringLiteral("Could not create %1: %2")
                    .arg(destination, QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    QFile output;
    if (!output.open(descriptor, QIODevice::WriteOnly, QFileDevice::AutoCloseHandle)) {
        ::close(descriptor);
        error = QStringLiteral("Could not write %1: %2").arg(destination, output.errorString());
        QFile::remove(destination);
        return false;
    }
    QByteArray buffer(1024 * 1024, Qt::Uninitialized);
    bool success = true;
    while (!input.atEnd()) {
        if (cancel && cancel->load()) {
            error = QStringLiteral("Cancelled while copying %1").arg(source);
            success = false;
            break;
        }
        const qint64 count = input.read(buffer.data(), buffer.size());
        if (count < 0) {
            error = QStringLiteral("Could not read %1: %2").arg(source, input.errorString());
            success = false;
            break;
        }
        qint64 written = 0;
        while (written < count) {
            const qint64 amount = output.write(buffer.constData() + written, count - written);
            if (amount <= 0) {
                error = QStringLiteral("Could not write %1: %2").arg(destination, output.errorString());
                success = false;
                break;
            }
            written += amount;
        }
        if (!success) break;
    }
    if (success && !output.flush()) {
        error = QStringLiteral("Could not flush %1: %2").arg(destination, output.errorString());
        success = false;
    }
    if (success && ::fsync(output.handle()) != 0) {
        error = QStringLiteral("Could not sync %1: %2")
                    .arg(destination, QString::fromLocal8Bit(std::strerror(errno)));
        success = false;
    }
    output.close();
    input.close();
    if (!success && !QFile::remove(destination))
        error += QStringLiteral("; partial output retained at %1").arg(destination);
    return success;
}

bool copyPathRecursiveImpl(const QString &source, const QString &destination,
                           const std::shared_ptr<std::atomic_bool> &cancel, QString &error,
                           dev_t rootDevice, bool root) {
    if (cancel && cancel->load()) { error = QStringLiteral("Cancelled"); return false; }
    const QByteArray sourceName = QFile::encodeName(source);
    struct stat status {};
    if (::lstat(sourceName.constData(), &status) != 0) {
        error = QStringLiteral("%1: %2").arg(source, QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    if (!root && !S_ISLNK(status.st_mode) && status.st_dev != rootDevice) {
        error = QStringLiteral("Refusing to cross a nested mount while copying: %1").arg(source);
        return false;
    }
    if (S_ISLNK(status.st_mode)) {
        QByteArray target(4096, '\0');
        const ssize_t length = ::readlink(sourceName.constData(), target.data(), target.size() - 1);
        if (length < 0 || length >= target.size() - 1 ||
            ::symlink(QByteArray(target.constData(), static_cast<int>(length)).constData(),
                      QFile::encodeName(destination).constData()) != 0) {
            error = QStringLiteral("Could not copy symbolic link %1: %2")
                        .arg(source, QString::fromLocal8Bit(std::strerror(errno)));
            return false;
        }
        return true;
    }
    if (S_ISDIR(status.st_mode)) {
        if (!QDir().mkdir(destination)) {
            error = QStringLiteral("Could not create directory: %1").arg(destination);
            return false;
        }
        const QDir directory(source);
        for (const QString &name : directory.entryList(QDir::AllEntries | QDir::Hidden | QDir::System |
                                                        QDir::NoDotAndDotDot)) {
            if (!copyPathRecursiveImpl(directory.filePath(name), QDir(destination).filePath(name),
                                       cancel, error, rootDevice, false)) {
                if (!cleanupCreatedPath(destination))
                    error += QStringLiteral("; partial destination retained at %1").arg(destination);
                return false;
            }
        }
        QFile::setPermissions(destination, QFile::permissions(source));
        return true;
    }
    if (!S_ISREG(status.st_mode)) {
        error = QStringLiteral("Unsupported special file: %1").arg(source);
        return false;
    }
    return copyRegularFile(source, destination, status.st_mode, cancel, error);
}

QString nestedMountPoint(const QString &root);

bool copyPathRecursive(const QString &source, const QString &destination,
                       const std::shared_ptr<std::atomic_bool> &cancel, QString &error) {
    const QString nested = nestedMountPoint(source);
    if (!nested.isEmpty()) {
        error = QStringLiteral("Refusing to cross a nested mount while copying: %1").arg(nested);
        return false;
    }
    struct stat status {};
    if (::lstat(QFile::encodeName(source).constData(), &status) != 0) {
        error = QStringLiteral("%1: %2").arg(source, QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    return copyPathRecursiveImpl(source, destination, cancel, error, status.st_dev, true);
}

bool containsNestedMount(const QString &path, dev_t rootDevice,
                         const std::shared_ptr<std::atomic_bool> &cancel, QString &error) {
    if (cancel && cancel->load()) { error = QStringLiteral("Cancelled"); return true; }
    struct stat status {};
    if (::lstat(QFile::encodeName(path).constData(), &status) != 0) {
        error = QStringLiteral("%1: %2").arg(path, QString::fromLocal8Bit(std::strerror(errno)));
        return true;
    }
    if (!S_ISLNK(status.st_mode) && status.st_dev != rootDevice) {
        error = QStringLiteral("Refusing to cross a nested mount while deleting: %1").arg(path);
        return true;
    }
    if (!S_ISDIR(status.st_mode) || S_ISLNK(status.st_mode)) return false;
    QDir directory(path);
    for (const QString &name : directory.entryList(QDir::AllEntries | QDir::Hidden | QDir::System |
                                                    QDir::NoDotAndDotDot)) {
        if (containsNestedMount(directory.filePath(name), rootDevice, cancel, error)) return true;
    }
    return false;
}

QString nestedMountPoint(const QString &root) {
    struct stat status {};
    if (::lstat(QFile::encodeName(root).constData(), &status) != 0 ||
        !S_ISDIR(status.st_mode) || S_ISLNK(status.st_mode)) return {};
    const QFileInfo rootInfo(root);
    const QString canonical = rootInfo.canonicalFilePath();
    const QString normalizedRoot = QDir::cleanPath(
        canonical.isEmpty() ? rootInfo.absoluteFilePath() : canonical);
    const QString prefix = normalizedRoot == QStringLiteral("/")
                               ? QStringLiteral("/") : normalizedRoot + QLatin1Char('/');
    for (const MountRecord &record : readMountRecords()) {
        const QString mount = QDir::cleanPath(record.path);
        if (mount != normalizedRoot && mount.startsWith(prefix))
            return mount;
    }
    return {};
}

bool removePathRecursive(const QString &path, const std::shared_ptr<std::atomic_bool> &cancel,
                         QString &error) {
    if (cancel && cancel->load()) {
        return false;
    }
    struct stat rootStatus {};
    if (::lstat(QFile::encodeName(path).constData(), &rootStatus) != 0) {
        error = QStringLiteral("%1: %2").arg(path, QString::fromLocal8Bit(std::strerror(errno)));
        return false;
    }
    const QString nested = nestedMountPoint(path);
    if (!nested.isEmpty()) {
        error = QStringLiteral("Refusing to cross a nested mount while deleting: %1").arg(nested);
        return false;
    }
    if (S_ISDIR(rootStatus.st_mode) && !S_ISLNK(rootStatus.st_mode) &&
        containsNestedMount(path, rootStatus.st_dev, cancel, error)) return false;
    std::function<bool(const QString &)> removeImpl = [&](const QString &current) {
        if (cancel && cancel->load()) { error = QStringLiteral("Cancelled"); return false; }
        struct stat status {};
        if (::lstat(QFile::encodeName(current).constData(), &status) != 0) {
            error = QStringLiteral("%1: %2").arg(current, QString::fromLocal8Bit(std::strerror(errno)));
            return false;
        }
        if (S_ISDIR(status.st_mode) && !S_ISLNK(status.st_mode)) {
            const QDir directory(current);
            for (const QString &name : directory.entryList(QDir::AllEntries | QDir::Hidden | QDir::System |
                                                            QDir::NoDotAndDotDot)) {
                if (!removeImpl(directory.filePath(name))) return false;
            }
            if (!QDir().rmdir(current)) {
                error = QStringLiteral("Could not remove directory: %1").arg(current);
                return false;
            }
            return true;
        }
        if (!QFile::remove(current)) {
            error = QStringLiteral("Could not remove: %1").arg(current);
            return false;
        }
        return true;
    };
    return removeImpl(path);
}

QString formatModified(qint64 milliseconds) {
    return milliseconds > 0
               ? QDateTime::fromMSecsSinceEpoch(milliseconds).toString(QStringLiteral("yyyy-MM-dd HH:mm"))
               : QString();
}

bool validBaseName(const QString &name) {
    return !name.isEmpty() && name != QStringLiteral(".") && name != QStringLiteral("..") &&
           !name.contains(QLatin1Char('/')) && !name.contains(QChar(0));
}

bool skipSearchPath(const QString &path, const QStringList &explicitScopes) {
    const QString normalized = QDir::cleanPath(path);
    for (const QString &prefix : {QStringLiteral("/proc"), QStringLiteral("/sys"), QStringLiteral("/dev")}) {
        if (normalized == prefix || normalized.startsWith(prefix + QLatin1Char('/'))) {
            return true;
        }
    }
    if (normalized == QStringLiteral("/run/user") ||
        normalized.startsWith(QStringLiteral("/run/user/"))) {
        const bool gvfs = normalized.contains(QStringLiteral("/gvfs/")) || normalized.endsWith(QStringLiteral("/gvfs"));
        if (gvfs) {
            for (const QString &scope : explicitScopes) {
                const QString normalizedScope = QDir::cleanPath(scope);
                if (normalized == normalizedScope || normalized.startsWith(normalizedScope + QLatin1Char('/'))) {
                    return false;
                }
            }
        }
        return true;
    }
    return normalized == QStringLiteral("/run/lock") || normalized.startsWith(QStringLiteral("/run/lock/"));
}

} // namespace

CallbackWidget::CallbackWidget(QWidget *parent) : QWidget(parent) {
    setAttribute(Qt::WA_NativeWindow);
    setAutoFillBackground(true);
    QPalette colors = palette();
    colors.setColor(QPalette::Window, Qt::black);
    setPalette(colors);
}

void CallbackWidget::mouseDoubleClickEvent(QMouseEvent *event) {
    if (doubleClickCallback) {
        doubleClickCallback();
    }
    event->accept();
}

void CallbackWidget::resizeEvent(QResizeEvent *event) {
    QWidget::resizeEvent(event);
    if (resizeCallback) {
        resizeCallback();
    }
}

MediaExplorerWindow::MediaExplorerWindow(AppConfig config, QWidget *parent)
    : QMainWindow(parent), config_(std::move(config)) {
    workerPool_.setMaxThreadCount(qBound(4, QThread::idealThreadCount(), 8));
    metadataPool_.setMaxThreadCount(4);
    mappingPool_.setMaxThreadCount(2);
    setWindowTitle(QString::fromLatin1(kApplicationName));
    resize(1220, 760);
    setMinimumSize(760, 480);
    buildUi();
    buildMenus();
    qApp->installEventFilter(this);

    playerTimer_ = new QTimer(this);
    playerTimer_->setInterval(250);
    connect(playerTimer_, &QTimer::timeout, this, [this] { pollPlayer(); });

    if (!config_.startPath.isEmpty()) navigate(config_.startPath);
    else showRoots();
    if (QGuiApplication::platformName() == QStringLiteral("xcb"))
        QTimer::singleShot(0, this, [this] { loadPendingCombineJobs(); });
}

MediaExplorerWindow::~MediaExplorerWindow() {
    qApp->removeEventFilter(this);
    if (scanCancel_) scanCancel_->store(true);
    if (searchCancel_) searchCancel_->store(true);
    if (fileOperationCancel_) fileOperationCancel_->store(true);
    if (mediaProcessCancel_) mediaProcessCancel_->store(true);
    if (vlcPlayer_) {
        libvlc_media_player_stop(vlcPlayer_);
        libvlc_media_player_release(vlcPlayer_);
    }
    if (currentMedia_) libvlc_media_release(currentMedia_);
    if (vlcInstance_) libvlc_release(vlcInstance_);
}

void MediaExplorerWindow::buildUi() {
    stack_ = new QStackedWidget(this);
    setCentralWidget(stack_);

    browserPage_ = new QWidget(stack_);
    auto *browserLayout = new QVBoxLayout(browserPage_);
    browserLayout->setContentsMargins(6, 6, 6, 6);
    auto *locationLayout = new QHBoxLayout;
    auto *up = new QPushButton(QStringLiteral("Up"), browserPage_);
    auto *mounts = new QPushButton(QStringLiteral("Mounts"), browserPage_);
    location_ = new QLineEdit(browserPage_);
    location_->setObjectName(QStringLiteral("location"));
    location_->setClearButtonEnabled(true);
    auto *go = new QPushButton(QStringLiteral("Go"), browserPage_);
    connect(up, &QPushButton::clicked, this, [this] { goUp(); });
    connect(mounts, &QPushButton::clicked, this, [this] { showRoots(); });
    connect(go, &QPushButton::clicked, this, [this] { goToLocation(); });
    connect(location_, &QLineEdit::returnPressed, this, [this] { goToLocation(); });
    locationLayout->addWidget(up);
    locationLayout->addWidget(mounts);
    locationLayout->addWidget(location_, 1);
    locationLayout->addWidget(go);
    browserLayout->addLayout(locationLayout);

    table_ = new QTableWidget(0, 6, browserPage_);
    table_->setObjectName(QStringLiteral("fileTable"));
    table_->setHorizontalHeaderLabels({QStringLiteral("Name"), QStringLiteral("Type"),
                                       QStringLiteral("Size"), QStringLiteral("Modified"),
                                       QStringLiteral("Resolution"), QStringLiteral("Duration")});
    table_->setSelectionBehavior(QAbstractItemView::SelectRows);
    table_->setSelectionMode(QAbstractItemView::ExtendedSelection);
    table_->setEditTriggers(QAbstractItemView::NoEditTriggers);
    table_->setAlternatingRowColors(true);
    table_->setContextMenuPolicy(Qt::CustomContextMenu);
    table_->horizontalHeader()->setSectionResizeMode(0, QHeaderView::Stretch);
    for (int column = 1; column < 6; ++column) {
        table_->horizontalHeader()->setSectionResizeMode(column, QHeaderView::Interactive);
    }
    table_->setColumnWidth(1, 145);
    table_->setColumnWidth(2, 105);
    table_->setColumnWidth(3, 145);
    table_->setColumnWidth(4, 115);
    table_->setColumnWidth(5, 100);
    connect(table_, &QTableWidget::cellDoubleClicked, this,
            [this](int row, int) { activateRow(row); });
    connect(table_->horizontalHeader(), &QHeaderView::sectionClicked, this, [this](int column) {
        if (column == sortColumn_) sortAscending_ = !sortAscending_;
        else { sortColumn_ = column; sortAscending_ = true; }
        sortAndFill(true);
    });
    connect(table_, &QTableWidget::customContextMenuRequested, this, [this](const QPoint &position) {
        QMenu menu(this);
        menu.addAction(QStringLiteral("Play selected"), this, [this] { playSelected(); });
        menu.addSeparator();
        menu.addAction(QStringLiteral("Copy"), this, [this] { copySelected(); });
        menu.addAction(QStringLiteral("Cut"), this, [this] { cutSelected(); });
        menu.addAction(QStringLiteral("Paste"), this, [this] { paste(); });
        menu.addAction(QStringLiteral("Rename"), this, [this] { renameSelected(); });
        menu.addAction(QStringLiteral("Delete"), this, [this] { deleteSelected(); });
        menu.addSeparator();
        menu.addAction(QStringLiteral("New folder"), this, [this] { newFolder(); });
        menu.addAction(QStringLiteral("Open in system file manager"), this, [this] { openExternal(); });
        menu.exec(table_->viewport()->mapToGlobal(position));
    });
    browserLayout->addWidget(table_, 1);
    stack_->addWidget(browserPage_);

    playerPage_ = new QWidget(stack_);
    auto *playerLayout = new QVBoxLayout(playerPage_);
    playerLayout->setContentsMargins(5, 5, 5, 5);
    playerTitle_ = new QLabel(playerPage_);
    playerTitle_->setObjectName(QStringLiteral("playerTitle"));
    playerTitle_->setTextInteractionFlags(Qt::TextSelectableByMouse);
    playerLayout->addWidget(playerTitle_);
    auto *splitter = new QSplitter(Qt::Horizontal, playerPage_);
    videoViewport_ = new CallbackWidget(splitter);
    videoFrame_ = new CallbackWidget(videoViewport_);
    videoFrame_->setObjectName(QStringLiteral("videoFrame"));
    videoFrame_->setFocusPolicy(Qt::StrongFocus);
    videoViewport_->resizeCallback = [this] { updateVideoGeometry(); };
    videoViewport_->doubleClickCallback = [this] { toggleFullscreen(); };
    videoFrame_->doubleClickCallback = [this] { toggleFullscreen(); };
    splitter->addWidget(videoViewport_);
    playlistWidget_ = new QListWidget(splitter);
    playlistWidget_->setMinimumWidth(220);
    playlistWidget_->setMaximumWidth(420);
    connect(playlistWidget_, &QListWidget::itemDoubleClicked, this,
            [this](QListWidgetItem *) { playIndex(playlistWidget_->currentRow()); });
    splitter->addWidget(playlistWidget_);
    splitter->setStretchFactor(0, 1);
    playerLayout->addWidget(splitter, 1);

    seekSlider_ = new QSlider(Qt::Horizontal, playerPage_);
    seekSlider_->setObjectName(QStringLiteral("seekSlider"));
    seekSlider_->setRange(0, 1000);
    connect(seekSlider_, &QSlider::sliderPressed, this, [this] { seeking_ = true; });
    connect(seekSlider_, &QSlider::sliderReleased, this, [this] {
        if (vlcPlayer_) libvlc_media_player_set_position(vlcPlayer_, seekSlider_->value() / 1000.0f);
        seeking_ = false;
    });
    playerLayout->addWidget(seekSlider_);

    playerControls_ = new QWidget(playerPage_);
    auto *controls = new QHBoxLayout(playerControls_);
    controls->setContentsMargins(0, 0, 0, 0);
    auto *back = new QPushButton(QStringLiteral("Back"), playerControls_);
    auto *previous = new QPushButton(QStringLiteral("Previous"), playerControls_);
    pauseButton_ = new QPushButton(QStringLiteral("Pause"), playerControls_);
    auto *next = new QPushButton(QStringLiteral("Next"), playerControls_);
    timeLabel_ = new QLabel(QStringLiteral("0:00 / 0:00"), playerControls_);
    volumeSlider_ = new QSlider(Qt::Horizontal, playerControls_);
    volumeSlider_->setRange(0, 200);
    volumeSlider_->setValue(100);
    volumeSlider_->setMaximumWidth(150);
    auto *fullscreen = new QPushButton(QStringLiteral("Fullscreen"), playerControls_);
    connect(back, &QPushButton::clicked, this, [this] { stopPlayback(); });
    connect(previous, &QPushButton::clicked, this, [this] { previousVideo(); });
    connect(pauseButton_, &QPushButton::clicked, this, [this] {
        if (vlcPlayer_ && libvlc_media_player_is_playing(vlcPlayer_)) pausePlayback(); else resumePlayback();
    });
    connect(next, &QPushButton::clicked, this, [this] { nextVideo(); });
    connect(volumeSlider_, &QSlider::valueChanged, this, [this](int value) { setVolume(value); });
    connect(fullscreen, &QPushButton::clicked, this, [this] { toggleFullscreen(); });
    const QList<QWidget *> transportButtons{back, previous, pauseButton_, next};
    for (QWidget *widget : transportButtons) controls->addWidget(widget);
    controls->addWidget(timeLabel_);
    controls->addStretch(1);
    controls->addWidget(new QLabel(QStringLiteral("Volume"), playerControls_));
    controls->addWidget(volumeSlider_);
    controls->addWidget(fullscreen);
    playerLayout->addWidget(playerControls_);
    stack_->addWidget(playerPage_);

    progress_ = new QProgressBar(this);
    progress_->setMaximumWidth(240);
    progress_->hide();
    cancelButton_ = new QPushButton(QStringLiteral("Cancel"), this);
    cancelButton_->setMaximumWidth(90);
    cancelButton_->hide();
    connect(cancelButton_, &QPushButton::clicked, this, [this] { cancelBusy(); });
    statusBar()->addPermanentWidget(progress_);
    statusBar()->addPermanentWidget(cancelButton_);
}

QAction *MediaExplorerWindow::addAction(QMenu *menu, const QString &text,
                                        const std::function<void()> &callback,
                                        const QKeySequence &shortcut) {
    // Keyboard commands are owned by eventFilter so mode-dependent duplicates
    // (notably Ctrl+V and Ctrl+P) cannot be triggered by hidden QAction state.
    QString label = text;
    if (!shortcut.isEmpty()) {
        label += QLatin1Char('\t') + shortcut.toString(QKeySequence::NativeText);
    }
    auto *action = menu->addAction(label);
    connect(action, &QAction::triggered, this, [callback](bool) { callback(); });
    return action;
}

void MediaExplorerWindow::buildMenus() {
    QMenu *file = menuBar()->addMenu(QStringLiteral("&File"));
    addAction(file, QStringLiteral("&Play selected"), [this] { playSelected(); }, QKeySequence(QStringLiteral("Ctrl+P")));
    addAction(file, QStringLiteral("&Copy"), [this] { copySelected(); }, QKeySequence::Copy);
    addAction(file, QStringLiteral("Cu&t"), [this] { cutSelected(); }, QKeySequence::Cut);
    addAction(file, QStringLiteral("&Paste"), [this] { paste(); }, QKeySequence::Paste);
    addAction(file, QStringLiteral("&Rename"), [this] { renameSelected(); }, QKeySequence(QStringLiteral("F2")));
    addAction(file, QStringLiteral("&Delete"), [this] { deleteSelected(); }, QKeySequence::Delete);
    addAction(file, QStringLiteral("New &folder"), [this] { newFolder(); }, QKeySequence(QStringLiteral("Ctrl+Shift+N")));
    addAction(file, QStringLiteral("Open in system file manager"), [this] { openExternal(); });
    file->addSeparator();
    addAction(file, QStringLiteral("E&xit"), [this] { close(); }, QKeySequence::Quit);

    QMenu *navigateMenu = menuBar()->addMenu(QStringLiteral("&Navigate"));
    addAction(navigateMenu, QStringLiteral("&Up"), [this] { goUp(); }, QKeySequence(QStringLiteral("Backspace")));
    addAction(navigateMenu, QStringLiteral("&Mounts"), [this] { showRoots(); }, QKeySequence(QStringLiteral("Ctrl+Home")));
    addAction(navigateMenu, QStringLiteral("&Location"), [this] {
        if (location_->isReadOnly()) { location_->setReadOnly(false); location_->clear(); }
        location_->setFocus();
        location_->selectAll();
    }, QKeySequence(QStringLiteral("Ctrl+L")));
    addAction(navigateMenu, QStringLiteral("&Refresh"), [this] { refresh(); }, QKeySequence::Refresh);
    addAction(navigateMenu, QStringLiteral("&Search / refine"), [this] { promptSearch(); }, QKeySequence::Find);

    QMenu *tools = menuBar()->addMenu(QStringLiteral("&Tools"));
    addAction(tools, QStringLiteral("Combine selected videos"), [this] { combineSelected(); },
              QKeySequence(QStringLiteral("Ctrl++")));
    addAction(tools, QStringLiteral("Submit selected to Topaz queue"), [this] { submitTopazQueue(); },
              QKeySequence(QStringLiteral("Ctrl+U")));
    addAction(tools, QStringLiteral("Submit selected to IW3 queue"), [this] { submitIw3Queue(); },
              QKeySequence(QStringLiteral("Ctrl+3")));
    tools->addSeparator();
    addAction(tools, QStringLiteral("Video tools"), [this] { showVideoTools(); },
              QKeySequence(QStringLiteral("Ctrl+V")));

    QMenu *playback = menuBar()->addMenu(QStringLiteral("&Playback"));
    addAction(playback, QStringLiteral("Pause"), [this] { pausePlayback(); });
    addAction(playback, QStringLiteral("Resume"), [this] { resumePlayback(); });
    addAction(playback, QStringLiteral("Previous"), [this] { previousVideo(); }, QKeySequence(QStringLiteral("PgUp")));
    addAction(playback, QStringLiteral("Next"), [this] { nextVideo(); }, QKeySequence(QStringLiteral("PgDown")));
    addAction(playback, QStringLiteral("Loop current video"), [this] { toggleLooping(); },
              QKeySequence(QStringLiteral("Ctrl+L")));
    addAction(playback, QStringLiteral("Fullscreen"), [this] { toggleFullscreen(); }, QKeySequence(QStringLiteral("F11")));
    addAction(playback, QStringLiteral("Video properties"), [this] { showVideoProperties(); });
    addAction(playback, QStringLiteral("Stop / return to browser"), [this] { stopPlayback(); });

    QMenu *help = menuBar()->addMenu(QStringLiteral("&Help"));
    addAction(help, QStringLiteral("Media Explorer &Help"), [this] { showHelp(); }, QKeySequence::HelpContents);
    addAction(help, QStringLiteral("&About"), [this] { showAbout(); });
}

QStringList MediaExplorerWindow::browserShortcutMatrix() {
    return {QStringLiteral("Enter=open/play"), QStringLiteral("Left|Backspace|Alt+Left=up"),
            QStringLiteral("Ctrl+A=select-videos"), QStringLiteral("Ctrl+P=play"),
            QStringLiteral("Ctrl+F=search/refine"), QStringLiteral("Ctrl+Up|Ctrl+Down=reorder"),
            QStringLiteral("Ctrl++=combine"), QStringLiteral("Ctrl+U=topaz"),
            QStringLiteral("Ctrl+3=iw3"), QStringLiteral("Ctrl+C|Ctrl+X|Ctrl+V=clipboard"),
            QStringLiteral("Delete=delete"), QStringLiteral("Escape=cancel"),
            QStringLiteral("F2=rename"), QStringLiteral("Ctrl+Shift+N=new-folder"),
            QStringLiteral("F5=refresh"), QStringLiteral("Ctrl+L=location"),
            QStringLiteral("Ctrl+Home=mounts"), QStringLiteral("F1=help")};
}

QStringList MediaExplorerWindow::playbackShortcutMatrix() {
    return {QStringLiteral("F1=help"), QStringLiteral("Enter=fullscreen"),
            QStringLiteral("Space=pause"), QStringLiteral("Tab=resume"),
            QStringLiteral("Escape=exit"), QStringLiteral("Ctrl+G=playlist"),
            QStringLiteral("Ctrl+P=properties"), QStringLiteral("Ctrl+V=tools"),
            QStringLiteral("Ctrl+R=rename-on-exit"), QStringLiteral("Ctrl+C=copy-on-exit"),
            QStringLiteral("Ctrl+L=loop-current"), QStringLiteral("Ctrl+Z|Ctrl+X=zoom"),
            QStringLiteral("Delete=delete-on-exit"),
            QStringLiteral("Alt+Arrows|Alt+Home=pan"), QStringLiteral("Up|Down=volume"),
            QStringLiteral("Left|Right=seek"), QStringLiteral("Shift+Left|Shift+Right=seek-60"),
            QStringLiteral("Ctrl+Left|Ctrl+Right=previous/next")};
}

bool MediaExplorerWindow::editableOrModalContext(QObject *watched) const {
    if (QApplication::activeModalWidget() || QApplication::activePopupWidget() ||
        qobject_cast<QMenu *>(watched)) return true;
    QWidget *focus = QApplication::focusWidget();
    if (!focus) focus = qobject_cast<QWidget *>(watched);
    while (focus) {
        if (qobject_cast<QLineEdit *>(focus) || qobject_cast<QTextEdit *>(focus) ||
            qobject_cast<QPlainTextEdit *>(focus) || qobject_cast<QAbstractSpinBox *>(focus)) return true;
        if (auto *combo = qobject_cast<QComboBox *>(focus); combo && combo->isEditable()) return true;
        focus = focus->parentWidget();
    }
    return false;
}

bool MediaExplorerWindow::eventFilter(QObject *watched, QEvent *event) {
    if (event->type() == QEvent::Show &&
        (qobject_cast<QDialog *>(watched) || qobject_cast<QMenu *>(watched)) &&
        !playbackDialogs_.contains(watched) && beginPlaybackPauseHold()) {
        playbackDialogs_.insert(watched);
    } else if ((event->type() == QEvent::Hide || event->type() == QEvent::Destroy) &&
               playbackDialogs_.remove(watched) > 0) {
        endPlaybackPauseHold();
    }
    if (event->type() != QEvent::KeyPress || editableOrModalContext(watched)) {
        return QMainWindow::eventFilter(watched, event);
    }
    auto *key = static_cast<QKeyEvent *>(event);
    if (stack_->currentWidget() == playerPage_) {
        if (dispatchPlaybackKey(key)) return true;
    } else if (dispatchBrowserKey(key)) {
        return true;
    }
    return QMainWindow::eventFilter(watched, event);
}

MediaExplorerWindow::KeyCommand MediaExplorerWindow::browserCommandFor(
    int key, Qt::KeyboardModifiers modifiers) {
    const bool control = modifiers.testFlag(Qt::ControlModifier);
    const bool shift = modifiers.testFlag(Qt::ShiftModifier);
    if (key == Qt::Key_F1) return KeyCommand::Help;
    if (key == Qt::Key_Return || key == Qt::Key_Enter) return KeyCommand::Activate;
    if (key == Qt::Key_Backspace || key == Qt::Key_Left) return KeyCommand::NavigateUp;
    if (key == Qt::Key_Escape) return KeyCommand::CancelFileOperation;
    if (key == Qt::Key_Delete) return KeyCommand::DeleteSelected;
    if (key == Qt::Key_F2) return KeyCommand::RenameSelected;
    if (key == Qt::Key_F5) return KeyCommand::Refresh;
    if (control && shift && key == Qt::Key_N) return KeyCommand::NewFolder;
    if (!control) return KeyCommand::None;
    switch (key) {
    case Qt::Key_A: return KeyCommand::SelectAllVideos;
    case Qt::Key_P: return KeyCommand::PlaySelected;
    case Qt::Key_F: return KeyCommand::Search;
    case Qt::Key_C: return KeyCommand::CopySelected;
    case Qt::Key_X: return KeyCommand::CutSelected;
    case Qt::Key_V: return KeyCommand::Paste;
    case Qt::Key_U: return KeyCommand::SubmitTopaz;
    case Qt::Key_3: return KeyCommand::SubmitIw3;
    case Qt::Key_Up: return KeyCommand::ReorderUp;
    case Qt::Key_Down: return KeyCommand::ReorderDown;
    case Qt::Key_Plus:
    case Qt::Key_Equal: return KeyCommand::Combine;
    case Qt::Key_L: return KeyCommand::EditLocation;
    case Qt::Key_Home: return KeyCommand::ShowRoots;
    default: return KeyCommand::None;
    }
}

MediaExplorerWindow::KeyCommand MediaExplorerWindow::playbackCommandFor(
    int key, Qt::KeyboardModifiers modifiers) {
    const bool control = modifiers.testFlag(Qt::ControlModifier);
    const bool shift = modifiers.testFlag(Qt::ShiftModifier);
    const bool alt = modifiers.testFlag(Qt::AltModifier);
    if (key == Qt::Key_F1) return KeyCommand::Help;
    if (key == Qt::Key_Return || key == Qt::Key_Enter) return KeyCommand::ToggleFullscreen;
    if (key == Qt::Key_Space) return KeyCommand::Pause;
    if (key == Qt::Key_Tab) return KeyCommand::Resume;
    if (key == Qt::Key_Escape) return KeyCommand::ExitPlayback;
    if (alt) {
        if (key == Qt::Key_Left) return KeyCommand::PanLeft;
        if (key == Qt::Key_Right) return KeyCommand::PanRight;
        if (key == Qt::Key_Up) return KeyCommand::PanUp;
        if (key == Qt::Key_Down) return KeyCommand::PanDown;
        if (key == Qt::Key_Home) return KeyCommand::CenterPan;
    }
    if (control) {
        if (key == Qt::Key_G) return KeyCommand::Playlist;
        if (key == Qt::Key_P) return KeyCommand::VideoProperties;
        if (key == Qt::Key_V) return KeyCommand::VideoTools;
        if (key == Qt::Key_R) return KeyCommand::RenameDeferred;
        if (key == Qt::Key_C) return KeyCommand::CopyDeferred;
        if (key == Qt::Key_L) return KeyCommand::ToggleLooping;
        if (key == Qt::Key_Z) return KeyCommand::ZoomIn;
        if (key == Qt::Key_X) return KeyCommand::ZoomOut;
        if (key == Qt::Key_Left) return KeyCommand::Previous;
        if (key == Qt::Key_Right) return KeyCommand::Next;
    }
    if (key == Qt::Key_Delete) return KeyCommand::RemoveDeferred;
    if (key == Qt::Key_Up) return KeyCommand::VolumeUp;
    if (key == Qt::Key_Down) return KeyCommand::VolumeDown;
    if (key == Qt::Key_Left) return shift ? KeyCommand::SeekBack60 : KeyCommand::SeekBack10;
    if (key == Qt::Key_Right) return shift ? KeyCommand::SeekForward60 : KeyCommand::SeekForward10;
    return KeyCommand::None;
}

bool MediaExplorerWindow::dispatchBrowserKey(QKeyEvent *event) {
    switch (browserCommandFor(event->key(), event->modifiers())) {
    case KeyCommand::Help: showHelp(); return true;
    case KeyCommand::Activate: activateSelection(); return true;
    case KeyCommand::NavigateUp: goUp(); return true;
    case KeyCommand::CancelFileOperation:
        if (fileOperationCancel_) {
            fileOperationCancel_->store(true);
            statusBar()->showMessage(QStringLiteral("Cancelling file operation..."));
        }
        return true;
    case KeyCommand::DeleteSelected: deleteSelected(); return true;
    case KeyCommand::RenameSelected: renameSelected(); return true;
    case KeyCommand::Refresh: refresh(); return true;
    case KeyCommand::NewFolder: newFolder(); return true;
    case KeyCommand::SelectAllVideos: selectAllVideos(); return true;
    case KeyCommand::PlaySelected: playSelected(); return true;
    case KeyCommand::Search: promptSearch(); return true;
    case KeyCommand::CopySelected: copySelected(); return true;
    case KeyCommand::CutSelected: cutSelected(); return true;
    case KeyCommand::Paste: paste(); return true;
    case KeyCommand::SubmitTopaz: submitTopazQueue(); return true;
    case KeyCommand::SubmitIw3: submitIw3Queue(); return true;
    case KeyCommand::ReorderUp: moveSelectedRow(-1); return true;
    case KeyCommand::ReorderDown: moveSelectedRow(1); return true;
    case KeyCommand::Combine: combineSelected(); return true;
    case KeyCommand::EditLocation:
        if (location_->isReadOnly()) { location_->setReadOnly(false); location_->clear(); }
        location_->setFocus(); location_->selectAll(); return true;
    case KeyCommand::ShowRoots: showRoots(); return true;
    default: return false;
    }
}

bool MediaExplorerWindow::dispatchPlaybackKey(QKeyEvent *event) {
    if (exitingPlayback_) return true;
    switch (playbackCommandFor(event->key(), event->modifiers())) {
    case KeyCommand::Help: showHelp(); return true;
    case KeyCommand::ToggleFullscreen: toggleFullscreen(); return true;
    case KeyCommand::Pause: pausePlayback(); return true;
    case KeyCommand::Resume: resumePlayback(); return true;
    case KeyCommand::ExitPlayback: stopPlayback(); return true;
    case KeyCommand::PanLeft: panVideo(-1, 0); return true;
    case KeyCommand::PanRight: panVideo(1, 0); return true;
    case KeyCommand::PanUp: panVideo(0, -1); return true;
    case KeyCommand::PanDown: panVideo(0, 1); return true;
    case KeyCommand::CenterPan: centerVideoPan(); return true;
    case KeyCommand::Playlist: showPlaylistChooser(); return true;
    case KeyCommand::VideoProperties: showVideoProperties(); return true;
    case KeyCommand::VideoTools: showVideoTools(); return true;
    case KeyCommand::RenameDeferred: renameCurrentDeferred(); return true;
    case KeyCommand::CopyDeferred: copyCurrentDeferred(); return true;
    case KeyCommand::ToggleLooping: toggleLooping(); return true;
    case KeyCommand::ZoomIn: adjustZoom(true); return true;
    case KeyCommand::ZoomOut: adjustZoom(false); return true;
    case KeyCommand::Previous: previousVideo(); return true;
    case KeyCommand::Next: nextVideo(); return true;
    case KeyCommand::RemoveDeferred: removeCurrentDeferred(); return true;
    case KeyCommand::VolumeUp:
        volumeSlider_->setValue(qMin(200, volumeSlider_->value() + 5)); return true;
    case KeyCommand::VolumeDown:
        volumeSlider_->setValue(qMax(0, volumeSlider_->value() - 5)); return true;
    case KeyCommand::SeekBack10: seekBy(-10000); return true;
    case KeyCommand::SeekForward10: seekBy(10000); return true;
    case KeyCommand::SeekBack60: seekBy(-60000); return true;
    case KeyCommand::SeekForward60: seekBy(60000); return true;
    default: return false;
    }
}

void MediaExplorerWindow::setBusy(const QString &owner, const QString &message,
                                  const std::shared_ptr<std::atomic_bool> &cancel,
                                  int maximum) {
    busyOwner_ = owner;
    busyCancel_ = cancel;
    statusBar()->showMessage(message);
    progress_->setRange(0, maximum > 0 ? maximum : 0);
    progress_->setValue(0);
    progress_->show();
    cancelButton_->setVisible(static_cast<bool>(cancel));
}

void MediaExplorerWindow::finishBusy(const QString &owner, const QString &message) {
    if (busyOwner_ != owner) return;
    busyOwner_.clear();
    busyCancel_.reset();
    progress_->hide();
    cancelButton_->hide();
    if (!message.isEmpty()) statusBar()->showMessage(message, 7000);
    else statusBar()->clearMessage();
}

void MediaExplorerWindow::cancelBusy() {
    if (busyCancel_) {
        busyCancel_->store(true);
        statusBar()->showMessage(QStringLiteral("Cancelling %1...").arg(busyOwner_));
    }
}

void MediaExplorerWindow::invalidateScan() {
    ++scanGeneration_;
    if (scanCancel_) scanCancel_->store(true);
    scanCancel_.reset();
}

void MediaExplorerWindow::invalidateSearch() {
    ++searchToken_;
    if (searchCancel_) searchCancel_->store(true);
    searchCancel_.reset();
}

void MediaExplorerWindow::invalidateMappingProbe() {
    ++mappingToken_;
    mappingPool_.clear();
}

void MediaExplorerWindow::clearMetadataQueue() {
    ++metadataGeneration_;
    metadataPool_.clear();
    metadataPending_.clear();
}

void MediaExplorerWindow::showRoots() {
    if (fileOperationCancel_) {
        statusBar()->showMessage(QStringLiteral("Wait for the active file operation or cancel it."), 5000);
        return;
    }
    if (mediaProcessCancel_) {
        statusBar()->showMessage(QStringLiteral("Wait for the active media operation or cancel it."), 5000);
        return;
    }
    invalidateScan();
    invalidateSearch();
    invalidateMappingProbe();
    clearMetadataQueue();
    if (busyOwner_ == QStringLiteral("scan") || busyOwner_ == QStringLiteral("search") ||
        busyOwner_ == QStringLiteral("mapping")) finishBusy(busyOwner_);
    viewMode_ = ViewMode::Roots;
    currentDirectory_.clear();
    searchTerms_.clear();
    searchScopes_.clear();
    location_->setText(QStringLiteral("Mounts and mappings"));
    location_->setReadOnly(true);
    entries_ = discoverMounts(config_);
    sortAndFill();
    statusBar()->showMessage(QStringLiteral("%1 location(s)").arg(entries_.size()), 5000);
}

void MediaExplorerWindow::navigate(const QString &path, const QString &selectPath) {
    if (fileOperationCancel_) {
        statusBar()->showMessage(QStringLiteral("Wait for the active file operation or cancel it."), 5000);
        return;
    }
    if (mediaProcessCancel_) {
        statusBar()->showMessage(QStringLiteral("Wait for the active media operation or cancel it."), 5000);
        return;
    }
    const QString requested = path.trimmed();
    const QString cleaned = requested.isEmpty()
                                ? QString()
                                : QDir::cleanPath(QFileInfo(requested).absoluteFilePath());
    if (cleaned.isEmpty()) return;
    invalidateScan();
    invalidateSearch();
    invalidateMappingProbe();
    clearMetadataQueue();
    if (busyOwner_ == QStringLiteral("scan") || busyOwner_ == QStringLiteral("search") ||
        busyOwner_ == QStringLiteral("mapping")) finishBusy(busyOwner_);

    viewMode_ = ViewMode::Folder;
    currentDirectory_ = cleaned;
    searchTerms_.clear();
    searchScopes_.clear();
    location_->setReadOnly(false);
    location_->setText(cleaned);
    entries_.clear();
    fillTable();

    const quint64 generation = scanGeneration_;
    scanCancel_ = std::make_shared<std::atomic_bool>(false);
    const auto cancel = scanCancel_;
    setBusy(QStringLiteral("scan"), QStringLiteral("Reading %1...").arg(cleaned), cancel);
    const AppConfig config = config_;
    QPointer<MediaExplorerWindow> guard(this);
    workerPool_.start(new LambdaRunnable([guard, generation, cleaned, selectPath, cancel, config] {
        QVector<Entry> found;
        QString error;
        bool cancelled = false;
        QDir directory(cleaned);
        if (!directory.exists()) {
            error = QStringLiteral("Directory does not exist or is not accessible: %1").arg(cleaned);
        } else {
            QDir::Filters filters = QDir::AllEntries | QDir::NoDotAndDotDot | QDir::System;
            if (config.showHidden) filters |= QDir::Hidden;
            const QFileInfoList infos = directory.entryInfoList(filters, QDir::NoSort);
            found.reserve(infos.size());
            for (const QFileInfo &info : infos) {
                if (cancel->load()) { cancelled = true; break; }
                const bool directoryEntry = (!info.isSymLink() || config.followSymlinks) && info.isDir();
                if (!directoryEntry && !isVideoFile(info.filePath(), config.videoExtensions)) continue;
                Entry entry;
                entry.name = info.fileName();
                entry.path = info.filePath();
                entry.directory = directoryEntry;
                entry.kind = directoryEntry
                                 ? (info.isSymLink() ? QStringLiteral("Folder link") : QStringLiteral("Folder"))
                                 : fileKind(info.filePath());
                entry.size = directoryEntry ? -1 : info.size();
                entry.modifiedMs = info.lastModified().toMSecsSinceEpoch();
                entry.source = EntrySource::Folder;
                found.append(entry);
            }
        }
        if (!guard) return;
        QMetaObject::invokeMethod(guard, [guard, generation, cleaned, found = std::move(found),
                                         error, cancelled, selectPath]() mutable {
            if (guard) guard->scanFinished(generation, cleaned, std::move(found), error,
                                           cancelled, selectPath);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::scanFinished(quint64 generation, const QString &path,
                                       QVector<Entry> entries, const QString &error,
                                       bool cancelled, const QString &selectPath) {
    if (generation != scanGeneration_ || viewMode_ != ViewMode::Folder ||
        currentDirectory_ != path) return;
    scanCancel_.reset();
    if (!error.isEmpty()) {
        finishBusy(QStringLiteral("scan"));
        showError(QStringLiteral("Open folder"), error);
        showRoots();
        return;
    }
    if (cancelled) {
        finishBusy(QStringLiteral("scan"), QStringLiteral("Folder scan cancelled"));
        return;
    }
    entries_ = std::move(entries);
    sortAndFill();
    if (!selectPath.isEmpty()) selectPaths({selectPath});
    finishBusy(QStringLiteral("scan"), QStringLiteral("%1 item(s)").arg(entries_.size()));
    queueMetadata();
}

void MediaExplorerWindow::refresh() {
    if (viewMode_ == ViewMode::Roots) showRoots();
    else if (viewMode_ == ViewMode::Folder) navigate(currentDirectory_);
    else if (viewMode_ == ViewMode::Search || viewMode_ == ViewMode::Searching)
        startSearch(searchScopes_, searchTerms_, false);
}

void MediaExplorerWindow::goToLocation() {
    if (location_->isReadOnly()) return;
    const QString path = expandPath(location_->text());
    if (path.isEmpty()) return;
    navigate(path);
}

void MediaExplorerWindow::goUp() {
    if (viewMode_ == ViewMode::Search || viewMode_ == ViewMode::Searching) {
        if (searchReturnRoots_ || searchReturnDirectory_.isEmpty()) showRoots();
        else navigate(searchReturnDirectory_);
        return;
    }
    if (viewMode_ == ViewMode::Roots || currentDirectory_.isEmpty() ||
        currentDirectory_ == QStringLiteral("/")) {
        showRoots();
        return;
    }
    const QString child = currentDirectory_;
    QDir parent(currentDirectory_);
    if (!parent.cdUp()) showRoots();
    else navigate(parent.absolutePath(), child);
}

QVector<Entry> MediaExplorerWindow::selectedEntries() const {
    QVector<int> rows;
    for (const QModelIndex &index : table_->selectionModel()->selectedRows()) rows.append(index.row());
    std::sort(rows.begin(), rows.end());
    QVector<Entry> selected;
    selected.reserve(rows.size());
    for (const int row : std::as_const(rows)) {
        if (row >= 0 && row < entries_.size()) selected.append(entries_.at(row));
    }
    return selected;
}

QStringList MediaExplorerWindow::selectedVideoPaths() const {
    QStringList paths;
    for (const Entry &entry : selectedEntries()) {
        if (!entry.directory && isVideoFile(entry.path, config_.videoExtensions)) paths.append(entry.path);
    }
    return paths;
}

void MediaExplorerWindow::activateSelection() {
    const QVector<Entry> selected = selectedEntries();
    if (selected.isEmpty()) return;
    if (selected.first().directory) activateRow(rowByPath_.value(selected.first().path, -1));
    else playSelected();
}

void MediaExplorerWindow::activateRow(int row) {
    if (row < 0 || row >= entries_.size()) return;
    const Entry entry = entries_.at(row);
    if (!entry.directory) {
        playSelected();
        return;
    }
    switch (entry.source) {
    case EntrySource::MappingStable:
    case EntrySource::MappingGvfs:
    case EntrySource::MappingAutomount:
    case EntrySource::MappingDisconnected:
    case EntrySource::Mount:
        activateMapping(entry);
        break;
    default:
        navigate(entry.path);
        break;
    }
}

void MediaExplorerWindow::activateMapping(const Entry &entry) {
    invalidateMappingProbe();
    const quint64 token = mappingToken_;
    setBusy(QStringLiteral("mapping"), QStringLiteral("Checking %1...").arg(entry.name));
    QPointer<MediaExplorerWindow> guard(this);
    mappingPool_.start(new LambdaRunnable([guard, token, entry] {
        const MountProbeResult result = probeStableMount(entry.path, 20000);
        if (!guard) return;
        QMetaObject::invokeMethod(guard, [guard, token, entry, result] {
            if (guard) guard->mappingProbeFinished(token, entry, result);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::mappingProbeFinished(quint64 token, const Entry &entry,
                                               const MountProbeResult &result) {
    if (token != mappingToken_) return;
    const bool usable = entry.source == EntrySource::MappingGvfs
                            ? result.accessible
                            : result.accessible && result.state == MountState::Mounted;
    if (usable) {
        finishBusy(QStringLiteral("mapping"));
        navigate(entry.path);
        return;
    }
    finishBusy(QStringLiteral("mapping"));
    if (!entry.uri.isEmpty()) openMappingUri(entry, result.error);
    else showError(QStringLiteral("Open mount"),
                   result.error.isEmpty() ? QStringLiteral("The mount is not accessible.") : result.error);
}

void MediaExplorerWindow::openMappingUri(const Entry &entry, const QString &stableError) {
    const QString gvfs = gvfsPathForUri(entry.uri);
    if (!gvfs.isEmpty() && (entry.source != EntrySource::MappingGvfs || gvfs != entry.path)) {
        Entry gvfsEntry = entry;
        gvfsEntry.path = gvfs;
        gvfsEntry.source = EntrySource::MappingGvfs;
        activateMapping(gvfsEntry);
        return;
    }
    if (!QDesktopServices::openUrl(QUrl(entry.uri))) {
        showError(QStringLiteral("Connect mapping"),
                  QStringLiteral("Could not open %1.%2").arg(
                      entry.uri, stableError.isEmpty() ? QString() : QLatin1Char('\n') + stableError));
        return;
    }
    statusBar()->showMessage(QStringLiteral("Complete the connection prompt for %1, then retry.")
                                 .arg(entry.name), 9000);
    QPointer<MediaExplorerWindow> guard(this);
    for (const int delay : {2500, 7000}) {
        QTimer::singleShot(delay, this, [guard] {
            if (guard && guard->viewMode_ == ViewMode::Roots) guard->showRoots();
        });
    }
}

void MediaExplorerWindow::selectAllVideos() {
    table_->clearSelection();
    for (int row = 0; row < entries_.size(); ++row) {
        if (!entries_.at(row).directory && isVideoFile(entries_.at(row).path, config_.videoExtensions))
            table_->selectionModel()->select(table_->model()->index(row, 0),
                                             QItemSelectionModel::Select | QItemSelectionModel::Rows);
    }
}

void MediaExplorerWindow::selectPaths(const QStringList &paths) {
    const QSet<QString> wanted(paths.cbegin(), paths.cend());
    table_->clearSelection();
    for (int row = 0; row < entries_.size(); ++row) {
        if (wanted.contains(entries_.at(row).path))
            table_->selectionModel()->select(table_->model()->index(row, 0),
                                             QItemSelectionModel::Select | QItemSelectionModel::Rows);
    }
}

void MediaExplorerWindow::moveSelectedRow(int direction) {
    const QModelIndexList rows = table_->selectionModel()->selectedRows();
    if (rows.size() != 1 || direction == 0) return;
    const int oldRow = rows.first().row();
    const int newRow = oldRow + direction;
    if (oldRow < 0 || newRow < 0 || newRow >= entries_.size()) return;
    entries_.swapItemsAt(oldRow, newRow);
    fillTable();
    table_->selectRow(newRow);
    table_->setCurrentCell(newRow, 0);
}

QVariant MediaExplorerWindow::sortValue(const Entry &entry, int column) const {
    switch (column) {
    case 1: return entry.kind.toLower();
    case 2: return entry.size;
    case 3: return entry.modifiedMs;
    case 4: return entry.resolution;
    case 5: return entry.durationSeconds;
    default: return entry.name.toLower();
    }
}

void MediaExplorerWindow::sortAndFill(bool keepSelection) {
    QStringList selected;
    if (keepSelection) {
        for (const Entry &entry : selectedEntries()) selected.append(entry.path);
    }
    const int column = sortColumn_;
    const bool ascending = sortAscending_;
    std::stable_sort(entries_.begin(), entries_.end(), [column, ascending](const Entry &left,
                                                                          const Entry &right) {
        if (left.directory != right.directory) return left.directory;
        int comparison = 0;
        switch (column) {
        case 1: comparison = QString::localeAwareCompare(left.kind, right.kind); break;
        case 2: comparison = left.size == right.size ? 0 : (left.size < right.size ? -1 : 1); break;
        case 3: comparison = left.modifiedMs == right.modifiedMs ? 0 : (left.modifiedMs < right.modifiedMs ? -1 : 1); break;
        case 4: comparison = QString::localeAwareCompare(left.resolution, right.resolution); break;
        case 5: comparison = left.durationSeconds == right.durationSeconds ? 0 : (left.durationSeconds < right.durationSeconds ? -1 : 1); break;
        default: comparison = QString::localeAwareCompare(left.name, right.name); break;
        }
        if (comparison == 0) comparison = QString::localeAwareCompare(left.name, right.name);
        return ascending ? comparison < 0 : comparison > 0;
    });
    fillTable();
    if (keepSelection) selectPaths(selected);
}

void MediaExplorerWindow::fillTable() {
    table_->setSortingEnabled(false);
    table_->setRowCount(entries_.size());
    rowByPath_.clear();
    for (int row = 0; row < entries_.size(); ++row) {
        Entry &entry = entries_[row];
        const auto cached = metadataCache_.constFind(entry.path);
        if (cached != metadataCache_.cend() && cached->modifiedMs == entry.modifiedMs &&
            cached->size == entry.size && cached->error.isEmpty()) {
            entry.resolution = cached->width > 0
                                   ? QStringLiteral("%1x%2").arg(cached->width).arg(cached->height)
                                   : QString();
            entry.durationSeconds = cached->duration;
        }
        const QStringList values{entry.name, entry.kind,
                                 entry.directory ? QString() : formatBytes(entry.size),
                                 formatModified(entry.modifiedMs), entry.resolution,
                                 formatDuration(entry.durationSeconds)};
        for (int column = 0; column < values.size(); ++column) {
            auto *item = new QTableWidgetItem(values.at(column));
            if (column == 0) item->setData(Qt::UserRole, entry.path);
            if (column == 2 || column == 5) item->setTextAlignment(Qt::AlignRight | Qt::AlignVCenter);
            table_->setItem(row, column, item);
        }
        rowByPath_.insert(entry.path, row);
    }
    table_->setSortingEnabled(false);
}

void MediaExplorerWindow::queueMetadata() {
    if (!config_.ffprobeAvailable || config_.metadataPrefetchLimit <= 0) return;
    const quint64 generation = metadataGeneration_;
    const int availableSlots = qMax(0, config_.metadataPrefetchLimit - metadataPending_.size());
    int queued = 0;
    for (const Entry &entry : std::as_const(entries_)) {
        if (queued >= availableSlots) break;
        if (entry.directory || !isVideoFile(entry.path, config_.videoExtensions)) continue;
        const auto cached = metadataCache_.constFind(entry.path);
        if (cached != metadataCache_.cend() && cached->modifiedMs == entry.modifiedMs &&
            cached->size == entry.size) continue;
        if (metadataPending_.contains(entry.path)) continue;
        metadataPending_.insert(entry.path);
        ++queued;
        const AppConfig config = config_;
        QPointer<MediaExplorerWindow> guard(this);
        metadataPool_.start(new LambdaRunnable([guard, generation, entry, config] {
            const Metadata metadata = probeMetadata(entry, config);
            if (!guard) return;
            QMetaObject::invokeMethod(guard, [guard, generation, path = entry.path, metadata] {
                if (guard) guard->metadataFinished(generation, path, metadata);
            }, Qt::QueuedConnection);
        }));
    }
}

MediaExplorerWindow::Metadata MediaExplorerWindow::probeMetadata(
    const Entry &entry, const AppConfig &config,
    const std::shared_ptr<std::atomic_bool> &cancel) {
    Metadata metadata;
    metadata.modifiedMs = entry.modifiedMs;
    metadata.size = entry.size;
    if (!config.ffprobeAvailable) {
        metadata.error = QStringLiteral("ffprobe is disabled");
        return metadata;
    }
    const QString program = executableResolution(config.ffprobePath);
    if (program.isEmpty()) {
        metadata.error = QStringLiteral("ffprobe executable was not found");
        return metadata;
    }
    QStringList arguments = config.ffprobeArgs;
    arguments << QStringLiteral("-v") << QStringLiteral("error")
              << QStringLiteral("-show_entries")
              << QStringLiteral("stream=codec_type,codec_name,width,height,duration:format=duration")
              << QStringLiteral("-of") << QStringLiteral("json") << entry.path;
    QProcess process;
    process.setProgram(program);
    process.setArguments(arguments);
    process.setProcessChannelMode(QProcess::SeparateChannels);
    process.start();
    if (!process.waitForStarted(5000)) {
        metadata.error = QStringLiteral("Could not start ffprobe: %1").arg(process.errorString());
        return metadata;
    }
    int elapsed = 0;
    while (!process.waitForFinished(100)) {
        elapsed += 100;
        if ((cancel && cancel->load()) || elapsed >= 30000) {
            process.kill();
            process.waitForFinished(2000);
            metadata.error = cancel && cancel->load() ? QStringLiteral("Cancelled")
                                                       : QStringLiteral("ffprobe timed out");
            return metadata;
        }
    }
    if (process.exitStatus() != QProcess::NormalExit || process.exitCode() != 0) {
        metadata.error = QString::fromUtf8(process.readAllStandardError()).trimmed();
        if (metadata.error.isEmpty()) metadata.error = QStringLiteral("ffprobe failed");
        return metadata;
    }
    QJsonParseError parseError;
    const QJsonDocument document = QJsonDocument::fromJson(process.readAllStandardOutput(), &parseError);
    if (parseError.error != QJsonParseError::NoError || !document.isObject()) {
        metadata.error = QStringLiteral("Invalid ffprobe response: %1").arg(parseError.errorString());
        return metadata;
    }
    const QJsonObject root = document.object();
    metadata.duration = root.value(QStringLiteral("format")).toObject()
                            .value(QStringLiteral("duration")).toString().toDouble();
    for (const QJsonValue value : root.value(QStringLiteral("streams")).toArray()) {
        const QJsonObject stream = value.toObject();
        if (stream.value(QStringLiteral("codec_type")).toString() != QStringLiteral("video")) continue;
        metadata.width = stream.value(QStringLiteral("width")).toInt();
        metadata.height = stream.value(QStringLiteral("height")).toInt();
        metadata.codec = stream.value(QStringLiteral("codec_name")).toString();
        if (metadata.duration <= 0.0)
            metadata.duration = stream.value(QStringLiteral("duration")).toString().toDouble();
        break;
    }
    return metadata;
}

void MediaExplorerWindow::metadataFinished(quint64 generation, const QString &path,
                                           const Metadata &metadata) {
    if (generation != metadataGeneration_ ||
        (viewMode_ != ViewMode::Folder && viewMode_ != ViewMode::Search)) return;
    metadataPending_.remove(path);
    metadataCache_.insert(path, metadata);
    const int row = rowByPath_.value(path, -1);
    if (row < 0 || row >= entries_.size()) return;
    Entry &entry = entries_[row];
    if (entry.modifiedMs != metadata.modifiedMs || entry.size != metadata.size) return;
    if (metadata.error.isEmpty()) {
        entry.resolution = metadata.width > 0
                               ? QStringLiteral("%1x%2").arg(metadata.width).arg(metadata.height)
                               : QString();
        entry.durationSeconds = metadata.duration;
        if (table_->item(row, 4)) table_->item(row, 4)->setText(entry.resolution);
        if (table_->item(row, 5)) table_->item(row, 5)->setText(formatDuration(entry.durationSeconds));
    }
}

void MediaExplorerWindow::showVideoProperties() {
    QString path;
    if (stack_->currentWidget() == playerPage_ && playlistIndex_ >= 0 &&
        playlistIndex_ < playlist_.size()) path = playlist_.at(playlistIndex_);
    else {
        const QStringList selected = selectedVideoPaths();
        if (!selected.isEmpty()) path = selected.first();
    }
    if (path.isEmpty()) {
        showError(QStringLiteral("Video properties"), QStringLiteral("Select a video first."));
        return;
    }
    const QFileInfo info(path);
    Entry entry;
    entry.path = path;
    entry.name = info.fileName();
    entry.size = info.size();
    entry.modifiedMs = info.lastModified().toMSecsSinceEpoch();
    const auto cached = metadataCache_.constFind(path);
    if (cached != metadataCache_.cend() && cached->size == entry.size &&
        cached->modifiedMs == entry.modifiedMs) {
        showPropertiesDialog(path, cached.value());
        return;
    }
    statusBar()->showMessage(QStringLiteral("Reading video properties..."));
    const AppConfig config = config_;
    QPointer<MediaExplorerWindow> guard(this);
    metadataPool_.start(new LambdaRunnable([guard, entry, config] {
        const Metadata metadata = probeMetadata(entry, config);
        if (!guard) return;
        QMetaObject::invokeMethod(guard, [guard, entry, metadata] {
            if (!guard) return;
            guard->metadataCache_.insert(entry.path, metadata);
            guard->statusBar()->clearMessage();
            guard->showPropertiesDialog(entry.path, metadata);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::showPropertiesDialog(const QString &path, const Metadata &metadata) {
    const QFileInfo info(path);
    QStringList lines{
        QStringLiteral("Path: %1").arg(path),
        QStringLiteral("Type: %1").arg(fileKind(path)),
        QStringLiteral("Size: %1").arg(formatBytes(info.size())),
        QStringLiteral("Modified: %1").arg(QLocale().toString(info.lastModified(), QLocale::LongFormat))};
    if (metadata.error.isEmpty()) {
        lines << QStringLiteral("Resolution: %1x%2").arg(metadata.width).arg(metadata.height)
              << QStringLiteral("Duration: %1").arg(formatDuration(metadata.duration))
              << QStringLiteral("Video codec: %1").arg(metadata.codec.isEmpty()
                                                          ? QStringLiteral("unknown") : metadata.codec);
    } else {
        lines << QStringLiteral("Metadata: %1").arg(metadata.error);
    }
    QMessageBox::information(this, QStringLiteral("Video properties — %1").arg(info.fileName()),
                             lines.join(QLatin1Char('\n')));
}

void MediaExplorerWindow::promptSearch() {
    if (stack_->currentWidget() == playerPage_) return;
    bool accepted = false;
    const QString initial;
    const QString text = QInputDialog::getText(
        this, QStringLiteral("Search videos"),
        viewMode_ == ViewMode::Search
            ? QStringLiteral("Additional keyword to AND with the current search:")
            : QStringLiteral("Keyword to find in the video file name:"),
        QLineEdit::Normal, initial, &accepted);
    if (!accepted) return;
    const QString keyword = text.trimmed().toLower();
    if (keyword.isEmpty()) {
        showError(QStringLiteral("Search"), QStringLiteral("Enter a search keyword."));
        return;
    }
    QStringList scopes;
    if (viewMode_ == ViewMode::Search) {
        scopes = searchScopes_;
    } else {
        for (const Entry &entry : selectedEntries()) {
            if (entry.source != EntrySource::MappingDisconnected)
                scopes.append(entry.path);
        }
        if (scopes.isEmpty()) {
            if (!currentDirectory_.isEmpty()) scopes.append(currentDirectory_);
            else {
                for (const Entry &entry : std::as_const(entries_)) {
                    if (entry.directory && entry.source != EntrySource::MappingDisconnected &&
                        entry.source != EntrySource::MappingAutomount)
                        scopes.append(entry.path);
                }
                if (scopes.isEmpty()) scopes.append(QDir::homePath());
            }
        }
        searchReturnRoots_ = currentDirectory_.isEmpty();
        searchReturnDirectory_ = currentDirectory_;
    }
    startSearch(scopes, {keyword}, viewMode_ == ViewMode::Search);
}

void MediaExplorerWindow::startSearch(const QStringList &scopes, const QStringList &terms,
                                      bool refine) {
    if (fileOperationCancel_) {
        statusBar()->showMessage(QStringLiteral("Wait for the active file operation or cancel it."), 5000);
        return;
    }
    if (mediaProcessCancel_) {
        statusBar()->showMessage(QStringLiteral("Wait for the active media operation or cancel it."), 5000);
        return;
    }
    QStringList normalizedTerms;
    for (const QString &term : terms) {
        const QString normalized = term.trimmed().toLower();
        if (!normalized.isEmpty() && !normalizedTerms.contains(normalized)) normalizedTerms.append(normalized);
    }
    if (normalizedTerms.isEmpty()) return;
    if (refine && viewMode_ == ViewMode::Search) {
        for (const QString &term : std::as_const(normalizedTerms)) {
            if (!searchTerms_.contains(term)) searchTerms_.append(term);
        }
        QVector<Entry> filtered;
        for (const Entry &entry : std::as_const(entries_)) {
            const QString name = QFileInfo(entry.path).fileName().toLower();
            bool matches = true;
            for (const QString &term : std::as_const(searchTerms_)) matches = matches && name.contains(term);
            if (matches) filtered.append(entry);
        }
        entries_ = std::move(filtered);
        location_->setText(QStringLiteral("Search: %1").arg(searchTerms_.join(QStringLiteral(" AND "))));
        sortAndFill();
        statusBar()->showMessage(QStringLiteral("Refined search: %1 match(es)").arg(entries_.size()), 6000);
        queueMetadata();
        return;
    }

    invalidateScan();
    invalidateSearch();
    invalidateMappingProbe();
    clearMetadataQueue();
    if (busyOwner_ == QStringLiteral("scan") || busyOwner_ == QStringLiteral("search") ||
        busyOwner_ == QStringLiteral("mapping")) finishBusy(busyOwner_);
    QSet<QString> seenScopes;
    searchScopes_.clear();
    for (const QString &scope : scopes) {
        const QString absolute = QDir::cleanPath(QFileInfo(scope).absoluteFilePath());
        if (!seenScopes.contains(absolute)) { seenScopes.insert(absolute); searchScopes_.append(absolute); }
    }
    searchTerms_ = normalizedTerms;
    viewMode_ = ViewMode::Searching;
    entries_.clear();
    fillTable();
    location_->setReadOnly(true);
    location_->setText(QStringLiteral("Search: %1").arg(searchTerms_.join(QStringLiteral(" AND "))));
    const quint64 token = searchToken_;
    searchCancel_ = std::make_shared<std::atomic_bool>(false);
    const auto cancel = searchCancel_;
    setBusy(QStringLiteral("search"), QStringLiteral("Searching..."), cancel);
    const AppConfig config = config_;
    const QStringList searchScopes = searchScopes_;
    const QStringList searchTerms = searchTerms_;
    QPointer<MediaExplorerWindow> guard(this);
    workerPool_.start(new LambdaRunnable([guard, token, cancel, config, searchScopes, searchTerms] {
        QVector<Entry> results;
        QStringList errors;
        int directories = 0;
        int files = 0;
        QStringList stack;
        for (auto it = searchScopes.crbegin(); it != searchScopes.crend(); ++it) stack.append(*it);
        QSet<QString> seen;
        while (!stack.isEmpty() && !cancel->load()) {
            const QString folder = stack.takeLast();
            if (skipSearchPath(folder, searchScopes)) continue;
            const QFileInfo scopeInfo(folder);
            if (scopeInfo.isFile()) {
                if (isVideoFile(folder, config.videoExtensions)) {
                    ++files;
                    const QString folded = scopeInfo.fileName().toLower();
                    bool matches = true;
                    for (const QString &term : searchTerms) matches = matches && folded.contains(term);
                    if (matches) {
                        Entry entry;
                        entry.name = scopeInfo.absoluteFilePath();
                        entry.path = scopeInfo.absoluteFilePath();
                        entry.kind = fileKind(entry.path);
                        entry.size = scopeInfo.size();
                        entry.modifiedMs = scopeInfo.lastModified().toMSecsSinceEpoch();
                        entry.source = EntrySource::Search;
                        results.append(entry);
                    }
                }
                continue;
            }
            const QString canonical = QFileInfo(folder).canonicalFilePath();
            const QString identity = canonical.isEmpty() ? QDir::cleanPath(folder) : canonical;
            if (seen.contains(identity)) continue;
            seen.insert(identity);
            ++directories;
            if (guard && (directories == 1 || directories % 20 == 0)) {
                const int checked = files;
                const int matches = results.size();
                QMetaObject::invokeMethod(guard, [guard, token, folder, checked, matches] {
                    if (guard && token == guard->searchToken_ && guard->viewMode_ == ViewMode::Searching)
                        guard->statusBar()->showMessage(
                            QStringLiteral("Searching %1 — %2 videos, %3 matches")
                                .arg(folder).arg(checked).arg(matches));
                }, Qt::QueuedConnection);
            }
            QDir directory(folder);
            if (!directory.exists()) {
                if (errors.size() < 25) errors.append(QStringLiteral("Not accessible: %1").arg(folder));
                continue;
            }
            QDir::Filters filters = QDir::AllEntries | QDir::NoDotAndDotDot | QDir::System;
            if (config.showHidden) filters |= QDir::Hidden;
            const QFileInfoList infos = directory.entryInfoList(filters, QDir::NoSort);
            for (const QFileInfo &info : infos) {
                if (cancel->load()) break;
                if (info.isDir()) {
                    if (!info.isSymLink() || config.followSymlinks) stack.append(info.filePath());
                    continue;
                }
                if (!isVideoFile(info.filePath(), config.videoExtensions)) continue;
                ++files;
                const QString folded = info.fileName().toLower();
                bool matches = true;
                for (const QString &term : searchTerms) matches = matches && folded.contains(term);
                if (!matches) continue;
                Entry entry;
                entry.name = info.absoluteFilePath();
                entry.path = info.absoluteFilePath();
                entry.kind = fileKind(entry.path);
                entry.size = info.size();
                entry.modifiedMs = info.lastModified().toMSecsSinceEpoch();
                entry.source = EntrySource::Search;
                results.append(entry);
            }
        }
        if (!guard) return;
        const bool cancelled = cancel->load();
        QMetaObject::invokeMethod(guard, [guard, token, results = std::move(results), directories,
                                         files, errors, cancelled]() mutable {
            if (guard) guard->searchFinished(token, std::move(results), directories, files,
                                             errors, cancelled);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::searchFinished(quint64 token, QVector<Entry> entries, int directories,
                                         int files, QStringList errors, bool cancelled) {
    if (token != searchToken_ || viewMode_ != ViewMode::Searching) return;
    searchCancel_.reset();
    viewMode_ = ViewMode::Search;
    entries_ = std::move(entries);
    sortAndFill();
    const QString state = cancelled ? QStringLiteral("cancelled; partial") : QStringLiteral("complete");
    QString message = QStringLiteral("Search %1: %2 matches, %3 videos in %4 folders")
                          .arg(state).arg(entries_.size()).arg(files).arg(directories);
    if (!errors.isEmpty()) message += QStringLiteral(" (%1 inaccessible folders)").arg(errors.size());
    finishBusy(QStringLiteral("search"), message);
    queueMetadata();
}

bool MediaExplorerWindow::fileChangesAllowed(const QString &action) const {
    if (stack_->currentWidget() == playerPage_) {
        showError(action, QStringLiteral("Exit playback before changing browser files."));
        return false;
    }
    if (fileOperationCancel_) {
        showError(action, QStringLiteral("Another file operation is already running."));
        return false;
    }
    if (mediaProcessCancel_) {
        showError(action, QStringLiteral("A media operation is running. Wait for it or cancel it first."));
        return false;
    }
    if (viewMode_ == ViewMode::Folder || viewMode_ == ViewMode::Search) return true;
    showError(action, QStringLiteral("Open a folder first. Mount roots and mapped-drive roots "
                                     "cannot be modified from the roots view."));
    return false;
}

void MediaExplorerWindow::copySelected() {
    if (!fileChangesAllowed(QStringLiteral("Copy"))) return;
    clipboardPaths_.clear();
    for (const Entry &entry : selectedEntries()) clipboardPaths_.append(entry.path);
    if (clipboardPaths_.isEmpty()) return;
    clipboardMode_ = ClipboardMode::Copy;
    statusBar()->showMessage(QStringLiteral("Copied %1 item(s) to the Media Explorer clipboard")
                                 .arg(clipboardPaths_.size()), 5000);
}

void MediaExplorerWindow::cutSelected() {
    if (!fileChangesAllowed(QStringLiteral("Cut"))) return;
    clipboardPaths_.clear();
    for (const Entry &entry : selectedEntries()) clipboardPaths_.append(entry.path);
    if (clipboardPaths_.isEmpty()) return;
    clipboardMode_ = ClipboardMode::Move;
    statusBar()->showMessage(QStringLiteral("Cut %1 item(s); open a folder and paste")
                                 .arg(clipboardPaths_.size()), 5000);
}

void MediaExplorerWindow::paste() {
    if (stack_->currentWidget() == playerPage_) {
        showError(QStringLiteral("Paste"), QStringLiteral("Exit playback before pasting files."));
        return;
    }
    if (mediaProcessCancel_) {
        showError(QStringLiteral("Paste"), QStringLiteral("Wait for the media operation to finish first."));
        return;
    }
    if (viewMode_ != ViewMode::Folder || currentDirectory_.isEmpty()) {
        showError(QStringLiteral("Paste"), QStringLiteral("Open the destination folder before pasting."));
        return;
    }
    QStringList sources;
    for (const QString &path : std::as_const(clipboardPaths_)) {
        if (pathLexists(path)) sources.append(path);
    }
    if (sources.isEmpty()) {
        showError(QStringLiteral("Paste"), QStringLiteral("The clipboard is empty or its items no longer exist."));
        return;
    }
    runFileOperation(clipboardMode_ == ClipboardMode::Copy ? QStringLiteral("copy")
                                                            : QStringLiteral("move"),
                     sources, currentDirectory_);
}

void MediaExplorerWindow::deleteSelected() {
    if (!fileChangesAllowed(QStringLiteral("Delete"))) return;
    QStringList paths;
    for (const Entry &entry : selectedEntries()) paths.append(entry.path);
    if (paths.isEmpty()) return;
    const bool trash = config_.useTrash && !executableResolution(QStringLiteral("gio")).isEmpty();
    const QString title = trash ? QStringLiteral("Move to Trash") : QStringLiteral("Permanently delete");
    QString prompt = trash
                         ? QStringLiteral("Move %1 selected item(s) to the Trash?").arg(paths.size())
                         : QStringLiteral("Permanently delete %1 selected item(s)?\n\nThis cannot be undone.")
                               .arg(paths.size());
    const QMessageBox::StandardButton answer = trash
        ? QMessageBox::question(this, title, prompt, QMessageBox::Yes | QMessageBox::Cancel,
                                QMessageBox::Cancel)
        : QMessageBox::warning(this, title, prompt, QMessageBox::Yes | QMessageBox::Cancel,
                               QMessageBox::Cancel);
    if (answer == QMessageBox::Yes)
        runFileOperation(trash ? QStringLiteral("trash") : QStringLiteral("delete"), paths);
}

void MediaExplorerWindow::newFolder() {
    if (stack_->currentWidget() == playerPage_) {
        showError(QStringLiteral("New folder"), QStringLiteral("Exit playback before creating a folder."));
        return;
    }
    if (mediaProcessCancel_) {
        showError(QStringLiteral("New folder"), QStringLiteral("Wait for the media operation to finish first."));
        return;
    }
    if (viewMode_ != ViewMode::Folder || currentDirectory_.isEmpty()) {
        showError(QStringLiteral("New folder"), QStringLiteral("Open a destination folder first."));
        return;
    }
    bool accepted = false;
    const QString name = QInputDialog::getText(this, QStringLiteral("New folder"),
                                                QStringLiteral("Folder name:"), QLineEdit::Normal,
                                                {}, &accepted).trimmed();
    if (!accepted) return;
    if (!validBaseName(name)) {
        showError(QStringLiteral("New folder"),
                  QStringLiteral("Use one non-empty name without '/', '.'/'..', or a NUL character."));
        return;
    }
    const QString path = QDir(currentDirectory_).filePath(name);
    if (pathLexists(path)) {
        showError(QStringLiteral("New folder"), QStringLiteral("An item with that name already exists."));
        return;
    }
    if (!QDir().mkdir(path)) {
        showError(QStringLiteral("New folder"), QStringLiteral("Could not create %1").arg(path));
        return;
    }
    navigate(currentDirectory_, path);
}

void MediaExplorerWindow::renameSelected() {
    if (!fileChangesAllowed(QStringLiteral("Rename"))) return;
    const QVector<Entry> selected = selectedEntries();
    if (selected.size() != 1) {
        showError(QStringLiteral("Rename"), QStringLiteral("Select exactly one item."));
        return;
    }
    bool accepted = false;
    const QString name = QInputDialog::getText(this, QStringLiteral("Rename"),
                                                QStringLiteral("New name:"), QLineEdit::Normal,
                                                QFileInfo(selected.first().path).fileName(),
                                                &accepted).trimmed();
    if (!accepted) return;
    if (!validBaseName(name)) {
        showError(QStringLiteral("Rename"), QStringLiteral("Use one non-empty name without '/' or NUL."));
        return;
    }
    const QString target = QDir(QFileInfo(selected.first().path).absolutePath()).filePath(name);
    if (pathLexists(target)) {
        showError(QStringLiteral("Rename"), QStringLiteral("An item with that name already exists."));
        return;
    }
    const bool renamed = selected.first().directory
                             ? QDir().rename(selected.first().path, target)
                             : QFile::rename(selected.first().path, target);
    if (!renamed) {
        showError(QStringLiteral("Rename"), QStringLiteral("Could not rename the selected item."));
        return;
    }
    refresh();
}

void MediaExplorerWindow::openExternal() {
    const QVector<Entry> selected = selectedEntries();
    QString path;
    if (!selected.isEmpty()) path = selected.first().directory
                                         ? selected.first().path
                                         : QFileInfo(selected.first().path).absolutePath();
    else path = currentDirectory_;
    if (path.isEmpty()) return;
    if (!QDesktopServices::openUrl(QUrl::fromLocalFile(path)))
        showError(QStringLiteral("Open file manager"),
                  QStringLiteral("No system file manager accepted:\n%1").arg(path));
}

void MediaExplorerWindow::runFileOperation(const QString &operation, const QStringList &sources,
                                           const QString &destination,
                                           const QVector<DeferredAction> &deferred) {
    if (fileOperationCancel_) {
        showError(QStringLiteral("File operation"),
                  QStringLiteral("Another file operation is already running."));
        return;
    }
    if (mediaProcessCancel_) {
        showError(QStringLiteral("File operation"),
                  QStringLiteral("Wait for the active media operation before changing files."));
        return;
    }
    QStringList protectedPaths{QStringLiteral("/"), QDir::cleanPath(config_.mappedRoot)};
    for (auto it = config_.mappedShares.cbegin(); it != config_.mappedShares.cend(); ++it)
        protectedPaths.append(QDir(config_.mappedRoot).filePath(it.key()));
    for (const MountRecord &mount : readMountRecords()) protectedPaths.append(QDir::cleanPath(mount.path));
    auto isProtected = [&protectedPaths](const QString &path) {
        const QString clean = QDir::cleanPath(QFileInfo(path).absoluteFilePath());
        return protectedPaths.contains(clean);
    };
    for (const QString &source : sources) {
        if (isProtected(source)) {
            showError(QStringLiteral("Protected location"),
                      QStringLiteral("Media Explorer will not modify a filesystem or mount root:\n%1")
                          .arg(source));
            return;
        }
    }
    for (const DeferredAction &action : deferred) {
        if (isProtected(action.source)) {
            showError(QStringLiteral("Protected location"),
                      QStringLiteral("Media Explorer will not modify a filesystem or mount root:\n%1")
                          .arg(action.source));
            return;
        }
    }

    fileOperationCancel_ = std::make_shared<std::atomic_bool>(false);
    const auto cancel = fileOperationCancel_;
    const int total = deferred.isEmpty() ? sources.size() : deferred.size();
    setBusy(QStringLiteral("fileop"), QStringLiteral("Starting %1...").arg(operation), cancel,
            qMax(1, total));
    const bool useTrashForDeferred = config_.useTrash;
    QPointer<MediaExplorerWindow> guard(this);
    workerPool_.start(new LambdaRunnable([guard, operation, sources, destination, deferred, cancel,
                                          useTrashForDeferred] {
        QStringList errors;
        int completed = 0;
        const int totalItems = deferred.isEmpty() ? sources.size() : deferred.size();
        const QString gio = executableResolution(QStringLiteral("gio"));
        auto report = [guard, totalItems](const QString &name, int current) {
            if (!guard) return;
            QMetaObject::invokeMethod(guard, [guard, name, current, totalItems] {
                if (!guard || guard->busyOwner_ != QStringLiteral("fileop")) return;
                guard->progress_->setRange(0, qMax(1, totalItems));
                guard->progress_->setValue(qMax(0, current - 1));
                guard->statusBar()->showMessage(QStringLiteral("Processing %1 of %2: %3")
                                                    .arg(current).arg(totalItems).arg(name));
            }, Qt::QueuedConnection);
        };
        auto movePath = [&cancel](const QString &source, const QString &target, QString &error) {
            const QString nested = nestedMountPoint(source);
            if (!nested.isEmpty()) {
                error = QStringLiteral("Refusing to move a tree containing a nested mount: %1").arg(nested);
                return false;
            }
            if (renameNoReplace(source, target) == 0) return true;
            const int renameError = errno;
            if (renameError == EEXIST) {
                error = QStringLiteral("Destination already exists: %1").arg(target);
                return false;
            }
            if (renameError != EXDEV && renameError != ENOSYS && renameError != EINVAL &&
                renameError != EOPNOTSUPP
#ifdef ENOTSUP
                && renameError != ENOTSUP
#endif
            ) {
                error = QStringLiteral("Could not move %1: %2")
                            .arg(source, QString::fromLocal8Bit(std::strerror(renameError)));
                return false;
            }
            if (!copyPathRecursive(source, target, cancel, error)) return false;
            if (cancel->load()) {
                error = QStringLiteral("Move cancelled after copy; source and destination were retained (%1)")
                            .arg(target);
                return false;
            }
            // Source removal is a commit phase. Do not interrupt it midway and
            // leave an unknowable, partially deleted tree.
            if (!removePathRecursive(source, {}, error)) {
                error += QStringLiteral("; copied destination retained at %1").arg(target);
                return false;
            }
            return true;
        };
        auto deletePath = [&cancel, &gio](const QString &source, bool trash, QString &error) {
            const QString nested = nestedMountPoint(source);
            if (!nested.isEmpty()) {
                error = QStringLiteral("Refusing to delete a tree containing a nested mount: %1")
                            .arg(nested);
                return false;
            }
            if (!trash) return removePathRecursive(source, cancel, error);
            if (gio.isEmpty()) { error = QStringLiteral("gio is unavailable"); return false; }
            QProcess process;
            process.start(gio, {QStringLiteral("trash"), QStringLiteral("--"), source});
            if (!process.waitForStarted(5000)) { error = process.errorString(); return false; }
            while (!process.waitForFinished(100)) {
                if (cancel->load()) { process.kill(); process.waitForFinished(2000); return false; }
            }
            if (process.exitStatus() != QProcess::NormalExit || process.exitCode() != 0) {
                error = QString::fromUtf8(process.readAllStandardError()).trimmed();
                if (error.isEmpty()) error = QStringLiteral("gio trash failed");
                return false;
            }
            return true;
        };

        if (!deferred.isEmpty()) {
            for (int index = 0; index < deferred.size() && !cancel->load(); ++index) {
                const DeferredAction action = deferred.at(index);
                report(QFileInfo(action.source).fileName(), index + 1);
                QString error;
                bool ok = false;
                if (action.kind == DeferredKind::Delete)
                    ok = deletePath(action.source, useTrashForDeferred && !gio.isEmpty(), error);
                else if (action.kind == DeferredKind::Rename)
                    ok = movePath(action.source, action.destination, error);
                else
                    ok = copyPathRecursive(action.source, action.destination, cancel, error);
                if (ok) ++completed;
                else if (!cancel->load() || error.contains(QStringLiteral("retained"), Qt::CaseInsensitive))
                    errors.append(error.isEmpty() ? action.source : error);
            }
        } else {
            for (int index = 0; index < sources.size() && !cancel->load(); ++index) {
                const QString source = sources.at(index);
                report(QFileInfo(source).fileName(), index + 1);
                QString error;
                bool ok = false;
                if (operation == QStringLiteral("trash") || operation == QStringLiteral("delete")) {
                    ok = deletePath(source, operation == QStringLiteral("trash"), error);
                } else {
                    if (destination.isEmpty()) {
                        error = QStringLiteral("No destination folder was provided");
                    } else if (isWithin(destination, source)) {
                        error = QStringLiteral("Cannot copy a folder inside itself: %1").arg(source);
                    } else {
                        const QString target = availableDestination(
                            QDir(destination).filePath(QFileInfo(source).fileName()));
                        ok = operation == QStringLiteral("move")
                                 ? movePath(source, target, error)
                                 : copyPathRecursive(source, target, cancel, error);
                    }
                }
                if (ok) ++completed;
                else if (!cancel->load() || error.contains(QStringLiteral("retained"), Qt::CaseInsensitive))
                    errors.append(error.isEmpty() ? source : error);
            }
        }
        if (!guard) return;
        const bool cancelled = cancel->load();
        QMetaObject::invokeMethod(guard, [guard, operation, completed, totalItems, errors, cancelled] {
            if (guard) guard->fileOperationFinished(operation, completed, totalItems, errors, cancelled);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::fileOperationFinished(const QString &operation, int completed, int total,
                                                const QStringList &errors, bool cancelled) {
    fileOperationCancel_.reset();
    if (clipboardMode_ == ClipboardMode::Move && operation == QStringLiteral("move") &&
        !cancelled && completed == total && errors.isEmpty())
        clipboardPaths_.clear();
    QString message = QStringLiteral("%1: %2 of %3 completed")
                          .arg(operation.left(1).toUpper() + operation.mid(1)).arg(completed).arg(total);
    if (cancelled) message += QStringLiteral(" (cancelled)");
    finishBusy(QStringLiteral("fileop"), message);
    if (!errors.isEmpty()) {
        QStringList shown = errors.mid(0, 12);
        if (errors.size() > 12) shown.append(QStringLiteral("...and %1 more").arg(errors.size() - 12));
        showError(QStringLiteral("File operation errors"), shown.join(QLatin1Char('\n')));
    }
    if (stack_->currentWidget() == browserPage_ && !deferredActions_.isEmpty()) {
        const QVector<DeferredAction> deferred = deferredActions_;
        deferredActions_.clear();
        runFileOperation(QStringLiteral("post-playback"), {}, {}, deferred);
        return;
    }
    refresh();
}

bool MediaExplorerWindow::ensurePlayer() {
    if (vlcPlayer_) return true;
    if (QGuiApplication::platformName() != QStringLiteral("xcb")) {
        showError(QStringLiteral("VLC unavailable"),
                  QStringLiteral("Embedded libVLC playback requires Qt's xcb platform. "
                                 "Restart with QT_QPA_PLATFORM=xcb under X11/XWayland."));
        return false;
    }
    QVector<QByteArray> argumentStorage;
    argumentStorage.reserve(config_.vlcArgs.size());
    for (const QString &argument : std::as_const(config_.vlcArgs))
        argumentStorage.append(argument.toUtf8());
    QVector<const char *> arguments;
    arguments.reserve(argumentStorage.size());
    for (const QByteArray &argument : std::as_const(argumentStorage))
        arguments.append(argument.constData());
    vlcInstance_ = libvlc_new(arguments.size(), arguments.isEmpty() ? nullptr : arguments.constData());
    if (!vlcInstance_) {
        showError(QStringLiteral("VLC unavailable"),
                  QStringLiteral("Could not initialize libVLC. Verify VLC and its plugins are installed."));
        return false;
    }
    vlcPlayer_ = libvlc_media_player_new(vlcInstance_);
    if (!vlcPlayer_) {
        libvlc_release(vlcInstance_);
        vlcInstance_ = nullptr;
        showError(QStringLiteral("VLC unavailable"), QStringLiteral("Could not create a VLC media player."));
        return false;
    }
    libvlc_video_set_key_input(vlcPlayer_, 0);
    libvlc_video_set_mouse_input(vlcPlayer_, 0);
    libvlc_audio_set_volume(vlcPlayer_, volumeSlider_->value());
    return true;
}

bool MediaExplorerWindow::startPlaybackForTest(const QStringList &paths) {
    QVector<Entry> entries;
    for (const QString &path : paths) {
        Entry entry;
        entry.name = QFileInfo(path).fileName();
        entry.path = path;
        entry.kind = fileKind(path);
        entries.append(entry);
    }
    playEntries(entries);
    return vlcPlayer_ && stack_->currentWidget() == playerPage_ && !playlist_.isEmpty();
}

int MediaExplorerWindow::playerStateForTest() const {
    return vlcPlayer_ ? static_cast<int>(libvlc_media_player_get_state(vlcPlayer_)) : -1;
}

qint64 MediaExplorerWindow::playerTimeForTest() const {
    return vlcPlayer_ ? libvlc_media_player_get_time(vlcPlayer_) : -1;
}

qint64 MediaExplorerWindow::playerLengthForTest() const {
    return vlcPlayer_ ? libvlc_media_player_get_length(vlcPlayer_) : -1;
}

quint64 MediaExplorerWindow::playbackGenerationForTest() const {
    return playbackMediaGeneration_;
}

bool MediaExplorerWindow::setPlayerTimeForTest(qint64 milliseconds) {
    if (!vlcPlayer_ || stack_->currentWidget() != playerPage_) return false;
    const libvlc_time_t length = libvlc_media_player_get_length(vlcPlayer_);
    if (length <= 0) return false;
    libvlc_media_player_set_time(vlcPlayer_, qBound<libvlc_time_t>(0, milliseconds, length));
    return true;
}

QSize MediaExplorerWindow::videoSizeForTest() const {
    unsigned width = 0;
    unsigned height = 0;
    if (!vlcPlayer_ || libvlc_video_get_size(vlcPlayer_, 0, &width, &height) != 0) return {};
    return QSize(static_cast<int>(width), static_cast<int>(height));
}

void MediaExplorerWindow::stopPlaybackForTest() {
    stopPlayback();
}

void MediaExplorerWindow::playSelected() {
    playEntries(selectedEntries());
}

void MediaExplorerWindow::playEntries(const QVector<Entry> &entries) {
    if (fileOperationCancel_) {
        showError(QStringLiteral("Playback"),
                  QStringLiteral("Wait for the active file operation or cancel it before playing."));
        return;
    }
    QStringList paths;
    for (const Entry &entry : entries) {
        if (!entry.directory && isVideoFile(entry.path, config_.videoExtensions)) paths.append(entry.path);
    }
    if (paths.isEmpty()) {
        statusBar()->showMessage(QStringLiteral("Select one or more videos to play"), 5000);
        return;
    }
    if (!ensurePlayer()) return;
    if (!isMaximized()) showMaximized();
    looping_ = false;
    playlist_ = paths;
    playlistIndex_ = 0;
    deferredActions_.clear();
    playlistWidget_->clear();
    for (const QString &path : std::as_const(playlist_)) {
        auto *item = new QListWidgetItem(QFileInfo(path).fileName(), playlistWidget_);
        item->setToolTip(path);
    }
    stack_->setCurrentWidget(playerPage_);
    videoFrame_->setFocus(Qt::OtherFocusReason);
    playerTimer_->start();
    playIndex(0);
}

void MediaExplorerWindow::playIndex(int index, bool resetView) {
    if (!vlcPlayer_ || !vlcInstance_ || index < 0 || index >= playlist_.size()) return;
    const QString path = playlist_.at(index);
    libvlc_media_t *media = libvlc_media_new_path(vlcInstance_, QFile::encodeName(path).constData());
    if (!media) {
        showError(QStringLiteral("Playback"), QStringLiteral("libVLC could not open:\n%1").arg(path));
        return;
    }
    libvlc_media_t *oldMedia = currentMedia_;
    libvlc_media_player_set_media(vlcPlayer_, media);
    currentMedia_ = media;
    if (oldMedia) libvlc_media_release(oldMedia);
    ++playbackMediaGeneration_;
    playbackEndPending_ = false;
    playlistIndex_ = index;
    if (resetView) {
        zoom_ = 1.0;
        panOffset_ = {};
    }
    updateVideoGeometry();
    playlistWidget_->setCurrentRow(index);
    updatePlaybackTitle();
    embedVideo();
    if (libvlc_media_player_play(vlcPlayer_) == -1) {
        showError(QStringLiteral("Playback"), QStringLiteral("libVLC refused to start playback."));
        return;
    }
    if (playbackPauseHolds_ > 0 || playbackPauseRestorePending_) {
        pausePlayback(false);
        QTimer::singleShot(0, this, [this] {
            if (playbackPauseHolds_ > 0 || playbackPauseRestorePending_)
                pausePlayback(false);
        });
    } else {
        pauseButton_->setText(QStringLiteral("Pause"));
    }
    lastVlcState_ = -1;
    playbackErrorShown_ = false;
}

void MediaExplorerWindow::embedVideo() {
    if (!vlcPlayer_) return;
    videoFrame_->winId();
    libvlc_media_player_set_xwindow(vlcPlayer_, static_cast<uint32_t>(videoFrame_->winId()));
}

void MediaExplorerWindow::pausePlayback(bool userRequested) {
    if (vlcPlayer_ && stack_->currentWidget() == playerPage_) {
        if (userRequested && (playbackPauseHolds_ > 0 || playbackPauseRestorePending_))
            resumeAfterPlaybackPause_ = false;
        libvlc_media_player_set_pause(vlcPlayer_, 1);
        pauseButton_->setText(QStringLiteral("Resume"));
    }
}

void MediaExplorerWindow::resumePlayback(bool userRequested) {
    if (vlcPlayer_ && stack_->currentWidget() == playerPage_) {
        if (userRequested && (playbackPauseHolds_ > 0 || playbackPauseRestorePending_))
            resumeAfterPlaybackPause_ = true;
        if (playbackPauseHolds_ > 0 || playbackPauseRestorePending_) {
            libvlc_media_player_set_pause(vlcPlayer_, 1);
            pauseButton_->setText(QStringLiteral("Resume"));
            return;
        }
        libvlc_media_player_set_pause(vlcPlayer_, 0);
        pauseButton_->setText(QStringLiteral("Pause"));
    }
}

bool MediaExplorerWindow::beginPlaybackPauseHold() {
    if (!vlcPlayer_ || stack_->currentWidget() != playerPage_ || exitingPlayback_) return false;
    ++playbackPauseRestoreToken_;
    if (playbackPauseHolds_ == 0) {
        if (!playbackPauseRestorePending_) {
            const libvlc_state_t state = libvlc_media_player_get_state(vlcPlayer_);
            resumeAfterPlaybackPause_ = libvlc_media_player_is_playing(vlcPlayer_) > 0 ||
                                        state == libvlc_Playing || state == libvlc_Opening ||
                                        state == libvlc_Buffering;
        }
        playbackPauseRestorePending_ = false;
    }
    ++playbackPauseHolds_;
    pausePlayback(false);
    return true;
}

void MediaExplorerWindow::endPlaybackPauseHold() {
    if (playbackPauseHolds_ <= 0) return;
    --playbackPauseHolds_;
    if (playbackPauseHolds_ != 0) return;
    playbackPauseRestorePending_ = true;
    const quint64 token = ++playbackPauseRestoreToken_;
    QTimer::singleShot(0, this, [this, token] {
        if (token != playbackPauseRestoreToken_ || playbackPauseHolds_ != 0 ||
            !playbackPauseRestorePending_) return;
        playbackPauseRestorePending_ = false;
        const bool shouldResume = resumeAfterPlaybackPause_;
        resumeAfterPlaybackPause_ = false;
        if (vlcPlayer_ && libvlc_media_player_get_state(vlcPlayer_) == libvlc_Ended) {
            playbackEndPending_ = true;
            playbackEndGeneration_ = playbackMediaGeneration_;
        }
        if (playbackEndPending_) {
            processPendingPlaybackEnd();
        } else if (shouldResume && vlcPlayer_ && stack_->currentWidget() == playerPage_ &&
                   !exitingPlayback_) {
            resumePlayback(false);
        }
    });
}

void MediaExplorerWindow::updatePlaybackTitle() {
    if (!vlcPlayer_ || stack_->currentWidget() != playerPage_ || playlistIndex_ < 0 ||
        playlistIndex_ >= playlist_.size()) return;
    const QString prefix = playlist_.size() <= 1
        ? QStringLiteral("(Single File)")
        : QStringLiteral("(Play List %1 of %2)").arg(playlistIndex_ + 1).arg(playlist_.size());
    const QString fileName = QFileInfo(playlist_.at(playlistIndex_)).fileName();
    const libvlc_time_t current = qMax<libvlc_time_t>(0, libvlc_media_player_get_time(vlcPlayer_));
    const libvlc_time_t length = qMax<libvlc_time_t>(0, libvlc_media_player_get_length(vlcPlayer_));
    QString title = QStringLiteral("%1 %2  %3 / %4")
                        .arg(prefix, fileName, formatClock(current), formatClock(length));
    if (zoom_ > 1.001)
        title += QStringLiteral("  [Zoom %1%]").arg(qRound(zoom_ * 100.0));
    if (looping_ && !fullscreen_) title += QStringLiteral("  [Looping]");
    playerTitle_->setText(title);
    playerTitle_->setToolTip(playlist_.at(playlistIndex_));
    setWindowTitle(title);
}

void MediaExplorerWindow::toggleLooping() {
    if (stack_->currentWidget() != playerPage_) return;
    looping_ = !looping_;
    updatePlaybackTitle();
    statusBar()->showMessage(looping_ ? QStringLiteral("Looping current video")
                                      : QStringLiteral("Looping off"),
                             2500);
    if (vlcPlayer_ && libvlc_media_player_get_state(vlcPlayer_) == libvlc_Ended) {
        playbackEndPending_ = true;
        playbackEndGeneration_ = playbackMediaGeneration_;
        processPendingPlaybackEnd();
    }
}

int MediaExplorerWindow::playbackIndexAfterEnd(int currentIndex, int playlistSize,
                                               bool looping) {
    if (currentIndex < 0 || currentIndex >= playlistSize) return -1;
    if (looping) return currentIndex;
    return currentIndex + 1 < playlistSize ? currentIndex + 1 : -1;
}

void MediaExplorerWindow::processPendingPlaybackEnd() {
    if (!playbackEndPending_ || playbackPauseHolds_ > 0 || playbackPauseRestorePending_)
        return;
    const quint64 endedGeneration = playbackEndGeneration_;
    const int endedIndex = playlistIndex_;
    const QString endedPath = playlist_.value(endedIndex);
    playbackEndPending_ = false;
    QTimer::singleShot(0, this, [this, endedGeneration, endedIndex, endedPath] {
        if (endedGeneration != playbackMediaGeneration_ || !vlcPlayer_ ||
            stack_->currentWidget() != playerPage_ || exitingPlayback_ ||
            playlistIndex_ != endedIndex || playlist_.value(endedIndex) != endedPath ||
            libvlc_media_player_get_state(vlcPlayer_) != libvlc_Ended)
            return;
        if (playbackPauseHolds_ > 0 || playbackPauseRestorePending_) {
            playbackEndPending_ = true;
            playbackEndGeneration_ = endedGeneration;
            return;
        }

        // The loop state is deliberately read here, rather than when Ended was
        // observed, so a Ctrl+L received before this queued action takes effect.
        const int targetIndex = playbackIndexAfterEnd(endedIndex, playlist_.size(), looping_);
        if (targetIndex >= 0) {
            playIndex(targetIndex, targetIndex != endedIndex);
        } else {
            pauseButton_->setText(QStringLiteral("Replay"));
        }
    });
}

void MediaExplorerWindow::previousVideo() {
    if (stack_->currentWidget() == playerPage_ && playlistIndex_ > 0) playIndex(playlistIndex_ - 1);
}

void MediaExplorerWindow::nextVideo() {
    if (stack_->currentWidget() == playerPage_ && playlistIndex_ + 1 < playlist_.size())
        playIndex(playlistIndex_ + 1);
}

void MediaExplorerWindow::seekBy(qint64 milliseconds) {
    if (!vlcPlayer_ || stack_->currentWidget() != playerPage_) return;
    const libvlc_time_t current = qMax<libvlc_time_t>(0, libvlc_media_player_get_time(vlcPlayer_));
    const libvlc_time_t length = libvlc_media_player_get_length(vlcPlayer_);
    libvlc_time_t target = qMax<libvlc_time_t>(0, current + milliseconds);
    if (length > 0) target = qMin(target, length);
    libvlc_media_player_set_time(vlcPlayer_, target);
}

void MediaExplorerWindow::setVolume(int value) {
    if (vlcPlayer_) libvlc_audio_set_volume(vlcPlayer_, qBound(0, value, 200));
}

void MediaExplorerWindow::pollPlayer() {
    if (!vlcPlayer_ || stack_->currentWidget() != playerPage_) return;
    const libvlc_time_t current = libvlc_media_player_get_time(vlcPlayer_);
    const libvlc_time_t length = libvlc_media_player_get_length(vlcPlayer_);
    if (!seeking_) {
        const float position = libvlc_media_player_get_position(vlcPlayer_);
        if (position >= 0.0f) seekSlider_->setValue(qBound(0, qRound(position * 1000.0f), 1000));
    }
    timeLabel_->setText(QStringLiteral("%1 / %2").arg(formatClock(current), formatClock(length)));
    const libvlc_state_t state = libvlc_media_player_get_state(vlcPlayer_);
    if ((playbackPauseHolds_ > 0 || playbackPauseRestorePending_) && state == libvlc_Playing)
        pausePlayback(false);
    if (state == libvlc_Ended && lastVlcState_ != static_cast<int>(libvlc_Ended)) {
        playbackEndPending_ = true;
        playbackEndGeneration_ = playbackMediaGeneration_;
    } else if (state == libvlc_Error && !playbackErrorShown_) {
        playbackErrorShown_ = true;
        showError(QStringLiteral("Playback"), QStringLiteral("libVLC reported a playback error."));
    }
    processPendingPlaybackEnd();
    updatePlaybackTitle();
    lastVlcState_ = static_cast<int>(state);
}

void MediaExplorerWindow::toggleFullscreen() {
    if (stack_->currentWidget() != playerPage_) return;
    if (!fullscreen_) wasMaximizedBeforeFullscreen_ = isMaximized();
    fullscreen_ = !fullscreen_;
    playerTitle_->setVisible(!fullscreen_);
    playlistWidget_->setVisible(!fullscreen_);
    seekSlider_->setVisible(!fullscreen_);
    playerControls_->setVisible(!fullscreen_);
    menuBar()->setVisible(!fullscreen_);
    statusBar()->setVisible(!fullscreen_);
    if (fullscreen_) showFullScreen();
    else if (wasMaximizedBeforeFullscreen_) showMaximized();
    else showNormal();
    updatePlaybackTitle();
    QTimer::singleShot(0, this, [this] { updateVideoGeometry(); embedVideo(); });
}

void MediaExplorerWindow::stopPlayback() {
    if (stack_->currentWidget() != playerPage_) return;
    exitingPlayback_ = true;
    playbackDialogs_.clear();
    playbackPauseHolds_ = 0;
    resumeAfterPlaybackPause_ = false;
    playbackPauseRestorePending_ = false;
    ++playbackPauseRestoreToken_;
    looping_ = false;
    playbackEndPending_ = false;
    ++playbackMediaGeneration_;
    if (vlcPlayer_) libvlc_media_player_stop(vlcPlayer_);
    playerTimer_->stop();
    if (fullscreen_) {
        fullscreen_ = false;
        playerTitle_->show();
        playlistWidget_->show();
        seekSlider_->show();
        playerControls_->show();
        menuBar()->show();
        statusBar()->show();
        if (wasMaximizedBeforeFullscreen_) showMaximized();
        else showNormal();
    }
    stack_->setCurrentWidget(browserPage_);
    setWindowTitle(QString::fromLatin1(kApplicationName));
    seekSlider_->setValue(0);
    timeLabel_->setText(QStringLiteral("0:00 / 0:00"));
    publishPendingMediaOutputs();
    if (!deferredActions_.isEmpty() && !fileOperationCancel_ && !mediaProcessCancel_) {
        const QVector<DeferredAction> deferred = deferredActions_;
        deferredActions_.clear();
        runFileOperation(QStringLiteral("post-playback"), {}, {}, deferred);
    } else if (deferredActions_.isEmpty()) {
        refresh();
    } else {
        statusBar()->showMessage(QStringLiteral("Deferred playback actions are waiting for the active background operation."));
    }
    exitingPlayback_ = false;
}

void MediaExplorerWindow::removeCurrentDeferred() {
    if (playlistIndex_ < 0 || playlistIndex_ >= playlist_.size()) return;
    const QString path = playlist_.at(playlistIndex_);
    const bool trash = config_.useTrash && !executableResolution(QStringLiteral("gio")).isEmpty();
    const QString question = trash
        ? QStringLiteral("Move this video to the Trash when playback exits?\n\n%1").arg(path)
        : QStringLiteral("Permanently delete this video when playback exits?\n\n%1\n\nThis cannot be undone.")
              .arg(path);
    if (QMessageBox::warning(this, QStringLiteral("Delete after playback"), question,
                             QMessageBox::Yes | QMessageBox::Cancel,
                             QMessageBox::Cancel) != QMessageBox::Yes) return;
    deferredActions_.append({DeferredKind::Delete, path, {}});
    playlist_.removeAt(playlistIndex_);
    delete playlistWidget_->takeItem(playlistIndex_);
    if (playlist_.isEmpty()) stopPlayback();
    else playIndex(qMin(playlistIndex_, playlist_.size() - 1));
}

void MediaExplorerWindow::renameCurrentDeferred() {
    if (playlistIndex_ < 0 || playlistIndex_ >= playlist_.size()) return;
    const QString source = playlist_.at(playlistIndex_);
    const QString destination = QFileDialog::getSaveFileName(
        this, QStringLiteral("Rename file after playback"), source,
        QStringLiteral("Video files (*.*)"));
    if (!destination.isEmpty() && destination != source) {
        if (pathLexists(destination))
            showError(QStringLiteral("Rename"), QStringLiteral("The destination already exists."));
        else {
            deferredActions_.append({DeferredKind::Rename, source, destination});
            statusBar()->showMessage(QStringLiteral("Rename queued until playback exits"), 5000);
        }
    }
}

void MediaExplorerWindow::copyCurrentDeferred() {
    if (playlistIndex_ < 0 || playlistIndex_ >= playlist_.size()) return;
    const QString source = playlist_.at(playlistIndex_);
    const QString destination = QFileDialog::getSaveFileName(
        this, QStringLiteral("Copy file after playback"), source,
        QStringLiteral("Video files (*.*)"));
    if (!destination.isEmpty() && destination != source) {
        if (pathLexists(destination))
            showError(QStringLiteral("Copy"), QStringLiteral("The destination already exists."));
        else {
            deferredActions_.append({DeferredKind::Copy, source, destination});
            statusBar()->showMessage(QStringLiteral("Copy queued until playback exits"), 5000);
        }
    }
}

void MediaExplorerWindow::adjustZoom(bool zoomIn) {
    if (stack_->currentWidget() != playerPage_) return;
    zoom_ = qBound(1.0, zoom_ + (zoomIn ? 0.25 : -0.25), 4.0);
    if (qFuzzyCompare(zoom_, 1.0)) panOffset_ = {};
    updateVideoGeometry();
    updatePlaybackTitle();
}

void MediaExplorerWindow::panVideo(int dx, int dy) {
    if (stack_->currentWidget() != playerPage_ || zoom_ <= 1.0) return;
    panOffset_ += QPoint(dx * 40, dy * 40);
    updateVideoGeometry();
}

void MediaExplorerWindow::centerVideoPan() {
    panOffset_ = {};
    updateVideoGeometry();
}

void MediaExplorerWindow::updateVideoGeometry() {
    if (!videoViewport_ || !videoFrame_) return;
    const QSize viewport = videoViewport_->size();
    QSize frame(qMax(1, qRound(viewport.width() * zoom_)),
                qMax(1, qRound(viewport.height() * zoom_)));
    const int maxX = qMax(0, (frame.width() - viewport.width()) / 2);
    const int maxY = qMax(0, (frame.height() - viewport.height()) / 2);
    panOffset_.setX(qBound(-maxX, panOffset_.x(), maxX));
    panOffset_.setY(qBound(-maxY, panOffset_.y(), maxY));
    const QPoint centered((viewport.width() - frame.width()) / 2,
                          (viewport.height() - frame.height()) / 2);
    videoFrame_->setGeometry(QRect(centered + panOffset_, frame));
}

void MediaExplorerWindow::showPlaylistChooser() {
    if (playlist_.isEmpty()) return;
    QDialog dialog(this);
    dialog.setWindowTitle(QStringLiteral("Playlist"));
    dialog.resize(520, 420);
    auto *layout = new QVBoxLayout(&dialog);
    auto *label = new QLabel(QStringLiteral("Up/Down previews immediately; Enter or Escape closes."), &dialog);
    auto *list = new QListWidget(&dialog);
    for (int index = 0; index < playlist_.size(); ++index) {
        auto *item = new QListWidgetItem(
            QStringLiteral("%1. %2").arg(index + 1).arg(QFileInfo(playlist_.at(index)).fileName()), list);
        item->setToolTip(playlist_.at(index));
    }
    list->setCurrentRow(qMax(0, playlistIndex_));
    connect(list, &QListWidget::currentRowChanged, &dialog, [this](int row) {
        if (row >= 0 && row < playlist_.size()) playIndex(row);
    });
    connect(list, &QListWidget::itemActivated, &dialog, [&dialog](QListWidgetItem *) {
        dialog.accept();
    });
    auto *buttons = new QDialogButtonBox(QDialogButtonBox::Close, &dialog);
    connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);
    connect(buttons, &QDialogButtonBox::accepted, &dialog, &QDialog::accept);
    layout->addWidget(label);
    layout->addWidget(list, 1);
    layout->addWidget(buttons);
    list->setFocus();
    dialog.exec();
}

void MediaExplorerWindow::showVideoTools() {
    if (stack_->currentWidget() != playerPage_ || playlistIndex_ < 0 ||
        playlistIndex_ >= playlist_.size()) {
        showError(QStringLiteral("Video tools"), QStringLiteral("Play a video first."));
        return;
    }
    const bool canUpscale = !config_.upscaleDirectory.isEmpty();
    const bool canFfmpeg = config_.ffmpegAvailable;
    if (!canUpscale && !canFfmpeg) {
        showError(QStringLiteral("Video tools"),
                  QStringLiteral("No video tools are configured. Enable ffmpeg or set upscale_directory."));
        return;
    }
    QDialog dialog(this);
    dialog.setWindowTitle(QStringLiteral("Video tools"));
    auto *layout = new QVBoxLayout(&dialog);
    layout->addWidget(new QLabel(QStringLiteral("Choose an action (press 1–4, or Escape to cancel):"), &dialog));
    int choice = 0;
    auto addChoice = [&dialog, layout, &choice](int number, const QString &label) {
        auto *button = new QPushButton(QStringLiteral("%1  %2").arg(number).arg(label), &dialog);
        QObject::connect(button, &QPushButton::clicked, &dialog, [&dialog, &choice, number] {
            choice = number;
            dialog.accept();
        });
        auto *shortcut = new QShortcut(QKeySequence(number == 1 ? Qt::Key_1
                                                    : number == 2 ? Qt::Key_2
                                                    : number == 3 ? Qt::Key_3 : Qt::Key_4), &dialog);
        QObject::connect(shortcut, &QShortcut::activated, button, &QPushButton::click);
        layout->addWidget(button);
    };
    if (canUpscale) addChoice(1, QStringLiteral("Submit for upscaling"));
    if (canFfmpeg) {
        addChoice(2, QStringLiteral("Trim everything before the current position"));
        addChoice(3, QStringLiteral("Trim everything after the current position"));
        addChoice(4, QStringLiteral("Create a horizontally flipped copy"));
    }
    dialog.exec();
    if (choice == 1 && canUpscale) {
        const QString source = playlist_.at(playlistIndex_);
        const QString requested = QDir(config_.upscaleDirectory).filePath(QFileInfo(source).fileName());
        const QString destination = MediaTools::uniqueOutputPath(requested, {});
        if (destination.isEmpty()) showError(QStringLiteral("Submit for upscaling"),
                                              QStringLiteral("Could not choose a destination name."));
        else {
            deferredActions_.append({DeferredKind::Copy, source, destination});
            QMessageBox::information(this, QStringLiteral("Submit for upscaling"),
                                     QStringLiteral("The video will be copied after playback exits:\n%1")
                                         .arg(destination));
        }
    } else if (choice == 2 && canFfmpeg) {
        startFfmpegTool(FfmpegOperation::TrimFront);
    } else if (choice == 3 && canFfmpeg) {
        startFfmpegTool(FfmpegOperation::TrimEnd);
    } else if (choice == 4 && canFfmpeg) {
        startFfmpegTool(FfmpegOperation::HorizontalFlip);
    }
}

void MediaExplorerWindow::startFfmpegTool(FfmpegOperation operation) {
    if (!config_.ffmpegAvailable) {
        showError(QStringLiteral("FFmpeg tools"), QStringLiteral("ffmpeg is disabled in the configuration."));
        return;
    }
    if (mediaProcessCancel_) {
        showError(QStringLiteral("FFmpeg tools"), QStringLiteral("Another media operation is running."));
        return;
    }
    if (playlistIndex_ < 0 || playlistIndex_ >= playlist_.size() || !vlcPlayer_) return;
    const QString program = executableResolution(config_.ffmpegPath);
    if (program.isEmpty()) {
        showError(QStringLiteral("FFmpeg tools"), QStringLiteral("The configured ffmpeg executable was not found."));
        return;
    }
    const QString source = playlist_.at(playlistIndex_);
    const qint64 position = qMax<libvlc_time_t>(0, libvlc_media_player_get_time(vlcPlayer_));
    const qint64 length = qMax<libvlc_time_t>(0, libvlc_media_player_get_length(vlcPlayer_));
    QString tag;
    QString title;
    QStringList toolArguments;
    if (operation == FfmpegOperation::TrimFront) {
        tag = QStringLiteral("_trimfront");
        title = QStringLiteral("Trim front");
    } else if (operation == FfmpegOperation::TrimEnd) {
        tag = QStringLiteral("_trimend");
        title = QStringLiteral("Trim end");
    } else {
        tag = QStringLiteral("_hflip");
        title = QStringLiteral("Horizontal flip");
    }
    if (operation != FfmpegOperation::HorizontalFlip && position <= 0) {
        showError(title, QStringLiteral("Seek beyond the start of the video before trimming."));
        return;
    }
    const QString output = MediaTools::uniqueOutputPath(source, tag);
    if (output.isEmpty()) {
        showError(title, QStringLiteral("Could not choose a collision-free output name."));
        return;
    }
    switch (operation) {
    case FfmpegOperation::TrimFront:
        toolArguments = MediaTools::trimFrontArguments(source, output, position); break;
    case FfmpegOperation::TrimEnd:
        toolArguments = MediaTools::trimEndArguments(source, output, position); break;
    case FfmpegOperation::HorizontalFlip:
        toolArguments = MediaTools::horizontalFlipArguments(source, output); break;
    }
    QStringList arguments = config_.ffmpegArgs;
    arguments.append(toolArguments);
    const double expected = operation == FfmpegOperation::TrimFront
                                ? qMax<qint64>(0, length - position) / 1000.0
                                : operation == FfmpegOperation::TrimEnd
                                      ? position / 1000.0 : length / 1000.0;
    startMediaProcess(title, program, arguments, output, expected, {}, {}, {}, true);
}

void MediaExplorerWindow::combineSelected() {
    if (viewMode_ != ViewMode::Folder && viewMode_ != ViewMode::Search) {
        showError(QStringLiteral("Combine videos"), QStringLiteral("Open a folder or search results first."));
        return;
    }
    if (!config_.videoCombineAvailable || !config_.ffmpegAvailable) {
        showError(QStringLiteral("Combine videos"),
                  QStringLiteral("Video combine or ffmpeg is disabled in the configuration."));
        return;
    }
    if (mediaProcessCancel_ || fileOperationCancel_) {
        showError(QStringLiteral("Combine videos"), QStringLiteral("Another background operation is running."));
        return;
    }
    if (!ensureCombineLock(true)) return;
    const QStringList sources = selectedVideoPaths();
    if (sources.size() < 2) {
        showError(QStringLiteral("Combine videos"), QStringLiteral("Select at least two videos."));
        return;
    }
    const QString program = executableResolution(config_.ffmpegPath);
    if (program.isEmpty()) {
        showError(QStringLiteral("Combine videos"), QStringLiteral("The configured ffmpeg executable was not found."));
        return;
    }
    const QString suggested = MediaTools::uniqueOutputPath(sources.first(), QStringLiteral("_combined"));
    const QString output = QFileDialog::getSaveFileName(
        this, QStringLiteral("Combined video output"), suggested, QStringLiteral("Video files (*.*)"));
    if (output.isEmpty()) return;
    if (pathLexists(output)) {
        showError(QStringLiteral("Combine videos"), QStringLiteral("The selected output already exists."));
        return;
    }
    QTemporaryDir temporary(QDir(QFileInfo(output).absolutePath()).filePath(
        QStringLiteral(".media-explorer-combine-XXXXXX")));
    if (!temporary.isValid()) {
        showError(QStringLiteral("Combine videos"), QStringLiteral("Could not create a hidden per-job work directory."));
        return;
    }
    temporary.setAutoRemove(false);
    const QString workingDirectory = temporary.path();
    const QString listPath = QDir(workingDirectory).filePath(QStringLiteral("inputs.ffconcat"));
    QString concatError;
    const QString concat = MediaTools::ffconcatContent(sources, &concatError);
    QSaveFile listFile(listPath);
    if (concat.isEmpty() || !listFile.open(QIODevice::WriteOnly | QIODevice::Text) ||
        listFile.write(concat.toUtf8()) != concat.toUtf8().size() || !listFile.commit()) {
        QDir(workingDirectory).removeRecursively();
        showError(QStringLiteral("Combine videos"), concatError.isEmpty()
                                                     ? QStringLiteral("Could not write the temporary concat list.")
                                                     : concatError);
        return;
    }

    CombineJobState::Job job;
    job.title = QStringLiteral("Combine videos");
    job.outputPath = QFileInfo(output).absoluteFilePath();
    job.workingDirectory = QFileInfo(workingDirectory).absoluteFilePath();
    job.listPath = QFileInfo(listPath).absoluteFilePath();
    const QString outputSuffix = QFileInfo(output).suffix();
    job.encodedPath = QDir(job.workingDirectory).filePath(
        QStringLiteral("encoded-output") +
        (outputSuffix.isEmpty() ? QStringLiteral(".mkv") : QLatin1Char('.') + outputSuffix));
    for (const QString &source : sources) job.sources.append(QFileInfo(source).absoluteFilePath());
    QString manifestError;
    if (!combineStore_.create(job, &manifestError)) {
        if (job.manifestPath.isEmpty())
            cleanupOwnedWorkDirectory(workingDirectory, QFileInfo(output).absolutePath());
        showError(QStringLiteral("Combine videos"),
                  QStringLiteral("Could not persist the pending combine job: %1%2")
                      .arg(manifestError,
                           job.manifestPath.isEmpty()
                               ? QString()
                               : QStringLiteral("\n\nThe committed manifest and work directory were retained for recovery.")));
        return;
    }
    beginCombineJob(job);
}

bool MediaExplorerWindow::ensureCombineLock(bool showDialog) {
    if (combineLock_ && combineLock_->isLocked()) return true;
    QString error;
    if (!combineStore_.initialize(&error)) {
        if (showDialog) showError(QStringLiteral("Combine recovery"), error);
        else statusBar()->showMessage(QStringLiteral("Combine recovery disabled: %1").arg(error), 12000);
        return false;
    }
    auto lock = std::make_unique<QLockFile>(
        QDir(combineStore_.applicationStateDirectory()).filePath(QStringLiteral("combine.lock")));
    lock->setStaleLockTime(30000);
    if (!lock->tryLock(0)) {
        const QString message = QStringLiteral("Another Media Explorer instance owns combine-job recovery. "
                                               "This instance will not create or resume combine jobs.");
        if (showDialog) showError(QStringLiteral("Combine videos"), message);
        else statusBar()->showMessage(message, 12000);
        return false;
    }
    combineLock_ = std::move(lock);
    return true;
}

void MediaExplorerWindow::beginCombineJob(CombineJobState::Job job) {
    if (mediaProcessCancel_ || fileOperationCancel_ || hasActiveCombineJob_) {
        combineResumeQueue_.prepend(job);
        return;
    }
    hasActiveCombineJob_ = true;
    activeCombineJob_ = job;
    mediaProcessCancel_ = std::make_shared<std::atomic_bool>(false);
    const auto cancel = mediaProcessCancel_;
    setBusy(QStringLiteral("media"), QStringLiteral("Checking pending combine inputs..."), cancel,
            job.sources.size());
    const AppConfig config = config_;
    QPointer<MediaExplorerWindow> guard(this);
    workerPool_.start(new LambdaRunnable([guard, job, config, cancel] {
        QString error;
        if (job.stage == CombineJobState::Stage::Publishing) {
            bool encodedExists = false;
            bool publishedExists = false;
            if (!ownedRegularFileState(job.encodedPath, encodedExists, error)) {
                // Preserve the manifest and all paths for explicit inspection.
            } else if (encodedExists && !combineWorkspaceIsOwned(job, error)) {
                // Retain everything when the encoded file's parent is unsafe.
            } else if (encodedExists) {
                const QFileInfo recoveredInfo(job.encodedPath);
                Entry recoveredEntry;
                recoveredEntry.path = job.encodedPath;
                recoveredEntry.size = recoveredInfo.size();
                recoveredEntry.modifiedMs = recoveredInfo.lastModified().toMSecsSinceEpoch();
                const Metadata recovered = probeMetadata(recoveredEntry, config, cancel);
                if (!recovered.error.isEmpty() || recovered.width <= 0 || recovered.height <= 0) {
                    error = QStringLiteral("Recovered encoded output failed ffprobe verification: %1")
                                .arg(recovered.error.isEmpty()
                                         ? QStringLiteral("video dimensions unavailable")
                                         : recovered.error);
                } else {
                    if (!guard) return;
                    QMetaObject::invokeMethod(guard, [guard, job] {
                        if (!guard || !guard->hasActiveCombineJob_ ||
                            guard->activeCombineJob_.id != job.id) return;
                        guard->mediaProcessFinished(job.title, job.outputPath, job.encodedPath,
                                                    job.workingDirectory, job.workingDirectory,
                                                    false, true, false,
                                                    QStringLiteral("Recovered verified combine output"));
                    }, Qt::QueuedConnection);
                    return;
                }
            } else if (!ownedRegularFileState(job.publishPath, publishedExists, error)) {
                // With no encoded file, only a verified owned publication can
                // prove that the crash happened after the atomic rename.
            } else if (publishedExists) {
                const QFileInfo recoveredInfo(job.publishPath);
                Entry recoveredEntry;
                recoveredEntry.path = job.publishPath;
                recoveredEntry.size = recoveredInfo.size();
                recoveredEntry.modifiedMs = recoveredInfo.lastModified().toMSecsSinceEpoch();
                const Metadata recovered = probeMetadata(recoveredEntry, config, cancel);
                if (!recovered.error.isEmpty() || recovered.width <= 0 || recovered.height <= 0) {
                    error = QStringLiteral("Recovered published output failed ffprobe verification: %1")
                                .arg(recovered.error.isEmpty()
                                         ? QStringLiteral("video dimensions unavailable")
                                         : recovered.error);
                } else {
                    if (!guard) return;
                    QMetaObject::invokeMethod(guard, [guard, job] {
                        if (!guard || !guard->hasActiveCombineJob_ ||
                            guard->activeCombineJob_.id != job.id) return;
                        guard->mediaProcessCancel_.reset();
                        guard->finishBusy(QStringLiteral("media"),
                                          QStringLiteral("Recovered completed combine: %1")
                                              .arg(job.publishPath));
                        guard->finishActiveCombineJob(true, false);
                        if (guard->stack_->currentWidget() == guard->browserPage_)
                            guard->refresh();
                    }, Qt::QueuedConnection);
                    return;
                }
            } else {
                error = QStringLiteral("Neither the verified encoded file nor the recorded published output exists.");
            }
        } else if (!combineWorkspaceIsOwned(job, error)) {
            // The callback below records the job as failed without touching an
            // untrusted replacement of the manifest-owned work directory.
        }

        int width = -1;
        int height = -1;
        double duration = 0.0;
        for (int index = 0; error.isEmpty() && index < job.sources.size() && !cancel->load(); ++index) {
            const QFileInfo info(job.sources.at(index));
            struct stat sourceStatus {};
            if (::lstat(QFile::encodeName(job.sources.at(index)).constData(), &sourceStatus) != 0 ||
                !S_ISREG(sourceStatus.st_mode) || S_ISLNK(sourceStatus.st_mode)) {
                error = QStringLiteral("Combine source is missing or is not a regular file: %1")
                            .arg(job.sources.at(index));
                break;
            }
            Entry entry;
            entry.path = job.sources.at(index);
            entry.size = info.size();
            entry.modifiedMs = info.lastModified().toMSecsSinceEpoch();
            const Metadata metadata = probeMetadata(entry, config, cancel);
            if (!metadata.error.isEmpty() || metadata.width <= 0 || metadata.height <= 0) {
                error = QStringLiteral("Could not inspect %1: %2")
                            .arg(entry.path, metadata.error.isEmpty()
                                                 ? QStringLiteral("resolution unavailable") : metadata.error);
                break;
            }
            if (width < 0) { width = metadata.width; height = metadata.height; }
            else if (width != metadata.width || height != metadata.height) {
                error = QStringLiteral("All videos must have exactly equal dimensions.\n"
                                       "Expected %1x%2, but %3 is %4x%5.")
                            .arg(width).arg(height).arg(entry.path).arg(metadata.width).arg(metadata.height);
                break;
            }
            if (metadata.duration > 0.0) duration += metadata.duration;
            if (guard) {
                QMetaObject::invokeMethod(guard, [guard, index, count = job.sources.size()] {
                    if (guard && guard->busyOwner_ == QStringLiteral("media")) {
                        guard->progress_->setRange(0, count);
                        guard->progress_->setValue(index + 1);
                    }
                }, Qt::QueuedConnection);
            }
        }
        if (error.isEmpty() && !cancel->load() &&
            (!rebuildCombineList(job, error) || !removeOwnedCombineEncodedFile(job, error))) {
            // Keep the first preparation error and retain the owned workspace.
        }
        if (!guard) return;
        const bool cancelled = cancel->load();
        QMetaObject::invokeMethod(guard, [guard, job, config, cancel, error, duration,
                                         width, height, cancelled]() mutable {
            if (!guard || guard->mediaProcessCancel_ != cancel || !guard->hasActiveCombineJob_ ||
                guard->activeCombineJob_.id != job.id) return;
            guard->mediaProcessCancel_.reset();
            if (cancelled || !error.isEmpty()) {
                if (cancelled) guard->finishBusy(QStringLiteral("media"), QStringLiteral("Combine cancelled"));
                else {
                    guard->finishBusy(QStringLiteral("media"), QStringLiteral("Combine validation failed"));
                    guard->showError(QStringLiteral("Combine videos"), error);
                }
                guard->finishActiveCombineJob(false, cancelled, error);
                return;
            }
            guard->activeCombineJob_.expectedWidth = width;
            guard->activeCombineJob_.expectedHeight = height;
            guard->activeCombineJob_.expectedDurationMs = qMax<qint64>(0, qRound64(duration * 1000.0));
            QString stateError;
            if (!guard->combineStore_.setStage(guard->activeCombineJob_,
                                               CombineJobState::Stage::InputsValidated,
                                               guard->activeCombineJob_.attempt, &stateError) ||
                !guard->combineStore_.setStage(guard->activeCombineJob_,
                                               CombineJobState::Stage::StreamCopyRunning,
                                               guard->activeCombineJob_.attempt + 1, &stateError)) {
                guard->finishBusy(QStringLiteral("media"), QStringLiteral("Combine state update failed"));
                guard->showError(QStringLiteral("Combine videos"), stateError);
                guard->finishActiveCombineJob(false, false, stateError);
                return;
            }
            const QString program = executableResolution(config.ffmpegPath);
            if (program.isEmpty()) {
                const QString toolError = QStringLiteral("The configured ffmpeg executable was not found.");
                guard->finishBusy(QStringLiteral("media"));
                guard->showError(QStringLiteral("Combine videos"), toolError);
                guard->finishActiveCombineJob(false, false, toolError);
                return;
            }
            QStringList primary = config.ffmpegArgs;
            primary.append(MediaTools::combineStreamCopyArguments(
                guard->activeCombineJob_.listPath, guard->activeCombineJob_.outputPath));
            QStringList fallback = config.ffmpegArgs;
            fallback.append(MediaTools::combineX264AacArguments(
                guard->activeCombineJob_.listPath, guard->activeCombineJob_.outputPath));
            guard->finishBusy(QStringLiteral("media"));
            guard->startMediaProcess(QStringLiteral("Combine videos"), program, primary,
                                     guard->activeCombineJob_.outputPath,
                                     duration, program, fallback,
                                     guard->activeCombineJob_.workingDirectory, false,
                                     guard->activeCombineJob_.encodedPath);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::loadPendingCombineJobs() {
    if (!ensureCombineLock(false)) return;
    const CombineJobState::ScanResult scan = combineStore_.loadPending();
    QStringList recoveryErrors = scan.errors;
    QVector<CombineJobState::Job> failedJobs;
    combineResumeQueue_.clear();
    for (const CombineJobState::Job &job : scan.jobs) {
        if (job.stage == CombineJobState::Stage::Completed) {
            QString removeError;
            if (!cleanupOwnedWorkDirectory(job.workingDirectory,
                                           QFileInfo(job.outputPath).absolutePath())) {
                recoveryErrors.append(QStringLiteral("Completed combine work directory could not be safely removed: %1")
                                          .arg(job.workingDirectory));
            } else if (!combineStore_.remove(job, &removeError)) {
                recoveryErrors.append(removeError);
            }
        } else if (job.stage == CombineJobState::Stage::Failed) {
            failedJobs.append(job);
        } else {
            combineResumeQueue_.append(job);
        }
    }
    if (!scan.quarantinedPaths.isEmpty())
        recoveryErrors.append(QStringLiteral("%1 malformed manifest(s) were quarantined")
                                  .arg(scan.quarantinedPaths.size()));
    if (!recoveryErrors.isEmpty())
        statusBar()->showMessage(QStringLiteral("Combine recovery: %1")
                                     .arg(recoveryErrors.join(QStringLiteral("; "))), 12000);
    if (!failedJobs.isEmpty()) {
        QMessageBox prompt(QMessageBox::Warning, QStringLiteral("Combine recovery"),
                           QStringLiteral("%1 failed combine job(s) were retained. Manifests remain under:\n%2\n\n"
                                          "Their app-owned work directories remain beside the requested outputs. "
                                          "Retry only if the earlier problem (for example a disconnected mount) is resolved.")
                               .arg(failedJobs.size()).arg(combineStore_.applicationStateDirectory()),
                           QMessageBox::NoButton, this);
        QPushButton *retry = prompt.addButton(QStringLiteral("Retry failed jobs"),
                                              QMessageBox::AcceptRole);
        prompt.addButton(QStringLiteral("Keep for inspection"), QMessageBox::RejectRole);
        prompt.exec();
        if (prompt.clickedButton() == retry) {
            for (const CombineJobState::Job &job : std::as_const(failedJobs))
                combineResumeQueue_.append(job);
        }
    }
    resumeNextCombineJob();
}

void MediaExplorerWindow::resumeNextCombineJob() {
    if (combineResumeQueue_.isEmpty() || hasActiveCombineJob_) return;
    if (mediaProcessCancel_ || fileOperationCancel_ || !busyOwner_.isEmpty()) {
        QTimer::singleShot(1000, this, [this] { resumeNextCombineJob(); });
        return;
    }
    const CombineJobState::Job job = combineResumeQueue_.takeFirst();
    statusBar()->showMessage(QStringLiteral("Resuming pending combine job %1").arg(job.id), 8000);
    beginCombineJob(job);
}

void MediaExplorerWindow::finishActiveCombineJob(bool published, bool cancelled,
                                                  const QString &error) {
    if (!hasActiveCombineJob_) return;
    QString stateError;
    if (published) {
        if (!combineStore_.setStage(activeCombineJob_, CombineJobState::Stage::Completed,
                                    qMax(1, activeCombineJob_.attempt), &stateError)) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("Output was published, but its manifest could not be completed: %1")
                          .arg(stateError));
        } else if (!cleanupOwnedWorkDirectory(activeCombineJob_.workingDirectory,
                                              QFileInfo(activeCombineJob_.outputPath).absolutePath())) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("Output was published, but its per-job work directory could not be safely removed. "
                                     "The completed manifest was retained for cleanup on the next launch."));
        } else if (!combineStore_.remove(activeCombineJob_, &stateError)) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("Output was published, but its pending manifest remains: %1")
                          .arg(stateError));
        }
    } else if (cancelled) {
        if (!combineStore_.setStage(activeCombineJob_, CombineJobState::Stage::Failed,
                                    activeCombineJob_.attempt, &stateError)) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("The cancelled combine manifest could not be made terminal: %1")
                          .arg(stateError));
        } else if (!cleanupOwnedWorkDirectory(activeCombineJob_.workingDirectory,
                                              QFileInfo(activeCombineJob_.outputPath).absolutePath())) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("The cancelled combine work directory was not removed because it failed safety validation. "
                                     "Its terminal manifest was retained."));
        } else if (!combineStore_.remove(activeCombineJob_, &stateError)) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("The cancelled combine work was removed, but its terminal manifest remains: %1")
                          .arg(stateError));
        }
        Q_UNUSED(error)
    } else {
        if (!combineStore_.setStage(activeCombineJob_, CombineJobState::Stage::Failed,
                                    activeCombineJob_.attempt, &stateError)) {
            showError(QStringLiteral("Combine recovery"),
                      QStringLiteral("The failed combine manifest could not be updated: %1")
                          .arg(stateError));
        }
        Q_UNUSED(error)
    }
    hasActiveCombineJob_ = false;
    activeCombineJob_ = {};
    QTimer::singleShot(0, this, [this] { resumeNextCombineJob(); });
}

bool MediaExplorerWindow::publishActiveCombineOutput(QString &publishedPath, QString &error) {
    publishedPath.clear();
    error.clear();
    if (!hasActiveCombineJob_) {
        error = QStringLiteral("No tracked combine job is active.");
        return false;
    }
    bool encodedExists = false;
    if (!ownedRegularFileState(activeCombineJob_.encodedPath, encodedExists, error) ||
        !encodedExists) {
        if (error.isEmpty()) error = QStringLiteral("The verified combine output is missing.");
        return false;
    }
    for (int attempt = 0; attempt < 100; ++attempt) {
        const QString candidate = MediaTools::uniqueOutputPath(activeCombineJob_.outputPath, {});
        if (candidate.isEmpty()) break;
        QString stateError;
        if (!combineStore_.setPublishPath(activeCombineJob_, candidate, &stateError) ||
            !combineStore_.setStage(activeCombineJob_, CombineJobState::Stage::Publishing,
                                    qMax(1, activeCombineJob_.attempt), &stateError)) {
            error = QStringLiteral("Could not persist the exact publication candidate: %1\n\n"
                                   "The verified encoded file was retained at:\n%2")
                        .arg(stateError, activeCombineJob_.encodedPath);
            return false;
        }
        if (renameNoReplace(activeCombineJob_.encodedPath, candidate) == 0) {
            publishedPath = candidate;
            return true;
        }
        if (errno == EEXIST) continue;
        error = QStringLiteral("Could not atomically publish the finished combine: %1\n\n"
                               "The verified encoded file was retained at:\n%2")
                    .arg(QString::fromLocal8Bit(std::strerror(errno)), activeCombineJob_.encodedPath);
        return false;
    }
    error = QStringLiteral("Could not find a collision-free output name. "
                           "The verified encoded file was retained at:\n%1")
                .arg(activeCombineJob_.encodedPath);
    return false;
}

void MediaExplorerWindow::startMediaProcess(const QString &title, const QString &program,
                                            const QStringList &arguments, const QString &outputPath,
                                            double expectedSeconds, const QString &fallbackProgram,
                                            const QStringList &fallbackArguments,
                                            const QString &temporaryPath, bool deferPublication,
                                            const QString &encodedPathOverride) {
    if (mediaProcessCancel_) {
        showError(title, QStringLiteral("Another media operation is already running."));
        return;
    }
    if (fileOperationCancel_) {
        showError(title, QStringLiteral("Wait for the active file operation before starting FFmpeg."));
        return;
    }
    if (activeMediaOutput_ == outputPath ||
        std::any_of(pendingMediaOutputs_.cbegin(), pendingMediaOutputs_.cend(),
                    [&outputPath](const PendingMediaOutput &pending) {
                        return pending.suggestedPath == outputPath;
                    })) {
        showError(title, QStringLiteral("That operation is already queued for this video."));
        return;
    }
    activeMediaOutput_ = outputPath;
    mediaProcessCancel_ = std::make_shared<std::atomic_bool>(false);
    const auto cancel = mediaProcessCancel_;
    setBusy(QStringLiteral("media"), QStringLiteral("%1: starting...").arg(title), cancel, 1000);
    QPointer<MediaExplorerWindow> guard(this);
    const AppConfig config = config_;
    workerPool_.start(new LambdaRunnable([guard, title, program, arguments, outputPath,
                                          expectedSeconds, fallbackProgram, fallbackArguments,
                                          temporaryPath, deferPublication, encodedPathOverride,
                                          cancel, config] {
        const QString outputParent = QFileInfo(outputPath).absolutePath();
        std::unique_ptr<QTemporaryDir> ownedWork;
        QString workingDirectory;
        QString encodedPath;
        if (!encodedPathOverride.isEmpty()) {
            workingDirectory = temporaryPath;
            encodedPath = encodedPathOverride;
        } else {
            ownedWork = std::make_unique<QTemporaryDir>(QDir(outputParent).filePath(
                QStringLiteral(".media-explorer-process-XXXXXX")));
            ownedWork->setAutoRemove(false);
            workingDirectory = ownedWork->isValid() ? ownedWork->path() : QString();
            encodedPath = workingDirectory.isEmpty()
                              ? QString()
                              : QDir(workingDirectory).filePath(QFileInfo(outputPath).fileName());
        }
        QStringList primaryArguments = arguments;
        QStringList secondaryArguments = fallbackArguments;
        auto replaceArgument = [](QStringList &items, const QString &from, const QString &to) {
            for (QString &item : items) if (item == from) item = to;
        };
        replaceArgument(primaryArguments, outputPath, encodedPath);
        replaceArgument(secondaryArguments, outputPath, encodedPath);
        QString isolatedSource;
        QString isolatedCopy;
        if (deferPublication && !workingDirectory.isEmpty()) {
            for (int index = 0; index + 1 < arguments.size(); ++index) {
                if (arguments.at(index) == QStringLiteral("-i") &&
                    !arguments.at(index + 1).endsWith(QStringLiteral(".ffconcat"))) {
                    isolatedSource = arguments.at(index + 1);
                    isolatedCopy = QDir(workingDirectory).filePath(QFileInfo(isolatedSource).fileName());
                    replaceArgument(primaryArguments, isolatedSource, isolatedCopy);
                    replaceArgument(secondaryArguments, isolatedSource, isolatedCopy);
                    break;
                }
            }
        }
        auto run = [guard, &title, expectedSeconds, &cancel](const QString &executable,
                                                             const QStringList &args,
                                                             QString &details) {
            QProcess process;
            process.setProgram(executable);
            process.setArguments(args);
            process.setProcessChannelMode(QProcess::SeparateChannels);
            process.start();
            if (!process.waitForStarted(8000)) {
                details = process.errorString();
                return false;
            }
            QByteArray diagnostics;
            static const QRegularExpression timePattern(
                QStringLiteral(R"(time=(\d+):(\d+):(\d+(?:\.\d+)?))"));
            while (!process.waitForFinished(150)) {
                const QByteArray block = process.readAllStandardError();
                diagnostics.append(block);
                if (diagnostics.size() > 1024 * 1024) diagnostics.remove(0, diagnostics.size() - 1024 * 1024);
                if (expectedSeconds > 0.0 && guard) {
                    const QString text = QString::fromUtf8(diagnostics);
                    QRegularExpressionMatchIterator matches = timePattern.globalMatch(text);
                    double seconds = -1.0;
                    while (matches.hasNext()) {
                        const QRegularExpressionMatch match = matches.next();
                        seconds = match.captured(1).toDouble() * 3600.0 +
                                  match.captured(2).toDouble() * 60.0 + match.captured(3).toDouble();
                    }
                    if (seconds >= 0.0) {
                        const int value = qBound(0, qRound(seconds / expectedSeconds * 1000.0), 1000);
                        QMetaObject::invokeMethod(guard, [guard, title, value] {
                            if (guard && guard->busyOwner_ == QStringLiteral("media")) {
                                guard->progress_->setRange(0, 1000);
                                guard->progress_->setValue(value);
                                guard->statusBar()->showMessage(QStringLiteral("%1: %2%")
                                                                   .arg(title).arg(value / 10.0, 0, 'f', 1));
                            }
                        }, Qt::QueuedConnection);
                    }
                }
                if (cancel->load()) {
                    process.terminate();
                    if (!process.waitForFinished(2000)) { process.kill(); process.waitForFinished(2000); }
                    diagnostics.append(process.readAllStandardError());
                    details = QString::fromUtf8(diagnostics).trimmed();
                    return false;
                }
            }
            diagnostics.append(process.readAllStandardError());
            details = QString::fromUtf8(diagnostics).trimmed();
            return process.exitStatus() == QProcess::NormalExit && process.exitCode() == 0;
        };

        QString details;
        bool success = !workingDirectory.isEmpty();
        if (!success) details = QStringLiteral("Could not create a hidden per-job work directory beside the output.");
        if (success && !isolatedSource.isEmpty()) {
            QString copyError;
            success = copyPathRecursive(isolatedSource, isolatedCopy, cancel, copyError);
            if (!success) details = QStringLiteral("Could not isolate the playing source: %1").arg(copyError);
        }
        if (success) success = run(program, primaryArguments, details);
        if (!success && !cancel->load() && !fallbackProgram.isEmpty()) {
            QFile::remove(encodedPath);
            QString fallbackDetails;
            bool fallbackAllowed = true;
            if (guard && !temporaryPath.isEmpty()) {
                QMetaObject::invokeMethod(guard, [guard, temporaryPath, &fallbackAllowed, &fallbackDetails] {
                    if (!guard || !guard->hasActiveCombineJob_ ||
                        guard->activeCombineJob_.workingDirectory != temporaryPath) return;
                    QString stateError;
                    fallbackAllowed = guard->combineStore_.setStage(
                        guard->activeCombineJob_, CombineJobState::Stage::TranscodeRunning,
                        qMax(1, guard->activeCombineJob_.attempt), &stateError);
                    if (!fallbackAllowed) fallbackDetails = stateError;
                }, Qt::BlockingQueuedConnection);
            }
            success = fallbackAllowed && run(fallbackProgram, secondaryArguments, fallbackDetails);
            if (!fallbackDetails.isEmpty()) details += QStringLiteral("\nFallback:\n") + fallbackDetails;
        }
        if (success) {
            const QFileInfo output(encodedPath);
            success = output.isFile() && output.size() > 0;
            if (!success) details += QStringLiteral("\nFFmpeg did not create a non-empty output file.");
        }
        if (success) {
            const QString ffprobe = config.ffprobeAvailable
                                        ? executableResolution(config.ffprobePath) : QString();
            if (ffprobe.isEmpty()) {
                success = false;
                details += QStringLiteral("\nffprobe is required to verify the encoded output.");
            } else {
                QProcess verify;
                QStringList verifyArguments = config.ffprobeArgs;
                verifyArguments.append(MediaTools::ffprobeVideoArguments(encodedPath));
                verify.start(ffprobe, verifyArguments);
                if (!verify.waitForStarted(5000)) success = false;
                int waited = 0;
                while (success && !verify.waitForFinished(100)) {
                    waited += 100;
                    if (cancel->load() || waited >= 30000) {
                        verify.kill();
                        verify.waitForFinished(2000);
                        success = false;
                    }
                }
                const QByteArray verification = verify.readAllStandardOutput();
                if (success && (verify.exitStatus() != QProcess::NormalExit || verify.exitCode() != 0 ||
                                !verification.contains("width=") || !verification.contains("height=")))
                    success = false;
                if (!success) details += QStringLiteral("\nThe encoded output failed ffprobe verification.");
            }
        }
        const bool cancelled = cancel->load();
        if (!success) {
            if (encodedPathOverride.isEmpty()) {
                QFile::remove(encodedPath);
                if (!workingDirectory.isEmpty()) cleanupOwnedWorkDirectory(workingDirectory, outputParent);
            }
        }
        if (!guard) return;
        QMetaObject::invokeMethod(guard, [guard, title, outputPath, encodedPath, workingDirectory,
                                         temporaryPath, deferPublication, success, cancelled, details] {
            if (guard) guard->mediaProcessFinished(title, outputPath, encodedPath,
                                                   workingDirectory, temporaryPath,
                                                   deferPublication, success, cancelled, details);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::mediaProcessFinished(const QString &title, const QString &outputPath,
                                               const QString &encodedPath,
                                               const QString &workingDirectory,
                                               const QString &temporaryPath,
                                               bool deferPublication, bool success,
                                               bool cancelled, const QString &details) {
    const bool trackedCombine = hasActiveCombineJob_ &&
                                activeCombineJob_.workingDirectory == temporaryPath;
    mediaProcessCancel_.reset();
    activeMediaOutput_.clear();
    if (trackedCombine) {
        if (success) {
            QString published;
            QString publishError;
            if (publishActiveCombineOutput(published, publishError)) {
                finishActiveCombineJob(true, false);
                finishBusy(QStringLiteral("media"),
                           QStringLiteral("%1 complete: %2").arg(title, published));
                QMessageBox::information(this, title, QStringLiteral("Created:\n%1").arg(published));
            } else {
                finishActiveCombineJob(false, false, publishError);
                finishBusy(QStringLiteral("media"), QStringLiteral("%1 publish failed").arg(title));
                showError(title, publishError);
            }
        } else if (cancelled) {
            finishActiveCombineJob(false, true, details);
            finishBusy(QStringLiteral("media"), QStringLiteral("%1 cancelled").arg(title));
        } else {
            finishActiveCombineJob(false, false, details);
            finishBusy(QStringLiteral("media"), QStringLiteral("%1 failed").arg(title));
            QString shown = details.trimmed();
            if (shown.size() > 6000) shown = shown.right(6000);
            showError(title, shown.isEmpty() ? QStringLiteral("The media operation failed.") : shown);
        }
        if (stack_->currentWidget() == browserPage_ && !fileOperationCancel_ && !mediaProcessCancel_)
            refresh();
        return;
    }
    if (!temporaryPath.isEmpty()) {
        cleanupOwnedWorkDirectory(temporaryPath, QFileInfo(outputPath).absolutePath());
    }
    if (success) {
        PendingMediaOutput pending{title, outputPath, encodedPath, workingDirectory,
                                   temporaryPath, details};
        if (deferPublication && stack_->currentWidget() == playerPage_) {
            pendingMediaOutputs_.append(pending);
            finishBusy(QStringLiteral("media"),
                       QStringLiteral("%1 encoded; exit playback to publish it").arg(title));
        } else {
            QString published;
            QString publishError;
            if (publishMediaOutput(pending, published, publishError)) {
                finishBusy(QStringLiteral("media"), QStringLiteral("%1 complete: %2").arg(title, published));
                QMessageBox::information(this, title, QStringLiteral("Created:\n%1").arg(published));
            } else {
                finishBusy(QStringLiteral("media"), QStringLiteral("%1 publish failed").arg(title));
                showError(title, publishError);
            }
        }
    } else if (cancelled) {
        finishBusy(QStringLiteral("media"), QStringLiteral("%1 cancelled").arg(title));
    } else {
        finishBusy(QStringLiteral("media"), QStringLiteral("%1 failed").arg(title));
        QString shown = details.trimmed();
        if (shown.size() > 6000) shown = shown.right(6000);
        showError(title, shown.isEmpty() ? QStringLiteral("The media operation failed.") : shown);
    }
    if (stack_->currentWidget() == browserPage_ && !deferredActions_.isEmpty() &&
        !fileOperationCancel_ && !mediaProcessCancel_) {
        const QVector<DeferredAction> deferred = deferredActions_;
        deferredActions_.clear();
        runFileOperation(QStringLiteral("post-playback"), {}, {}, deferred);
    } else if (stack_->currentWidget() == browserPage_ && !fileOperationCancel_) {
        refresh();
    }
}

bool MediaExplorerWindow::publishMediaOutput(PendingMediaOutput output,
                                             QString &publishedPath, QString &error) {
    publishedPath.clear();
    error.clear();
    for (int attempt = 0; attempt < 100; ++attempt) {
        const QString candidate = MediaTools::uniqueOutputPath(output.suggestedPath, {});
        if (candidate.isEmpty()) break;
        if (renameNoReplace(output.encodedPath, candidate) == 0) {
            publishedPath = candidate;
            cleanupOwnedWorkDirectory(output.workingDirectory,
                                      QFileInfo(output.suggestedPath).absolutePath());
            return true;
        }
        if (errno == EEXIST) continue;
        error = QStringLiteral("Could not atomically publish the finished output: %1\n\n"
                               "The verified encoded file was retained at:\n%2")
                    .arg(QString::fromLocal8Bit(std::strerror(errno)), output.encodedPath);
        return false;
    }
    error = QStringLiteral("Could not find a collision-free output name. "
                           "The verified encoded file was retained at:\n%1")
                .arg(output.encodedPath);
    return false;
}

void MediaExplorerWindow::publishPendingMediaOutputs() {
    if (pendingMediaOutputs_.isEmpty()) return;
    const QVector<PendingMediaOutput> pending = pendingMediaOutputs_;
    pendingMediaOutputs_.clear();
    QStringList published;
    QStringList failures;
    for (const PendingMediaOutput &output : pending) {
        QString path;
        QString error;
        if (publishMediaOutput(output, path, error)) published.append(path);
        else failures.append(error);
    }
    if (!failures.isEmpty()) showError(QStringLiteral("FFmpeg output publication"),
                                        failures.join(QStringLiteral("\n\n")));
    if (!published.isEmpty()) statusBar()->showMessage(
        QStringLiteral("Published %1 FFmpeg output(s)").arg(published.size()), 9000);
}

void MediaExplorerWindow::submitTopazQueue() {
    if (!fileChangesAllowed(QStringLiteral("Topaz queue"))) return;
    const QStringList sources = selectedVideoPaths();
    if (sources.isEmpty()) {
        showError(QStringLiteral("Topaz queue"), QStringLiteral("Select one or more videos."));
        return;
    }
    if (config_.topazUpscaleQueue.isEmpty()) {
        showError(QStringLiteral("Topaz queue"),
                  QStringLiteral("topaz_upscale_queue is not configured."));
        return;
    }
    QueueOptions options;
    if (promptTopazOptions(options)) startQueueSubmit(sources, options);
}

void MediaExplorerWindow::submitIw3Queue() {
    if (!fileChangesAllowed(QStringLiteral("IW3 queue"))) return;
    const QStringList sources = selectedVideoPaths();
    if (sources.isEmpty()) {
        showError(QStringLiteral("IW3 queue"), QStringLiteral("Select one or more videos."));
        return;
    }
    if (config_.topazUpscaleQueue.isEmpty()) {
        showError(QStringLiteral("IW3 queue"),
                  QStringLiteral("topaz_upscale_queue is not configured (both queue consumers use it)."));
        return;
    }
    QueueOptions options;
    options.iw3 = true;
    if (promptIw3Options(options)) startQueueSubmit(sources, options);
}

bool MediaExplorerWindow::promptTopazOptions(QueueOptions &options) {
    bool accepted = false;
    const QString target = QInputDialog::getItem(
        this, QStringLiteral("Topaz queue"), QStringLiteral("Target resolution:"),
        {QStringLiteral("4K (3840x2160)"), QStringLiteral("8K (7680x4320)")},
        0, false, &accepted);
    if (!accepted) return false;
    options.iw3 = false;
    options.target = target.startsWith(QStringLiteral("8K")) ? QStringLiteral("8k")
                                                              : QStringLiteral("4k");
    const QVector<MediaTools::TopazProfilePreset> presets = MediaTools::topazProfilePresets();
    QStringList labels;
    for (const auto &preset : presets) labels.append(preset.displayName);
    const QString label = QInputDialog::getItem(this, QStringLiteral("Topaz queue"),
                                                QStringLiteral("Enhancement profile:"), labels,
                                                0, false, &accepted);
    if (!accepted) return false;
    const int index = labels.indexOf(label);
    if (index < 0) return false;
    const auto preset = presets.at(index);
    options.profile = preset.schemaName;
    options.grain = preset.grain;
    options.grainSize = preset.grainSize;
    if (preset.profile == MediaTools::TopazProfile::GeneralGrain ||
        preset.profile == MediaTools::TopazProfile::RepairGrain) {
        options.grain = QInputDialog::getDouble(this, QStringLiteral("Topaz queue"),
                                                 QStringLiteral("Grain amount (0.0–1.0):"),
                                                 options.grain, 0.0, 1.0, 3, &accepted);
        if (!accepted) return false;
        options.grainSize = QInputDialog::getInt(this, QStringLiteral("Topaz queue"),
                                                  QStringLiteral("Grain size:"),
                                                  options.grainSize, 1, 5, 1, &accepted);
        if (!accepted) return false;
    }
    return true;
}

bool MediaExplorerWindow::promptIw3Options(QueueOptions &options) {
    const QVector<MediaTools::Iw3PresetModel> presets = MediaTools::iw3Presets();
    QStringList labels;
    for (const auto &preset : presets) labels.append(preset.displayName);
    bool accepted = false;
    const QString label = QInputDialog::getItem(this, QStringLiteral("IW3 queue"),
                                                QStringLiteral("3D conversion preset:"), labels,
                                                qMin(1, labels.size() - 1), false, &accepted);
    if (!accepted) return false;
    const int index = labels.indexOf(label);
    if (index < 0) return false;
    const MediaTools::Iw3Options preset = presets.at(index).options;
    options.iw3 = true;
    options.divergence = preset.divergence;
    options.method = preset.method;
    options.depthModel = preset.depthModel;
    options.tta = preset.tta;
    options.depthAa = preset.depthAA;
    return true;
}

void MediaExplorerWindow::startQueueSubmit(const QStringList &sources,
                                           const QueueOptions &options) {
    if (fileOperationCancel_ || mediaProcessCancel_) {
        showError(QStringLiteral("Queue submission"),
                  QStringLiteral("Another file or media operation is running."));
        return;
    }
    const QString queueDirectory = config_.topazUpscaleQueue;
    fileOperationCancel_ = std::make_shared<std::atomic_bool>(false);
    const auto cancel = fileOperationCancel_;
    setBusy(QStringLiteral("fileop"),
            options.iw3 ? QStringLiteral("Submitting to IW3 queue...")
                        : QStringLiteral("Submitting to Topaz queue..."),
            cancel, sources.size() * 1000);
    QPointer<MediaExplorerWindow> guard(this);
    workerPool_.start(new LambdaRunnable([guard, sources, options, queueDirectory, cancel] {
        QStringList errors;
        int completed = 0;
        MediaTools::TopazOptions topaz;
        MediaTools::Iw3Options iw3;
        if (!options.iw3) {
            topaz.target = options.target == QStringLiteral("8k")
                               ? MediaTools::TopazTarget::K8 : MediaTools::TopazTarget::K4;
            for (const auto &preset : MediaTools::topazProfilePresets()) {
                if (preset.schemaName == options.profile) { topaz.profile = preset.profile; break; }
            }
            topaz.grain = options.grain;
            topaz.grainSize = options.grainSize;
        } else {
            iw3.divergence = options.divergence;
            iw3.convergence = 0.5;
            iw3.method = options.method;
            iw3.depthModel = options.depthModel;
            iw3.tta = options.tta;
            iw3.depthAA = options.depthAa;
            iw3.emaNormalize = options.method != QStringLiteral("row_flow");
            iw3.sceneDetect = iw3.emaNormalize;
        }
        for (int index = 0; index < sources.size() && !cancel->load(); ++index) {
            const QString source = sources.at(index);
            const auto builder = options.iw3
                ? MediaTools::QueueJsonBuilder([iw3](const QString &original, const QString &queued) {
                      return MediaTools::buildIw3JobJson(original, queued, iw3);
                  })
                : MediaTools::QueueJsonBuilder([topaz](const QString &original, const QString &queued) {
                      return MediaTools::buildTopazJobJson(original, queued, topaz);
                  });
            const MediaTools::QueueSubmitResult result = MediaTools::submitQueueJob(
                source, queueDirectory, builder,
                [cancel] { return cancel->load(); },
                [guard, index, count = sources.size(), source](qint64 copied, qint64 total) {
                    if (!guard || total <= 0) return;
                    const int value = index * 1000 + qBound(0, qRound(copied * 1000.0 / total), 1000);
                    QMetaObject::invokeMethod(guard, [guard, value, count, source] {
                        if (guard && guard->busyOwner_ == QStringLiteral("fileop")) {
                            guard->progress_->setRange(0, count * 1000);
                            guard->progress_->setValue(value);
                            guard->statusBar()->showMessage(
                                QStringLiteral("Queueing %1").arg(QFileInfo(source).fileName()));
                        }
                    }, Qt::QueuedConnection);
                });
            if (result.succeeded()) ++completed;
            else if (result.status == MediaTools::QueueSubmitStatus::Failed) {
                QString message = QStringLiteral("%1: %2").arg(source, result.error);
                if (!result.retainedPaths.isEmpty())
                    message += QStringLiteral(" (retained: %1)").arg(result.retainedPaths.join(QStringLiteral(", ")));
                errors.append(message);
                break;
            }
        }
        if (!guard) return;
        const bool cancelled = cancel->load();
        QMetaObject::invokeMethod(guard, [guard, completed, total = sources.size(), errors, cancelled,
                                         iw3Mode = options.iw3] {
            if (guard) guard->fileOperationFinished(iw3Mode ? QStringLiteral("IW3 queue")
                                                            : QStringLiteral("Topaz queue"),
                                                     completed, total, errors, cancelled);
        }, Qt::QueuedConnection);
    }));
}

void MediaExplorerWindow::showHelp() {
    QDialog dialog(this);
    dialog.setWindowTitle(QStringLiteral("Media Explorer Help"));
    dialog.resize(780, 650);
    auto *layout = new QVBoxLayout(&dialog);
    auto *text = new QTextBrowser(&dialog);
    text->setOpenExternalLinks(true);
    text->setHtml(QStringLiteral(
        "<h2>Media Explorer</h2>"
        "<p>Open a mount or mapped drive, browse folders, search video names, and play files "
        "with embedded VLC. Folders always sort before videos.</p>"
        "<h3>Browser keyboard commands</h3>"
        "<table cellspacing='5'>"
        "<tr><td>Enter</td><td>Open folder or play selected videos</td></tr>"
        "<tr><td>Left / Backspace / Alt+Left</td><td>Parent folder; exit search</td></tr>"
        "<tr><td>Ctrl+A</td><td>Select all visible videos</td></tr>"
        "<tr><td>Ctrl+P</td><td>Play selected videos</td></tr>"
        "<tr><td>Ctrl+F</td><td>Recursive search; in results, append an AND keyword</td></tr>"
        "<tr><td>Ctrl+Up / Ctrl+Down</td><td>Reorder one selected playlist row</td></tr>"
        "<tr><td>Ctrl++</td><td>Combine selected videos</td></tr>"
        "<tr><td>Ctrl+U / Ctrl+3</td><td>Submit Topaz / IW3 queue jobs</td></tr>"
        "<tr><td>Ctrl+C / Ctrl+X / Ctrl+V</td><td>Copy / cut / paste</td></tr>"
        "<tr><td>Delete</td><td>Trash or delete after confirmation</td></tr>"
        "<tr><td>Escape</td><td>Cancel the active file operation only</td></tr>"
        "<tr><td>F2 / Ctrl+Shift+N / F5</td><td>Rename / new folder / refresh</td></tr>"
        "<tr><td>Ctrl+L / Ctrl+Home</td><td>Edit location / show mounts</td></tr>"
        "</table>"
        "<h3>Playback keyboard commands</h3>"
        "<table cellspacing='5'>"
        "<tr><td>Enter / Escape</td><td>Fullscreen / exit playback</td></tr>"
        "<tr><td>Space / Tab</td><td>Pause / resume (not toggle)</td></tr>"
        "<tr><td>Left / Right</td><td>Seek 10 seconds; hold Shift for 60 seconds</td></tr>"
        "<tr><td>Ctrl+Left / Ctrl+Right</td><td>Previous / next video</td></tr>"
        "<tr><td>Up / Down</td><td>Volume ±5 (0–200)</td></tr>"
        "<tr><td>Ctrl+G / Ctrl+P</td><td>Playlist chooser / video properties</td></tr>"
        "<tr><td>Ctrl+L</td><td>Toggle looping of the current video</td></tr>"
        "<tr><td>Ctrl+V</td><td>Video tools: 1 upscale, 2 trim front, 3 trim end, 4 flip</td></tr>"
        "<tr><td>Ctrl+R / Ctrl+C / Delete</td><td>Rename / copy / trash after playback</td></tr>"
        "<tr><td>Ctrl+Z / Ctrl+X</td><td>Zoom in / out</td></tr>"
        "<tr><td>Alt+Arrows / Alt+Home</td><td>Pan / recenter zoomed video</td></tr>"
        "</table>"
        "<p>Opening a dialog or popup menu during playback temporarily pauses the video. "
        "Closing the last popup restores the prior state, so a video that was already paused "
        "stays paused. In windowed playback, the top title shows the current file, time, and "
        "zoom percentage. When looping is enabled, <b>[Looping]</b> appears at the right end "
        "of that line except in fullscreen.</p>"
        "<p>FFmpeg encodes in a hidden per-job directory beside the source, verifies with ffprobe, "
        "and atomically publishes without replacing existing files. Playback edits publish only "
        "after playback exits. Queue JSON is published atomically after its video copy.</p>"));
    auto *buttons = new QDialogButtonBox(QDialogButtonBox::Close, &dialog);
    connect(buttons, &QDialogButtonBox::rejected, &dialog, &QDialog::reject);
    layout->addWidget(text, 1);
    layout->addWidget(buttons);
    dialog.exec();
}

void MediaExplorerWindow::showAbout() {
    QMessageBox::about(this, QStringLiteral("About Media Explorer"),
                       QStringLiteral("Media Explorer %1\n\n"
                                      "Native Qt 5 / C++17 Linux media browser with libVLC, "
                                      "ffmpeg, mapped-drive discovery, and queue integration.")
                           .arg(QString::fromLatin1(kVersion)));
}

void MediaExplorerWindow::showError(const QString &title, const QString &message) const {
    QMessageBox::critical(const_cast<MediaExplorerWindow *>(this), title, message);
}

void MediaExplorerWindow::closeEvent(QCloseEvent *event) {
    if (fileOperationCancel_ || mediaProcessCancel_) {
        event->ignore();
        QMessageBox::information(this, QStringLiteral("Background operation active"),
                                 QStringLiteral("Media Explorer cannot close while a file or media "
                                                "operation is active. Wait for it to finish, or use Cancel."));
        return;
    }
    if (stack_->currentWidget() == playerPage_) stopPlayback();
    if (fileOperationCancel_ || mediaProcessCancel_) {
        event->ignore();
        QMessageBox::information(this, QStringLiteral("Finishing playback actions"),
                                 QStringLiteral("Deferred playback actions must finish before closing."));
        return;
    }
    publishPendingMediaOutputs();
    if (scanCancel_) scanCancel_->store(true);
    if (searchCancel_) searchCancel_->store(true);
    invalidateMappingProbe();
    clearMetadataQueue();
    workerPool_.clear();
    event->accept();
}
