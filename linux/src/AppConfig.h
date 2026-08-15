#pragma once

#include <QHash>
#include <QSet>
#include <QString>
#include <QStringList>

struct AppConfig {
    QString mappedRoot = QStringLiteral("/mnt/media-explorer");
    QHash<QString, QString> mappedShares;
    QString startPath;
    bool ffprobeAvailable = true;
    QString ffprobePath = QStringLiteral("ffprobe");
    QStringList ffprobeArgs;
    bool ffmpegAvailable = true;
    QString ffmpegPath = QStringLiteral("ffmpeg");
    QStringList ffmpegArgs;
    bool videoCombineAvailable = true;
    QString upscaleDirectory;
    QString topazUpscaleQueue;
    QString loggingPath;
    QStringList vlcArgs{QStringLiteral("--no-video-title-show"), QStringLiteral("--quiet")};
    QSet<QString> videoExtensions{
        QStringLiteral(".mp4"), QStringLiteral(".mkv"), QStringLiteral(".mov"),
        QStringLiteral(".avi"), QStringLiteral(".wmv"), QStringLiteral(".m4v"),
        QStringLiteral(".ts"), QStringLiteral(".m2ts"), QStringLiteral(".webm"),
        QStringLiteral(".flv"), QStringLiteral(".rm")};
    bool showHidden = false;
    bool followSymlinks = false;
    int metadataPrefetchLimit = 500;
    bool useTrash = true;
    QString configPath;

    static bool load(const QString &explicitPath, AppConfig &config, QString &error);
};

QString expandPath(const QString &value);
QString normalizeConfigKey(QString key);
bool parseConfigBool(const QString &value, bool fallback);
bool isVideoFile(const QString &path, const QSet<QString> &extensions);
QString formatBytes(qint64 bytes);
QString formatDuration(double seconds);
QString formatClock(qint64 milliseconds);
QString fileKind(const QString &path);
QString executableResolution(const QString &command);
