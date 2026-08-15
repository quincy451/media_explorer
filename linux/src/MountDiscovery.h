#pragma once

#include "AppConfig.h"

#include <QString>
#include <QVector>

enum class EntrySource {
    Folder,
    Search,
    Home,
    Filesystem,
    MappingStable,
    MappingGvfs,
    MappingAutomount,
    MappingDisconnected,
    Mount
};

struct Entry {
    QString name;
    QString path;
    bool directory = false;
    QString kind;
    qint64 size = 0;
    qint64 modifiedMs = 0;
    QString resolution;
    double durationSeconds = -1.0;
    EntrySource source = EntrySource::Folder;
    QString uri;
};

struct MountRecord {
    QString path;
    QString filesystem;
};

enum class MountState { Mounted, Automount, Unmounted };

struct MountProbeResult {
    QString path;
    MountState state = MountState::Unmounted;
    bool accessible = false;
    QString error;
};

QVector<MountRecord> readMountRecords();
MountState mappingMountState(const QString &path);
QString gvfsPathForUri(const QString &uri);
QVector<Entry> discoverMounts(const AppConfig &config);
MountProbeResult probeStableMount(const QString &path, int timeoutMs = 20000);
QString entrySourceName(EntrySource source);
