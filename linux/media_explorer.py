#!/usr/bin/env python3
"""Media Explorer for Linux.

A PyQt5/libVLC desktop companion to the original Win32 Media Explorer.  The
program deliberately uses only the Python standard library beyond PyQt5 and
python-vlc so it can be deployed cleanly on Debian-family machines.
"""

from __future__ import annotations

import configparser
import datetime as _datetime
import json
import mimetypes
import os
import re
import shlex
import shutil
import subprocess
import sys
import sysconfig
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional, Sequence
from urllib.parse import unquote


# libVLC 3 embeds in an X11 window.  Debian's Wayland desktops normally make
# XWayland available; choosing xcb keeps embedding reliable.  A caller can
# explicitly override the platform.  SSH self-tests do not need a display.
_SELF_TEST_REQUESTED = "--self-test" in sys.argv or "--smoke-test" in sys.argv
if _SELF_TEST_REQUESTED:
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
elif sys.platform.startswith("linux"):
    os.environ.setdefault("QT_QPA_PLATFORM", "xcb")

# PyInstaller correctly collects libvlc.so, but VLC's plugin tree remains a
# system package on Linux.  A frozen process does not always discover that
# tree on its own (libvlc_new then returns NULL), so supply the distro path
# without overriding an administrator-provided value.
if sys.platform.startswith("linux") and not os.environ.get("VLC_PLUGIN_PATH"):
    _multiarch = sysconfig.get_config_var("MULTIARCH")
    _vlc_plugin_candidates = []
    if _multiarch:
        _vlc_plugin_candidates.append(Path("/usr/lib") / str(_multiarch) / "vlc" / "plugins")
    _vlc_plugin_candidates.extend((Path("/usr/lib/vlc/plugins"), Path("/usr/lib64/vlc/plugins")))
    for _candidate in _vlc_plugin_candidates:
        if _candidate.is_dir():
            os.environ["VLC_PLUGIN_PATH"] = os.fspath(_candidate)
            break

try:
    from PyQt5.QtCore import QObject, QRunnable, QThreadPool, QTimer, Qt, QUrl, pyqtSignal
    from PyQt5.QtGui import QColor, QDesktopServices, QKeySequence
    from PyQt5.QtWidgets import (
        QAbstractItemView,
        QAction,
        QApplication,
        QDialog,
        QDialogButtonBox,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QHeaderView,
        QInputDialog,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMenu,
        QMessageBox,
        QProgressBar,
        QPushButton,
        QShortcut,
        QSlider,
        QSplitter,
        QStackedWidget,
        QTableWidget,
        QTableWidgetItem,
        QTextBrowser,
        QVBoxLayout,
        QWidget,
    )
except ImportError as exc:  # pragma: no cover - useful diagnostic on target
    print(
        "Media Explorer requires PyQt5. On Debian run: "
        "sudo apt install python3-pyqt5",
        file=sys.stderr,
    )
    raise SystemExit(2) from exc

try:
    import vlc  # type: ignore
except Exception:
    vlc = None


APP_NAME = "Media Explorer"
APP_VERSION = "1.0.0-linux"
ORGANIZATION = "MediaExplorer"
VIDEO_EXTENSIONS = {
    ".mp4",
    ".mkv",
    ".mov",
    ".avi",
    ".wmv",
    ".m4v",
    ".ts",
    ".m2ts",
    ".webm",
    ".flv",
    ".rm",
}


def is_video_file(path: os.PathLike[str] | str, extensions: Iterable[str] = VIDEO_EXTENSIONS) -> bool:
    """Return whether *path* has a configured video suffix (case-insensitive)."""

    return Path(os.fspath(path)).suffix.casefold() in {str(ext).casefold() for ext in extensions}


def parse_bool(value: object, default: bool = False) -> bool:
    if value is None:
        return default
    text = str(value).strip().casefold()
    if text in {"1", "true", "yes", "on", "y", "enabled"}:
        return True
    if text in {"0", "false", "no", "off", "n", "disabled"}:
        return False
    return default


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


@dataclass
class AppConfig:
    mapped_root: Path = Path("/mnt/media-explorer")
    mapped_shares: dict[str, str] = field(default_factory=dict)
    start_path: str = ""
    ffprobe_available: bool = True
    ffprobe_path: str = "ffprobe"
    ffprobe_args: list[str] = field(default_factory=list)
    ffmpeg_available: bool = True
    ffmpeg_path: str = "ffmpeg"
    ffmpeg_args: list[str] = field(default_factory=list)
    vlc_args: list[str] = field(default_factory=lambda: ["--no-video-title-show", "--quiet"])
    video_extensions: set[str] = field(default_factory=lambda: set(VIDEO_EXTENSIONS))
    show_hidden: bool = False
    follow_symlinks: bool = False
    metadata_prefetch_limit: int = 500
    use_trash: bool = True
    config_path: Optional[Path] = None

    @classmethod
    def load(cls, explicit: Optional[str] = None) -> "AppConfig":
        config = cls()
        candidates: list[Path] = []
        if explicit is not None:
            explicit_path = Path(os.path.expandvars(explicit)).expanduser()
            if not explicit_path.is_file():
                raise ValueError(f"Configuration file does not exist or is not a file: {explicit_path}")
            candidates.append(explicit_path)
        else:
            env_path = os.environ.get("MEDIA_EXPLORER_CONFIG")
            if env_path is not None:
                configured_path = Path(os.path.expandvars(env_path)).expanduser()
                if not configured_path.is_file():
                    raise ValueError(
                        "MEDIA_EXPLORER_CONFIG does not name an existing file: "
                        f"{configured_path}"
                    )
                candidates.append(configured_path)
        candidates.extend(
            [
                Path.home() / ".config" / "media-explorer" / "mediaexplorer.ini",
                application_dir() / "mediaexplorer.ini",
            ]
        )
        selected = next((path for path in candidates if path.is_file()), None)
        if selected is None:
            config.ffprobe_available = shutil.which(config.ffprobe_path) is not None
            config.ffmpeg_available = shutil.which(config.ffmpeg_path) is not None
            return config

        parser = configparser.ConfigParser(interpolation=None)
        try:
            raw = selected.read_text(encoding="utf-8-sig")
            # Accept the sectionless key=value format used by the Win32 build.
            if not re.search(r"^\s*\[[^]]+\]", raw, flags=re.MULTILINE):
                raw = "[media_explorer]\n" + raw
            parser.read_string(raw, source=str(selected))
        except (OSError, UnicodeError, configparser.Error) as exc:
            raise ValueError(f"Cannot read configuration {selected}: {exc}") from exc

        values: dict[str, str] = {}
        for section in parser.sections():
            for key, value in parser.items(section):
                values[key.replace("_", "").replace("-", "").casefold()] = value.strip()

        def get(*names: str, default: str = "") -> str:
            for name in names:
                normalized = name.replace("_", "").replace("-", "").casefold()
                if normalized in values:
                    return values[normalized]
            return default

        mapped_root = get("mapped_root", "mapped_drive_root", default=str(config.mapped_root))
        config.mapped_root = Path(os.path.expandvars(os.path.expanduser(mapped_root)))
        # mapping_v = smb://server/share (also accepts mapped_drive_v).  The
        # directory /mnt/media-explorer/v is only considered connected when it
        # is an actual mount; an empty mount-point placeholder is not enough.
        for normalized_key, value in values.items():
            match = re.fullmatch(r"(?:mapping|mappeddrive|mappedshare)([a-z0-9_-]+)", normalized_key)
            if match and value:
                config.mapped_shares[match.group(1).strip("_-").casefold()] = value
        config.start_path = os.path.expandvars(os.path.expanduser(get("start_path")))
        config.ffprobe_path = os.path.expandvars(os.path.expanduser(get("ffprobe_path", default="ffprobe")))
        config.ffmpeg_path = os.path.expandvars(os.path.expanduser(get("ffmpeg_path", default="ffmpeg")))
        config.ffprobe_available = parse_bool(
            get("ffprobe_available", default=str(config.ffprobe_available)), config.ffprobe_available
        )
        config.ffmpeg_available = parse_bool(
            get("ffmpeg_available", default=str(config.ffmpeg_available)), config.ffmpeg_available
        )
        config.show_hidden = parse_bool(get("show_hidden"), config.show_hidden)
        config.follow_symlinks = parse_bool(get("follow_symlinks"), config.follow_symlinks)
        config.use_trash = parse_bool(get("use_trash"), config.use_trash)
        try:
            config.metadata_prefetch_limit = max(
                0, min(10000, int(get("metadata_prefetch_limit", default="500")))
            )
        except ValueError:
            config.metadata_prefetch_limit = 500

        def split_args(text: str, fallback: list[str]) -> list[str]:
            if not text:
                return fallback
            try:
                return shlex.split(text)
            except ValueError:
                return fallback

        config.ffprobe_args = split_args(get("ffprobe_args"), [])
        config.ffmpeg_args = split_args(get("ffmpeg_args"), [])
        config.vlc_args = split_args(get("vlc_args"), config.vlc_args)
        extensions = get("video_extensions")
        if extensions:
            parsed = {
                (item if item.startswith(".") else f".{item}").casefold()
                for item in re.split(r"[,;\s]+", extensions)
                if item.strip()
            }
            if parsed:
                config.video_extensions = parsed
        config.config_path = selected
        return config


@dataclass
class Entry:
    name: str
    path: str
    is_dir: bool
    kind: str
    size: int = 0
    mtime: float = 0.0
    mtime_ns: int = 0
    resolution: str = ""
    duration_seconds: Optional[float] = None
    source: str = "folder"
    uri: str = ""


def format_size(size: int) -> str:
    if size < 0:
        return ""
    if size < 1024:
        return f"{size} B"
    value = float(size)
    for unit in ("KiB", "MiB", "GiB", "TiB", "PiB"):
        value /= 1024.0
        if value < 1024.0 or unit == "PiB":
            return f"{value:.1f} {unit}"
    return ""


def format_duration(seconds: Optional[float]) -> str:
    if seconds is None or seconds < 0:
        return ""
    total = int(round(seconds))
    hours, rem = divmod(total, 3600)
    minutes, secs = divmod(rem, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{secs:02d}"
    return f"{minutes}:{secs:02d}"


def format_clock(milliseconds: int) -> str:
    if milliseconds < 0:
        milliseconds = 0
    return format_duration(milliseconds / 1000.0) or "0:00"


def format_modified(timestamp: float) -> str:
    if not timestamp:
        return ""
    try:
        return _datetime.datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M")
    except (OSError, OverflowError, ValueError):
        return ""


def file_kind(path: str) -> str:
    suffix = Path(path).suffix
    mime, _ = mimetypes.guess_type(path)
    if mime and mime.startswith("video/"):
        return f"{suffix[1:].upper()} video" if suffix else "Video"
    return f"{suffix[1:].upper()} file" if suffix else "File"


_OCTAL_ESCAPE = re.compile(r"\\([0-7]{3})")


def _unescape_mount_field(value: str) -> str:
    return _OCTAL_ESCAPE.sub(lambda match: chr(int(match.group(1), 8)), value)


PSEUDO_FILESYSTEMS = {
    "autofs",
    "bpf",
    "cgroup",
    "cgroup2",
    "configfs",
    "debugfs",
    "devpts",
    "devtmpfs",
    "efivarfs",
    "fusectl",
    "hugetlbfs",
    "mqueue",
    "proc",
    "pstore",
    "securityfs",
    "sysfs",
    "tracefs",
}
NETWORK_FILESYSTEMS = {"cifs", "smb3", "nfs", "nfs4", "fuse.sshfs", "sshfs", "davfs", "davfs2"}


def discover_mounts(config: AppConfig) -> list[Entry]:
    """Discover useful Linux roots and mapped shares without third-party modules."""

    found: dict[str, Entry] = {}

    def add(path: Path, name: str, kind: str, source: str, uri: str = "", require_directory: bool = True) -> None:
        normalized = os.path.abspath(os.fspath(path))
        if normalized in found or (require_directory and not os.path.isdir(normalized)):
            return
        found[normalized] = Entry(name, normalized, True, kind, source=source, uri=uri)

    add(Path.home(), "Home", "Home folder", "home")
    add(Path("/"), "Filesystem", "Filesystem", "filesystem")

    mount_info = Path("/proc/self/mountinfo")
    try:
        lines = mount_info.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        lines = []
    mount_records: list[tuple[str, str]] = []
    for line in lines:
        fields = line.split()
        if "-" not in fields or len(fields) < 7:
            continue
        separator = fields.index("-")
        if separator + 2 >= len(fields):
            continue
        mountpoint = _unescape_mount_field(fields[4])
        fs_type = fields[separator + 1]
        mount_records.append((os.path.normpath(mountpoint), fs_type))

    mounted_paths = {path for path, fs_type in mount_records if fs_type not in PSEUDO_FILESYSTEMS}

    def gvfs_path_for_uri(uri: str) -> Optional[Path]:
        url = QUrl(uri)
        if not url.isValid() or url.scheme().casefold() not in {"smb", "cifs"}:
            return None
        host = url.host().casefold()
        encoded_parts = url.path(QUrl.FullyEncoded).strip("/").split("/")
        if not host or not encoded_parts or not encoded_parts[0]:
            return None
        try:
            uri_parts = [unquote(part, errors="strict") for part in encoded_parts]
        except UnicodeDecodeError:
            return None
        if any(
            not part or part in {".", ".."} or "/" in part or "\\" in part or "\0" in part
            for part in uri_parts
        ):
            return None
        share = uri_parts[0].casefold()
        subpath = uri_parts[1:]
        try:
            gvfs_root = Path(f"/run/user/{os.getuid()}/gvfs")
            candidates = list(gvfs_root.iterdir()) if gvfs_root.is_dir() else []
        except OSError:
            candidates = []
        for candidate in candidates:
            prefix, separator, raw_fields = candidate.name.partition(":")
            if not separator or prefix.casefold() not in {"smb-share", "cifs-share"}:
                continue
            fields: dict[str, str] = {}
            try:
                for component in raw_fields.split(","):
                    key, assignment, value = component.partition("=")
                    if assignment:
                        fields[key.casefold()] = unquote(value, errors="strict").casefold()
            except UnicodeDecodeError:
                continue
            if fields.get("server") != host or fields.get("share") != share:
                continue
            resolved = candidate.joinpath(*subpath)
            try:
                if resolved.is_dir():
                    return resolved
            except OSError:
                continue
        return None

    # Named mappings have stable Windows-like labels.  Prefer their configured
    # mount point, then a desktop/GVFS mount created by opening the smb:// URI.
    mapped_root = config.mapped_root
    mapping_names: set[str] = set(config.mapped_shares)
    if mapped_root.is_dir():
        try:
            mapping_names.update(child.name.casefold() for child in mapped_root.iterdir() if child.is_dir())
        except OSError:
            pass
    for mapping_name in sorted(mapping_names):
        mountpoint = mapped_root / mapping_name
        uri = config.mapped_shares.get(mapping_name, "")
        normalized_mountpoint = os.path.normpath(os.path.abspath(os.fspath(mountpoint)))
        gvfs_path = gvfs_path_for_uri(uri) if uri else None
        display = f"{mapping_name.upper()}:" if len(mapping_name) <= 3 else mapping_name
        if normalized_mountpoint in mounted_paths:
            add(mountpoint, display, "Mapped drive", "mapped", uri)
        elif gvfs_path is not None:
            add(gvfs_path, display, "Mapped drive (GVFS)", "mapped", uri)
        else:
            add(
                mountpoint,
                f"{display} (disconnected)",
                "Mapped drive (disconnected)",
                "mapping-disconnected",
                uri,
                require_directory=False,
            )

    for mountpoint, fs_type in mount_records:
        if fs_type in PSEUDO_FILESYSTEMS or mountpoint == "/":
            continue
        interesting_prefix = mountpoint.startswith(("/mnt/", "/media/", "/run/media/"))
        if not interesting_prefix and fs_type not in NETWORK_FILESYSTEMS:
            continue
        label = Path(mountpoint).name or mountpoint
        kind = "Network mount" if fs_type in NETWORK_FILESYSTEMS else f"Mounted {fs_type}"
        add(Path(mountpoint), label, kind, "mount")

    priority = {"home": 0, "filesystem": 1, "mapped": 2, "mapping-disconnected": 3, "mount": 4}
    return sorted(found.values(), key=lambda entry: (priority.get(entry.source, 9), entry.name.casefold()))


class ScanSignals(QObject):
    finished = pyqtSignal(object)


class FolderScanTask(QRunnable):
    def __init__(self, path: str, generation: int, config: AppConfig, cancel: threading.Event):
        super().__init__()
        self.path = path
        self.generation = generation
        self.config = config
        self.cancel = cancel
        self.signals = ScanSignals()

    def run(self) -> None:
        entries: list[Entry] = []
        error = ""
        try:
            with os.scandir(self.path) as iterator:
                for item in iterator:
                    if self.cancel.is_set():
                        break
                    if not self.config.show_hidden and item.name.startswith("."):
                        continue
                    try:
                        is_dir = item.is_dir(follow_symlinks=self.config.follow_symlinks)
                        if not is_dir and not is_video_file(item.name, self.config.video_extensions):
                            continue
                        stat_result = item.stat(follow_symlinks=self.config.follow_symlinks)
                        entries.append(
                            Entry(
                                item.name,
                                item.path,
                                is_dir,
                                "Folder" if is_dir else file_kind(item.path),
                                0 if is_dir else stat_result.st_size,
                                stat_result.st_mtime,
                                getattr(stat_result, "st_mtime_ns", int(stat_result.st_mtime * 1e9)),
                            )
                        )
                    except (FileNotFoundError, PermissionError, OSError):
                        continue
        except (PermissionError, FileNotFoundError, NotADirectoryError, OSError) as exc:
            error = str(exc)
        self.signals.finished.emit(
            {
                "generation": self.generation,
                "path": self.path,
                "entries": entries,
                "cancelled": self.cancel.is_set(),
                "error": error,
            }
        )


class MetadataSignals(QObject):
    finished = pyqtSignal(object)


class MetadataTask(QRunnable):
    def __init__(self, entry: Entry, config: AppConfig, generation: int = 0):
        super().__init__()
        self.entry = entry
        self.config = config
        self.generation = generation
        self.signals = MetadataSignals()

    def run(self) -> None:
        result = {
            "generation": self.generation,
            "path": self.entry.path,
            "mtime_ns": self.entry.mtime_ns,
            "size": self.entry.size,
            "width": 0,
            "height": 0,
            "duration": None,
            "codec": "",
            "error": "",
        }
        if not self.config.ffprobe_available:
            result["error"] = "ffprobe is disabled"
            self.signals.finished.emit(result)
            return
        command = [
            self.config.ffprobe_path,
            *self.config.ffprobe_args,
            "-v",
            "error",
            "-select_streams",
            "v:0",
            "-show_entries",
            "stream=width,height,codec_name,duration:format=duration",
            "-of",
            "json",
            self.entry.path,
        ]
        try:
            completed = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=30,
                check=False,
            )
            if completed.returncode != 0:
                result["error"] = completed.stderr.strip() or f"ffprobe exited {completed.returncode}"
            else:
                payload = json.loads(completed.stdout or "{}")
                streams = payload.get("streams") or []
                if streams:
                    stream = streams[0]
                    result["width"] = int(stream.get("width") or 0)
                    result["height"] = int(stream.get("height") or 0)
                    result["codec"] = str(stream.get("codec_name") or "")
                    duration = stream.get("duration")
                    if duration not in (None, "N/A"):
                        result["duration"] = float(duration)
                if result["duration"] is None:
                    duration = (payload.get("format") or {}).get("duration")
                    if duration not in (None, "N/A"):
                        result["duration"] = float(duration)
        except (OSError, subprocess.SubprocessError, ValueError, TypeError, json.JSONDecodeError) as exc:
            result["error"] = str(exc)
        self.signals.finished.emit(result)


class SearchSignals(QObject):
    progress = pyqtSignal(str, int, int)
    finished = pyqtSignal(object)


def _skip_virtual_search_path(path: str, explicit_scopes: Sequence[str] = ()) -> bool:
    normalized = os.path.normpath(path)
    for prefix in ("/proc", "/sys", "/dev"):
        if normalized == prefix or normalized.startswith(prefix + os.sep):
            return True
    # /run/user is normally runtime noise, but GVFS network shares live there.
    # Permit a GVFS tree when the user explicitly selected that mapping.
    if normalized == "/run/user" or normalized.startswith("/run/user" + os.sep):
        if f"{os.sep}gvfs{os.sep}" in normalized or normalized.endswith(f"{os.sep}gvfs"):
            for scope in explicit_scopes:
                normalized_scope = os.path.normpath(scope)
                if normalized == normalized_scope or normalized.startswith(normalized_scope + os.sep):
                    return False
        return True
    for prefix in ("/run/lock",):
        if normalized == prefix or normalized.startswith(prefix + os.sep):
            return True
    return False


class SearchTask(QRunnable):
    def __init__(
        self,
        scopes: Sequence[str],
        terms: Sequence[str],
        config: AppConfig,
        cancel: threading.Event,
        token: int,
    ):
        super().__init__()
        self.scopes = list(scopes)
        self.terms = [term.casefold() for term in terms]
        self.config = config
        self.cancel = cancel
        self.token = token
        self.signals = SearchSignals()

    def run(self) -> None:
        results: list[Entry] = []
        errors: list[str] = []
        directories = 0
        files = 0
        stack = list(reversed(self.scopes))
        seen: set[tuple[int, int]] = set()
        while stack and not self.cancel.is_set():
            folder = stack.pop()
            if _skip_virtual_search_path(folder, self.scopes):
                continue
            directories += 1
            if directories == 1 or directories % 20 == 0:
                self.signals.progress.emit(folder, files, len(results))
            try:
                folder_stat = os.stat(folder, follow_symlinks=self.config.follow_symlinks)
                identity = (folder_stat.st_dev, folder_stat.st_ino)
                if identity in seen:
                    continue
                seen.add(identity)
                with os.scandir(folder) as iterator:
                    for item in iterator:
                        if self.cancel.is_set():
                            break
                        if not self.config.show_hidden and item.name.startswith("."):
                            continue
                        try:
                            if item.is_dir(follow_symlinks=self.config.follow_symlinks):
                                stack.append(item.path)
                                continue
                            if not is_video_file(item.name, self.config.video_extensions):
                                continue
                            files += 1
                            folded = item.name.casefold()
                            if not all(term in folded for term in self.terms):
                                continue
                            stat_result = item.stat(follow_symlinks=self.config.follow_symlinks)
                            results.append(
                                Entry(
                                    os.path.abspath(item.path),
                                    os.path.abspath(item.path),
                                    False,
                                    file_kind(item.path),
                                    stat_result.st_size,
                                    stat_result.st_mtime,
                                    getattr(stat_result, "st_mtime_ns", int(stat_result.st_mtime * 1e9)),
                                    source="search",
                                )
                            )
                        except (FileNotFoundError, PermissionError, OSError):
                            continue
            except (FileNotFoundError, PermissionError, NotADirectoryError, OSError) as exc:
                if len(errors) < 25:
                    errors.append(f"{folder}: {exc}")
        self.signals.finished.emit(
            {
                "token": self.token,
                "entries": results,
                "cancelled": self.cancel.is_set(),
                "directories": directories,
                "files": files,
                "errors": errors,
            }
        )


class FileOperationSignals(QObject):
    progress = pyqtSignal(str, int, int)
    finished = pyqtSignal(object)


def _available_destination(destination: str) -> str:
    if not os.path.lexists(destination):
        return destination
    parent = os.path.dirname(destination)
    name = os.path.basename(destination)
    stem, suffix = os.path.splitext(name)
    candidate = os.path.join(parent, f"{stem} (copy){suffix}")
    counter = 2
    while os.path.lexists(candidate):
        candidate = os.path.join(parent, f"{stem} ({counter}){suffix}")
        counter += 1
    return candidate


def _is_within(child: str, parent: str) -> bool:
    try:
        return os.path.commonpath([os.path.realpath(child), os.path.realpath(parent)]) == os.path.realpath(parent)
    except ValueError:
        return False


class FileOperationTask(QRunnable):
    def __init__(
        self,
        operation: str,
        sources: Sequence[str],
        destination: Optional[str],
        cancel: threading.Event,
        use_gio: bool = False,
    ):
        super().__init__()
        self.operation = operation
        self.sources = list(sources)
        self.destination = destination
        self.cancel = cancel
        self.use_gio = use_gio
        self.signals = FileOperationSignals()

    def run(self) -> None:
        errors: list[str] = []
        completed_count = 0
        total = len(self.sources)
        for index, source in enumerate(self.sources, 1):
            if self.cancel.is_set():
                break
            self.signals.progress.emit(os.path.basename(source) or source, index, total)
            try:
                if self.operation in {"copy", "move"}:
                    assert self.destination is not None
                    target = _available_destination(os.path.join(self.destination, os.path.basename(source)))
                    if os.path.isdir(source) and _is_within(target, source):
                        raise OSError("cannot copy or move a folder into itself")
                    if self.operation == "copy":
                        if os.path.isdir(source) and not os.path.islink(source):
                            shutil.copytree(source, target, symlinks=True)
                        else:
                            shutil.copy2(source, target, follow_symlinks=False)
                    else:
                        shutil.move(source, target)
                elif self.operation == "trash":
                    completed = subprocess.run(
                        ["gio", "trash", source],
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        check=False,
                    )
                    if completed.returncode != 0:
                        raise OSError(completed.stderr.strip() or f"gio exited {completed.returncode}")
                elif self.operation == "delete":
                    if os.path.isdir(source) and not os.path.islink(source):
                        shutil.rmtree(source)
                    else:
                        os.unlink(source)
                else:
                    raise ValueError(f"unknown operation: {self.operation}")
                completed_count += 1
            except Exception as exc:  # retain remaining operations and report every failure
                errors.append(f"{source}: {exc}")
        self.signals.finished.emit(
            {
                "operation": self.operation,
                "completed": completed_count,
                "total": total,
                "cancelled": self.cancel.is_set(),
                "errors": errors,
            }
        )


class VideoFrame(QWidget):
    double_clicked = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_NativeWindow, True)
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor("black"))
        self.setPalette(palette)
        self.setMinimumSize(320, 180)

    def mouseDoubleClickEvent(self, event) -> None:  # type: ignore[no-untyped-def]
        self.double_clicked.emit()
        event.accept()


class PropertiesDialog(QDialog):
    def __init__(self, title: str, rows: Sequence[tuple[str, str]], parent: QWidget):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(560, 260)
        layout = QVBoxLayout(self)
        form = QFormLayout()
        for label, value in rows:
            field = QLineEdit(value)
            field.setReadOnly(True)
            field.setCursorPosition(0)
            form.addRow(label, field)
        layout.addLayout(form)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


class MediaExplorerWindow(QMainWindow):
    HEADERS = ("Name", "Type", "Size", "Modified", "Resolution", "Duration")

    def __init__(self, config: AppConfig):
        super().__init__()
        self.config = config
        self.pool = QThreadPool.globalInstance()
        self.pool.setMaxThreadCount(max(4, min(8, os.cpu_count() or 4)))
        # ffprobe calls can be slow on network storage.  Keeping them on their
        # own small pool prevents hundreds of metadata jobs from delaying a
        # folder scan, search, or file operation.
        self.metadata_pool = QThreadPool(self)
        self.metadata_pool.setMaxThreadCount(4)
        self.entries: list[Entry] = []
        self.entry_by_path: dict[str, Entry] = {}
        self.row_by_path: dict[str, int] = {}
        self.current_dir: Optional[str] = None
        self.view_mode = "roots"
        self.sort_column = 0
        self.sort_ascending = True
        self.scan_generation = 0
        self.scan_cancel: Optional[threading.Event] = None
        self.search_token = 0
        self.search_cancel: Optional[threading.Event] = None
        self.search_terms: list[str] = []
        self.search_scopes: list[str] = []
        self.search_return_dir: Optional[str] = None
        self.search_return_roots = True
        self.mapping_refresh_token = 0
        self.metadata_cache: dict[str, tuple[int, int, int, int, Optional[float], str, str]] = {}
        self.metadata_generation = 0
        self.metadata_pending: set[tuple[int, str, int, int]] = set()
        self.clipboard_paths: list[str] = []
        self.clipboard_mode = "copy"
        self.file_operation_cancel: Optional[threading.Event] = None
        self.busy_owner = ""
        self.cancel_callback: Optional[Callable[[], None]] = None
        self._fullscreen = False
        self._last_window_state = None
        self._playlist: list[str] = []
        self._playlist_index = -1
        self._vlc_instance = None
        self._player = None
        self._current_media = None
        self._last_player_state = None
        self._seeking = False
        self._playback_error_shown = False

        self.setWindowTitle(APP_NAME)
        self.resize(1220, 760)
        self.setMinimumSize(760, 480)
        self._build_ui()
        self._build_actions()
        self._build_shortcuts()

        self.player_timer = QTimer(self)
        self.player_timer.setInterval(250)
        self.player_timer.timeout.connect(self._poll_player)
        self.metadata_sort_timer = QTimer(self)
        self.metadata_sort_timer.setSingleShot(True)
        self.metadata_sort_timer.timeout.connect(lambda: self._sort_and_fill(True))

        start = config.start_path
        if start and os.path.isdir(start):
            self.navigate(start, add_history=False)
        else:
            self.show_roots(add_history=False)

    # ----- UI construction -------------------------------------------------
    def _build_ui(self) -> None:
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.browser_page = QWidget()
        browser_layout = QVBoxLayout(self.browser_page)
        browser_layout.setContentsMargins(6, 6, 6, 6)
        location_layout = QHBoxLayout()
        self.up_button = QPushButton("Up")
        self.up_button.clicked.connect(self.go_up)
        self.roots_button = QPushButton("Mounts")
        self.roots_button.clicked.connect(lambda: self.show_roots())
        self.location = QLineEdit()
        self.location.setClearButtonEnabled(True)
        self.location.returnPressed.connect(self.go_to_location)
        self.go_button = QPushButton("Go")
        self.go_button.clicked.connect(self.go_to_location)
        location_layout.addWidget(self.up_button)
        location_layout.addWidget(self.roots_button)
        location_layout.addWidget(self.location, 1)
        location_layout.addWidget(self.go_button)
        browser_layout.addLayout(location_layout)

        self.table = QTableWidget(0, len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_table_menu)
        self.table.cellDoubleClicked.connect(self._activate_row)
        self.table.horizontalHeader().sectionClicked.connect(self._header_clicked)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for column in range(1, len(self.HEADERS)):
            header.setSectionResizeMode(column, QHeaderView.Interactive)
        self.table.setColumnWidth(1, 145)
        self.table.setColumnWidth(2, 105)
        self.table.setColumnWidth(3, 145)
        self.table.setColumnWidth(4, 115)
        self.table.setColumnWidth(5, 100)
        browser_layout.addWidget(self.table, 1)
        self.stack.addWidget(self.browser_page)

        self.player_page = QWidget()
        player_layout = QVBoxLayout(self.player_page)
        player_layout.setContentsMargins(5, 5, 5, 5)
        self.player_title = QLabel("")
        self.player_title.setTextInteractionFlags(Qt.TextSelectableByMouse)
        player_layout.addWidget(self.player_title)
        splitter = QSplitter(Qt.Horizontal)
        self.video_frame = VideoFrame()
        self.video_frame.double_clicked.connect(self.toggle_fullscreen)
        splitter.addWidget(self.video_frame)
        self.playlist_widget = QListWidget()
        self.playlist_widget.setMinimumWidth(220)
        self.playlist_widget.setMaximumWidth(420)
        self.playlist_widget.itemDoubleClicked.connect(self._playlist_double_clicked)
        splitter.addWidget(self.playlist_widget)
        splitter.setStretchFactor(0, 1)
        player_layout.addWidget(splitter, 1)

        self.seek_slider = QSlider(Qt.Horizontal)
        self.seek_slider.setRange(0, 1000)
        self.seek_slider.sliderPressed.connect(self._seek_pressed)
        self.seek_slider.sliderReleased.connect(self._seek_released)
        player_layout.addWidget(self.seek_slider)
        self.player_controls = QWidget()
        controls = QHBoxLayout(self.player_controls)
        controls.setContentsMargins(0, 0, 0, 0)
        self.back_button = QPushButton("Back")
        self.back_button.clicked.connect(self.stop_playback)
        self.previous_button = QPushButton("Previous")
        self.previous_button.clicked.connect(self.previous_video)
        self.pause_button = QPushButton("Pause")
        self.pause_button.clicked.connect(self.toggle_pause)
        self.next_button = QPushButton("Next")
        self.next_button.clicked.connect(self.next_video)
        self.time_label = QLabel("0:00 / 0:00")
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 200)
        self.volume_slider.setValue(100)
        self.volume_slider.setMaximumWidth(150)
        self.volume_slider.valueChanged.connect(self._set_volume)
        self.fullscreen_button = QPushButton("Fullscreen")
        self.fullscreen_button.clicked.connect(self.toggle_fullscreen)
        for widget in (self.back_button, self.previous_button, self.pause_button, self.next_button):
            controls.addWidget(widget)
        controls.addWidget(self.time_label)
        controls.addStretch(1)
        controls.addWidget(QLabel("Volume"))
        controls.addWidget(self.volume_slider)
        controls.addWidget(self.fullscreen_button)
        player_layout.addWidget(self.player_controls)
        self.stack.addWidget(self.player_page)

        self.progress = QProgressBar()
        self.progress.setMaximumWidth(220)
        self.progress.hide()
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setMaximumWidth(90)
        self.cancel_button.clicked.connect(self._cancel_busy)
        self.cancel_button.hide()
        self.statusBar().addPermanentWidget(self.progress)
        self.statusBar().addPermanentWidget(self.cancel_button)

    def _action(self, text: str, slot: Callable[[], None], shortcut: Optional[str] = None) -> QAction:
        action = QAction(text, self)
        action.triggered.connect(slot)
        if shortcut:
            action.setShortcut(QKeySequence(shortcut))
        return action

    def _build_actions(self) -> None:
        file_menu = self.menuBar().addMenu("&File")
        self.play_action = self._action("&Play selected", self.play_or_properties, "Ctrl+P")
        self.copy_action = self._action("&Copy", self.copy_selected, "Ctrl+C")
        self.cut_action = self._action("Cu&t", self.cut_selected, "Ctrl+X")
        self.paste_action = self._action("&Paste", self.paste, "Ctrl+V")
        self.rename_action = self._action("&Rename", self.rename_selected, "F2")
        self.delete_action = self._action("&Delete", self.delete_selected, "Delete")
        self.new_folder_action = self._action("New &folder", self.new_folder, "Ctrl+Shift+N")
        self.open_external_action = self._action("Open in system file manager", self.open_external)
        for action in (
            self.play_action,
            self.copy_action,
            self.cut_action,
            self.paste_action,
            self.rename_action,
            self.delete_action,
            self.new_folder_action,
            self.open_external_action,
        ):
            file_menu.addAction(action)
        file_menu.addSeparator()
        file_menu.addAction(self._action("E&xit", self.close, "Ctrl+Q"))

        navigate_menu = self.menuBar().addMenu("&Navigate")
        navigate_menu.addAction(self._action("&Up", self.go_up, "Backspace"))
        navigate_menu.addAction(self._action("&Mounts", self.show_roots, "Ctrl+Home"))
        navigate_menu.addAction(self._action("&Location", self.focus_location, "Ctrl+L"))
        navigate_menu.addAction(self._action("&Refresh", self.refresh, "F5"))
        navigate_menu.addAction(self._action("&Search", self.prompt_search, "Ctrl+F"))

        playback_menu = self.menuBar().addMenu("&Playback")
        playback_menu.addAction(self._action("Pause / resume", self.toggle_pause, "Space"))
        playback_menu.addAction(self._action("Previous", self.previous_video, "PgUp"))
        playback_menu.addAction(self._action("Next", self.next_video, "PgDown"))
        playback_menu.addAction(self._action("Fullscreen", self.toggle_fullscreen, "F11"))
        playback_menu.addAction(self._action("Video properties", self.show_video_properties))
        playback_menu.addAction(self._action("Stop / return to browser", self.stop_playback))

        help_menu = self.menuBar().addMenu("&Help")
        help_menu.addAction(self._action("Media Explorer &Help", self.show_help, "F1"))
        help_menu.addAction(self._action("&About", self.show_about))

    def _shortcut(self, sequence: str, slot: Callable[[], None], playback_only: bool = False) -> None:
        shortcut = QShortcut(QKeySequence(sequence), self)
        shortcut.setContext(Qt.ApplicationShortcut)
        shortcut.activated.connect(slot)
        if playback_only:
            shortcut.setEnabled(False)
            self.player_shortcuts.append(shortcut)
        self.shortcuts.append(shortcut)

    def _build_shortcuts(self) -> None:
        self.shortcuts: list[QShortcut] = []
        self.player_shortcuts: list[QShortcut] = []
        self._shortcut("Return", self.activate_selection)
        self._shortcut("Enter", self.activate_selection)
        self._shortcut("Alt+Left", self.go_up)
        self._shortcut("Ctrl+A", self.select_all_videos)
        self._shortcut("Escape", self._escape)
        self._shortcut("Left", lambda: self.seek_by(-10000), True)
        self._shortcut("Right", lambda: self.seek_by(10000), True)
        self._shortcut("Shift+Left", lambda: self.seek_by(-60000), True)
        self._shortcut("Shift+Right", lambda: self.seek_by(60000), True)
        self._shortcut("Up", lambda: self.change_volume(5), True)
        self._shortcut("Down", lambda: self.change_volume(-5), True)
        self._shortcut("F", self.toggle_fullscreen, True)
        self._shortcut("P", self.previous_video, True)
        self._shortcut("N", self.next_video, True)
        self._shortcut("Ctrl+Left", self.previous_video, True)
        self._shortcut("Ctrl+Right", self.next_video, True)
        self._shortcut("Tab", self.toggle_pause, True)
        self._shortcut("Ctrl+G", self.focus_playlist, True)

    # ----- browser ---------------------------------------------------------
    def _set_busy(self, owner: str, message: str, cancel: Optional[Callable[[], None]], maximum: int = 0) -> None:
        self.busy_owner = owner
        self.cancel_callback = cancel
        self.statusBar().showMessage(message)
        self.progress.setRange(0, maximum)
        self.progress.setValue(0)
        self.progress.show()
        if cancel is not None:
            self.cancel_button.setEnabled(True)
        self.cancel_button.setVisible(cancel is not None)

    def _finish_busy(self, owner: str, message: str = "") -> None:
        if self.busy_owner != owner:
            return
        self.busy_owner = ""
        self.cancel_callback = None
        self.progress.hide()
        self.cancel_button.hide()
        self.statusBar().showMessage(message, 7000 if message else 0)

    def _cancel_busy(self) -> None:
        if self.cancel_callback:
            self.cancel_callback()
            self.statusBar().showMessage("Cancelling after the current item…")
            self.cancel_button.setEnabled(False)

    def _invalidate_scan(self) -> None:
        if self.scan_cancel:
            self.scan_cancel.set()
        self.scan_cancel = None
        self.scan_generation += 1

    def _invalidate_search(self) -> None:
        if self.search_cancel:
            self.search_cancel.set()
        self.search_cancel = None
        self.search_token += 1

    def _clear_metadata_queue(self) -> None:
        self.metadata_generation += 1
        self.metadata_pool.clear()
        self.metadata_pending.clear()
        if hasattr(self, "metadata_sort_timer"):
            self.metadata_sort_timer.stop()

    def show_roots(self, checked: bool = False, add_history: bool = True) -> None:
        del checked, add_history
        previous_busy = self.busy_owner
        self._invalidate_scan()
        self._invalidate_search()
        self._clear_metadata_queue()
        self.view_mode = "roots"
        self.current_dir = None
        self.search_terms = []
        self.location.setText("Computer / mounts")
        self.entries = discover_mounts(self.config)
        self._sort_and_fill()
        self.setWindowTitle(f"{APP_NAME} — Mounts")
        message = f"{len(self.entries)} locations"
        if previous_busy in {"scan", "search"}:
            self._finish_busy(previous_busy, message)
        elif not self.busy_owner:
            self.statusBar().showMessage(message, 7000)

    def navigate(self, path: str, add_history: bool = True, select_path: str = "") -> None:
        del add_history
        expanded = os.path.abspath(os.path.expandvars(os.path.expanduser(path)))
        if not os.path.isdir(expanded):
            self._error("Open folder", f"The folder is not available:\n{expanded}")
            return
        self.mapping_refresh_token += 1
        self._invalidate_scan()
        self._invalidate_search()
        self._clear_metadata_queue()
        generation = self.scan_generation
        cancel = threading.Event()
        self.scan_cancel = cancel
        self.current_dir = expanded
        self.view_mode = "folder"
        self.search_terms = []
        self.location.setText(expanded)
        self.entries = []
        self._fill_table([])
        self.setWindowTitle(f"{APP_NAME} — {expanded}")
        self._set_busy("scan", f"Opening {expanded}…", cancel.set)
        task = FolderScanTask(expanded, generation, self.config, cancel)
        task.signals.finished.connect(lambda result, wanted=select_path: self._scan_finished(result, wanted))
        self.pool.start(task)

    def _scan_finished(self, result: dict, select_path: str = "") -> None:
        if (
            result["generation"] != self.scan_generation
            or result["path"] != self.current_dir
            or self.view_mode != "folder"
        ):
            return
        self.scan_cancel = None
        self.cancel_button.setEnabled(True)
        if result["cancelled"]:
            self._finish_busy("scan", "Folder load cancelled")
            return
        if result["error"]:
            self._finish_busy("scan")
            self._error("Open folder", f"Could not read:\n{result['path']}\n\n{result['error']}")
            return
        self.entries = result["entries"]
        self._sort_and_fill()
        self._finish_busy("scan", f"{len(self.entries)} folders and videos")
        if select_path:
            self._select_paths([select_path])
        self._queue_metadata()

    def refresh(self) -> None:
        if self.stack.currentWidget() is self.player_page:
            return
        if self.view_mode == "roots":
            self.show_roots()
        elif self.view_mode == "folder" and self.current_dir:
            self.navigate(self.current_dir)
        elif self.view_mode == "search":
            self._start_search(self.search_scopes, self.search_terms, keep_return=True)

    def go_to_location(self) -> None:
        value = self.location.text().strip()
        if value.casefold() in {"computer", "mounts", "computer / mounts"}:
            self.show_roots()
            return
        if value.startswith("file://"):
            value = QUrl(value).toLocalFile()
        self.navigate(value)

    def focus_location(self) -> None:
        if self.stack.currentWidget() is self.browser_page:
            self.location.setFocus()
            self.location.selectAll()

    def go_up(self) -> None:
        if self.stack.currentWidget() is self.player_page:
            self.stop_playback()
            return
        if self.view_mode == "search":
            if self.search_return_roots:
                self.show_roots()
            elif self.search_return_dir:
                self.navigate(self.search_return_dir)
            return
        if self.current_dir:
            parent = os.path.dirname(self.current_dir.rstrip(os.sep)) or os.sep
            if os.path.normpath(parent) == os.path.normpath(self.current_dir):
                self.show_roots()
            else:
                self.navigate(parent, select_path=self.current_dir)
        else:
            self.show_roots()

    def _sort_value(self, entry: Entry, column: int):
        if column == 0:
            return entry.name.casefold()
        if column == 1:
            return entry.kind.casefold()
        if column == 2:
            return entry.size
        if column == 3:
            return entry.mtime
        if column == 4:
            match = re.match(r"(\d+)x(\d+)", entry.resolution)
            return int(match.group(1)) * int(match.group(2)) if match else -1
        if column == 5:
            return entry.duration_seconds if entry.duration_seconds is not None else -1.0
        return entry.name.casefold()

    def _sort_and_fill(self, keep_selection: bool = False) -> None:
        selected = [entry.path for entry in self.selected_entries()] if keep_selection else []
        directories = [entry for entry in self.entries if entry.is_dir]
        files = [entry for entry in self.entries if not entry.is_dir]
        directories.sort(key=lambda item: self._sort_value(item, self.sort_column), reverse=not self.sort_ascending)
        files.sort(key=lambda item: self._sort_value(item, self.sort_column), reverse=not self.sort_ascending)
        self.entries = directories + files
        self._fill_table(self.entries)
        if selected:
            self._select_paths(selected)

    def _fill_table(self, entries: Sequence[Entry]) -> None:
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(len(entries))
        self.entry_by_path = {entry.path: entry for entry in entries}
        self.row_by_path = {}
        for row, entry in enumerate(entries):
            self.row_by_path[entry.path] = row
            values = (
                entry.name,
                entry.kind,
                "" if entry.is_dir else format_size(entry.size),
                format_modified(entry.mtime),
                entry.resolution,
                format_duration(entry.duration_seconds),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setData(Qt.UserRole, entry.path)
                if column in (2, 4, 5):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                if column == 0:
                    item.setToolTip(entry.path)
                self.table.setItem(row, column, item)
        header = self.table.horizontalHeader()
        header.setSortIndicatorShown(True)
        header.setSortIndicator(self.sort_column, Qt.AscendingOrder if self.sort_ascending else Qt.DescendingOrder)
        self.table.setUpdatesEnabled(True)

    def _header_clicked(self, column: int) -> None:
        if column == self.sort_column:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = column
            self.sort_ascending = True
        self._sort_and_fill(True)

    def selected_entries(self) -> list[Entry]:
        rows = sorted({index.row() for index in self.table.selectionModel().selectedRows()})
        return [self.entries[row] for row in rows if 0 <= row < len(self.entries)]

    def select_all_videos(self) -> None:
        if self.stack.currentWidget() is not self.browser_page:
            return
        self.table.clearSelection()
        first = -1
        for row, entry in enumerate(self.entries):
            if entry.is_dir:
                continue
            for column in range(self.table.columnCount()):
                item = self.table.item(row, column)
                if item is not None:
                    item.setSelected(True)
            if first < 0:
                first = row
        if first >= 0:
            self.table.scrollToItem(self.table.item(first, 0))

    def _select_paths(self, paths: Sequence[str]) -> None:
        wanted = set(paths)
        self.table.clearSelection()
        first = -1
        for path in wanted:
            row = self.row_by_path.get(path, -1)
            if row >= 0:
                self.table.selectRow(row)
                if first < 0:
                    first = row
        if first >= 0:
            self.table.scrollToItem(self.table.item(first, 0))

    def _activate_row(self, row: int, column: int = 0) -> None:
        del column
        if not 0 <= row < len(self.entries):
            return
        entry = self.entries[row]
        if entry.source == "mapping-disconnected":
            self._activate_disconnected_mapping(entry)
            return
        if entry.is_dir:
            self.navigate(entry.path)
            return
        selected = self.selected_entries()
        if entry not in selected:
            selected = [entry]
        self.play_entries(selected)

    def activate_selection(self) -> None:
        if self.location.hasFocus() and self.stack.currentWidget() is self.browser_page:
            self.go_to_location()
            return
        if self.stack.currentWidget() is self.player_page:
            if self.playlist_widget.hasFocus() and self.playlist_widget.currentRow() >= 0:
                self.play_index(self.playlist_widget.currentRow())
            else:
                self.toggle_pause()
            return
        selected = self.selected_entries()
        if not selected:
            row = self.table.currentRow()
            if row >= 0:
                selected = [self.entries[row]]
        if len(selected) == 1 and selected[0].source == "mapping-disconnected":
            self._activate_disconnected_mapping(selected[0])
        elif len(selected) == 1 and selected[0].is_dir:
            self.navigate(selected[0].path)
        else:
            self.play_entries(selected)

    def _show_table_menu(self, position) -> None:  # type: ignore[no-untyped-def]
        menu = QMenu(self)
        menu.addAction(self.play_action)
        menu.addSeparator()
        menu.addAction(self.copy_action)
        menu.addAction(self.cut_action)
        menu.addAction(self.paste_action)
        menu.addAction(self.rename_action)
        menu.addAction(self.delete_action)
        menu.addSeparator()
        menu.addAction(self.new_folder_action)
        menu.addAction(self.open_external_action)
        menu.exec_(self.table.viewport().mapToGlobal(position))

    def _activate_disconnected_mapping(self, entry: Entry) -> None:
        if not entry.uri:
            self._error(
                "Mapped drive disconnected",
                f"{entry.name} is not mounted. Configure its smb:// URI in mediaexplorer.ini "
                "or mount the share at:\n{entry.path}",
            )
            return
        if not QDesktopServices.openUrl(QUrl(entry.uri)):
            self._error("Connect mapped drive", f"The desktop could not open:\n{entry.uri}")
            return
        self.statusBar().showMessage(
            f"Connection requested for {entry.name}. Complete any credential prompt; locations will refresh.", 10000
        )
        self.mapping_refresh_token += 1
        token = self.mapping_refresh_token
        QTimer.singleShot(2500, lambda wanted=token: self._refresh_mapping_roots(wanted))
        QTimer.singleShot(7000, lambda wanted=token: self._refresh_mapping_roots(wanted))

    def _refresh_mapping_roots(self, token: int) -> None:
        if token != self.mapping_refresh_token or self.view_mode != "roots":
            return
        self.entries = discover_mounts(self.config)
        self._sort_and_fill()
        self.statusBar().showMessage(f"{len(self.entries)} locations", 7000)

    # ----- metadata --------------------------------------------------------
    def _queue_metadata(self) -> None:
        cached_sort_changed = False
        generation = self.metadata_generation
        limit = self.config.metadata_prefetch_limit
        for entry in self.entries:
            if entry.is_dir:
                continue
            cached = self.metadata_cache.get(entry.path)
            if cached and cached[0] == entry.mtime_ns and cached[1] == entry.size:
                entry.resolution = f"{cached[2]}x{cached[3]}" if cached[2] and cached[3] else ""
                entry.duration_seconds = cached[4]
                row = self.row_by_path.get(entry.path)
                if row is not None:
                    resolution_item = self.table.item(row, 4)
                    duration_item = self.table.item(row, 5)
                    if resolution_item is not None:
                        resolution_item.setText(entry.resolution)
                    if duration_item is not None:
                        duration_item.setText(format_duration(entry.duration_seconds))
                cached_sort_changed = True
                continue
            key = (generation, entry.path, entry.mtime_ns, entry.size)
            if key in self.metadata_pending or len(self.metadata_pending) >= limit:
                continue
            self.metadata_pending.add(key)
            task = MetadataTask(entry, self.config, generation)
            task.signals.finished.connect(self._metadata_finished)
            self.metadata_pool.start(task)
        if cached_sort_changed and self.sort_column in (4, 5):
            self._sort_and_fill(True)

    def _metadata_finished(self, result: dict) -> None:
        generation = int(result.get("generation", self.metadata_generation))
        key = (generation, result["path"], result["mtime_ns"], result["size"])
        self.metadata_pending.discard(key)
        self.metadata_cache[result["path"]] = (
            result["mtime_ns"],
            result["size"],
            result["width"],
            result["height"],
            result["duration"],
            result["codec"],
            result["error"],
        )
        if generation != self.metadata_generation:
            return
        entry = self.entry_by_path.get(result["path"])
        if not entry or entry.mtime_ns != result["mtime_ns"] or entry.size != result["size"]:
            return
        entry.resolution = (
            f"{result['width']}x{result['height']}" if result["width"] and result["height"] else ""
        )
        entry.duration_seconds = result["duration"]
        row = self.row_by_path.get(entry.path)
        if row is not None:
            self.table.item(row, 4).setText(entry.resolution)
            self.table.item(row, 5).setText(format_duration(entry.duration_seconds))
        if self.sort_column in (4, 5):
            self.metadata_sort_timer.start(350)

    # ----- recursive search ------------------------------------------------
    def prompt_search(self) -> None:
        if self.stack.currentWidget() is self.player_page:
            return
        initial = " ".join(self.search_terms) if self.view_mode == "search" else ""
        text, accepted = QInputDialog.getText(
            self,
            "Search videos",
            "All terms must occur in the file name (space-separated):",
            QLineEdit.Normal,
            initial,
        )
        if not accepted:
            return
        try:
            terms = [term for term in shlex.split(text) if term]
        except ValueError as exc:
            self._error("Search", f"Invalid quoted search text: {exc}")
            return
        if not terms:
            self._error("Search", "Enter at least one search term.")
            return
        if self.view_mode == "search":
            scopes = self.search_scopes
        elif self.current_dir:
            selected_dirs = [
                entry.path for entry in self.selected_entries()
                if entry.is_dir and entry.source != "mapping-disconnected"
            ]
            scopes = selected_dirs or [self.current_dir]
            self.search_return_dir = self.current_dir
            self.search_return_roots = False
        else:
            selected_dirs = [
                entry.path for entry in self.selected_entries()
                if entry.is_dir and entry.source != "mapping-disconnected"
            ]
            scopes = selected_dirs or [str(Path.home())]
            self.search_return_dir = None
            self.search_return_roots = True
        self._start_search(scopes, terms, keep_return=True)

    def _start_search(self, scopes: Sequence[str], terms: Sequence[str], keep_return: bool) -> None:
        del keep_return
        self.mapping_refresh_token += 1
        self._invalidate_scan()
        self._invalidate_search()
        self._clear_metadata_queue()
        token = self.search_token
        cancel = threading.Event()
        self.search_cancel = cancel
        self.search_scopes = list(dict.fromkeys(os.path.abspath(path) for path in scopes))
        self.search_terms = list(terms)
        self.view_mode = "searching"
        self.entries = []
        self._fill_table([])
        title_terms = " AND ".join(self.search_terms)
        self.location.setText(f"Search: {title_terms}")
        self.setWindowTitle(f"{APP_NAME} — Search — {title_terms}")
        self.cancel_button.setEnabled(True)
        self._set_busy("search", "Searching…", cancel.set)
        task = SearchTask(self.search_scopes, self.search_terms, self.config, cancel, token)
        task.signals.progress.connect(
            lambda folder, files, matches, wanted=token: self._search_progress(
                folder, files, matches, wanted
            )
        )
        task.signals.finished.connect(self._search_finished)
        self.pool.start(task)

    def _search_progress(self, folder: str, files: int, matches: int, token: Optional[int] = None) -> None:
        if (
            self.busy_owner != "search"
            or self.view_mode != "searching"
            or (token is not None and token != self.search_token)
        ):
            return
        self.statusBar().showMessage(f"Searching {folder} — {files:,} videos checked, {matches:,} matches")

    def _search_finished(self, result: dict) -> None:
        if result["token"] != self.search_token or self.view_mode != "searching":
            return
        self.search_cancel = None
        self.cancel_button.setEnabled(True)
        self.view_mode = "search"
        self.entries = result["entries"]
        self._sort_and_fill()
        state = "cancelled; partial results" if result["cancelled"] else "complete"
        message = (
            f"Search {state}: {len(self.entries):,} matches, "
            f"{result['files']:,} videos in {result['directories']:,} folders"
        )
        if result["errors"]:
            message += f" ({len(result['errors'])} inaccessible folders)"
        self._finish_busy("search", message)
        self._queue_metadata()

    # ----- file operations -------------------------------------------------
    def _file_changes_allowed(self, action: str) -> bool:
        if self.view_mode in {"folder", "search"}:
            return True
        self._error(action, "Open a folder first. Mount roots and mapped-drive roots cannot be modified here.")
        return False

    def copy_selected(self) -> None:
        if not self._file_changes_allowed("Copy"):
            return
        paths = [entry.path for entry in self.selected_entries()]
        if paths:
            self.clipboard_paths = paths
            self.clipboard_mode = "copy"
            self.statusBar().showMessage(f"Copied {len(paths)} item(s) to Media Explorer clipboard", 5000)

    def cut_selected(self) -> None:
        if not self._file_changes_allowed("Cut"):
            return
        paths = [entry.path for entry in self.selected_entries()]
        if paths:
            self.clipboard_paths = paths
            self.clipboard_mode = "move"
            self.statusBar().showMessage(f"Cut {len(paths)} item(s); choose a folder and paste", 5000)

    def paste(self) -> None:
        if not self.clipboard_paths:
            self.statusBar().showMessage("The Media Explorer clipboard is empty", 5000)
            return
        if self.view_mode != "folder" or not self.current_dir:
            self._error("Paste", "Open the destination folder before pasting.")
            return
        sources = [path for path in self.clipboard_paths if os.path.exists(path)]
        if not sources:
            self._error("Paste", "The copied or cut items no longer exist.")
            return
        operation = self.clipboard_mode
        self._run_file_operation(operation, sources, self.current_dir)

    def delete_selected(self) -> None:
        if not self._file_changes_allowed("Delete"):
            return
        paths = [entry.path for entry in self.selected_entries()]
        if not paths:
            return
        gio = shutil.which("gio")
        if self.config.use_trash and gio:
            answer = QMessageBox.question(
                self,
                "Move to Trash",
                f"Move {len(paths)} selected item(s) to the Trash?",
                QMessageBox.Yes | QMessageBox.Cancel,
                QMessageBox.Cancel,
            )
            if answer == QMessageBox.Yes:
                self._run_file_operation("trash", paths, None, use_gio=True)
            return
        answer = QMessageBox.warning(
            self,
            "Permanently delete",
            f"Trash is unavailable or disabled. Permanently delete {len(paths)} selected item(s)?\n\n"
            "This cannot be undone.",
            QMessageBox.Yes | QMessageBox.Cancel,
            QMessageBox.Cancel,
        )
        if answer == QMessageBox.Yes:
            self._run_file_operation("delete", paths, None)

    def _run_file_operation(
        self, operation: str, sources: Sequence[str], destination: Optional[str], use_gio: bool = False
    ) -> None:
        if self.file_operation_cancel is not None:
            self._error("File operation", "Another file operation is already running.")
            return
        cancel = threading.Event()
        self.file_operation_cancel = cancel
        self.cancel_button.setEnabled(True)
        self._set_busy("fileop", f"Starting {operation}…", cancel.set, len(sources))
        task = FileOperationTask(operation, sources, destination, cancel, use_gio)
        task.signals.progress.connect(self._file_operation_progress)
        task.signals.finished.connect(self._file_operation_finished)
        self.pool.start(task)

    def _file_operation_progress(self, name: str, current: int, total: int) -> None:
        if self.busy_owner != "fileop":
            return
        self.progress.setRange(0, total)
        self.progress.setValue(max(0, current - 1))
        self.statusBar().showMessage(f"Processing {current} of {total}: {name}")

    def _file_operation_finished(self, result: dict) -> None:
        self.file_operation_cancel = None
        self.cancel_button.setEnabled(True)
        if self.clipboard_mode == "move" and result["operation"] == "move" and not result["errors"]:
            self.clipboard_paths = []
        message = f"{result['operation'].capitalize()}: {result['completed']} of {result['total']} completed"
        if result["cancelled"]:
            message += " (cancelled)"
        self._finish_busy("fileop", message)
        if result["errors"]:
            shown = "\n".join(result["errors"][:12])
            if len(result["errors"]) > 12:
                shown += f"\n…and {len(result['errors']) - 12} more"
            self._error("File operation errors", shown)
        self.refresh()

    def new_folder(self) -> None:
        if self.view_mode != "folder" or not self.current_dir:
            self._error("New folder", "Open a destination folder first.")
            return
        name, accepted = QInputDialog.getText(self, "New folder", "Folder name:")
        name = name.strip()
        if not accepted:
            return
        if not self._valid_basename(name):
            self._error("New folder", "Use a single non-empty folder name without '/' or a NUL character.")
            return
        path = os.path.join(self.current_dir, name)
        try:
            os.mkdir(path)
        except OSError as exc:
            self._error("New folder", str(exc))
            return
        self.navigate(self.current_dir, select_path=path)

    def rename_selected(self) -> None:
        if not self._file_changes_allowed("Rename"):
            return
        selected = self.selected_entries()
        if len(selected) != 1:
            self._error("Rename", "Select exactly one item.")
            return
        entry = selected[0]
        name, accepted = QInputDialog.getText(self, "Rename", "New name:", QLineEdit.Normal, os.path.basename(entry.path))
        name = name.strip()
        if not accepted:
            return
        if not self._valid_basename(name):
            self._error("Rename", "Use a single non-empty name without '/' or a NUL character.")
            return
        target = os.path.join(os.path.dirname(entry.path), name)
        if os.path.lexists(target):
            self._error("Rename", "An item with that name already exists.")
            return
        try:
            os.rename(entry.path, target)
        except OSError as exc:
            self._error("Rename", str(exc))
            return
        self.refresh()

    @staticmethod
    def _valid_basename(name: str) -> bool:
        return bool(name and name not in {".", ".."} and "/" not in name and "\0" not in name)

    def open_external(self) -> None:
        selected = self.selected_entries()
        path = selected[0].path if selected else self.current_dir
        if not path:
            return
        if not os.path.isdir(path):
            path = os.path.dirname(path)
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
            self._error("Open file manager", f"No desktop file manager accepted:\n{path}")

    # ----- VLC player ------------------------------------------------------
    def _ensure_player(self) -> bool:
        if self._player is not None:
            return True
        if vlc is None:
            self._error(
                "VLC unavailable",
                "python-vlc could not be imported. Install python3-vlc and VLC, then restart Media Explorer.",
            )
            return False
        try:
            self._vlc_instance = vlc.Instance(*self.config.vlc_args)
            self._player = self._vlc_instance.media_player_new()
            self._player.audio_set_volume(self.volume_slider.value())
        except Exception as exc:
            self._vlc_instance = None
            self._player = None
            self._error("VLC unavailable", f"Could not initialize libVLC:\n{exc}")
            return False
        return True

    def play_or_properties(self) -> None:
        if self.stack.currentWidget() is self.player_page:
            self.show_video_properties()
        else:
            self.play_entries(self.selected_entries())

    def play_entries(self, entries: Sequence[Entry]) -> None:
        paths = [entry.path for entry in entries if not entry.is_dir and is_video_file(entry.path, self.config.video_extensions)]
        if not paths:
            self.statusBar().showMessage("Select one or more videos to play", 5000)
            return
        if not self._ensure_player():
            return
        self._playlist = paths
        self._playlist_index = 0
        self.playlist_widget.clear()
        for path in paths:
            item = QListWidgetItem(os.path.basename(path))
            item.setToolTip(path)
            self.playlist_widget.addItem(item)
        self.stack.setCurrentWidget(self.player_page)
        for shortcut in self.player_shortcuts:
            shortcut.setEnabled(True)
        self.player_timer.start()
        self.play_index(0)

    def _embed_video(self) -> None:
        if self._player is None:
            return
        try:
            window_id = int(self.video_frame.winId())
            if sys.platform.startswith("linux"):
                self._player.set_xwindow(window_id)
            elif sys.platform == "win32":
                self._player.set_hwnd(window_id)
            elif sys.platform == "darwin":
                self._player.set_nsobject(window_id)
        except Exception as exc:
            self.statusBar().showMessage(f"VLC video embedding failed: {exc}", 10000)

    def play_index(self, index: int) -> None:
        if self._player is None or self._vlc_instance is None or not 0 <= index < len(self._playlist):
            return
        path = self._playlist[index]
        if not os.path.isfile(path):
            self._error("Playback", f"The video is no longer available:\n{path}")
            return
        self._playlist_index = index
        self.playlist_widget.setCurrentRow(index)
        self.player_title.setText(f"({index + 1} of {len(self._playlist)})  {path}")
        self.setWindowTitle(f"{APP_NAME} — Playing — {os.path.basename(path)}")
        try:
            old_media = self._current_media
            media = self._vlc_instance.media_new_path(path)
            self._player.set_media(media)
            self._current_media = media
            QTimer.singleShot(0, self._embed_video)
            result = self._player.play()
            if result == -1:
                raise RuntimeError("libVLC refused to start playback")
            if old_media is not None:
                try:
                    old_media.release()
                except Exception:
                    pass
            self.pause_button.setText("Pause")
            self._last_player_state = None
            self._playback_error_shown = False
        except Exception as exc:
            self._error("Playback", str(exc))

    def _playlist_double_clicked(self, item: QListWidgetItem) -> None:
        del item
        self.play_index(self.playlist_widget.currentRow())

    def toggle_pause(self) -> None:
        if self.stack.currentWidget() is not self.player_page or self._player is None:
            return
        try:
            if self._player.is_playing():
                self._player.pause()
                self.pause_button.setText("Resume")
            else:
                self._player.play()
                self.pause_button.setText("Pause")
        except Exception as exc:
            self.statusBar().showMessage(f"Playback control failed: {exc}", 7000)

    def previous_video(self) -> None:
        if self.stack.currentWidget() is self.player_page and self._playlist_index > 0:
            self.play_index(self._playlist_index - 1)

    def next_video(self) -> None:
        if self.stack.currentWidget() is self.player_page and self._playlist_index + 1 < len(self._playlist):
            self.play_index(self._playlist_index + 1)

    def seek_by(self, delta_ms: int) -> None:
        if self.stack.currentWidget() is not self.player_page or self._player is None:
            return
        try:
            current = max(0, self._player.get_time())
            length = self._player.get_length()
            target = current + delta_ms
            if length > 0:
                target = min(target, length)
            self._player.set_time(max(0, target))
        except Exception:
            pass

    def _seek_pressed(self) -> None:
        self._seeking = True

    def _seek_released(self) -> None:
        if self._player is not None:
            try:
                self._player.set_position(self.seek_slider.value() / 1000.0)
            except Exception:
                pass
        self._seeking = False

    def _set_volume(self, value: int) -> None:
        if self._player is not None:
            try:
                self._player.audio_set_volume(value)
            except Exception:
                pass

    def change_volume(self, delta: int) -> None:
        if self.stack.currentWidget() is self.player_page:
            self.volume_slider.setValue(max(0, min(self.volume_slider.maximum(), self.volume_slider.value() + delta)))

    def _poll_player(self) -> None:
        if self._player is None or self.stack.currentWidget() is not self.player_page:
            return
        try:
            current = self._player.get_time()
            length = self._player.get_length()
            if not self._seeking:
                position = self._player.get_position()
                if position >= 0:
                    self.seek_slider.setValue(max(0, min(1000, int(position * 1000))))
            self.time_label.setText(f"{format_clock(current)} / {format_clock(length)}")
            state = self._player.get_state()
            if vlc is not None and state == vlc.State.Ended and self._last_player_state != vlc.State.Ended:
                if self._playlist_index + 1 < len(self._playlist):
                    QTimer.singleShot(0, self.next_video)
                else:
                    self.pause_button.setText("Replay")
            elif vlc is not None and state == vlc.State.Error and not self._playback_error_shown:
                self._playback_error_shown = True
                self._error("Playback", "libVLC reported a playback error for this file.")
            self._last_player_state = state
        except Exception:
            pass

    def focus_playlist(self) -> None:
        if self.stack.currentWidget() is self.player_page:
            self.playlist_widget.setFocus()

    def toggle_fullscreen(self) -> None:
        if self.stack.currentWidget() is not self.player_page:
            return
        self._fullscreen = not self._fullscreen
        self.menuBar().setVisible(not self._fullscreen)
        self.statusBar().setVisible(not self._fullscreen)
        self.player_title.setVisible(not self._fullscreen)
        self.player_controls.setVisible(not self._fullscreen)
        self.playlist_widget.setVisible(not self._fullscreen)
        self.seek_slider.setVisible(not self._fullscreen)
        if self._fullscreen:
            self.showFullScreen()
        else:
            self.showNormal()
        QTimer.singleShot(80, self._embed_video)

    def stop_playback(self) -> None:
        if self.stack.currentWidget() is not self.player_page:
            return
        if self._fullscreen:
            self.toggle_fullscreen()
        self.player_timer.stop()
        if self._player is not None:
            try:
                self._player.stop()
            except Exception:
                pass
        self.stack.setCurrentWidget(self.browser_page)
        for shortcut in self.player_shortcuts:
            shortcut.setEnabled(False)
        if self.view_mode == "search":
            self.setWindowTitle(f"{APP_NAME} — Search — {' AND '.join(self.search_terms)}")
        elif self.current_dir:
            self.setWindowTitle(f"{APP_NAME} — {self.current_dir}")
        else:
            self.setWindowTitle(f"{APP_NAME} — Mounts")

    def show_video_properties(self) -> None:
        if not self._playlist or not 0 <= self._playlist_index < len(self._playlist):
            return
        path = self._playlist[self._playlist_index]
        try:
            stat_result = os.stat(path)
        except OSError as exc:
            self._error("Video properties", str(exc))
            return
        cached = self.metadata_cache.get(path)
        if (
            cached
            and cached[0] == getattr(stat_result, "st_mtime_ns", int(stat_result.st_mtime * 1e9))
            and cached[1] == stat_result.st_size
        ):
            self._open_properties(path, stat_result, cached)
            return
        entry = Entry(
            os.path.basename(path),
            path,
            False,
            file_kind(path),
            stat_result.st_size,
            stat_result.st_mtime,
            getattr(stat_result, "st_mtime_ns", int(stat_result.st_mtime * 1e9)),
        )
        task = MetadataTask(entry, self.config)
        task.signals.finished.connect(lambda result, st=stat_result: self._properties_metadata_ready(result, st))
        self.metadata_pool.start(task)
        self.statusBar().showMessage("Reading video properties…", 3000)

    def _properties_metadata_ready(self, result: dict, stat_result) -> None:  # type: ignore[no-untyped-def]
        cached = (
            result["mtime_ns"], result["size"], result["width"], result["height"],
            result["duration"], result["codec"], result["error"],
        )
        self.metadata_cache[result["path"]] = cached
        self._open_properties(result["path"], stat_result, cached)

    def _open_properties(self, path: str, stat_result, cached: tuple) -> None:  # type: ignore[no-untyped-def]
        resolution = f"{cached[2]}x{cached[3]}" if cached[2] and cached[3] else "Unknown"
        rows = [
            ("File", path),
            ("Type", file_kind(path)),
            ("Size", f"{format_size(stat_result.st_size)} ({stat_result.st_size:,} bytes)"),
            ("Modified", format_modified(stat_result.st_mtime)),
            ("Resolution", resolution),
            ("Duration", format_duration(cached[4]) or "Unknown"),
            ("Video codec", cached[5] or "Unknown"),
        ]
        if cached[6]:
            rows.append(("ffprobe", cached[6]))
        PropertiesDialog("Video properties", rows, self).exec_()

    def _escape(self) -> None:
        if self._fullscreen:
            self.toggle_fullscreen()
        elif self.stack.currentWidget() is self.player_page:
            self.stop_playback()
        elif self.busy_owner and self.cancel_callback:
            self._cancel_busy()
        elif self.view_mode == "search":
            self.go_up()

    # ----- help and lifecycle ---------------------------------------------
    def show_help(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Media Explorer Help")
        dialog.resize(760, 650)
        layout = QVBoxLayout(dialog)
        browser = QTextBrowser()
        config_location = str(self.config.config_path or (Path.home() / ".config/media-explorer/mediaexplorer.ini"))
        browser.setHtml(
            f"""
            <h2>{APP_NAME} {APP_VERSION}</h2>
            <p>Browse local mounts and mapped shares, search recursively for videos, and play a selection as a VLC playlist.</p>
            <h3>Browsing</h3>
            <p><b>Enter / double-click</b> opens a folder or plays selected videos. <b>Backspace</b> goes up.
            <b>Ctrl+L</b> focuses the path. <b>Ctrl+F</b> searches recursively; every term must occur in a file name.
            <b>F5</b> refreshes.</p>
            <h3>Files</h3>
            <p><b>Ctrl+C / Ctrl+X / Ctrl+V</b> copy, move, and paste. <b>F2</b> renames.
            <b>Ctrl+Shift+N</b> creates a folder. <b>Delete</b> asks before moving to Trash or permanently deleting.</p>
            <h3>Playback</h3>
            <p><b>Space</b> pauses/resumes. <b>Left/Right</b> seeks 10 seconds; add Shift for 60 seconds.
            <b>Page Up/Page Down</b> or <b>P/N</b> changes video. <b>Up/Down</b> changes volume.
            <b>F or F11</b> toggles fullscreen. <b>Ctrl+G</b> focuses the playlist. <b>Escape</b> leaves fullscreen or playback.
            During playback, <b>Ctrl+P</b> shows video properties.</p>
            <h3>Configuration</h3>
            <p>Loaded from: <code>{config_location}</code></p>
            <p>Mapped-drive root: <code>{self.config.mapped_root}</code><br>
            ffprobe: <code>{self.config.ffprobe_path}</code> ({'enabled' if self.config.ffprobe_available else 'disabled'})<br>
            ffmpeg: <code>{self.config.ffmpeg_path}</code> ({'enabled' if self.config.ffmpeg_available else 'disabled'})</p>
            """
        )
        layout.addWidget(browser)
        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        dialog.exec_()

    def show_about(self) -> None:
        QMessageBox.about(
            self,
            f"About {APP_NAME}",
            f"{APP_NAME} {APP_VERSION}\n\nLinux PyQt5/libVLC edition.",
        )

    def _error(self, title: str, message: str) -> None:
        QMessageBox.critical(self, title, message)

    def closeEvent(self, event) -> None:  # type: ignore[no-untyped-def]
        if self.file_operation_cancel is not None:
            self._error(
                "File operation in progress",
                "Media Explorer will remain open until the active copy, move, or delete operation finishes. "
                "Use Cancel and wait for it to stop before closing.",
            )
            event.ignore()
            return
        self.mapping_refresh_token += 1
        self._invalidate_scan()
        self._invalidate_search()
        self._clear_metadata_queue()
        self.player_timer.stop()
        if self._player is not None:
            try:
                self._player.stop()
                self._player.release()
            except Exception:
                pass
        if self._current_media is not None:
            try:
                self._current_media.release()
            except Exception:
                pass
        if self._vlc_instance is not None:
            try:
                self._vlc_instance.release()
            except Exception:
                pass
        event.accept()


def _executable_status(command: str) -> dict[str, object]:
    if os.sep in command:
        path = os.path.abspath(os.path.expanduser(command))
        return {"configured": command, "resolved": path, "available": os.path.isfile(path) and os.access(path, os.X_OK)}
    resolved = shutil.which(command)
    return {"configured": command, "resolved": resolved, "available": bool(resolved)}


def run_self_test(config_path: Optional[str] = None) -> int:
    """Perform a deterministic, noninteractive SSH/deployment smoke test."""

    report: dict[str, object] = {"application": APP_NAME, "version": APP_VERSION}
    failures: list[str] = []
    try:
        config = AppConfig.load(config_path)
        report["config"] = {
            "loaded_from": str(config.config_path) if config.config_path else "defaults",
            "mapped_root": str(config.mapped_root),
            "ffprobe_available_flag": config.ffprobe_available,
            "ffmpeg_available_flag": config.ffmpeg_available,
        }
    except Exception as exc:
        report["config_error"] = str(exc)
        failures.append("configuration")
        config = AppConfig()

    extension_checks = {
        "movie.mp4": is_video_file("movie.mp4", config.video_extensions),
        "UPPER.MKV": is_video_file("UPPER.MKV", config.video_extensions),
        "not-video.txt": not is_video_file("not-video.txt", config.video_extensions),
    }
    report["video_extensions"] = extension_checks
    if not all(extension_checks.values()):
        failures.append("video extension recognition")

    try:
        mounts = discover_mounts(config)
        report["mounts"] = {"count": len(mounts), "paths": [entry.path for entry in mounts]}
        if not mounts:
            failures.append("mount discovery")
    except Exception as exc:
        report["mount_error"] = str(exc)
        failures.append("mount discovery")

    report["ffprobe"] = _executable_status(config.ffprobe_path)
    report["ffmpeg"] = _executable_status(config.ffmpeg_path)
    if config.ffprobe_available and not report["ffprobe"]["available"]:  # type: ignore[index]
        failures.append("ffprobe lookup")
    if config.ffmpeg_available and not report["ffmpeg"]["available"]:  # type: ignore[index]
        failures.append("ffmpeg lookup")

    if vlc is None:
        report["libvlc"] = {"available": False, "error": "python-vlc import failed"}
        failures.append("libVLC initialization")
    else:
        test_instance = None
        test_player = None
        try:
            version_value = vlc.libvlc_get_version()
            if isinstance(version_value, bytes):
                version_text = version_value.decode("utf-8", errors="replace")
            else:
                version_text = str(version_value)
            test_instance = vlc.Instance(*config.vlc_args)
            if test_instance is None:
                raise RuntimeError("vlc.Instance returned no instance")
            test_player = test_instance.media_player_new()
            if test_player is None:
                raise RuntimeError("libVLC could not create a media player")
            report["libvlc"] = {"available": True, "version": version_text, "player_created": True}
        except Exception as exc:
            report["libvlc"] = {"available": False, "error": str(exc)}
            failures.append("libVLC initialization")
        finally:
            if test_player is not None:
                try:
                    test_player.release()
                except Exception:
                    pass
            if test_instance is not None:
                try:
                    test_instance.release()
                except Exception:
                    pass

    app = QApplication.instance() or QApplication([sys.argv[0], "--self-test"])
    app.setApplicationName(APP_NAME)
    app.processEvents()
    report["qapplication"] = {"created": True, "platform": QApplication.platformName()}
    app.quit()
    report["ok"] = not failures
    report["failures"] = failures
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0 if not failures else 1


def parse_command_line(argv: Sequence[str]) -> tuple[Optional[str], bool]:
    config_path: Optional[str] = None
    self_test = False
    index = 1
    while index < len(argv):
        argument = argv[index]
        if argument in {"--self-test", "--smoke-test"}:
            self_test = True
        elif argument == "--config":
            index += 1
            if index >= len(argv):
                raise ValueError("--config requires a path")
            config_path = argv[index]
        elif argument == "--version":
            print(f"{APP_NAME} {APP_VERSION}")
            raise SystemExit(0)
        else:
            raise ValueError(f"unknown argument: {argument}")
        index += 1
    return config_path, self_test


def main(argv: Optional[Sequence[str]] = None) -> int:
    arguments = list(argv if argv is not None else sys.argv)
    try:
        config_path, self_test = parse_command_line(arguments)
    except ValueError as exc:
        print(f"media-explorer: {exc}", file=sys.stderr)
        return 2
    if self_test:
        return run_self_test(config_path)
    try:
        config = AppConfig.load(config_path)
    except ValueError as exc:
        print(f"media-explorer: {exc}", file=sys.stderr)
        return 2
    app = QApplication(arguments[:1])
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName(ORGANIZATION)
    if hasattr(app, "setDesktopFileName"):
        app.setDesktopFileName("media-explorer")
    window = MediaExplorerWindow(config)
    window.show()
    return app.exec_()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        raise SystemExit(1)
