# Media Explorer for Linux

This directory contains a native Linux edition of Media Explorer for Debian 13
and compatible distributions. It is written in Python 3 with PyQt5 and embeds
VLC through `python-vlc`. The Windows C++ source and build products are not used
or modified.

## What is included

- A mount/root view for Home, `/`, removable/network mounts, and stable mapped
  drive locations under `/mnt/media-explorer` (configurable).
- Async folder loading with folders-first sorting and Name, Type, Size,
  Modified, Resolution, and Duration columns. Folder views intentionally show
  directories and supported video files, matching the Win32 application.
- Cancellable, recursive, case-insensitive AND-term video search. Permission
  failures are skipped and summarized without aborting the whole search.
- Bounded background `ffprobe` work for resolution, duration, and codec data.
- Copy, cut, paste, rename, new-folder, and prompted deletion. Deletion uses the
  desktop Trash through `gio` by default; permanent deletion has a separate,
  explicit warning.
- Embedded libVLC playback with a visible playlist, seek/time display,
  pause/resume, previous/next, volume, fullscreen, video properties, automatic
  playlist advance, and keyboard controls.
- A PyInstaller one-directory build, icon, desktop entry, installer, and an
  offscreen `--self-test` suitable for SSH deployment checks.

## Debian 13 / arm64 dependencies

PyQt5 should come from Debian rather than being compiled from PyPI on arm64:

```bash
sudo apt update
sudo apt install python3 python3-venv python3-pyqt5 python3-vlc pyinstaller \
  vlc ffmpeg libglib2.0-bin xdg-utils
```

Or let the build helper install these packages:

```bash
bash build.sh --install-deps --clean
```

If Debian does not provide the desired PyInstaller version, `build.sh` creates
a `--system-site-packages` venv, reuses Debian's PyQt5, and installs only the
missing pure-Python/build packages. The finished bundle still expects the
machine's VLC/libVLC and normal graphics/system libraries.

## Build, test, and install

From this directory:

```bash
bash build.sh --clean
dist/MediaExplorer/MediaExplorer --self-test
sudo bash install.sh
```

The system install goes to `/opt/media-explorer`, registers the application,
creates the invoking desktop user's configuration, and puts `Media
Explorer.desktop` in that user's XDG Desktop directory. A no-root install is
also supported:

```bash
bash install.sh --user
```

To run directly from source:

```bash
QT_QPA_PLATFORM=xcb python3 media_explorer.py
```

The application selects Qt's `xcb` platform by default because libVLC 3 uses
an X11 window handle for embedded video. On a Wayland desktop it runs through
XWayland. `--self-test` automatically uses Qt's offscreen platform and does not
open a window.

## Configuration and mapped drives

The search order is:

1. `--config /path/to/file.ini`
2. `$MEDIA_EXPLORER_CONFIG`
3. `~/.config/media-explorer/mediaexplorer.ini`
4. `mediaexplorer.ini` beside the executable

See `mediaexplorer.ini.example` for every setting. Linux paths, executable
paths, boolean FFmpeg/ffprobe flags, extra argument lists, VLC arguments, video
extensions, hidden-file behavior, symlink policy, and mapped-drive roots are
supported. The old sectionless Win32 `key=value` form is accepted as well.

A persistent SMB mount can be placed at (for example)
`/mnt/media-explorer/v`. The root view checks `/proc/self/mountinfo`, so an empty
placeholder directory is shown as **disconnected**, not as a working drive.
Give it a connection URI in the INI file:

```ini
[media_explorer]
mapped_root = /mnt/media-explorer
mapping_v = smb://server/video
mapping_x = smb://server/archive
```

Double-clicking a disconnected mapping asks the desktop to open its `smb://`
URI. After credentials are supplied, Media Explorer also discovers the GVFS
mount under `/run/user/<uid>/gvfs` and retains the `V:`/`X:` label. For fully
noninteractive access, configure `/etc/fstab`/systemd mounts at the stable
paths instead. Credentials are intentionally not stored by Media Explorer.

## Keyboard reference

- Enter/double-click: open folder or play selected video(s)
- Backspace or Alt+Left: parent folder / leave search
- Ctrl+L: enter a path; Ctrl+Home: mount view; F5: refresh
- Ctrl+F: recursive search; Escape: cancel an active operation
- Ctrl+A: select all videos; Ctrl+C, Ctrl+X, Ctrl+V: copy, cut, paste
- Delete: prompted Trash/delete; F2: rename; Ctrl+Shift+N: new folder
- Ctrl+P: play selection; during playback, show properties
- Space: pause/resume; Left/Right: seek 10 seconds; Shift adds 60 seconds
- Page Up/Page Down, Ctrl+Left/Right, or P/N: previous/next; Up/Down: volume
- Tab or Space: pause/resume; volume supports VLC's 0–200% range
- F or F11: fullscreen; Ctrl+G: focus playlist; Escape: leave playback
- F1: Help

## SSH smoke test

`--self-test` and `--smoke-test` are aliases. They validate configuration,
case-insensitive video-extension recognition, mount discovery, configured
FFmpeg/ffprobe executable lookup, and QApplication creation/teardown. The
result is JSON and the process exits nonzero if an enabled executable is
missing or another validation fails:

```bash
/opt/media-explorer/MediaExplorer --self-test
```
