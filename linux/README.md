# Media Explorer for Linux

This is the native Linux edition of Media Explorer: a C++17 desktop application
built with Qt 5 Widgets and libVLC 3. It targets Debian 13 on arm64 and x86-64.
It does not use Python, PyQt, or PyInstaller.

## Features

- Home, filesystem, mounted-device, GVFS, and Windows-style mapped-drive roots.
- Stable `/mnt/media-explorer/<letter>` mount discovery without touching remote
  paths on the GUI thread. Activating an autofs mapping uses a bounded background
  probe; inaccessible mappings can open a configured `smb://` URI for the desktop
  credential prompt.
- Async folder scans, folders-first sorting, and Name, Type, Size, Modified,
  Resolution, and Duration columns. `ffprobe` metadata uses a separate four-thread
  pool so it cannot starve navigation or recursive search.
- Cancellable recursive video-name search. Each Ctrl+F refinement appends another
  case-insensitive AND keyword. Explicit GVFS scopes are allowed while unrelated
  `/run/user` content and kernel pseudo-filesystems are skipped.
- Cancellable chunked copy/move, cut/paste, rename, new folder, and prompted
  Trash/permanent deletion. Mount roots are protected and recursive operations do
  not cross nested filesystems.
- Embedded libVLC playback with playlist, seek bar, pause/resume, previous/next,
  0–200 volume, fullscreen, zoom/pan, properties, and deferred rename/copy/delete.
- FFmpeg trim-front, trim-end, horizontal-flip, and selected-video combine. Media
  is encoded in hidden per-job same-filesystem work directories, verified with ffprobe,
  then published atomically without replacing a raced or existing path. Playback
  outputs publish only after playback exits; stream-copy combine falls back to
  H.264/AAC. Combine state is persisted under the user's XDG state directory;
  interrupted prepared/running/publication jobs resume after restart without
  duplicating an already-published output. A per-state lock prevents two app
  instances from resuming the same job.
- Ctrl+U Topaz and Ctrl+3 IW3 queue submission with Windows-compatible preset JSON.
  The video is copied and synced first; a temporary JSON is then atomically
  published as the consumer trigger. Credentials are never stored by the app.

## Debian build dependencies

```bash
sudo apt update
sudo apt install build-essential pkg-config qtbase5-dev libvlc-dev vlc ffmpeg \
  libglib2.0-bin xdg-utils desktop-file-utils
```

Or install dependencies and build in one step:

```bash
./build.sh --install-deps --clean
```

## Build and test

```bash
make -j"$(nproc)"
make test
```

The executable is `dist/MediaExplorer/MediaExplorer`. `make test` runs its
noninteractive JSON self-test, the complete offscreen keyboard acceptance suite,
and the atomic combine-state unit test. Useful individual targets are:

```bash
make self-test
make keyboard-test
make combine-state-test
dist/MediaExplorer/MediaExplorer --self-test-shortcuts
```

The full self-test creates and tears down `QApplication`, loads configuration,
recognizes video extensions, discovers mounts without probing their contents,
locates enabled FFmpeg tools, tests pure media/queue helpers, and initializes,
creates a player from, and releases libVLC. It is safe to run over SSH and exits
nonzero on failure.

An opt-in live-display test exercises the application's real embedded libVLC
path. It needs a short video and an X11/XWayland display:

```bash
DISPLAY=:0 make playback-smoke PLAYBACK_CLIP=/path/to/short-video.mp4
# Soak example:
DISPLAY=:0 make playback-smoke PLAYBACK_CLIP=/path/to/short-video.mp4 \
  PLAYBACK_SMOKE_ARGS="--cycles 50"
```

Combine manifests live in `$XDG_STATE_HOME/media-explorer/pending` (or
`~/.local/state/media-explorer/pending`) with local private permissions. The
large same-filesystem work directory is a hidden, per-job directory beside the
requested output; fixed-mode CIFS mounts may not enforce private Unix mode bits.
Explicit cancellation safely abandons and removes an owned job. Genuine
validation, process, or publication failures are terminal and retained for
inspection rather than restarted automatically; the next startup offers an
explicit choice to retry them or keep them for inspection.

## Install

```bash
sudo ./install.sh
```

The default system install places the executable in `/opt/media-explorer`, the
registered entry in `/usr/share/applications`, the icon under the hicolor theme,
and a starter configuration in the invoking user's `~/.config/media-explorer`.
It then runs the self-test as that user.

When upgrading the former Python/PyInstaller Linux build, the installer removes
only an exact, non-symlink `_internal` directory beneath the selected prefix and
only after validating its PyInstaller/Python/PyQt markers. Unrecognized paths
are left untouched. The native installation does not ship Python artifacts.

For PCManFM/libfm desktops, an ordinary executable `Type=Application` file on the
Desktop triggers the misleading “executable script” warning, and neither a POSIX
symlink nor GIO `metadata::trusted` reliably suppresses it. The system installer
therefore creates a regular mode-0644 `Type=Link` wrapper on the user's Desktop,
pointing to the trusted registered entry at
`/usr/share/applications/media-explorer.desktop`. The bare native path is
intentional: libfm's trusted-target check does not accept a `file://` URI here.
An existing launcher is preserved outside the Desktop under
`~/.local/state/media-explorer/launcher-backups`. `--user` installs the
application entry under `~/.local` but skips this wrapper because libfm does not
treat that target as trusted. Use `--no-desktop-shortcut` to skip it for a system
install.

The launcher and executable force `QT_QPA_PLATFORM=xcb`; libVLC 3 embeds into an
X11 XID. On a Wayland labwc session this runs through XWayland, and it also works
in an XRDP X11 session. Self-tests force Qt's offscreen backend.

## Configuration

Configuration precedence is:

1. `--config /path/to/mediaexplorer.ini`
2. `$MEDIA_EXPLORER_CONFIG`
3. `~/.config/media-explorer/mediaexplorer.ini`
4. `mediaexplorer.ini` beside the executable

Both `[media_explorer]` INI files and the older sectionless `key=value` form are
accepted. Unknown `$VARIABLE` references remain intact; known variables and `~/`
are expanded. Inline `; comments` outside quoted values are supported. See
`mediaexplorer.ini.example` for the complete settings, including FFmpeg/ffprobe
paths and flags, VLC arguments, extension list, symlink policy, mapped roots,
upscale directory, and shared Topaz/IW3 queue.

A typical mapping configuration is:

```ini
[media_explorer]
mapped_root = /mnt/media-explorer
mapping_g = smb://server/share-g
mapping_x = smb://server/share-x
mapping_y = smb://server/share-y
mapping_z = smb://server/share-z
```

If `/mnt/media-explorer/g` is backed by systemd autofs, it is shown as on-demand.
Activation runs a finite `timeout ... stat -- path/.` worker, re-reads
`/proc/self/mountinfo`, and navigates only when the stable path is both mounted
and accessible. GVFS paths require accessibility but need not themselves be a
mountinfo mount point. Empty placeholder directories are disconnected, not
connected. Prefer system CIFS automounts for unattended access; use desktop GVFS
only when an interactive credential prompt is desired.

## Keyboard commands

Browser:

- Enter: open folder or play selected videos.
- Left, Backspace, Alt+Left: parent folder or leave search.
- Ctrl+A: select all visible videos; Ctrl+P: play.
- Ctrl+F: recursive search or append a refine keyword.
- Ctrl+Up / Ctrl+Down: reorder one selected visible row.
- Ctrl++ (Ctrl+=): combine selected videos in visible order.
- Ctrl+U / Ctrl+3: submit Topaz / IW3 jobs.
- Ctrl+C / Ctrl+X / Ctrl+V: copy / cut / paste.
- Delete: prompted Trash/delete; Escape: cancel the active file operation only.
- F2: rename; Ctrl+Shift+N: new folder; F5: refresh.
- Ctrl+L: edit location; Ctrl+Home: mount view; F1: help.

Playback:

- Enter: fullscreen; Escape: exit playback.
- Space: force pause; Tab: force resume.
- Left / Right: seek 10 seconds; Shift+Left / Shift+Right: seek 60 seconds.
- Ctrl+Left / Ctrl+Right: previous / next video.
- Up / Down: volume by 5, from 0 through 200.
- Ctrl+G: live playlist chooser; Ctrl+P: video properties.
- Ctrl+V: tools (`1` upscale, `2` trim front, `3` trim end, `4` horizontal flip).
- Ctrl+R / Ctrl+C / Delete: queue rename / copy / prompted deletion after playback.
- Ctrl+Z / Ctrl+X: zoom in / out; Alt+arrows / Alt+Home: pan / recenter.
- F1: help.

While a video is playing, any dialog or popup menu temporarily pauses it. Closing
the last popup resumes only if the video was playing beforehand; an already-paused
video stays paused. Windowed playback shows the file name, elapsed/total time, and
zoom percentage in the title at the top of the maximized window.

Menus display the same keys but do not register competing QAction shortcuts;
one centralized, mode-aware event dispatcher owns the complete matrix. Editable
fields, modal dialogs, and popup menus retain their normal keys.
