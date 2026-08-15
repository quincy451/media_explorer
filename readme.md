# MediaExplorer

MediaExplorer has separate native implementations for Windows and Linux:

- The Windows edition is written in C++17 using Win32 and libVLC. Its source,
  Visual Studio 2022 project files, and prebuilt runtime are at the repository
  root.
- The Linux edition is written in C++17 using Qt 5 Widgets and libVLC. Its
  source, GNU Make build, installer, and documentation are under
  [`linux/`](linux/README.md).

## Features

- Fast drive and folder browsing
- Recursive video search
- Video metadata (resolution, duration) with background loading
- Playlist playback using libVLC
- Keyboard shortcuts for playback and file operations
- Optional FFmpeg tools (trim, flip) if enabled in the configuration file
- Built-in FFmpeg video combining with stream-copy and H.264/AAC fallback
- Background status, progress, and cancellation for long operations

## Running the Windows Application

After downloading/cloning the repository, run:

```
bin/x64/Release/MediaExplorer.exe
```

All required VLC runtime files are included.

The optional trim, flip, combine, and extended-properties features use FFmpeg.
Install its current Windows build with:

```powershell
winget install --id Gyan.FFmpeg --exact
```

When no explicit `ffmpeg_path` or `ffprobe_path` is configured, MediaExplorer
uses WinGet's stable command links when available, then falls back to `PATH`.

## Running the Linux Application

See [`linux/README.md`](linux/README.md) for Debian dependencies, GNU Make
build commands, installation, mapped-drive configuration, and smoke tests.

## Building with Visual Studio 2022

1. Open `mediaexplorer.sln`
2. Select:
   - Configuration: Release
   - Platform: x64
3. Build the solution

libVLC headers, import libraries, and plugins are included under `vlclib/`.

## Configuration (mediaexplorer.ini)

Place this file next to MediaExplorer.exe to enable optional features:

```
upscaleDirectory = w:\upscale\autosubmit
ffmpegAvailable  = 1
ffprobeAvailable = 1
videoCombineAvailable = 1
loggingEnabled   = 1
loggingPath      = C:\mediaexplorer_logs
```

## Folder Structure (Simplified)

```
media_explorer/
    README.md
    MediaExplorer.cpp
    mediaexplorer.sln
    MediaExplorer.vcxproj
    linux/
        Makefile
        src/
            main.cpp
            MediaExplorerWindow.cpp
        build.sh
        install.sh
        README.md
    vlclib/
        include/
        lib/
        bin/
        plugins/
    bin/
        x64/
            Release/
                MediaExplorer.exe
                plugins/
```

## License

No license has been selected for this repository.
