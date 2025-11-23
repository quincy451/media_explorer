# MediaExplorer

MediaExplorer is a Windows desktop application written in C++17 using Win32 and libVLC. This repository includes full source code, Visual Studio 2022 project files, and prebuilt binaries (EXE, DLLs, and VLC plugins) so the application can run without installing additional dependencies.

## Features

- Fast drive and folder browsing
- Recursive video search
- Video metadata (resolution, duration) with background loading
- Playlist playback using libVLC
- Keyboard shortcuts for playback and file operations
- Optional FFmpeg tools (trim, flip) if enabled in the configuration file
- Optional video combining if external tool is provided
- Background worker windows for long operations

## Running the Application

After downloading/cloning the repository, run:

```
bin/x64/Release/MediaExplorer.exe
```

All required VLC runtime files are included.

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
    main.cpp
    mediaexplorer.sln
    MediaExplorer.vcxproj
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

Insert preferred license here (MIT recommended).

