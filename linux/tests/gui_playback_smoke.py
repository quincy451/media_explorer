#!/usr/bin/env python3
"""Short on-display integration check for the packaged Media Explorer player.

Usage: python3 gui_playback_smoke.py /path/to/media_explorer.py /path/to/video
The caller must supply DISPLAY/XAUTHORITY (this test is intended for SSH
deployment verification, not the normal build).
"""

from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "xcb")

from PyQt5.QtCore import QTimer
from PyQt5.QtWidgets import QApplication


def main() -> int:
    if len(sys.argv) != 3:
        print("usage: gui_playback_smoke.py MEDIA_EXPLORER_PY VIDEO", file=sys.stderr)
        return 2
    module_path = Path(sys.argv[1]).resolve()
    video_path = Path(sys.argv[2]).resolve()
    if not module_path.is_file() or not video_path.is_file():
        print("module or video does not exist", file=sys.stderr)
        return 2

    spec = importlib.util.spec_from_file_location("media_explorer_under_test", module_path)
    if spec is None or spec.loader is None:
        print("unable to load Media Explorer module", file=sys.stderr)
        return 2
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)

    app = QApplication(["media-explorer-gui-smoke"])
    config = module.AppConfig()
    config.vlc_args = ["--no-video-title-show", "--quiet", "--avcodec-hw=none"]
    window = module.MediaExplorerWindow(config)
    window.show()

    entry = module.Entry(
        video_path.name,
        str(video_path),
        False,
        module.file_kind(str(video_path)),
        video_path.stat().st_size,
        video_path.stat().st_mtime,
        video_path.stat().st_mtime_ns,
    )
    result = {"exit": 1}

    def start_playback() -> None:
        window.play_entries([entry])

    def inspect_playback() -> None:
        player = window._player
        if player is None:
            print("player was not created", file=sys.stderr)
        else:
            state = str(player.get_state())
            size = player.video_get_size(0)
            print(f"state={state} size={size}")
            if "Playing" in state and size and tuple(size) != (0, 0):
                result["exit"] = 0
        window.stop_playback()
        window.close()
        app.exit(result["exit"])

    QTimer.singleShot(350, start_playback)
    QTimer.singleShot(4500, inspect_playback)
    return app.exec_()


if __name__ == "__main__":
    raise SystemExit(main())
