#!/usr/bin/env python3
"""Non-destructive core integration checks for a Linux deployment."""

from __future__ import annotations

import importlib.util
import json
import os
import sys
import tempfile
import threading
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PyQt5.QtWidgets import QApplication


def load_module(path: Path):
    spec = importlib.util.spec_from_file_location("media_explorer_core_test", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def signal_result(task):
    captured = []
    task.signals.finished.connect(captured.append)
    task.run()
    if len(captured) != 1:
        raise AssertionError(f"worker emitted {len(captured)} results")
    return captured[0]


def main() -> int:
    source = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path(__file__).resolve().parents[1] / "media_explorer.py"
    module = load_module(source)
    report = {}

    with tempfile.TemporaryDirectory(prefix="media-explorer-core-") as raw_temp:
        root = Path(raw_temp)
        source_dir = root / "source"
        destination = root / "destination"
        source_dir.mkdir()
        destination.mkdir()
        (source_dir / "child").mkdir()
        (source_dir / "alpha beta.MP4").write_bytes(b"not-real-video")
        (source_dir / "alpha.txt").write_text("ignored", encoding="utf-8")

        config = module.AppConfig()
        config.metadata_prefetch_limit = 0
        folder_task = module.FolderScanTask(str(source_dir), 7, config, threading.Event())
        folder_result = signal_result(folder_task)
        names = sorted(entry.name for entry in folder_result["entries"])
        assert names == ["alpha beta.MP4", "child"], names
        report["folder_scan"] = names

        search_task = module.SearchTask(
            [str(source_dir)], ["alpha", "beta"], config, threading.Event(), 11
        )
        search_result = signal_result(search_task)
        assert [Path(entry.path).name for entry in search_result["entries"]] == ["alpha beta.MP4"]
        report["and_search_matches"] = len(search_result["entries"])

        copy_task = module.FileOperationTask(
            "copy", [str(source_dir / "alpha beta.MP4")], str(destination), threading.Event()
        )
        copy_result = signal_result(copy_task)
        assert copy_result["completed"] == 1 and (destination / "alpha beta.MP4").is_file()
        report["copy"] = copy_result["completed"]

        dangling = destination / "dangling.MP4"
        dangling.symlink_to(root / "outside-target")
        candidate = module._available_destination(str(dangling))
        assert candidate != str(dangling), "dangling symlink was treated as an unused destination"
        report["dangling_symlink_collision"] = Path(candidate).name

        missing_config = root / "does-not-exist.ini"
        try:
            module.AppConfig.load(str(missing_config))
        except ValueError:
            report["missing_explicit_config_rejected"] = True
        else:
            raise AssertionError("missing explicit configuration silently fell back")

        app = QApplication.instance() or QApplication(["media-explorer-core-smoke"])
        window = module.MediaExplorerWindow(config)
        app.processEvents()

        sentinel = list(window.entries)
        window.current_dir = str(source_dir)
        window.view_mode = "searching"
        window._scan_finished(
            {
                "generation": window.scan_generation,
                "path": str(source_dir),
                "entries": [module.Entry("stale", str(source_dir / "stale"), True, "Folder")],
                "cancelled": False,
                "error": "",
            }
        )
        assert window.entries == sentinel, "stale folder scan replaced a search view"
        window.view_mode = "folder"
        window._search_finished(
            {
                "token": window.search_token,
                "entries": [module.Entry("stale", str(source_dir / "stale.MP4"), False, "Video")],
                "cancelled": False,
                "directories": 1,
                "files": 1,
                "errors": [],
            }
        )
        assert window.entries == sentinel, "stale search replaced a folder view"
        report["stale_worker_results_ignored"] = True

        stat_result = (source_dir / "alpha beta.MP4").stat()
        cached_entry = module.Entry(
            "alpha beta.MP4",
            str(source_dir / "alpha beta.MP4"),
            False,
            "MP4 video",
            stat_result.st_size,
            stat_result.st_mtime,
            stat_result.st_mtime_ns,
        )
        window.entries = [cached_entry]
        window._sort_and_fill()
        window.metadata_cache[cached_entry.path] = (
            cached_entry.mtime_ns, cached_entry.size, 1920, 1080, 65.0, "h264", ""
        )
        window._queue_metadata()
        assert window.table.item(0, 4).text() == "1920x1080"
        assert window.table.item(0, 5).text() == "1:05"
        report["cached_metadata_rendered"] = True

        window.close()
        app.processEvents()
        report["window_constructed"] = True

    report["ok"] = True
    print(json.dumps(report, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
