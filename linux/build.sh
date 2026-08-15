#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"
INSTALL_DEPS=0
CLEAN=0

usage() {
    echo "Usage: $0 [--install-deps] [--clean] [--jobs N]" >&2
}

JOBS="$(getconf _NPROCESSORS_ONLN 2>/dev/null || echo 2)"
while (($#)); do
    case "$1" in
        --install-deps) INSTALL_DEPS=1 ;;
        --clean) CLEAN=1 ;;
        --jobs)
            shift; [[ $# -gt 0 && "$1" =~ ^[1-9][0-9]*$ ]] || { usage; exit 2; }
            JOBS="$1"
            ;;
        -h|--help) usage; exit 0 ;;
        *) usage; exit 2 ;;
    esac
    shift
done

if ((INSTALL_DEPS)); then
    command -v sudo >/dev/null 2>&1 || { echo "sudo is required for --install-deps" >&2; exit 2; }
    sudo apt-get update
    sudo apt-get install -y \
        build-essential pkg-config qtbase5-dev libvlc-dev vlc ffmpeg \
        libglib2.0-bin xdg-utils desktop-file-utils
fi

for command in make g++ pkg-config ffmpeg ffprobe vlc; do
    command -v "$command" >/dev/null 2>&1 || {
        echo "Missing required command: $command" >&2
        echo "Run: $0 --install-deps" >&2
        exit 2
    }
done
pkg-config --exists Qt5Widgets libvlc || {
    echo "Missing Qt5 Widgets or libVLC development files. Run: $0 --install-deps" >&2
    exit 2
}

cd -- "$SCRIPT_DIR"
if ((CLEAN)); then
    make clean
fi
make -j"$JOBS" all
make test
echo "Built and tested: ${SCRIPT_DIR}/dist/MediaExplorer/MediaExplorer"
