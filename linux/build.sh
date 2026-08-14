#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"
BUILD_VENV="${MEDIA_EXPLORER_BUILD_VENV:-${SCRIPT_DIR}/.build-venv}"
INSTALL_DEPS=0

usage() {
    echo "Usage: $0 [--install-deps] [--clean]" >&2
}

CLEAN=0
while (($#)); do
    case "$1" in
        --install-deps) INSTALL_DEPS=1 ;;
        --clean) CLEAN=1 ;;
        -h|--help) usage; exit 0 ;;
        *) usage; exit 2 ;;
    esac
    shift
done

if ((INSTALL_DEPS)); then
    if ! command -v sudo >/dev/null 2>&1; then
        echo "sudo is required for --install-deps" >&2
        exit 2
    fi
    sudo apt-get update
    sudo apt-get install -y \
        python3 python3-venv python3-pyqt5 python3-vlc python3-pip \
        pyinstaller vlc ffmpeg libglib2.0-bin xdg-utils
fi

for command in python3 ffprobe ffmpeg vlc; do
    if ! command -v "$command" >/dev/null 2>&1; then
        echo "Missing required command: $command" >&2
        echo "Run: $0 --install-deps" >&2
        exit 2
    fi
done

if ((CLEAN)); then
    # These targets are fixed inside this source directory; no broad/globbed
    # deletion is used.
    rm -rf -- "${SCRIPT_DIR}/build" "${SCRIPT_DIR}/dist"
fi

BUILD_PYTHON=python3
if ! python3 -c 'import PyQt5, vlc, PyInstaller' >/dev/null 2>&1; then
    if [[ ! -x "${BUILD_VENV}/bin/python" ]]; then
        python3 -m venv --system-site-packages "${BUILD_VENV}"
    fi
    BUILD_PYTHON="${BUILD_VENV}/bin/python"
    if ! "$BUILD_PYTHON" -c 'import vlc' >/dev/null 2>&1; then
        "$BUILD_PYTHON" -m pip install 'python-vlc>=3.0.20,<4'
    fi
    if ! "$BUILD_PYTHON" -c 'import PyInstaller' >/dev/null 2>&1; then
        "$BUILD_PYTHON" -m pip install 'PyInstaller>=6.13,<7'
    fi
fi

if ! "$BUILD_PYTHON" -c 'import PyQt5, vlc, PyInstaller' >/dev/null 2>&1; then
    echo "Python build modules are unavailable. Install python3-pyqt5, python3-vlc, and pyinstaller." >&2
    exit 2
fi

cd -- "$SCRIPT_DIR"
"$BUILD_PYTHON" -m PyInstaller --noconfirm --clean media-explorer.spec

if [[ ! -x "${SCRIPT_DIR}/dist/MediaExplorer/MediaExplorer" ]]; then
    echo "Build did not produce dist/MediaExplorer/MediaExplorer" >&2
    exit 1
fi

QT_QPA_PLATFORM=offscreen "${SCRIPT_DIR}/dist/MediaExplorer/MediaExplorer" --self-test
echo "Built: ${SCRIPT_DIR}/dist/MediaExplorer/MediaExplorer"
