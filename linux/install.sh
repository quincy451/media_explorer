#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"
SOURCE_DIR="${SCRIPT_DIR}/dist/MediaExplorer"
PREFIX="/opt/media-explorer"
USER_INSTALL=0
CREATE_DESKTOP=1

usage() {
    echo "Usage: $0 [--source DIR] [--prefix DIR] [--user] [--no-desktop-shortcut]" >&2
}

while (($#)); do
    case "$1" in
        --source)
            shift; [[ $# -gt 0 ]] || { usage; exit 2; }
            SOURCE_DIR="$1"
            ;;
        --prefix)
            shift; [[ $# -gt 0 ]] || { usage; exit 2; }
            PREFIX="$1"
            ;;
        --user)
            USER_INSTALL=1
            PREFIX="${HOME}/.local/opt/media-explorer"
            ;;
        --no-desktop-shortcut) CREATE_DESKTOP=0 ;;
        -h|--help) usage; exit 0 ;;
        *) usage; exit 2 ;;
    esac
    shift
done

SOURCE_DIR="$(cd -- "$SOURCE_DIR" 2>/dev/null && pwd -P)" || {
    echo "Build output does not exist: $SOURCE_DIR" >&2
    exit 2
}
if [[ ! -x "${SOURCE_DIR}/MediaExplorer" ]]; then
    echo "Missing executable: ${SOURCE_DIR}/MediaExplorer" >&2
    exit 2
fi

if ((USER_INSTALL)); then
    APPLICATION_DIR="${HOME}/.local/share/applications"
    ICON_DIR="${HOME}/.local/share/icons/hicolor/scalable/apps"
    CONFIG_DIR="${HOME}/.config/media-explorer"
    INSTALL=(install)
    COPY=(cp)
else
    if [[ ${EUID} -ne 0 ]]; then
        echo "System installation requires root. Use sudo, or pass --user." >&2
        exit 2
    fi
    APPLICATION_DIR="/usr/share/applications"
    ICON_DIR="/usr/share/icons/hicolor/scalable/apps"
    TARGET_USER="${SUDO_USER:-${USER}}"
    TARGET_HOME="$(getent passwd "$TARGET_USER" | cut -d: -f6)"
    CONFIG_DIR="${TARGET_HOME}/.config/media-explorer"
    INSTALL=(install)
    COPY=(cp)
fi

"${INSTALL[@]}" -d -m 755 "$PREFIX" "$APPLICATION_DIR" "$ICON_DIR"
"${COPY[@]}" -a -- "${SOURCE_DIR}/." "${PREFIX}/"
if ((!USER_INSTALL)); then
    # Do not preserve the staging user's ownership inside a system prefix.
    chown -R root:root "$PREFIX"
    chmod -R go-w "$PREFIX"
fi
"${INSTALL[@]}" -m 644 "${SCRIPT_DIR}/assets/media-explorer.svg" "${ICON_DIR}/media-explorer.svg"

DESKTOP_TMP="$(mktemp)"
trap 'rm -f -- "$DESKTOP_TMP"' EXIT
ESCAPED_EXEC="${PREFIX//|/\\|}/MediaExplorer"
ESCAPED_ICON="${ICON_DIR//|/\\|}/media-explorer.svg"
sed -e "s|@EXEC@|${ESCAPED_EXEC}|g" -e "s|@ICON@|${ESCAPED_ICON}|g" \
    "${SCRIPT_DIR}/MediaExplorer.desktop.in" >"$DESKTOP_TMP"
"${INSTALL[@]}" -m 644 "$DESKTOP_TMP" "${APPLICATION_DIR}/media-explorer.desktop"

"${INSTALL[@]}" -d -m 700 "$CONFIG_DIR"
if [[ ! -e "${CONFIG_DIR}/mediaexplorer.ini" ]]; then
    "${INSTALL[@]}" -m 600 "${SCRIPT_DIR}/mediaexplorer.ini.example" "${CONFIG_DIR}/mediaexplorer.ini"
fi
if ((!USER_INSTALL)); then
    chown -R "${TARGET_USER}:${TARGET_USER}" "$CONFIG_DIR"
fi

if ((CREATE_DESKTOP)); then
    if ((USER_INSTALL)); then
        DESKTOP_DIR="$(xdg-user-dir DESKTOP 2>/dev/null || true)"
        [[ -n "$DESKTOP_DIR" ]] || DESKTOP_DIR="${HOME}/Desktop"
        "${INSTALL[@]}" -d -m 755 "$DESKTOP_DIR"
        "${INSTALL[@]}" -m 755 "$DESKTOP_TMP" "${DESKTOP_DIR}/Media Explorer.desktop"
        gio set "${DESKTOP_DIR}/Media Explorer.desktop" metadata::trusted true >/dev/null 2>&1 || true
    else
        DESKTOP_DIR="$(runuser -u "$TARGET_USER" -- xdg-user-dir DESKTOP 2>/dev/null || true)"
        [[ -n "$DESKTOP_DIR" ]] || DESKTOP_DIR="${TARGET_HOME}/Desktop"
        "${INSTALL[@]}" -d -m 755 -o "$TARGET_USER" -g "$TARGET_USER" "$DESKTOP_DIR"
        "${INSTALL[@]}" -m 755 -o "$TARGET_USER" -g "$TARGET_USER" \
            "$DESKTOP_TMP" "${DESKTOP_DIR}/Media Explorer.desktop"
        TARGET_UID="$(id -u "$TARGET_USER")"
        runuser -u "$TARGET_USER" -- env \
            XDG_RUNTIME_DIR="/run/user/${TARGET_UID}" \
            DBUS_SESSION_BUS_ADDRESS="unix:path=/run/user/${TARGET_UID}/bus" \
            gio set "${DESKTOP_DIR}/Media Explorer.desktop" \
                metadata::trusted true >/dev/null 2>&1 || true
    fi
fi

update-desktop-database "$APPLICATION_DIR" >/dev/null 2>&1 || true
gtk-update-icon-cache -f "$(dirname "$(dirname "$ICON_DIR")")" >/dev/null 2>&1 || true
if ((USER_INSTALL)); then
    QT_QPA_PLATFORM=offscreen "${PREFIX}/MediaExplorer" --self-test
else
    TARGET_UID="$(id -u "$TARGET_USER")"
    runuser -u "$TARGET_USER" -- env QT_QPA_PLATFORM=offscreen \
        XDG_RUNTIME_DIR="/run/user/${TARGET_UID}" "${PREFIX}/MediaExplorer" --self-test
fi
echo "Installed Media Explorer in ${PREFIX}"
if ((CREATE_DESKTOP)); then
    echo "Desktop shortcut: ${DESKTOP_DIR}/Media Explorer.desktop"
fi
