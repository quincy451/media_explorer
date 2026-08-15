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
[[ -x "${SOURCE_DIR}/MediaExplorer" ]] || {
    echo "Missing executable: ${SOURCE_DIR}/MediaExplorer" >&2
    exit 2
}

[[ "$PREFIX" == /* ]] || {
    echo "Install prefix must be an absolute path: $PREFIX" >&2
    exit 2
}
PREFIX="$(readlink -m -- "$PREFIX")"
[[ -n "$PREFIX" && "$PREFIX" != "/" ]] || {
    echo "Refusing unsafe install prefix: $PREFIX" >&2
    exit 2
}

if ((USER_INSTALL)); then
    APPLICATION_DIR="${HOME}/.local/share/applications"
    ICON_DIR="${HOME}/.local/share/icons/hicolor/scalable/apps"
    CONFIG_DIR="${HOME}/.config/media-explorer"
    TARGET_USER="$(id -un)"
    TARGET_HOME="$HOME"
else
    [[ ${EUID} -eq 0 ]] || {
        echo "System installation requires root. Use sudo, or pass --user." >&2
        exit 2
    }
    APPLICATION_DIR="/usr/share/applications"
    ICON_DIR="/usr/share/icons/hicolor/scalable/apps"
    TARGET_USER="${SUDO_USER:-${USER}}"
    TARGET_HOME="$(getent passwd "$TARGET_USER" | cut -d: -f6)"
    [[ -n "$TARGET_HOME" ]] || { echo "Could not resolve home for $TARGET_USER" >&2; exit 2; }
    CONFIG_DIR="${TARGET_HOME}/.config/media-explorer"
fi

install -d -m 755 "$PREFIX" "$APPLICATION_DIR" "$ICON_DIR"
install -m 755 "${SOURCE_DIR}/MediaExplorer" "${PREFIX}/MediaExplorer"
install -m 644 "${SCRIPT_DIR}/assets/media-explorer.svg" "${ICON_DIR}/media-explorer.svg"

remove_legacy_pyinstaller_payload() {
    local legacy="${PREFIX}/_internal"
    local resolved_prefix resolved_legacy
    local -a python_libraries=()
    [[ -e "$legacy" || -L "$legacy" ]] || return 0
    if [[ -L "$legacy" || ! -d "$legacy" ]]; then
        echo "Leaving unrecognized non-directory legacy path untouched: $legacy" >&2
        return 0
    fi
    resolved_prefix="$(cd -- "$PREFIX" && pwd -P)"
    resolved_legacy="$(cd -- "$legacy" && pwd -P)"
    if [[ "$resolved_legacy" != "${resolved_prefix}/_internal" ||
          ! -f "${legacy}/base_library.zip" || -L "${legacy}/base_library.zip" ||
          ! -d "${legacy}/PyQt5" || -L "${legacy}/PyQt5" ]]; then
        echo "Leaving unrecognized _internal directory untouched: $legacy" >&2
        return 0
    fi
    shopt -s nullglob
    python_libraries=("${legacy}"/libpython*.so*)
    shopt -u nullglob
    ((${#python_libraries[@]} > 0)) || {
        echo "Leaving _internal without a libpython marker untouched: $legacy" >&2
        return 0
    }
    local library
    for library in "${python_libraries[@]}"; do
        [[ -f "$library" && ! -L "$library" ]] || {
            echo "Leaving _internal with an unsafe libpython marker untouched: $legacy" >&2
            return 0
        }
    done

    # mountinfo encodes whitespace and backslashes as octal sequences. Decode
    # its fifth field, canonicalize it, and refuse both different-device and
    # same-device bind mounts anywhere in the tree before recursive removal.
    local mount_id parent_id device mount_root encoded_target remainder decoded_target canonical_target
    local mount_status=1 mount_records=0
    if [[ ! -r /proc/self/mountinfo ]]; then
        mount_status=2
    else
        while IFS=' ' read -r mount_id parent_id device mount_root encoded_target remainder; do
            ((mount_records += 1))
            [[ -n "$encoded_target" ]] || { mount_status=2; break; }
            printf -v decoded_target '%b' "$encoded_target"
            IFS= read -r -d '' canonical_target < <(readlink -mz -- "$decoded_target") || {
                mount_status=2
                break
            }
            if [[ "$canonical_target" == "$resolved_legacy" ||
                  "$canonical_target" == "${resolved_legacy}/"* ]]; then
                mount_status=0
                break
            fi
        done </proc/self/mountinfo
        if ((mount_records == 0)); then
            mount_status=2
        fi
    fi
    if ((mount_status == 0)); then
        echo "Leaving _internal containing a mount point untouched: $legacy" >&2
        return 0
    fi
    if ((mount_status != 1)); then
        echo "Leaving _internal untouched because mount safety could not be verified: $legacy" >&2
        return 0
    fi

    rm --one-file-system -rf -- "$legacy"
    echo "Removed validated legacy PyInstaller payload: $legacy"
}

remove_legacy_pyinstaller_payload

escape_sed_replacement() {
    local value="$1"
    value="${value//\\/\\\\}"
    value="${value//&/\\&}"
    value="${value//|/\\|}"
    printf '%s' "$value"
}

DESKTOP_TMP="$(mktemp --suffix=.desktop)"
trap 'rm -f -- "$DESKTOP_TMP"' EXIT
ESCAPED_EXEC="$(escape_sed_replacement "${PREFIX}/MediaExplorer")"
ESCAPED_ICON="$(escape_sed_replacement "${ICON_DIR}/media-explorer.svg")"
sed -e "s|@EXEC@|${ESCAPED_EXEC}|g" -e "s|@ICON@|${ESCAPED_ICON}|g" \
    "${SCRIPT_DIR}/MediaExplorer.desktop.in" >"$DESKTOP_TMP"
if command -v desktop-file-validate >/dev/null 2>&1; then
    desktop-file-validate "$DESKTOP_TMP"
fi
install -m 644 "$DESKTOP_TMP" "${APPLICATION_DIR}/media-explorer.desktop"

install -d -m 700 "$CONFIG_DIR"
if [[ ! -e "${CONFIG_DIR}/mediaexplorer.ini" ]]; then
    install -m 600 "${SCRIPT_DIR}/mediaexplorer.ini.example" \
        "${CONFIG_DIR}/mediaexplorer.ini"
fi
if ((!USER_INSTALL)); then
    chown -R "${TARGET_USER}:${TARGET_USER}" "$CONFIG_DIR"
fi

install_system_desktop_wrapper() {
    local desktop_dir="$1"
    local owner="$2"
    local shortcut="${desktop_dir}/Media Explorer.desktop"
    local template="${SCRIPT_DIR}/MediaExplorer.desktop-link.in"
    local backup_dir="${TARGET_HOME}/.local/state/media-explorer/launcher-backups"
    local backup counter legacy_backup legacy_name temporary
    local -a legacy_backups=()

    install -d -m 755 -o "$owner" -g "$owner" "$desktop_dir"
    install -d -m 700 -o "$owner" -g "$owner" "$backup_dir"
    if [[ -e "${shortcut}.backup" || -L "${shortcut}.backup" ]]; then
        legacy_backups+=("${shortcut}.backup")
    fi
    shopt -s nullglob
    legacy_backups+=("${shortcut}.backup."[0-9]*)
    shopt -u nullglob
    for legacy_backup in "${legacy_backups[@]}"; do
        if [[ -d "$legacy_backup" && ! -L "$legacy_backup" ]]; then
            echo "Leaving unexpected backup directory on Desktop: $legacy_backup" >&2
            continue
        fi
        legacy_name="$(basename -- "$legacy_backup")"
        backup="${backup_dir}/${legacy_name}"
        counter=2
        while [[ -e "$backup" || -L "$backup" ]]; do
            backup="${backup_dir}/${legacy_name}.${counter}"
            ((counter += 1))
        done
        mv -- "$legacy_backup" "$backup"
        chown -h "$owner:$owner" "$backup"
        echo "Moved legacy Desktop backup to: $backup"
    done
    if [[ -f "$shortcut" && ! -L "$shortcut" ]] && cmp -s -- "$template" "$shortcut"; then
        chmod 644 "$shortcut"
        chown "$owner:$owner" "$shortcut"
        return
    fi
    if [[ -d "$shortcut" && ! -L "$shortcut" ]]; then
        echo "Cannot replace desktop launcher because it is a directory: $shortcut" >&2
        return 1
    fi
    if [[ -e "$shortcut" || -L "$shortcut" ]]; then
        backup="${backup_dir}/Media Explorer.desktop.backup"
        counter=2
        while [[ -e "$backup" || -L "$backup" ]]; do
            backup="${backup_dir}/Media Explorer.desktop.backup.${counter}"
            ((counter += 1))
        done
        mv -- "$shortcut" "$backup"
        chown -h "$owner:$owner" "$backup"
        echo "Backed up existing desktop launcher to: $backup"
    fi
    temporary="$(mktemp "${desktop_dir}/.media-explorer-launcher.XXXXXX")"
    install -m 644 -o "$owner" -g "$owner" "$template" "$temporary"
    mv -- "$temporary" "$shortcut"
}

if ((CREATE_DESKTOP)); then
    if ((USER_INSTALL)); then
        echo "Skipping the Desktop wrapper for --user install: PCManFM only trusts a " \
             "Type=Link target registered under /usr/share or /usr/local/share." >&2
    else
        DESKTOP_DIR="$(runuser -u "$TARGET_USER" -- xdg-user-dir DESKTOP 2>/dev/null || true)"
        [[ -n "$DESKTOP_DIR" ]] || DESKTOP_DIR="${TARGET_HOME}/Desktop"
        install_system_desktop_wrapper "$DESKTOP_DIR" "$TARGET_USER"
    fi
fi

update-desktop-database "$APPLICATION_DIR" >/dev/null 2>&1 || true
gtk-update-icon-cache -f "$(dirname "$(dirname "$ICON_DIR")")" >/dev/null 2>&1 || true
if ((USER_INSTALL)); then
    "${PREFIX}/MediaExplorer" --self-test
else
    TARGET_UID="$(id -u "$TARGET_USER")"
    runuser -u "$TARGET_USER" -- env XDG_RUNTIME_DIR="/run/user/${TARGET_UID}" \
        "${PREFIX}/MediaExplorer" --self-test
fi
echo "Installed Media Explorer in ${PREFIX}"
if ((!USER_INSTALL)) && ((CREATE_DESKTOP)); then
    echo "Desktop launcher: ${DESKTOP_DIR}/Media Explorer.desktop"
fi
