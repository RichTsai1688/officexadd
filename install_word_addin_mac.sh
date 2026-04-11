#!/usr/bin/env bash
set -euo pipefail

DEFAULT_ORIGIN="https://office-addin.labelnine.app"
DEFAULT_MANIFEST_URL="${DEFAULT_ORIGIN}/manifest.xml"
WEF_DIR="${HOME}/Library/Containers/com.microsoft.Word/Data/Documents/wef"
TARGET_NAME="officexadd-manifest.xml"
TARGET_PATH="${WEF_DIR}/${TARGET_NAME}"
TMP_MANIFEST="$(mktemp -t officexadd-manifest.XXXXXX.xml)"
RESTART_WORD=1
MANIFEST_SOURCE="${1:-$DEFAULT_MANIFEST_URL}"

cleanup() {
    rm -f "$TMP_MANIFEST"
}
trap cleanup EXIT

usage() {
    cat <<EOF
Usage:
  $0 [manifest-url-or-local-path] [--no-restart]

Examples:
  $0
  $0 https://office-addin.labelnine.app/manifest.xml
  $0 ./frontend/manifest.xml --no-restart
EOF
}

for arg in "$@"; do
    case "$arg" in
        --no-restart)
            RESTART_WORD=0
            ;;
        -h|--help)
            usage
            exit 0
            ;;
    esac
done

fetch_manifest() {
    if [[ "$MANIFEST_SOURCE" =~ ^https?:// ]]; then
        curl -fsSL "$MANIFEST_SOURCE" -o "$TMP_MANIFEST"
    else
        cp "$MANIFEST_SOURCE" "$TMP_MANIFEST"
    fi
}

validate_manifest() {
    if ! grep -q "<OfficeApp" "$TMP_MANIFEST"; then
        echo "Manifest does not look valid: $MANIFEST_SOURCE" >&2
        exit 1
    fi
}

install_manifest() {
    mkdir -p "$WEF_DIR"
    cp "$TMP_MANIFEST" "$TARGET_PATH"
}

restart_word() {
    if [[ "$RESTART_WORD" -eq 0 ]]; then
        return
    fi

    osascript <<'EOF' >/dev/null 2>&1 || true
tell application "Microsoft Word"
    if running then
        quit saving no
    end if
end tell
EOF

    open -a "Microsoft Word" || true
}

fetch_manifest
validate_manifest
install_manifest
restart_word

echo "Installed manifest to: $TARGET_PATH"
if [[ "$RESTART_WORD" -eq 1 ]]; then
    echo "Microsoft Word was restarted. Open Insert > My Add-ins > Developer Add-ins to load OfficeXAdd."
else
    echo "Restart Microsoft Word manually, then open Insert > My Add-ins > Developer Add-ins to load OfficeXAdd."
fi
