#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
FRONTEND_DIR="$ROOT_DIR/frontend"

if [[ -f "$ROOT_DIR/.env" ]]; then
    set -a
    # shellcheck disable=SC1091
    source "$ROOT_DIR/.env"
    set +a
fi

: "${OFFICEXADD_PUBLIC_ORIGIN:?OFFICEXADD_PUBLIC_ORIGIN is required. Set it to your Cloudflare Tunnel hostname, e.g. https://addin.example.com}"
: "${OFFICEXADD_API_TOKEN:?OFFICEXADD_API_TOKEN is required. Set it before rendering frontend assets.}"

origin_escaped=$(printf '%s' "$OFFICEXADD_PUBLIC_ORIGIN" | sed 's/[\\|&]/\\&/g')
token_escaped=$(printf '%s' "$OFFICEXADD_API_TOKEN" | sed 's/[\\|&]/\\&/g')

sed \
    -e "s|__OFFICEXADD_PUBLIC_ORIGIN__|${origin_escaped}|g" \
    -e "s|__OFFICEXADD_API_TOKEN__|${token_escaped}|g" \
    "$FRONTEND_DIR/config.template.js" > "$FRONTEND_DIR/config.js"

sed \
    -e "s|__OFFICEXADD_PUBLIC_ORIGIN__|${origin_escaped}|g" \
    "$FRONTEND_DIR/manifest.template.xml" > "$FRONTEND_DIR/manifest.xml"

sed \
    -e "s|__OFFICEXADD_PUBLIC_ORIGIN__|${origin_escaped}|g" \
    "$FRONTEND_DIR/manifest-powerpoint.template.xml" > "$FRONTEND_DIR/manifest-powerpoint.xml"

echo "Rendered frontend/config.js, frontend/manifest.xml, and frontend/manifest-powerpoint.xml for ${OFFICEXADD_PUBLIC_ORIGIN}"
