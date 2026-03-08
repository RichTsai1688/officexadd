#!/bin/sh
set -eu

: "${OFFICEXADD_PUBLIC_ORIGIN:=https://fcu.labelnine.app:2053}"

if [ -z "${OFFICEXADD_API_TOKEN:-}" ]; then
  echo "ERROR: OFFICEXADD_API_TOKEN is required to start nginx service." >&2
  echo "Run with: OFFICEXADD_API_TOKEN=your-token docker compose up -d" >&2
  exit 1
fi

export OFFICEXADD_API_TOKEN
export OFFICEXADD_PUBLIC_ORIGIN

TOKEN_ESCAPED=$(printf '%s' "$OFFICEXADD_API_TOKEN" | sed 's/[\\|&]/\\&/g')
sed "s|__OFFICEXADD_API_TOKEN__|${TOKEN_ESCAPED}|g" /etc/nginx/nginx.conf > /etc/nginx/nginx-rendered.conf

cat <<EOCONFIG > /usr/share/nginx/html/config.js
window.__OFFICEXADD_CONFIG__ = {
  apiBaseUrl: "${OFFICEXADD_PUBLIC_ORIGIN}",
  apiToken: "${OFFICEXADD_API_TOKEN}",
};
EOCONFIG

exec nginx -c /etc/nginx/nginx-rendered.conf -g 'daemon off;'
