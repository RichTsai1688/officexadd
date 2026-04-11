#!/bin/sh
set -eu

: "${OFFICEXADD_PUBLIC_ORIGIN:?OFFICEXADD_PUBLIC_ORIGIN is required. Set it to the Cloudflare Tunnel hostname for this add-in.}"

if [ -z "${OFFICEXADD_API_TOKEN:-}" ]; then
  echo "ERROR: OFFICEXADD_API_TOKEN is required to start nginx service." >&2
  echo "Run with: OFFICEXADD_API_TOKEN=your-token docker compose up -d" >&2
  exit 1
fi

export OFFICEXADD_API_TOKEN
export OFFICEXADD_PUBLIC_ORIGIN

TOKEN_ESCAPED=$(printf '%s' "$OFFICEXADD_API_TOKEN" | sed 's/[\\|&]/\\&/g')
sed "s|__OFFICEXADD_API_TOKEN__|${TOKEN_ESCAPED}|g" /etc/nginx/nginx.conf > /etc/nginx/nginx-rendered.conf

ORIGIN_ESCAPED=$(printf '%s' "$OFFICEXADD_PUBLIC_ORIGIN" | sed 's/[\\|&]/\\&/g')

sed \
  -e "s|__OFFICEXADD_PUBLIC_ORIGIN__|${ORIGIN_ESCAPED}|g" \
  -e "s|__OFFICEXADD_API_TOKEN__|${TOKEN_ESCAPED}|g" \
  /usr/share/nginx/html/config.template.js > /usr/share/nginx/html/config.js

sed \
  -e "s|__OFFICEXADD_PUBLIC_ORIGIN__|${ORIGIN_ESCAPED}|g" \
  /usr/share/nginx/html/manifest.template.xml > /usr/share/nginx/html/manifest.xml

exec nginx -c /etc/nginx/nginx-rendered.conf -g 'daemon off;'
