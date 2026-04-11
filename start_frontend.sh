#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
./render_frontend_assets.sh
cd frontend
npx http-server -p 3010 --cors
