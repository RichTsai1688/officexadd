#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_DIR="$ROOT_DIR/.pids"

PYTHON_BIN="python3"
if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    PYTHON_BIN="python"
fi

"$ROOT_DIR/stop.sh"

mkdir -p "$PID_DIR"

"$PYTHON_BIN" "$ROOT_DIR/backend/app.py" > "$ROOT_DIR/backend/server.log" 2>&1 &
echo $! > "$PID_DIR/backend.pid"

(cd "$ROOT_DIR/frontend" && npx http-server -p 3010 --cors > "$ROOT_DIR/frontend/frontend.log" 2>&1 & echo $! > "$PID_DIR/frontend.pid")

echo "Restarted backend (5010) and frontend (3010)."
