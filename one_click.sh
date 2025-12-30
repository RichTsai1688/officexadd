#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BACKEND_DIR="$ROOT_DIR/backend"
FRONTEND_DIR="$ROOT_DIR/frontend"
PID_DIR="$ROOT_DIR/.pids"

require_cmd() {
    local name="$1"
    if ! command -v "$name" >/dev/null 2>&1; then
        echo "Missing required command: $name"
        exit 1
    fi
}

require_cmd npm
require_cmd npx

PYTHON_BIN="python3"
if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
    PYTHON_BIN="python"
fi
require_cmd "$PYTHON_BIN"

if [[ ! -f "$BACKEND_DIR/.env" ]]; then
    if [[ -f "$BACKEND_DIR/.env.example" ]]; then
        cp "$BACKEND_DIR/.env.example" "$BACKEND_DIR/.env"
    fi
    echo "Missing backend/.env. Fill it in and re-run."
    exit 1
fi

echo "Installing backend dependencies..."
"$PYTHON_BIN" -m pip install -r "$BACKEND_DIR/requirements.txt"

echo "Installing frontend tools..."
npm install

mkdir -p "$PID_DIR"

if command -v lsof >/dev/null 2>&1; then
    if lsof -nP -iTCP:5001 -sTCP:LISTEN >/dev/null 2>&1; then
        echo "Port 5001 is already in use. Stop the current process or change the backend port."
        exit 1
    fi
    if lsof -nP -iTCP:3000 -sTCP:LISTEN >/dev/null 2>&1; then
        echo "Port 3000 is already in use. Stop the current process or change the frontend port."
        exit 1
    fi
fi

echo "Starting backend..."
"$PYTHON_BIN" "$BACKEND_DIR/app.py" > "$BACKEND_DIR/server.log" 2>&1 &
echo $! > "$PID_DIR/backend.pid"

echo "Starting frontend..."
(cd "$FRONTEND_DIR" && npx http-server -p 3000 --cors > "$FRONTEND_DIR/frontend.log" 2>&1 & echo $! > "$PID_DIR/frontend.pid")

echo "Starting sideload..."
npx office-addin-debugging start frontend/manifest.xml
