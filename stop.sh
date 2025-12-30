#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_DIR="$ROOT_DIR/.pids"

stop_pid() {
    local pid_file="$1"
    if [[ -f "$pid_file" ]]; then
        local pid
        pid="$(cat "$pid_file")"
        if [[ -n "$pid" ]] && kill -0 "$pid" >/dev/null 2>&1; then
            kill "$pid"
        fi
        rm -f "$pid_file"
    fi
}

stop_pid "$PID_DIR/frontend.pid"
stop_pid "$PID_DIR/backend.pid"

echo "Stopped background servers (if running)."
