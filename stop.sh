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

force_stop_port() {
    local port="$1"
    local pids
    pids=$(lsof -ti :"$port" || true)
    if [[ -n "$pids" ]]; then
        echo "Force killing processes on port $port: $pids"
        kill $pids >/dev/null 2>&1 || true
    fi
}

stop_pid "$PID_DIR/frontend.pid"
stop_pid "$PID_DIR/backend.pid"

# Fallback: ensure ports are freed even if pid files are stale or missing
force_stop_port 3010
force_stop_port 5010

echo "Stopped background servers (if running)."
