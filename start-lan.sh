#!/usr/bin/env bash
# Run the AIS LAN dashboard in THIS terminal (foreground).
# Leave the window open while the site is in use. Ctrl+C to stop.
#
#   cd ~/weighbridge-data-entry
#   ./start-lan.sh

set -euo pipefail
cd "$(dirname "$0")"

export PATH="${HOME}/.local/node/bin:${HOME}/.local/bin:${PATH}"

PORT="${PORT:-${AIS_DASHBOARD_PORT:-5000}}"

# Signature-only kiosk tablet(s) — comma-separated lists OK.
export AIS_SIGNATURE_ONLY_MACS="${AIS_SIGNATURE_ONLY_MACS:-e2:4c:e6:97:59:29}"
export AIS_SIGNATURE_ONLY_IPS="${AIS_SIGNATURE_ONLY_IPS:-192.168.1.171}"

# Login once → stay signed in (cookie, 7 days). Change password here.
export AIS_DASHBOARD_USER="${AIS_DASHBOARD_USER:-admin}"
export AIS_DASHBOARD_PASSWORD="${AIS_DASHBOARD_PASSWORD:-ais5626}"
export AIS_SESSION_HOURS="${AIS_SESSION_HOURS:-168}"
# AIS_AUTH_DISABLED=1  → no login

# API lines in this terminal: api | all | 0
export AIS_ACCESS_LOG="${AIS_ACCESS_LOG:-api}"

if ! command -v node >/dev/null 2>&1; then
  echo "ERROR: node not found."
  echo "  Expected at: \$HOME/.local/node/bin/node"
  echo "  PATH is: $PATH"
  exit 1
fi

if [[ ! -f server.js ]]; then
  echo "ERROR: server.js not found in $(pwd)"
  exit 1
fi

# Stop any previous AIS dashboard instance (node running server.js).
stop_old_servers() {
  local p cmd
  for p in $(pgrep -x node 2>/dev/null || true); do
    cmd="$(tr '\0' ' ' < "/proc/$p/cmdline" 2>/dev/null || true)"
    case "$cmd" in
      *server.js*) kill "$p" 2>/dev/null || true ;;
    esac
  done
  # Also free the port if something else is stuck on it (same user).
  if command -v fuser >/dev/null 2>&1; then
    fuser -k "${PORT}/tcp" >/dev/null 2>&1 || true
  elif command -v lsof >/dev/null 2>&1; then
    local pids
    pids="$(lsof -tiTCP:"${PORT}" -sTCP:LISTEN 2>/dev/null || true)"
    if [[ -n "${pids}" ]]; then
      # shellcheck disable=SC2086
      kill ${pids} 2>/dev/null || true
    fi
  fi
}

port_in_use() {
  if command -v ss >/dev/null 2>&1; then
    ss -tln 2>/dev/null | grep -qE ":${PORT}\\s"
  else
    return 1
  fi
}

stop_old_servers
# Wait up to ~3s for the port to free (avoids EADDRINUSE race).
for _ in 1 2 3 4 5 6; do
  port_in_use || break
  sleep 0.5
  stop_old_servers
done
if port_in_use; then
  echo "ERROR: port ${PORT} is still in use after stopping old servers."
  echo "  Try:  ss -tlnp | grep ${PORT}"
  echo "  Or:   kill the process listed, then re-run ./start-lan.sh"
  exit 1
fi

LAN_IP="$(hostname -I 2>/dev/null | awk '{print $1}')"
echo "=============================================="
echo "  AIS LAN dashboard — HOST TERMINAL"
echo "  Keep this window open. Ctrl+C to stop."
echo "----------------------------------------------"
echo "  This PC:  http://127.0.0.1:${PORT}"
if [ -n "${LAN_IP}" ]; then
  echo "  LAN:      http://${LAN_IP}:${PORT}"
fi
echo "  Node:     $(command -v node) ($(node -v 2>/dev/null || echo '?'))"
echo "  Folder:   $(pwd)"
echo "=============================================="
echo

exec node --no-warnings server.js
