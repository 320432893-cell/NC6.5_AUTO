#!/usr/bin/env bash
set -euo pipefail

SRC_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
DEST_DIR="${NC_AUTO_WINDOWS_DIR:-/mnt/h/python脚本/.venv/nc_auto_v2}"
INTERVAL="${NC_AUTO_SYNC_INTERVAL:-2}"
MODE="once"
DRY_RUN=0

usage() {
  cat <<'EOF'
Usage: tools/sync_to_windows.sh [--watch] [--dry-run] [--dest PATH]

Sync the WSL source repo to the Windows/H-drive JAB runtime mirror.

Options:
  --watch       Keep syncing every NC_AUTO_SYNC_INTERVAL seconds, default 2.
  --dry-run     Show what would change without writing files.
  --dest PATH   Override destination path.

Environment:
  NC_AUTO_WINDOWS_DIR      Destination path, default /mnt/h/python脚本/.venv/nc_auto_v2
  NC_AUTO_SYNC_INTERVAL    Watch interval seconds, default 2
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --watch)
      MODE="watch"
      shift
      ;;
    --dry-run)
      DRY_RUN=1
      shift
      ;;
    --dest)
      DEST_DIR="${2:?--dest requires a path}"
      shift 2
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $1" >&2
      usage >&2
      exit 2
      ;;
  esac
done

RSYNC_ARGS=(
  -a
  --delete
  --exclude .git
  --exclude .venv-local
  --exclude .agents
  --exclude .codex
  --exclude .vscode
  --exclude __pycache__
  --exclude '*.pyc'
  --exclude logs
  --exclude .pytest_cache
  --exclude tools/jab_probe_output.txt
)

if [[ "$DRY_RUN" -eq 1 ]]; then
  RSYNC_ARGS+=(--dry-run --itemize-changes)
fi

sync_once() {
  mkdir -p "$DEST_DIR"
  rsync "${RSYNC_ARGS[@]}" "$SRC_DIR/" "$DEST_DIR/"
  printf '[sync] %s -> %s\n' "$SRC_DIR" "$DEST_DIR"
}

if [[ "$MODE" == "watch" ]]; then
  echo "[sync] watch mode: interval=${INTERVAL}s, dest=${DEST_DIR}"
  while true; do
    sync_once
    sleep "$INTERVAL"
  done
else
  sync_once
fi
