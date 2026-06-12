#!/usr/bin/env bash
# remote_build.sh — host-side driver for the Phase 6 Stage 5 freeze + COM smoke on the VM.
#
# PyInstaller does not cross-compile, so the freeze runs on the Windows guest. The freeze and
# `uv sync` are COM-free, so they run over plain SSH (session 0). The frozen exe's COM smoke
# (--selftest binds OneNote) MUST run in the autologon interactive session, so it goes through
# the `onenote-selftest` scheduled task exactly like the Tier-2 loop (VM-ops lesson: never
# activate OneNote COM over plain SSH).
#
# Builds the freeze only; the Inno Setup installer is a separate step gated on iscc.exe being
# installed on the VM (see packaging/README.md). Mirrors the guest selftest exit code.
set -euo pipefail

GUEST_HOST="${GUEST_HOST:-192.168.122.13}"
GUEST_USER="${GUEST_USER:-dev}"
GUEST_SSH_PORT="${GUEST_SSH_PORT:-22}"
GUEST_REPO="${GUEST_REPO:-C:/onenote-mcp}"
RESULTS_DIR="${RESULTS_DIR:-./test-results}"
POLL_TIMEOUT_S="${POLL_TIMEOUT_S:-1800}"
POLL_INTERVAL_S="${POLL_INTERVAL_S:-10}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

ssh_guest() { ssh -p "$GUEST_SSH_PORT" "${GUEST_USER}@${GUEST_HOST}" "$@"; }
GUEST_REPO_WIN="${GUEST_REPO//\//\\}"
SENTINEL="${GUEST_REPO}/test-results/selftest_exit.txt"
SENTINEL_WIN="${SENTINEL//\//\\}"

mkdir -p "$RESULTS_DIR"

# --- 1. ship code to the guest ----------------------------------------------
echo ">> syncing code to guest (tar over ssh)"
ssh_guest "cmd /c \"if not exist ${GUEST_REPO_WIN} mkdir ${GUEST_REPO_WIN}\""
tar czf - \
  --exclude='.git' --exclude='.venv' --exclude='refs' \
  --exclude='test-results' --exclude='__pycache__' --exclude='*.pyc' \
  --exclude='dist' --exclude='build' \
  -C "$REPO_ROOT" . \
  | ssh_guest "tar -xzf - -C ${GUEST_REPO}"

# --- 2. install build deps + freeze (COM-free → plain SSH) ------------------
echo ">> installing build deps (uv sync --group packaging)"
ssh_guest "cmd /c \"cd /d ${GUEST_REPO_WIN} && uv sync --group packaging\""

echo ">> freezing server (PyInstaller onedir) — this can take a few minutes"
ssh_guest "cmd /c \"cd /d ${GUEST_REPO_WIN} && uv run pyinstaller packaging\\onenote-mcp.spec --noconfirm --distpath dist --workpath build\""

# --- 3. COM smoke of the frozen exe via the interactive task ----------------
echo ">> (re)registering onenote-selftest task"
ssh_guest "schtasks /create /f /tn onenote-selftest /sc once /st 00:00 /it /tr ${GUEST_REPO_WIN}\\packaging\\run_selftest.bat" >/dev/null
echo ">> clearing stale sentinel + triggering selftest"
ssh_guest "cmd /c \"if exist ${SENTINEL_WIN} del /q ${SENTINEL_WIN}\""
ssh_guest "schtasks /run /tn onenote-selftest" >/dev/null

echo ">> waiting for selftest (timeout ${POLL_TIMEOUT_S}s)"
deadline=$(( SECONDS + POLL_TIMEOUT_S ))
guest_rc=""
while (( SECONDS < deadline )); do
  if guest_rc="$(ssh_guest "cmd /c \"if exist ${SENTINEL_WIN} type ${SENTINEL_WIN}\"" 2>/dev/null)"; then
    guest_rc="$(printf '%s' "$guest_rc" | tr -d '[:space:]')"
    [[ -n "$guest_rc" ]] && break
  fi
  sleep "$POLL_INTERVAL_S"
done

# --- 4. collect the selftest log --------------------------------------------
echo ">> collecting selftest results"
ssh_guest "tar -czf - -C ${GUEST_REPO} test-results/selftest.log test-results/selftest_exit.txt" 2>/dev/null \
  | tar -xzf - -C "$REPO_ROOT" || echo "!! collection failed (did the selftest start?)"

if [[ -z "$guest_rc" ]]; then
  echo "!! timed out after ${POLL_TIMEOUT_S}s — see ${RESULTS_DIR}/selftest.log"
  exit 124
fi
echo ">> frozen-exe selftest finished with exit code ${guest_rc}; see ${RESULTS_DIR}/selftest.log"
exit "$guest_rc"
