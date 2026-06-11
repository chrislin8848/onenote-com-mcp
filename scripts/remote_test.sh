#!/usr/bin/env bash
# remote_test.sh — host-side driver for the Tier 2 (Windows VM) test loop (SPEC §2.4).
#
# Ships the repo to the guest, triggers the Tier 2 run in the AUTOLOGON interactive session
# (where COM can reach OneNote — a bare SSH command lands in session 0 and fails), waits for
# it, pulls back results + freshly-dumped fixtures, and exits with the guest's pytest code so
# the host dev loop / CI can gate on it.
#
# Transport is tar-over-SSH, NOT rsync: the Windows guest has neither rsync nor git, but ships
# bsdtar as tar.exe (Win10 1803+). Shell builtins are wrapped in `cmd /c` so the command works
# whether the guest's OpenSSH default shell is cmd or PowerShell; external tools (tar, schtasks,
# uv) are invoked directly. Code sync is extract-over (no --delete equivalent) — deleted files
# linger on the guest; wipe C:\onenote-mcp by hand if that ever matters (keeps .venv otherwise).
set -euo pipefail

# --- config (override via env) ----------------------------------------------
# VM IP can shift with DHCP; look it up with: virsh domifaddr --source agent win11-onenote
GUEST_HOST="${GUEST_HOST:-192.168.122.13}"
GUEST_USER="${GUEST_USER:-dev}"
GUEST_SSH_PORT="${GUEST_SSH_PORT:-22}"
GUEST_REPO="${GUEST_REPO:-C:/onenote-mcp}"          # forward slashes: bsdtar -C accepts them
RESULTS_DIR="${RESULTS_DIR:-./test-results}"
POLL_TIMEOUT_S="${POLL_TIMEOUT_S:-1800}"            # uv sync + COM round-trips can take minutes
POLL_INTERVAL_S="${POLL_INTERVAL_S:-10}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

ssh_guest() { ssh -p "$GUEST_SSH_PORT" "${GUEST_USER}@${GUEST_HOST}" "$@"; }
SENTINEL="${GUEST_REPO}/test-results/exit_code.txt"
SENTINEL_WIN="${SENTINEL//\//\\}"                   # backslash form for cmd builtins

mkdir -p "$RESULTS_DIR"

# --- 1. ship code to the guest (tar pipe) -----------------------------------
echo ">> syncing code to guest (tar over ssh)"
ssh_guest "cmd /c \"if not exist ${GUEST_REPO//\//\\} mkdir ${GUEST_REPO//\//\\}\""
tar czf - \
  --exclude='.git' --exclude='.venv' --exclude='refs' \
  --exclude='test-results' --exclude='__pycache__' --exclude='*.pyc' \
  -C "$REPO_ROOT" . \
  | ssh_guest "tar -xzf - -C ${GUEST_REPO}"

# --- 2. (re)register the interactive task, clear the sentinel, trigger -------
# /it → runs only when logged on, in the autologon interactive session. /f → idempotent.
# run_tier2.bat path has no spaces, so /tr needs no inner quoting.
echo ">> (re)registering onenote-tier2 task"
ssh_guest "schtasks /create /f /tn onenote-tier2 /sc once /st 00:00 /it /tr ${GUEST_REPO//\//\\}\\scripts\\run_tier2.bat" >/dev/null

echo ">> clearing stale sentinel + triggering run"
ssh_guest "cmd /c \"if exist ${SENTINEL_WIN} del /q ${SENTINEL_WIN}\""
ssh_guest "schtasks /run /tn onenote-tier2" >/dev/null

# --- 3. poll for the sentinel (task is async) -------------------------------
echo ">> waiting for Tier 2 to finish (timeout ${POLL_TIMEOUT_S}s)"
deadline=$(( SECONDS + POLL_TIMEOUT_S ))
guest_rc=""
while (( SECONDS < deadline )); do
  if guest_rc="$(ssh_guest "cmd /c \"if exist ${SENTINEL_WIN} type ${SENTINEL_WIN}\"" 2>/dev/null)"; then
    guest_rc="$(printf '%s' "$guest_rc" | tr -d '[:space:]')"
    [[ -n "$guest_rc" ]] && break
  fi
  sleep "$POLL_INTERVAL_S"
done

# --- 4. collect results + freshly-dumped fixtures ---------------------------
echo ">> collecting results + fixtures"
ssh_guest "tar -czf - -C ${GUEST_REPO} test-results tests/fixtures" 2>/dev/null \
  | tar -xzf - -C "$REPO_ROOT" || echo "!! collection failed (did the run start?)"

# --- 5. mirror the guest's exit code ----------------------------------------
if [[ -z "$guest_rc" ]]; then
  echo "!! timed out after ${POLL_TIMEOUT_S}s with no exit_code.txt — see ${RESULTS_DIR}/tier2.log"
  exit 124
fi
echo ">> Tier 2 finished with pytest exit code ${guest_rc}; see ${RESULTS_DIR}"
exit "$guest_rc"
