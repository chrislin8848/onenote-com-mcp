#!/usr/bin/env bash
# remote_test.sh — host-side driver for the Tier 2 (Windows VM) test loop (SPEC §2.4).
#
# STATUS: skeleton. Wired up in Phase 0b once the Windows VM exists. The control flow below is
# the intended contract; the TODOs are the parts that need the live VM to finalize.
#
# Why "trigger vs run" are split (SPEC §2.4): COM needs the autologon *interactive* session.
# SSHing in lands in a non-interactive session where COM against OneNote fails. So we SSH only
# to TRIGGER work; the work itself RUNS in the autologon session (Task Scheduler "only when
# logged on", or a long-running watcher in that session).
#
# Exit code mirrors the guest's pytest exit code, so the host dev loop / CI can gate on it.
set -euo pipefail

# --- config (override via env) ----------------------------------------------
GUEST_HOST="${GUEST_HOST:?set GUEST_HOST (libvirt guest IP/host reachable over SSH)}"
GUEST_USER="${GUEST_USER:-onenote}"
GUEST_SSH_PORT="${GUEST_SSH_PORT:-22}"
GUEST_REPO="${GUEST_REPO:-C:/onenote-mcp}"          # repo path on the guest
RESULTS_DIR="${RESULTS_DIR:-./test-results}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

ssh_guest() { ssh -p "$GUEST_SSH_PORT" "${GUEST_USER}@${GUEST_HOST}" "$@"; }

mkdir -p "$RESULTS_DIR"

# --- 1. ship code to the guest ----------------------------------------------
# TODO(Phase 0b): choose git push→guest pull OR rsync. rsync is simplest cross-tool:
echo ">> syncing code to guest"
rsync -az --delete \
  --exclude '.git' --exclude '.venv' --exclude 'refs' --exclude 'test-results' \
  -e "ssh -p ${GUEST_SSH_PORT}" \
  "${REPO_ROOT}/" "${GUEST_USER}@${GUEST_HOST}:${GUEST_REPO}/"

# --- 2. trigger the run in the AUTOLOGON interactive session ----------------
# TODO(Phase 0b): create the scheduled task on the guest once (runs only when logged on):
#   schtasks /create /tn onenote-tier2 /sc once /st 00:00 /it \
#     /tr "cmd /c cd /d C:\onenote-mcp && uv sync && uv run pytest -m windows --junitxml=test-results\tier2.xml -ra > test-results\tier2.log 2>&1"
echo ">> triggering Tier 2 run in interactive session"
ssh_guest "schtasks /run /tn onenote-tier2"

# TODO(Phase 0b): poll for completion — task is async. Watch for the JUnit file to appear/update,
# or query `schtasks /query /tn onenote-tier2 /fo list` until Status != Running.

# --- 3. collect results + freshly-dumped fixtures ---------------------------
echo ">> collecting results + fixtures"
rsync -az -e "ssh -p ${GUEST_SSH_PORT}" \
  "${GUEST_USER}@${GUEST_HOST}:${GUEST_REPO}/test-results/" "${RESULTS_DIR}/"
rsync -az -e "ssh -p ${GUEST_SSH_PORT}" \
  "${GUEST_USER}@${GUEST_HOST}:${GUEST_REPO}/tests/fixtures/" "${REPO_ROOT}/tests/fixtures/" || true

# --- 4. mirror the guest's exit code ----------------------------------------
# TODO(Phase 0b): read the real pytest exit code (e.g. from a sentinel file the task writes).
echo ">> Tier 2 complete; see ${RESULTS_DIR}"
exit 0
