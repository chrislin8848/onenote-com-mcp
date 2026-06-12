#!/usr/bin/env bash
# remote_dump.sh — host-side driver for a MANUAL fixture dump on the VM (Phase 5b intake).
#
# Same transport + session mechanics as remote_test.sh (tar-over-SSH; COM only works in the
# autologon interactive session, so the dump runs via a `schtasks /it` task, never bare SSH).
# Differences from the Tier-2 loop, both deliberate (PII lesson):
#   * the guest dumps into test-results/dump, NOT tests/fixtures — sanitized fixtures are
#     never clobbered; and
#   * collection pulls back test-results ONLY; a human inspects and moves files into
#     tests/fixtures by hand.
set -euo pipefail

GUEST_HOST="${GUEST_HOST:-192.168.122.13}"
GUEST_USER="${GUEST_USER:-dev}"
GUEST_SSH_PORT="${GUEST_SSH_PORT:-22}"
GUEST_REPO="${GUEST_REPO:-C:/onenote-mcp}"
RESULTS_DIR="${RESULTS_DIR:-./test-results}"
POLL_TIMEOUT_S="${POLL_TIMEOUT_S:-600}"
POLL_INTERVAL_S="${POLL_INTERVAL_S:-5}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

ssh_guest() { ssh -p "$GUEST_SSH_PORT" "${GUEST_USER}@${GUEST_HOST}" "$@"; }
SENTINEL="${GUEST_REPO}/test-results/dump_exit_code.txt"
SENTINEL_WIN="${SENTINEL//\//\\}"

mkdir -p "$RESULTS_DIR"

echo ">> syncing code to guest (tar over ssh)"
ssh_guest "cmd /c \"if not exist ${GUEST_REPO//\//\\} mkdir ${GUEST_REPO//\//\\}\""
tar czf - \
  --exclude='.git' --exclude='.venv' --exclude='refs' \
  --exclude='test-results' --exclude='__pycache__' --exclude='*.pyc' \
  -C "$REPO_ROOT" . \
  | ssh_guest "tar -xzf - -C ${GUEST_REPO}"

echo ">> (re)registering onenote-dump task"
ssh_guest "schtasks /create /f /tn onenote-dump /sc once /st 00:00 /it /tr ${GUEST_REPO//\//\\}\\scripts\\run_dump.bat" >/dev/null

echo ">> clearing stale sentinel + triggering dump"
ssh_guest "cmd /c \"if exist ${SENTINEL_WIN} del /q ${SENTINEL_WIN}\""
ssh_guest "schtasks /run /tn onenote-dump" >/dev/null

echo ">> waiting for dump to finish (timeout ${POLL_TIMEOUT_S}s)"
deadline=$(( SECONDS + POLL_TIMEOUT_S ))
guest_rc=""
while (( SECONDS < deadline )); do
  if guest_rc="$(ssh_guest "cmd /c \"if exist ${SENTINEL_WIN} type ${SENTINEL_WIN}\"" 2>/dev/null)"; then
    guest_rc="$(printf '%s' "$guest_rc" | tr -d '[:space:]')"
    [[ -n "$guest_rc" ]] && break
  fi
  sleep "$POLL_INTERVAL_S"
done

echo ">> collecting test-results (dump output + log)"
ssh_guest "tar -czf - -C ${GUEST_REPO} test-results" 2>/dev/null \
  | tar -xzf - -C "$REPO_ROOT" || echo "!! collection failed (did the run start?)"

if [[ -z "$guest_rc" ]]; then
  echo "!! timed out after ${POLL_TIMEOUT_S}s with no dump_exit_code.txt — see ${RESULTS_DIR}/dump.log"
  exit 124
fi
echo ">> dump finished with exit code ${guest_rc}; output in ${RESULTS_DIR}/dump — inspect before moving into tests/fixtures"
exit "$guest_rc"
