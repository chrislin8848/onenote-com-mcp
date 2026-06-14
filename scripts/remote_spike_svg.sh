#!/usr/bin/env bash
# remote_spike_svg.sh — THROWAWAY freeze gate for the insert_svg_image feature.
#
# Proves resvg_py (the chosen SVG->PNG rasterizer) PyInstaller-freezes cleanly on Windows AND
# renders Traditional-Chinese glyphs from the frozen exe. Everything here is COM-free (pure
# rasterization), so — unlike remote_build.sh's selftest — the freeze AND the run both go over
# plain SSH; no interactive scheduled task is needed.
#
# Ships only packaging/spike_svg, freezes it in an EPHEMERAL uv env (--no-project --with), runs
# the frozen exe, and collects spike_out.png + spike.log back to the host for eyeballing.
set -euo pipefail

GUEST_HOST="${GUEST_HOST:-192.168.122.13}"
GUEST_USER="${GUEST_USER:-dev}"
GUEST_SSH_PORT="${GUEST_SSH_PORT:-22}"
GUEST_REPO="${GUEST_REPO:-C:/onenote-mcp}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

ssh_guest() { ssh -p "$GUEST_SSH_PORT" "${GUEST_USER}@${GUEST_HOST}" "$@"; }
GUEST_REPO_WIN="${GUEST_REPO//\//\\}"

# --- 1. ship the spike to the guest -----------------------------------------
echo ">> syncing packaging/spike_svg to guest (tar over ssh)"
ssh_guest "cmd /c \"if not exist ${GUEST_REPO_WIN} mkdir ${GUEST_REPO_WIN}\""
tar czf - --exclude='dist' --exclude='build' --exclude='__pycache__' --exclude='spike_out.png' \
  -C "$REPO_ROOT" packaging/spike_svg \
  | ssh_guest "tar -xzf - -C ${GUEST_REPO}"

# --- 2. freeze the spike (COM-free → plain SSH, ephemeral uv env) ------------
echo ">> freezing spike (PyInstaller + resvg_py, ephemeral env) — a few minutes"
ssh_guest "cmd /c \"cd /d ${GUEST_REPO_WIN} && uv run --no-project --with resvg-py --with pyinstaller pyinstaller packaging\\spike_svg\\spike.spec --noconfirm --distpath packaging\\spike_svg\\dist --workpath packaging\\spike_svg\\build\""

# --- 3. run the FROZEN exe (still COM-free → plain SSH) ----------------------
echo ">> running the frozen spike exe"
ssh_guest "cmd /c \"${GUEST_REPO_WIN}\\packaging\\spike_svg\\run_spike.bat\""

# --- 4. collect the PNG + log + exit sentinel -------------------------------
echo ">> collecting spike results"
ssh_guest "tar -czf - -C ${GUEST_REPO} test-results/spike_out.png test-results/spike.log test-results/spike_exit.txt" 2>/dev/null \
  | tar -xzf - -C "$REPO_ROOT" || echo "!! collection failed (did the freeze/run start?)"

rc="$(tr -d '[:space:]' < "${REPO_ROOT}/test-results/spike_exit.txt" 2>/dev/null || true)"
echo ">> frozen spike exit code: ${rc:-<none>}; PNG at test-results/spike_out.png, log at test-results/spike.log"
[[ "${rc:-1}" == "0" ]] || exit "${rc:-1}"
