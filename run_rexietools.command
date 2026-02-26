#!/bin/bash
# Rexie Tools launcher (macOS)
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT_PATH="$SCRIPT_DIR/RexieTools.ps1"

# Find pwsh (works for Intel + Apple Silicon Homebrew)
PWSH="$(command -v pwsh || true)"
if [[ -z "$PWSH" ]]; then
  if [[ -x "/opt/homebrew/bin/pwsh" ]]; then
    PWSH="/opt/homebrew/bin/pwsh"
  elif [[ -x "/usr/local/bin/pwsh" ]]; then
    PWSH="/usr/local/bin/pwsh"
  else
    echo "PowerShell (pwsh) not found. Install PowerShell 7 and try again."
    echo "https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-macos"
    exit 1
  fi
fi

exec "$PWSH" -NoProfile -ExecutionPolicy Bypass -File "$SCRIPT_PATH"