#!/usr/bin/env bash
set -euo pipefail

MANIFEST_URL="https://stephenjoly.github.io/jolify-ppt/manifest.xml"
MANIFEST_NAME="a13454cd-574c-44ce-9c64-19dcd0ae477b.manifest.xml"
LOCAL_SERVICE_LABEL="com.stephenjoly.jolify.localserver"
LOCAL_INSTALL_ROOT="$HOME/Library/Application Support/JolifyLocal"
LOCAL_PLIST_PATH="$HOME/Library/LaunchAgents/$LOCAL_SERVICE_LABEL.plist"
LOCAL_CERT_NAME="Jolify Local Add-in"

echo ""
echo "  Jolify — Hosted Installer"
echo "  ─────────────────────────"
echo ""

# ── 1. Confirm PowerPoint is installed ───────────────────────────
PPT_APP=""
for candidate in \
  "/Applications/Microsoft PowerPoint.app" \
  "$HOME/Applications/Microsoft PowerPoint.app"; do
  if [ -d "$candidate" ]; then
    PPT_APP="$candidate"
    break
  fi
done

if [ -z "$PPT_APP" ]; then
  echo "  ✗  Microsoft PowerPoint was not found in /Applications."
  echo "     Please install PowerPoint (Microsoft 365) and try again."
  echo ""
  exit 1
fi

echo "  ✓  Found PowerPoint: $PPT_APP"

# ── 1b. Disable local mode if it exists ──────────────────────────
if [ -f "$LOCAL_PLIST_PATH" ] || [ -d "$LOCAL_INSTALL_ROOT" ]; then
  echo "  → Switching from local mode to hosted mode"
  launchctl bootout "gui/$(id -u)" "$LOCAL_PLIST_PATH" >/dev/null 2>&1 || true
  rm -f "$LOCAL_PLIST_PATH"
  rm -rf "$LOCAL_INSTALL_ROOT"

  hashes="$(security find-certificate -a -Z -c "$LOCAL_CERT_NAME" "$HOME/Library/Keychains/login.keychain-db" 2>/dev/null | awk '/SHA-1 hash:/ { print $3 }')"
  if [ -n "$hashes" ]; then
    while IFS= read -r hash; do
      [ -z "$hash" ] && continue
      security delete-certificate -Z "$hash" "$HOME/Library/Keychains/login.keychain-db" >/dev/null 2>&1 || true
    done <<< "$hashes"
  fi
fi

# ── 2. Locate or create the WEF folder ───────────────────────────
WEF_CANDIDATES=(
  "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Wef"
  "$HOME/Library/Application Support/Microsoft/Office/16.0/Wef"
)

WEF_DIR=""

for candidate in "${WEF_CANDIDATES[@]}"; do
  if [ -d "$candidate" ]; then
    WEF_DIR="$candidate"
    break
  fi
done

if [ -z "$WEF_DIR" ]; then
  if [ -d "$HOME/Library/Containers/com.microsoft.Powerpoint" ]; then
    WEF_DIR="${WEF_CANDIDATES[0]}"
  elif [ -d "$HOME/Library/Group Containers/UBF8T346G9.Office" ]; then
    WEF_DIR="${WEF_CANDIDATES[1]}"
  else
    WEF_DIR="${WEF_CANDIDATES[2]}"
  fi
  echo "  → Creating add-ins folder: $WEF_DIR"
  mkdir -p "$WEF_DIR"
fi

# ── 3. Check for existing installation ────────────────────────────
UPGRADE=false
if [ -f "$WEF_DIR/$MANIFEST_NAME" ]; then
  UPGRADE=true
  echo "  → Existing Jolify installation found — upgrading"
fi

echo "  → Installing to: $WEF_DIR"

# ── 4. Download and install the manifest ──────────────────────────
echo "  → Downloading manifest..."
if ! curl -fsSL "$MANIFEST_URL" -o "/tmp/$MANIFEST_NAME"; then
  echo ""
  echo "  ✗  Download failed. Check your internet connection and try again."
  exit 1
fi

cp "/tmp/$MANIFEST_NAME" "$WEF_DIR/$MANIFEST_NAME"
rm "/tmp/$MANIFEST_NAME"

echo "  ✓  Manifest installed"

# ── 5. Restart PowerPoint if upgrading ────────────────────────────
if [ "$UPGRADE" = true ]; then
  echo "  → Closing PowerPoint..."
  osascript -e 'tell application "Microsoft PowerPoint" to quit' 2>/dev/null || true
  sleep 2

  # Force-kill if it didn't quit gracefully
  if pgrep -xq "Microsoft PowerPoint"; then
    kill "$(pgrep -x 'Microsoft PowerPoint')" 2>/dev/null || true
    sleep 1
  fi

  echo "  → Reopening PowerPoint..."
  open -a "Microsoft PowerPoint"
  echo ""
  echo "  ✓  Jolify hosted mode has been upgraded! Look for the Jolify tab in the ribbon."
else
  echo ""
  echo "  ✓  Done! Restart PowerPoint and look for the Jolify tab in the ribbon."
fi

echo ""
