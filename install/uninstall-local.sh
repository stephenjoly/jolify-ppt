#!/usr/bin/env bash
set -euo pipefail

MANIFEST_NAME="a13454cd-574c-44ce-9c64-19dcd0ae477b.manifest.xml"
SERVICE_LABEL="com.stephenjoly.jolify.localserver"
CERT_NAME="Jolify Local Add-in"
INSTALL_ROOT="$HOME/Library/Application Support/JolifyLocal"
PLIST_PATH="$HOME/Library/LaunchAgents/$SERVICE_LABEL.plist"

echo ""
echo "  Jolify — Local Stable Uninstaller"
echo "  ─────────────────────────────────"
echo ""

WEF_CANDIDATES=(
  "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Wef"
  "$HOME/Library/Application Support/Microsoft/Office/16.0/Wef"
)

for candidate in "${WEF_CANDIDATES[@]}"; do
  if [ -f "$candidate/$MANIFEST_NAME" ]; then
    echo "  → Removing manifest from $candidate"
    rm "$candidate/$MANIFEST_NAME"
  fi
done

echo "  → Stopping launchd agent..."
launchctl bootout "gui/$(id -u)" "$PLIST_PATH" >/dev/null 2>&1 || true
rm -f "$PLIST_PATH"

if [ -d "$INSTALL_ROOT" ]; then
  echo "  → Removing local runtime..."
  rm -rf "$INSTALL_ROOT"
fi

hashes="$(security find-certificate -a -Z -c "$CERT_NAME" "$HOME/Library/Keychains/login.keychain-db" 2>/dev/null | awk '/SHA-1 hash:/ { print $3 }')"
if [ -n "$hashes" ]; then
  echo "  → Removing trusted local certificate..."
  while IFS= read -r hash; do
    [ -z "$hash" ] && continue
    security delete-certificate -Z "$hash" "$HOME/Library/Keychains/login.keychain-db" >/dev/null 2>&1 || true
  done <<< "$hashes"
fi

if pgrep -xq "Microsoft PowerPoint"; then
  echo "  → Restarting PowerPoint..."
  osascript -e 'tell application "Microsoft PowerPoint" to quit' 2>/dev/null || true
  sleep 2
  if pgrep -xq "Microsoft PowerPoint"; then
    kill "$(pgrep -x 'Microsoft PowerPoint')" 2>/dev/null || true
    sleep 1
  fi
  open -a "Microsoft PowerPoint"
fi

echo ""
echo "  ✓  Jolify local mode has been removed."
echo ""
