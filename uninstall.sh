#!/usr/bin/env bash
set -euo pipefail

MANIFEST_NAME="a13454cd-574c-44ce-9c64-19dcd0ae477b.manifest.xml"

echo ""
echo "  Jolify — Uninstaller"
echo "  ──────────────────────────────────────"
echo ""

# ── 1. Find the manifest ─────────────────────────────────────────
WEF_CANDIDATES=(
  "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Wef"
  "$HOME/Library/Application Support/Microsoft/Office/16.0/Wef"
)

FOUND=""
for candidate in "${WEF_CANDIDATES[@]}"; do
  if [ -f "$candidate/$MANIFEST_NAME" ]; then
    FOUND="$candidate/$MANIFEST_NAME"
    break
  fi
done

if [ -z "$FOUND" ]; then
  echo "  ✗  Jolify is not installed (manifest not found)."
  echo ""
  exit 0
fi

# ── 2. Remove the manifest ───────────────────────────────────────
echo "  → Removing: $FOUND"
rm "$FOUND"
echo "  ✓  Manifest removed"

# ── 3. Restart PowerPoint ────────────────────────────────────────
if pgrep -xq "Microsoft PowerPoint"; then
  echo "  → Closing PowerPoint..."
  osascript -e 'tell application "Microsoft PowerPoint" to quit' 2>/dev/null || true
  sleep 2

  if pgrep -xq "Microsoft PowerPoint"; then
    kill "$(pgrep -x 'Microsoft PowerPoint')" 2>/dev/null || true
    sleep 1
  fi

  echo "  → Reopening PowerPoint..."
  open -a "Microsoft PowerPoint"
  echo ""
  echo "  ✓  Jolify has been removed. PowerPoint restarted without the add-in."
else
  echo ""
  echo "  ✓  Jolify has been removed. The Jolify tab will be gone next time you open PowerPoint."
fi

echo ""
