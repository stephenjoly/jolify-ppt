#!/usr/bin/env bash
set -euo pipefail

MANIFEST_URL="https://stephenjoly.github.io/jolify-ppt/manifest.xml"
MANIFEST_NAME="a13454cd-574c-44ce-9c64-19dcd0ae477b.manifest.xml"

echo ""
echo "  Jolify — PowerPoint Add-in Installer"
echo "  ──────────────────────────────────────"
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

# ── 2. Locate or create the WEF folder ───────────────────────────
# Checked in order of likelihood; the folder may not exist yet if
# the user has never installed an add-in — we create it if needed.
WEF_CANDIDATES=(
  "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
  "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Wef"
  "$HOME/Library/Application Support/Microsoft/Office/16.0/Wef"
)

WEF_DIR=""

# Prefer whichever candidate already exists
for candidate in "${WEF_CANDIDATES[@]}"; do
  if [ -d "$candidate" ]; then
    WEF_DIR="$candidate"
    break
  fi
done

# If none exists yet, create the one that matches the install type.
# App Store builds sandbox under ~/Library/Containers; retail builds don't.
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

echo "  → Installing to: $WEF_DIR"

# ── 3. Download and install the manifest ─────────────────────────
echo "  → Downloading manifest..."
if ! curl -fsSL "$MANIFEST_URL" -o "/tmp/$MANIFEST_NAME"; then
  echo ""
  echo "  ✗  Download failed. Check your internet connection and try again."
  exit 1
fi

cp "/tmp/$MANIFEST_NAME" "$WEF_DIR/$MANIFEST_NAME"
rm "/tmp/$MANIFEST_NAME"

echo ""
echo "  ✓  Done! Restart PowerPoint and look for the Jolify tab in the ribbon."
echo ""
