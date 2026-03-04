#!/usr/bin/env bash
set -euo pipefail

MANIFEST_URL="https://stephenjoly.github.io/custom-ppt-addin/manifest.xml"
MANIFEST_NAME="jolify-manifest.xml"

# Possible WEF folder locations on macOS
WEF_APPSTORE="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"
WEF_RETAIL="$HOME/Library/Application Support/Microsoft/Office/16.0/Wef"

echo ""
echo "  Jolify — PowerPoint Add-in Installer"
echo "  ──────────────────────────────────────"
echo ""

# Detect WEF folder
if [ -d "$WEF_APPSTORE" ]; then
  WEF_DIR="$WEF_APPSTORE"
elif [ -d "$WEF_RETAIL" ]; then
  WEF_DIR="$WEF_RETAIL"
else
  echo "  ✗  Could not find a PowerPoint installation."
  echo "     Make sure PowerPoint for Mac (Microsoft 365) is installed"
  echo "     and has been opened at least once."
  echo ""
  exit 1
fi

echo "  → Downloading manifest..."
curl -fsSL "$MANIFEST_URL" -o "/tmp/$MANIFEST_NAME"

echo "  → Installing to: $WEF_DIR"
cp "/tmp/$MANIFEST_NAME" "$WEF_DIR/$MANIFEST_NAME"
rm "/tmp/$MANIFEST_NAME"

echo ""
echo "  ✓  Done! Restart PowerPoint and look for the Jolify tab in the ribbon."
echo ""
