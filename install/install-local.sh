#!/usr/bin/env bash
set -euo pipefail

BASE_URL="https://stephenjoly.github.io/jolify-ppt"
BUNDLE_URL="$BASE_URL/jolify-local-bundle.tar.gz"
MANIFEST_NAME="a13454cd-574c-44ce-9c64-19dcd0ae477b.manifest.xml"
SERVICE_LABEL="com.stephenjoly.jolify.localserver"
LOCAL_HOST="127.0.0.1"
LOCAL_PORT="38443"
CERT_NAME="Jolify Local Add-in"

INSTALL_ROOT="$HOME/Library/Application Support/JolifyLocal"
RUNTIME_DIR="$INSTALL_ROOT/runtime"
CERT_DIR="$INSTALL_ROOT/certs"
LOG_DIR="$INSTALL_ROOT/logs"
VENV_DIR="$INSTALL_ROOT/venv"
VENV_PYTHON="$VENV_DIR/bin/python"
PLIST_PATH="$HOME/Library/LaunchAgents/$SERVICE_LABEL.plist"
MANIFEST_PATH="$RUNTIME_DIR/manifest.xml"
REQUIREMENTS_PATH="$RUNTIME_DIR/requirements.txt"
PYTHON_BIN="$(command -v python3 || true)"

echo ""
echo "  Jolify — Local Stable Installer"
echo "  ───────────────────────────────"
echo ""

if [ "$(uname -s)" != "Darwin" ]; then
  echo "  ✗  Local mode currently supports macOS only."
  exit 1
fi

if [ -z "$PYTHON_BIN" ]; then
  echo "  ✗  python3 is required for local mode."
  echo "     Use the hosted installer instead, or install Python 3 and retry."
  exit 1
fi

for cmd in curl tar openssl security launchctl; do
  if ! command -v "$cmd" >/dev/null 2>&1; then
    echo "  ✗  Required command not found: $cmd"
    exit 1
  fi
done

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
  WEF_DIR="${WEF_CANDIDATES[0]}"
  mkdir -p "$WEF_DIR"
fi

TMP_DIR="$(mktemp -d)"
cleanup() {
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT

echo "  → Downloading local runtime bundle..."
curl -fsSL "$BUNDLE_URL" -o "$TMP_DIR/jolify-local-bundle.tar.gz"

echo "  → Installing runtime to $INSTALL_ROOT"
rm -rf "$RUNTIME_DIR"
mkdir -p "$RUNTIME_DIR" "$CERT_DIR" "$LOG_DIR" "$(dirname "$PLIST_PATH")"
tar -xzf "$TMP_DIR/jolify-local-bundle.tar.gz" -C "$RUNTIME_DIR"

if [ ! -f "$REQUIREMENTS_PATH" ]; then
  echo "  ✗  Local runtime bundle is missing Python requirements metadata."
  exit 1
fi

echo "  → Installing local Python environment..."
rm -rf "$VENV_DIR"
"$PYTHON_BIN" -m venv "$VENV_DIR"
"$VENV_PYTHON" -m pip install --upgrade pip >/dev/null
"$VENV_PYTHON" -m pip install --quiet -r "$REQUIREMENTS_PATH"

remove_existing_cert() {
  local hashes
  hashes="$(security find-certificate -a -Z -c "$CERT_NAME" "$HOME/Library/Keychains/login.keychain-db" 2>/dev/null | awk '/SHA-1 hash:/ { print $3 }')"
  if [ -n "$hashes" ]; then
    while IFS= read -r hash; do
      [ -z "$hash" ] && continue
      security delete-certificate -Z "$hash" "$HOME/Library/Keychains/login.keychain-db" >/dev/null 2>&1 || true
    done <<< "$hashes"
  fi
}

generate_cert() {
  local conf="$TMP_DIR/openssl.cnf"
  cat > "$conf" <<EOF
[req]
default_bits = 2048
distinguished_name = dn
x509_extensions = v3_req
prompt = no

[dn]
CN = $CERT_NAME

[v3_req]
subjectAltName = @alt_names
keyUsage = digitalSignature, keyEncipherment
extendedKeyUsage = serverAuth

[alt_names]
DNS.1 = localhost
IP.1 = 127.0.0.1
EOF

  openssl req -x509 -nodes -days 3650 \
    -newkey rsa:2048 \
    -keyout "$CERT_DIR/localhost-key.pem" \
    -out "$CERT_DIR/localhost-cert.pem" \
    -config "$conf" >/dev/null 2>&1
}

if [ ! -f "$CERT_DIR/localhost-cert.pem" ] || [ ! -f "$CERT_DIR/localhost-key.pem" ]; then
  echo "  → Generating local TLS certificate..."
  generate_cert
fi

echo "  → Trusting local TLS certificate in your login keychain..."
remove_existing_cert
security add-trusted-cert -d -r trustRoot \
  -k "$HOME/Library/Keychains/login.keychain-db" \
  "$CERT_DIR/localhost-cert.pem" >/dev/null

echo "  → Installing launchd agent..."
cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>$SERVICE_LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>$VENV_PYTHON</string>
    <string>$RUNTIME_DIR/local-server.py</string>
    <string>--root</string>
    <string>$RUNTIME_DIR/web</string>
    <string>--host</string>
    <string>$LOCAL_HOST</string>
    <string>--port</string>
    <string>$LOCAL_PORT</string>
    <string>--cert</string>
    <string>$CERT_DIR/localhost-cert.pem</string>
    <string>--key</string>
    <string>$CERT_DIR/localhost-key.pem</string>
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>WorkingDirectory</key>
  <string>$RUNTIME_DIR</string>
  <key>StandardOutPath</key>
  <string>$LOG_DIR/server.out.log</string>
  <key>StandardErrorPath</key>
  <string>$LOG_DIR/server.err.log</string>
</dict>
</plist>
EOF

launchctl bootout "gui/$(id -u)" "$PLIST_PATH" >/dev/null 2>&1 || true
launchctl bootstrap "gui/$(id -u)" "$PLIST_PATH"
launchctl kickstart -k "gui/$(id -u)/$SERVICE_LABEL"

echo "  → Waiting for local server to become healthy..."
for _ in $(seq 1 20); do
  if curl --silent --show-error --fail --cacert "$CERT_DIR/localhost-cert.pem" "https://$LOCAL_HOST:$LOCAL_PORT/healthz" >/dev/null 2>&1; then
    break
  fi
  sleep 1
done

if ! curl --silent --show-error --fail --cacert "$CERT_DIR/localhost-cert.pem" "https://$LOCAL_HOST:$LOCAL_PORT/healthz" >/dev/null 2>&1; then
  echo "  ✗  Local Jolify server did not become healthy."
  echo "     Check logs in $LOG_DIR"
  exit 1
fi

cp "$MANIFEST_PATH" "$WEF_DIR/$MANIFEST_NAME"
echo "  ✓  Local manifest installed to $WEF_DIR"

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
echo "  ✓  Jolify local mode is installed."
echo "     The local server is now managed by launchd and will restart automatically."
echo "     AI features still require internet access."
echo ""
