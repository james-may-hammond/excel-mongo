#!/bin/bash

PORT=8000
MANIFEST_SRC="fe/manifest.xml"
EXCEL_WEF_DIR="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
KEY_FILE="key.pem"
CERT_FILE="cert.pem"

echo ""
echo "=================================================="
echo "   Excel-MongoDB Connector — Setup & Run"
echo "=================================================="

# 1. Sideload manifest (real file copy)
echo ""
echo "[1/3] Sideloading manifest..."
mkdir -p "$EXCEL_WEF_DIR"
rm -f "$EXCEL_WEF_DIR/manifest.xml"
cp "$MANIFEST_SRC" "$EXCEL_WEF_DIR/manifest.xml"
echo "      Done → $EXCEL_WEF_DIR/manifest.xml"

# 2. SSL cert via mkcert (properly trusted by macOS + WKWebView)
echo "[2/3] Checking SSL certificates..."
if [ ! -f "$CERT_FILE" ] || [ ! -f "$KEY_FILE" ]; then
    echo "      Generating certificate with mkcert..."
    mkcert -cert-file "$CERT_FILE" -key-file "$KEY_FILE" localhost 127.0.0.1 ::1 2>&1
fi

# Check if mkcert CA is installed
CAROOT=$(mkcert -CAROOT)
if ! security find-certificate -a "$CAROOT/rootCA.pem" /Library/Keychains/System.keychain > /dev/null 2>&1; then
    echo ""
    echo "  ⚠️  IMPORTANT — Run this once to trust the certificate:"
    echo ""
    echo "  sudo security add-trusted-cert -d -r trustRoot \\"
    echo "    -k /Library/Keychains/System.keychain \\"
    echo "    \"$CAROOT/rootCA.pem\""
    echo ""
    echo "  Then re-run ./run.sh"
    echo ""
    exit 1
fi

echo "      Certificate trusted ✅"

# 3. Start server
echo "[3/3] Starting HTTPS server on https://localhost:$PORT ..."
echo ""

lsof -ti:$PORT | xargs kill -9 2>/dev/null
sleep 1

./.venv/bin/uvicorn be.main:app \
  --host localhost \
  --port $PORT \
  --ssl-keyfile "$KEY_FILE" \
  --ssl-certfile "$CERT_FILE" \
  --reload
