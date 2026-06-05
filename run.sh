#!/bin/bash
# run.sh — Start the Excel-MongoDB Connector backend (macOS)
# For Windows, see README.md

PORT=8000
KEY_FILE="key.pem"
CERT_FILE="cert.pem"

echo ""
echo "=================================================="
echo "   Excel-MongoDB Connector — Backend Startup"
echo "=================================================="

# ── 1. Check .env ──────────────────────────────────────
echo ""
echo "[1/3] Checking environment..."
if [ ! -f ".env" ]; then
    echo "      ⚠️  No .env file found. Creating from .env.example..."
    if [ -f ".env.example" ]; then
        cp .env.example .env
        echo "      Created .env — edit it to set MONGO_URI and DB_NAME before continuing."
        exit 1
    else
        echo "      ✗ No .env.example either. Create a .env file with MONGO_URI and DB_NAME."
        exit 1
    fi
fi
echo "      .env found ✅"

# ── 2. SSL certificate ─────────────────────────────────
echo "[2/3] Checking SSL certificate..."
if [ ! -f "$CERT_FILE" ] || [ ! -f "$KEY_FILE" ]; then
    echo "      Generating certificate with mkcert..."
    if ! command -v mkcert &> /dev/null; then
        echo "      ✗ mkcert not found. Install it first:"
        echo "          brew install mkcert && sudo mkcert -install"
        exit 1
    fi
    mkcert -cert-file "$CERT_FILE" -key-file "$KEY_FILE" localhost 127.0.0.1 ::1 2>&1
    echo "      cert.pem + key.pem generated."
fi

# Verify the mkcert root CA is trusted in the system keychain
CAROOT=$(mkcert -CAROOT 2>/dev/null)
if [ -n "$CAROOT" ] && ! security find-certificate -a "$CAROOT/rootCA.pem" /Library/Keychains/System.keychain > /dev/null 2>&1; then
    echo ""
    echo "  ⚠️  Certificate not yet trusted. Run this once:"
    echo ""
    echo "      sudo security add-trusted-cert -d -r trustRoot \\"
    echo "        -k /Library/Keychains/System.keychain \\"
    echo "        \"$CAROOT/rootCA.pem\""
    echo ""
    echo "  Then re-run ./run.sh"
    echo ""
    exit 1
fi
echo "      Certificate trusted ✅"

# ── 3. Start uvicorn ───────────────────────────────────
echo "[3/3] Starting server on https://localhost:$PORT ..."
echo ""
echo "  To load the add-in into Excel, run in a second terminal:"
echo "  ./node_modules/.bin/office-addin-debugging start fe/manifest.xml --no-debug --dev-server-port $PORT"
echo ""

# Kill anything already on the port
lsof -ti:$PORT | xargs kill -9 2>/dev/null
sleep 1

./.venv/bin/uvicorn be.main:app \
  --host localhost \
  --port $PORT \
  --ssl-keyfile "$KEY_FILE" \
  --ssl-certfile "$CERT_FILE" \
  --reload
