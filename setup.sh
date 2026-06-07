#!/usr/bin/env bash
# ===================================================================
#  One-time setup for the HP Gas sign-in bot (macOS / Linux).
#  Run:  ./setup.sh
#  It creates a private Python environment and installs everything.
# ===================================================================
set -e
cd "$(dirname "$0")"

echo
echo "[1/4] Creating the Python virtual environment (.venv)..."
python3 -m venv .venv

echo "[2/4] Activating the environment..."
# shellcheck disable=SC1091
source .venv/bin/activate

echo "[3/4] Installing dependencies..."
python -m pip install --upgrade pip
pip install -r requirements.txt

echo "[4/4] Registering Google Chrome for automation..."
python -m playwright install chrome

echo
echo "==================================================================="
echo " Setup complete!"
echo " 1. Put your .xlsx file (columns: conno, id, password, code) in the"
echo "    'input' folder."
echo " 2. Run ./run.sh to start the bot."
echo "==================================================================="
