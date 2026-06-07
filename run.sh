#!/usr/bin/env bash
# ===================================================================
#  Start the HP Gas sign-in bot (macOS / Linux).
#  Run:  ./run.sh   (after running ./setup.sh once)
#  To STOP early: press Ctrl+C, or close the terminal.
#  Progress is saved to the "output" folder after every row.
# ===================================================================
set -e
cd "$(dirname "$0")"

if [ ! -f ".venv/bin/activate" ]; then
  echo "The environment is missing. Please run ./setup.sh first."
  exit 1
fi

# shellcheck disable=SC1091
source .venv/bin/activate
python run_bot.py

echo
echo "Bot finished (or was stopped). Output is in the 'output' folder."
