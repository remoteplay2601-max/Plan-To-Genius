#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR="$SCRIPT_DIR/.venv"
PYTHON_BIN="$VENV_DIR/bin/python"
STREAMLIT_BIN="$VENV_DIR/bin/streamlit"
HOST="${HOST:-127.0.0.1}"
PORT="${PORT:-8501}"

if [ ! -x "$PYTHON_BIN" ]; then
  echo "Creation de l'environnement virtuel..."
  python3 -m venv "$VENV_DIR"
fi

if ! "$PYTHON_BIN" -c "import streamlit, pandas, openpyxl" >/dev/null 2>&1; then
  echo "Installation des dependances..."
  "$PYTHON_BIN" -m pip install -r requirements.txt
fi

exec "$STREAMLIT_BIN" run app.py \
  --server.headless true \
  --server.address "$HOST" \
  --server.port "$PORT" \
  "$@"
