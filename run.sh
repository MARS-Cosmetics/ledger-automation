#!/usr/bin/env bash
# Launch the Mars Recon tool in your browser.
# First run installs deps; subsequent runs skip if streamlit is already present.

set -e
cd "$(dirname "$0")"

PY=${PY:-python3}

if ! "$PY" -c "import streamlit" 2>/dev/null; then
  echo "Installing dependencies…"
  "$PY" -m pip install --user -r requirements.txt
fi

exec "$PY" -m streamlit run app.py
