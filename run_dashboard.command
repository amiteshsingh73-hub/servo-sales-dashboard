#!/bin/bash
# Double-click to start the dashboard (opens in your browser).
# The Python environment lives outside OneDrive (thousands of small files slow the Mac down).
cd "$(dirname "$0")"
V="$HOME/.local/venvs/sales_review_ho"
if [ ! -x "$V/bin/streamlit" ]; then mkdir -p "$HOME/.local/venvs" && python3 -m venv "$V" && "$V/bin/pip" install -q -r requirements.txt; fi
"$V/bin/streamlit" run app.py --server.headless false --browser.gatherUsageStats false
