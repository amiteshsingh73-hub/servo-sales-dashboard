#!/bin/bash
# Double-click to start the dashboard (opens in your browser).
cd "$(dirname "$0")"
if [ ! -d .venv ]; then python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt; fi
.venv/bin/streamlit run app.py --server.headless false --browser.gatherUsageStats false
