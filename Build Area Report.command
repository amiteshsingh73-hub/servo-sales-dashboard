#!/bin/bash
# Double-click: rebuilds "Reports/Secondary Sales - Segment Area Customer <Mon-YY>.xlsx" from the files in Sales Review/Data.
cd "$(dirname "$0")"
V="$HOME/.local/venvs/sales_review_ho"
if [ ! -x "$V/bin/python" ]; then mkdir -p "$HOME/.local/venvs" && python3 -m venv "$V" && "$V/bin/pip" install -q -r requirements.txt; fi
"$V/bin/python" area_report.py && open "../Reports"
echo; read -n 1 -s -r -p "Press any key to close"
