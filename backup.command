#!/bin/bash
# Double-click: zips code + Data + AISTO/FOC folders (no .venv) into ~/Documents/SalesReview_Backups/
cd "$(dirname "$0")/.."
ROOT="$(pwd)"; OUT="$HOME/Documents/SalesReview_Backups"; mkdir -p "$OUT"
STAMP="$(date +%d%b%y_%H%M)"
zip -rq "$OUT/SalesReview_$STAMP.zip" . -x "*/.venv/*" "*/__pycache__/*" "*/.cache/*" "*/logs/*" "*.app/*"
echo "Saved: $OUT/SalesReview_$STAMP.zip"; ls -lh "$OUT" | tail -5
