#!/bin/bash
# Serve the dashboard to colleagues on the same office network / VPN. Keep this window open.
cd "$(dirname "$0")"
IP=$(ipconfig getifaddr en0 || ipconfig getifaddr en1)
echo "Share this link:  http://$IP:8503   (only people on the same network can open it)"
"$HOME/.local/venvs/sales_review_ho/bin/streamlit" run app.py --server.headless true --server.address 0.0.0.0 --server.port 8503 --browser.gatherUsageStats false
