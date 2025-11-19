#!/usr/bin/env bash
set -e

# Start Streamlit in background so the container setup can finish.
nohup streamlit run app.py --server.address 0.0.0.0 --server.port 8000 > /tmp/streamlit.log 2>&1 &
sleep 1
echo "Streamlit started (logs: /tmp/streamlit.log)"
