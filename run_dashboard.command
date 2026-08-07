#!/bin/bash
# Double-click this file in Finder to launch the Bookkeeping Dashboard.
# It installs/updates dependencies automatically on first run, starts the
# local web server, and opens it in your browser.
#
# Once it's running, anyone else on the shop network can open it too, at:
#   http://<this Mac's name>.local:5001
# (Find this Mac's name under  System Settings > General > Sharing.)

cd "$(dirname "$0")"

# Activate project venv if present, otherwise use system python
if [ -d "venv/bin" ]; then
    source venv/bin/activate
    PY=python3
elif command -v python3 &>/dev/null; then
    PY=python3
else
    osascript -e 'display alert "Python 3 not found" message "Install Python 3 from python.org and try again."'
    exit 1
fi

# Install missing packages quietly
$PY -m pip install -q --require-virtualenv openpyxl rapidfuzz reportlab xlwings Flask 2>/dev/null || \
$PY -m pip install -q openpyxl rapidfuzz reportlab xlwings Flask

echo "Starting Funky's Electrical Dashboard…"
echo "Opening http://localhost:5001 — leave this window open while the dashboard is in use."

# Open the browser shortly after the server starts, then run the server
# in the foreground so closing this window stops the dashboard.
( sleep 1.5 && open "http://localhost:5001" ) &
$PY webapp/app.py
