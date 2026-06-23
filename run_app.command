#!/bin/bash
# Double-click this file in Finder to launch Invoice Maker.
# It will install/update dependencies automatically on first run.

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
$PY -m pip install -q --require-virtualenv customtkinter rapidfuzz openpyxl pillow reportlab xlwings 2>/dev/null || \
$PY -m pip install -q customtkinter rapidfuzz openpyxl pillow reportlab xlwings

$PY gui_automater.py
