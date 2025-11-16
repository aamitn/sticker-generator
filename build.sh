#!/usr/bin/env bash

echo "================================"
echo "  Sticker Generator Build Script"
echo "================================"

# -------------------------------------------------------------
# Switch to location of this script
# -------------------------------------------------------------
cd "$(dirname "$0")" || exit 1
echo "Working Directory: $(pwd)"

# -------------------------------------------------------------
# Create venv if missing
# -------------------------------------------------------------
if [ ! -d ".venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv .venv
else
    echo "Virtual environment already exists."
fi

# -------------------------------------------------------------
# Activate venv
# -------------------------------------------------------------
echo "Activating virtual environment..."
source .venv/bin/activate

# -------------------------------------------------------------
# Install dependencies
# -------------------------------------------------------------
echo "Installing dependencies..."
pip install --upgrade pip
pip install -r requirements.txt

# -------------------------------------------------------------
# Build executable using PyInstaller
# -------------------------------------------------------------
echo "Building executable with PyInstaller..."
pyinstaller --noconfirm --onefile --windowed --icon=icon.ico app.py

if [ $? -ne 0 ]; then
    echo "PyInstaller build failed!"
    exit 1
fi

echo "Executable built: dist/app"

# -------------------------------------------------------------
# macOS/Linux Installer Step
# -------------------------------------------------------------
echo "NOTE: Installer creation skipped."
echo "      Inno Setup works only on Windows."

# -------------------------------------------------------------
# Completed
# -------------------------------------------------------------
echo "======================================"
echo "  BUILD SUCCESSFUL!"
echo "  Executable located in ./dist/"
echo "======================================"
