#!/bin/bash
echo "Starting build process on Render..."

# Install LibreOffice
echo "Installing LibreOffice..."
sudo apt-get update
sudo apt-get install -y libreoffice

# Install Python dependencies if needed
echo "Setting up Python environment..."
python3 -m pip install --upgrade pip
python3 -m pip install pywin32 || echo "pywin32 install failed (non-Windows system)"

echo "Build complete"
