#!/bin/bash
# Install LibreOffice if not present
if ! command -v libreoffice &> /dev/null; then
    echo "LibreOffice not found, installing..."
    sudo apt-get update && sudo apt-get install -y libreoffice
fi
