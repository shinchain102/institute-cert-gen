#!/bin/bash

echo "Certificate Generator - Linux Setup"
echo "================================="

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python is not installed! Please install Python 3.x:"
    echo "sudo apt update && sudo apt install python3 python3-pip"
    exit 1
fi

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "pip is not installed! Please install pip:"
    echo "sudo apt install python3-pip"
    exit 1
fi

# Check if LibreOffice is installed
if ! command -v soffice &> /dev/null; then
    echo "LibreOffice is not installed!"
    echo "Please install LibreOffice:"
    echo "sudo apt update && sudo apt install libreoffice"
    exit 1
fi

echo "Installing required Python packages..."
pip3 install -r requirements.txt

echo "Starting Certificate Generator..."
xdg-open http://localhost:5000 2>/dev/null || open http://localhost:5000 2>/dev/null || echo "Please open http://localhost:5000 in your browser"
python3 main.py
