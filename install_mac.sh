#!/bin/bash

echo "=== Installing MSReportChecker (macOS) ==="

# Check and install Homebrew (with sudo)
if ! command -v brew &>/dev/null; then
    echo "Homebrew not found. Installing..."
    NONINTERACTIVE=1 /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    if ! command -v brew &>/dev/null; then
        echo "Homebrew installation failed. Please install it manually."
        exit 1
    fi
fi

# Install mkcert if not installed
if ! command -v mkcert &>/dev/null; then
    echo "Installing mkcert..."
    brew install mkcert || {
        echo "mkcert installation failed. Trying with sudo..."
        sudo brew install mkcert || {
            echo "mkcert failed even with sudo. Aborting."
            exit 1
        }
    }
fi

# Ensure mkcert is trusted (this always requires sudo)
echo "Installing mkcert CA (you may be prompted for your password)..."
sudo mkcert -install
mkcert -key-file key.pem -cert-file cert.pem localhost 127.0.0.1 ::1

# Detect python
if command -v python3 &>/dev/null; then
    PYTHON_CMD="python3"
elif command -v python &>/dev/null; then
    PYTHON_CMD="python"
else
    echo "Python not found. Please install Python."
    exit 1
fi

# Install Python dependencies
$PYTHON_CMD -m pip install --upgrade pip
$PYTHON_CMD -m pip install flask || {
    echo "Failed to install Flask. Please check your Python/pip setup."
    exit 1
}

chmod +x server.py
chown "$USER" key.pem cert.pem
chmod 600 key.pem cert.pem

# Copy manifest
echo "Copying manifest to Word add-in folder..."
mkdir -p ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
cp MSReportChecker.xml ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/MSReportChecker.xml

echo -e "\n\nInstallation complete!"
echo -e "\nStart './server.py' to launch the local HTTPS server.\n"
