Write-Host "=== Installing MSReportChecker (Windows) ==="

# Function to print formatted messages
function Print {
    param([string]$Message)
    Write-Host "`n$Message`n"
}

# Ensure script is run as Administrator
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "`nThis script must be run as Administrator. Right-click and choose 'Run with PowerShell'."
    Exit 1
}

# Check for Python
$python = Get-Command python -ErrorAction SilentlyContinue
if (!$python) {
    Print "Python not found. Please install Python and ensure it is in your PATH."
    Exit 1
}

# Install pip packages
Print "Installing Flask (via pip)..."
try {
    python -m pip install --upgrade pip
    python -m pip install flask
} catch {
    Print "Failed to install Flask. Please check your Python/pip setup."
    Exit 1
}

# Check for mkcert
$mkcertInstalled = $false
if (Get-Command mkcert -ErrorAction SilentlyContinue) {
    $mkcertInstalled = $true
} elseif (Test-Path ".\mkcert.exe") {
    $mkcertInstalled = $true
} else {
    Print "mkcert not found. Downloading binary from GitHub..."

    $mkcertUrl = "https://github.com/FiloSottile/mkcert/releases/download/v1.4.4/mkcert-v1.4.4-windows-amd64.exe"
    $destination = ".\mkcert.exe"

    try {
        Invoke-WebRequest -Uri $mkcertUrl -OutFile $destination -UseBasicParsing
        $mkcertInstalled = $true
    } catch {
        Print "Failed to download mkcert. Please check your internet connection or download it manually."
        Exit 1
    }
}

# Create certificates
Print "Installing mkcert root CA and generating localhost certificate..."

try {
    if (Get-Command mkcert -ErrorAction SilentlyContinue) {
        mkcert -install
        mkcert -key-file key.pem -cert-file cert.pem localhost 127.0.0.1 ::1
    } else {
        .\mkcert.exe -install
        .\mkcert.exe -key-file key.pem -cert-file cert.pem localhost 127.0.0.1 ::1
    }
} catch {
    Print "mkcert failed. You may need to install manually or run this again."
    Exit 1
}

# Display manifest instructions
Print "Please add the following path to your Trusted Add-In Catalogs in Word:"
Print (Resolve-Path .\MSReportChecker.xml).Path

Print ""
Print "Installation complete!"
Print ""
Print "Run 'python server.py' to launch the local HTTPS server on https://localhost:3000/"
