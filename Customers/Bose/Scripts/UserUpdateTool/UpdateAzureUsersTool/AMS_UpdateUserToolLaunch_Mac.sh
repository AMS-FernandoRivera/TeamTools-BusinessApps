#!/bin/bash

# Clear the terminal and set a title
clear
echo "======== AMS_UpdateUsersTool ========"

# Check for PowerShell executable
if ! command -v pwsh &> /dev/null; then
    echo "ERROR: PowerShell (pwsh) is not installed!"
    echo "Please install PowerShell (https://aka.ms/powershell) and try again."
    exit 1
fi

# Check for minimum PowerShell version (5.0 or higher)
PS_VERSION=$(pwsh -Command '$PSVersionTable.PSVersion.Major')
if [ "$PS_VERSION" -lt 5 ]; then
    echo "ERROR: PowerShell version 5.0 or higher is required!"
    echo "Current version detected: $PS_VERSION"
    echo "Please update PowerShell and try again."
    exit 1
fi

echo "PowerShell version $PS_VERSION detected. OK."

# Ensure elevated permissions
if [ "$EUID" -ne 0 ]; then
    echo "This script requires administrative privileges. Relaunching with elevated permissions..."
    sudo bash "$0" "$@"
    exit
fi

# Launch the PowerShell script
echo "Starting AMS_UpdateUsersTool..."
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
pwsh -NoExit -ExecutionPolicy Bypass -File "$SCRIPT_DIR/AMS_UpdateUsersTool.ps1"

echo "======== Script Execution Complete ========"
exit