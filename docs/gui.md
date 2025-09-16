# Visio Temp File Remover GUI

A standalone desktop application for finding and removing corrupted Visio temporary files (~$*.vssx) that cause errors.

## Overview

This GUI application provides a user-friendly interface for the Visio Temp File Remover tool without requiring a web server. It runs locally on your Windows machine and uses PowerShell commands to scan for and delete temporary Visio files.

## Features

- **No Server Required**: Runs locally without starting a web server
- **User-Friendly Interface**: Simple GUI with directory selection and file management
- **Safe File Operations**: Uses the same PowerShell scripts as the web version for consistency
- **Progress Feedback**: Visual progress indicators during scanning and deletion
- **File Selection**: Select specific files for deletion
- **Error Handling**: Comprehensive error handling and user feedback
- **Fallback Mechanisms**: Uses both PowerShell scripts and direct PowerShell commands for maximum compatibility
- **Improved Display**: Shows relative paths and optimized column widths for better readability

## Requirements

- Windows operating system
- Python 3.6 or higher
- PowerShell (included with Windows)
- Visio Temp File Remover PowerShell scripts (included in this repository)

## Installation

1. Ensure you have Python installed on your system
2. No additional Python packages are required (uses standard library)

## Usage

### Method 1: Using the batch file
1. Double-click on `run_gui.bat` in the root directory

### Method 2: Using Python directly
1. Open a command prompt in the project directory
2. Run: `python visio_gui.py`

### Using the Application
1. Select the directory you want to scan for Visio temp files
2. Click "Scan for Files"
3. Review the found files in the results list (paths are shown relative to the scan directory for better readability)
4. Select the files you want to delete
5. Click "Delete Selected Files"
6. Confirm the deletion when prompted

## How It Works

The GUI application uses multiple approaches to ensure compatibility:

1. **Primary Method**: Uses the same PowerShell scripts as the web version:
   - `Scan-VisioTempFiles.ps1` for finding files
   - `Remove-VisioTempFiles.ps1` for deleting files

2. **Fallback Method**: If the scripts fail, it uses direct PowerShell commands like the web version

This ensures consistent behavior with the web version while providing a desktop experience and maximum compatibility.

## Troubleshooting

If you encounter issues:
1. Ensure PowerShell is available on your system
2. Check that the PowerShell scripts are in the `scripts` directory
3. Verify that you have appropriate permissions for the directories you're scanning
4. Make sure Python is properly installed and accessible from the command line
5. Try running the application as Administrator if you get permission errors