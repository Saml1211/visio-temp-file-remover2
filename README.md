# Visio Temp File Remover

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)

A specialized web application designed to find and remove corrupted temporary Visio files that can cause errors in Microsoft Visio environments.

![Visio Temp File Remover Screenshot](https://via.placeholder.com/800x400?text=Visio+Temp+File+Remover)

## 🔍 Overview

The Visio Temp File Remover tool addresses a common problem in Microsoft Visio environments where temporary files (specifically those matching the pattern `~$*.vssx`) can cause errors and corruption in Visio shapes and templates. These files are often hidden by default and difficult to locate and remove through normal file management tools.

This project provides two interfaces for the same core functionality:
1. **Web Interface**: A web-based application for server environments
2. **Desktop GUI**: A standalone desktop application for local use

Both interfaces allow users to:
1. Scan directories for problematic temporary files
2. Review what will be deleted before taking action
3. Safely remove these files with proper permission handling

## ✨ Features

- **Smart Scanning**: Recursively scans specified directories for Visio temporary files matching `~$*.vssx` pattern
- **Hidden File Detection**: Uses PowerShell with proper escaping to find hidden system files that Windows Explorer might not display
- **Preview Before Deletion**: Lists all found files with full paths before any deletion occurs
- **Selective Deletion**: Choose which files to delete rather than removing all matches
- **User-Friendly Interface**: Simple interfaces that require no technical knowledge to operate
- **Secure Operations**: All file operations are performed with proper error handling
- **Dual Interface**: Available as both web application and standalone desktop GUI
- **Consistent Behavior**: Both interfaces use the same PowerShell scripts for identical functionality

## 🛠️ Technology Stack

**Web Interface**:
- **Backend**: Node.js with Express.js
- **Frontend**: HTML5, CSS3, and vanilla JavaScript
- **System Integration**: PowerShell commands for file operations
- **Styling**: Bootstrap for responsive design
- **HTTP Requests**: Fetch API for AJAX operations

**Desktop GUI**:
- **GUI Framework**: Python Tkinter
- **System Integration**: PowerShell commands for file operations
- **Packaging**: PyInstaller (for standalone executables)

**Shared Components**:
- **File Operations**: PowerShell scripts (`Scan-VisioTempFiles.ps1`, `Remove-VisioTempFiles.ps1`)
- **Configuration**: JSON configuration files

## 📋 Requirements

- Windows environment (PowerShell required)
- Node.js (v12.0.0 or higher) for web interface
- Python 3.6 or higher for GUI interface
- Administrative privileges (for accessing system folders)
- Web browser (Chrome, Firefox, Edge recommended)

### GUI Requirements
The GUI application uses only the Python standard library:
- tkinter (included with Python) for the graphical interface
- subprocess, json, os, threading, and pathlib (all standard library modules)
- No additional Python packages are required

## 📚 Documentation

For detailed documentation, please see the [Documentation Index](docs/index.md).

Key documents include:
- [Quick Start Guide](docs/quickstart.md) - For quickly setting up and running the web version
- [Installation Guide](docs/installation.md) - For installing and using the GUI version
- [GUI Documentation](docs/gui.md) - Information about the GUI application
- [Release Notes](docs/release-notes.md) - Details about changes in each release

## 🚀 Getting Started

This project provides two ways to use the Visio Temp File Remover tool:

### Web Interface (Original Method)
For quick setup and usage instructions, please refer to the [Quick Start Guide](docs/quickstart.md).

Basic steps include:
1. Clone this repository
2. Install dependencies with `npm install`
3. Start the server with `npm start` or by running `start.bat`
4. Access the web interface at http://localhost:3000

### Desktop GUI (Recommended for Local Use)
The project includes a standalone desktop GUI application that runs locally without a web server:

1. Download the latest release from the [Releases](https://github.com/Saml1211/visio-temp-file-remover2/releases) page
2. Extract the package to a folder of your choice
3. Double-click on `run_gui.bat` in the root directory, or run `python visio_gui.py` from the command line
4. Use the GUI to scan directories and remove Visio temp files

The GUI application provides the same functionality as the web interface but runs locally without needing to start a server. It uses the same PowerShell scripts for consistent behavior.

### Creating a Standalone Executable
To create a standalone executable (.exe) file:

1. Install PyInstaller: `pip install pyinstaller`
2. Create the executable: `pyinstaller --onefile --windowed --name "VisioTempFileRemover" --add-data "scripts;scripts" visio_gui.py`
3. The executable will be created in the `dist` folder
4. Copy the `scripts` folder to the same directory as the executable

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

```
MIT License

Copyright (c) 2023 Visio Temp File Remover

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
