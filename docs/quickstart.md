# Visio Temp File Remover - Quick Start Guide

This guide will help you quickly set up, install, and run the Visio Temp File Remover application to scan for and delete temporary Visio files that can cause errors.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Installation](#installation)
3. [Running the Application](#running-the-application)
4. [Using the Application](#using-the-application)
5. [Troubleshooting](#troubleshooting)

## Prerequisites

Before setting up the application, ensure you have the following installed:

- [Node.js](https://nodejs.org/) (v18.x or higher, LTS)
- [npm](https://www.npmjs.com/) (typically comes with Node.js)
- [PowerShell](https://docs.microsoft.com/en-us/powershell/) (v5.0 or higher) - Windows only

## Installation

1. Clone the repository:

```bash
git clone https://github.com/Saml1211/visio-temp-file-remover2.git
cd visio-temp-file-remover2
```

2. Install dependencies:

```bash
npm install
```

This will install all required dependencies including Express and body-parser.

## Running the Application

There are two ways to start the application:

### Method 1: Using the batch file

Simply double-click on the `start.bat` file in the root directory of the project, or run:

```bash
start.bat
```

### Method 2: Using Node.js directly

Open a terminal in the project directory and run:

```bash
node app.js
```

The server will start on port 3000 by default. You should see a message in the console:

```text
Server running at http://localhost:3000
Available endpoints:
  - GET / : Web interface
  - POST /api/scan : Scan for temporary Visio files
  - POST /api/delete : Delete selected temporary Visio files
```

## Using the Application

### Web Interface

1. Open your web browser and navigate to [http://localhost:3000](http://localhost:3000)
2. You'll see the Visio Temp File Remover interface with the following options:

### Scanning for Files

1. Enter the directory path you want to scan (e.g., `Z:\ENGINEERING TEMPLATES\VISIO SHAPES 2025`)
2. Click the "Scan for Files" button
3. The application will display all temporary Visio files (matching pattern `~$*.vssx`) found in the specified directory and its subdirectories

### Deleting Files

1. After scanning, select the files you wish to delete by checking their corresponding checkboxes
2. Click the "Delete Selected Files" button
3. Confirm the deletion when prompted
4. The application will show a success message once files are deleted

### API Usage

If you prefer using the API directly:

#### Scan API

```bash
curl -X POST -H "Content-Type: application/json" -d '{"directory":"Z:\\ENGINEERING TEMPLATES\\VISIO SHAPES 2025"}' http://localhost:3000/api/scan
```

#### Delete API

```bash
curl -X POST -H "Content-Type: application/json" -d '{"files":["Z:\\ENGINEERING TEMPLATES\\VISIO SHAPES 2025\\~$SCHEMATIC TEMPLATES.vssx"]}' http://localhost:3000/api/delete
```

## Troubleshooting

### Common Issues

#### Server Won't Start

If the server won't start, check the following:

1. Ensure Node.js is properly installed:
```bash
node --version
```

2. Verify that port 3000 is not already in use:
```powershell
# For Windows (PowerShell)
Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue

# For Unix-based systems
lsof -i :3000
```

3. Check for dependency issues:
```bash
npm install
```

#### Files Not Found

If the application can't find any files:

1. Ensure the directory path is correct and accessible
2. Verify that there are indeed `~$*.vssx` files in the specified directory
3. Try using absolute paths instead of relative paths

For Windows paths that include spaces or special characters, ensure they are properly escaped:
```
Z:\My Folder\With Spaces  # Correct format for the UI
Z:\My Folder\With Spaces  # Correct format for API JSON
```

#### Permission Issues

If you encounter permission errors when deleting files:

1. Ensure you have appropriate permissions to the directories and files
2. Try running the application with administrator privileges
3. Check if the files are currently in use by another application (like Visio)

#### PowerShell Execution Policy

If PowerShell commands fail to execute:

```powershell
# Check your current execution policy
Get-ExecutionPolicy

# If needed, set to a less restrictive policy (as Administrator)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Testing API Endpoints

You can test the API endpoints using the included PowerShell script:

```powershell
.\test-api.ps1
```

This script will:
1. Start the server
2. Test the scan API with a sample directory
3. Optionally test the delete API with a sample file
4. Stop the server

### Need More Help?

If you continue to encounter issues:

1. Check the console output for specific error messages
2. Examine the PowerShell logs for command execution errors
3. Review the Node.js server logs for API errors
4. Contact the development team with specific error details
