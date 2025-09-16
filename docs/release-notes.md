# Visio Temp File Remover GUI - Release Notes

## Version 1.0.0

### New Features
- **Standalone Desktop Application**: Complete GUI application built with Python Tkinter that runs locally without requiring a web server
- **User-Friendly Interface**: Intuitive graphical interface for scanning and removing Visio temporary files
- **Directory Selection**: Easy browsing and selection of directories to scan
- **File Display**: Tabular view showing file name, relative path, size, and last modified date
- **File Selection**: Select individual or multiple files for deletion
- **Safe Deletion**: Confirmation prompts before deleting files
- **Progress Indicators**: Visual feedback during scanning and deletion operations
- **Error Handling**: Comprehensive error handling with user-friendly messages

### Technical Features
- **Dual Approach**: Uses both PowerShell scripts and direct PowerShell commands for maximum compatibility
- **Fallback Mechanisms**: Automatically falls back to direct PowerShell commands if scripts fail
- **Consistent Behavior**: Same file detection and deletion logic as the web version
- **Cross-Platform Path Handling**: Proper handling of Windows file paths
- **JSON Parsing**: Robust parsing of PowerShell output in various formats
- **Threading**: Background processing to prevent UI freezing during operations

### Improvements
- **Optimized Display**: Relative paths shown instead of full paths for better readability
- **Column Widths**: Adjusted column widths for optimal information display
- **PowerShell Detection**: Improved detection and error handling for PowerShell availability
- **File Size Formatting**: Human-readable file sizes (B, KB, MB, GB)

### Requirements
- Windows operating system
- Python 3.6 or higher
- PowerShell (included with Windows)

### Installation
1. Extract the release package
2. Double-click `run_gui.bat` to start the application
3. Or run `python visio_gui.py` from the command line

### Usage
1. Select a directory to scan for Visio temporary files
2. Click "Scan for Files"
3. Review the results in the table
4. Select files to delete
5. Click "Delete Selected Files"
6. Confirm the deletion

### Files Included
- `visio_gui.py`: Main application file
- `run_gui.bat`: Batch file to easily start the application
- `docs/gui.md`: Documentation for the GUI application
- `docs/quickstart.md`: Quick start and usage guide
- `docs/build-gui.md`: Instructions for creating a standalone executable
- `scripts/Scan-VisioTempFiles.ps1`: PowerShell script for scanning files
- `scripts/Remove-VisioTempFiles.ps1`: PowerShell script for deleting files
- `config.json`: Configuration file
- `LICENSE`: License information

### Known Issues
None at this time.

### Future Enhancements (Planned)
- Configuration options for file patterns
- Logging of operations
- Recent directories list
- Export functionality for scan results