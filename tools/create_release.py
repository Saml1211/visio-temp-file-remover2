# Visio Temp File Remover GUI Release Packaging Script

import os
import shutil
import zipfile
from pathlib import Path

def create_release_package():
    """Create a release package for the Visio Temp File Remover GUI"""
    
    # Define package name and version
    package_name = "VisioTempFileRemover-GUI"
    # Prefer environment override, e.g., VTFR_GUI_VERSION=1.0.1
    version = os.getenv("VTFR_GUI_VERSION", "1.0.0")
    
    # Create release directory
    release_dir = Path("release")
    release_dir.mkdir(exist_ok=True)
    
    # Create package directory
    package_dir = release_dir / f"{package_name}-v{version}"
    if package_dir.exists():
        shutil.rmtree(package_dir)
    package_dir.mkdir()
    
    # Files to include in the release
    files_to_include = [
        "visio_gui.py",
        "run_gui.bat",
        "LICENSE",
        "config.json",
        "scripts/Scan-VisioTempFiles.ps1",
        "scripts/Remove-VisioTempFiles.ps1",
        "scripts/.placeholder"
    ]
    
    # Copy files to package directory
    for file_path in files_to_include:
        src_path = Path(file_path)
        if src_path.exists():
            dst_path = package_dir / file_path
            dst_path.parent.mkdir(parents=True, exist_ok=True)
            if src_path.is_dir():
                shutil.copytree(src_path, dst_path)
            elif src_path.is_file():
                shutil.copy2(src_path, dst_path)
    
    # Copy documentation files
    docs_to_include = [
        "docs/gui.md",
        "docs/installation.md",
        "docs/release-notes.md"
    ]
    
    docs_dir = package_dir / "docs"
    docs_dir.mkdir(exist_ok=True)
    
    for doc_path in docs_to_include:
        src_path = Path(doc_path)
        if src_path.exists():
            dst_path = package_dir / doc_path
            shutil.copy2(src_path, dst_path)
    
    # Create a simple installation guide
    install_guide = """# Visio Temp File Remover GUI - Installation Guide

## System Requirements
- Windows operating system (Windows 7 or later)
- Python 3.6 or higher installed
- PowerShell (included with Windows)

## Installation Steps
1. Extract this zip file to a folder of your choice
2. Ensure Python is installed and accessible from the command line
3. Double-click on `run_gui.bat` to start the application

## Usage
1. Run the application using `run_gui.bat`
2. Select the directory you want to scan for Visio temp files
3. Click "Scan for Files"
4. Review the found files in the results list
5. Select the files you want to delete
6. Click "Delete Selected Files"
7. Confirm the deletion when prompted

## Notes
- The application runs locally and does not require internet access
- No installation of additional Python packages is required
- The application uses the same PowerShell scripts as the web version for consistency
- For detailed instructions, see docs/installation.md
- For GUI information, see docs/gui.md
"""
    
    with open(package_dir / "INSTALLATION.md", "w", encoding="utf-8", newline="\n") as f:
        f.write(install_guide)
    
    # Create a release zip file
    zip_filename = f"{package_name}-v{version}.zip"
    zip_path = release_dir / zip_filename
    
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as zipf:
        for root, _, files in os.walk(package_dir):
            for file in files:
                file_path = Path(root) / file
                arc_path = file_path.relative_to(release_dir)
                zipf.write(file_path, arc_path)
    
    print(f"Release package created: {zip_path}")
    print(f"Package contents are in: {package_dir}")

if __name__ == "__main__":
    create_release_package()