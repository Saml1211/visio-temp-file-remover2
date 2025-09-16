1. Install the required Python packages:
   ```powershell
   py -3 -m pip install --upgrade pip
   py -3 -m pip install pyinstaller
   ```

2. Create the standalone executable:
   ```powershell
   py -3 -m PyInstaller --onefile --windowed --name "VisioTempFileRemover" --add-data "scripts;scripts" visio_gui.py
   ```

3. The executable will be created in the `dist` folder

## Running the Executable

After creating the executable:
1. Run the `VisioTempFileRemover.exe` file directly (the scripts folder is bundled)
2. The application uses a `resource_path()` helper to load bundled resources correctly

## Notes

- The executable will be quite large due to the Python runtime
- You may need to run the executable as Administrator depending on the directories you're scanning
- Windows SmartScreen may warn about the executable since it's not signed
- PyInstaller `--add-data` uses `;` on Windows and `:` on macOS/Linux