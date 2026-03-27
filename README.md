# Ballistic Installer USB Copier

This folder contains the Python source for the desktop app that copies selected installer files from `F:\ballistic target` to one or more USB drives.

## Usability Features

- Automatically refreshes and shows likely USB target drives
- Uses checkbox-based drive selection instead of typing drive letters
- Remembers the last source folder, destination folder, selected files, and selected drives
- Shows selected file count and total size before copying
- Lets you choose whether to erase selected USB drives first or copy without wiping
- Includes a preview step that shows the exact files to copy and what will be deleted or overwritten
- Uses a stronger confirmation dialog that names the selected target drives before copying
- Uses a clearer review-first flow with larger summaries and a stronger start-copy action

## Files

- `main.py`: Tkinter desktop app
- `build-exe.ps1`: Rebuilds the standalone `.exe` with PyInstaller

## Run the Python app

```powershell
python .\main.py
```

## Build the standalone EXE

```powershell
.\build-exe.ps1
```

The built executable is written to:

```text
.\dist\BallisticInstallerUsbCopier.exe
```

Saved settings are stored in:

```text
%USERPROFILE%\AppData\Local\BallisticInstallerUsbCopier\settings.json
```
