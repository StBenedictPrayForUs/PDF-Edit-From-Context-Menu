# PDF Splitter (Personal Windows Utility)

A local Windows-only PDF splitting utility with:
- tray startup (always available)
- File Explorer right-click action for `.pdf`
- visual page list with split starts
- per-page rotation in 90-degree steps
- editable output names with defaults

## 1) Install

From this folder in PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
./scripts/install.ps1
```

This will:
- install dependencies from `requirements.txt`
- register `Split PDF...` in your PDF right-click menu (current user)
- create a startup shortcut so the tray app launches at logon
- launch the tray app now

## 2) Use

- Right-click any PDF in File Explorer -> `Split PDF...`
- In the app:
  - check `Start split here` on pages that begin a new output file
  - edit section names on the right
  - optionally rotate selected pages with `Rotate -90` / `Rotate +90`
  - click `Export Splits`
- Output files are written next to the source PDF.
- Existing names are not overwritten; `_1`, `_2`, etc. are appended.

## 3) Optional CLI

From this project root:

```powershell
python -m app.tray_runtime
python -m app.launcher "C:\path\to\file.pdf"
```

## 4) Uninstall

```powershell
./scripts/uninstall.ps1
```

This removes startup + right-click integration.

## Troubleshooting

- If right-click does nothing, restart the tray app:
  1. Right-click tray icon -> `Quit`
  2. Run `pythonw .\\run_tray.py` from project root
- Check logs at:
  - `%LOCALAPPDATA%\\PDFSplitter\\app.log`
