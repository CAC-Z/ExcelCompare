# Excel Compare

Excel data comparison tool with a simple GUI. Select two Excel files, pick index columns, map columns to compare, and export a report with mismatches.

## Features

- GUI-based comparison (PyQt5)
- Fast header preview for column selection
- Column mapping with auto-match for same names
- Normalized comparison to reduce false mismatches
- Export result report (summary + details)

## Requirements

- Python 3.9+
- Windows (tested)

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python excel_compare.py
```

## Build (PyInstaller)

```powershell
pwsh ./build.ps1 -Mode onefile -NoUPX
```

## Project Files

- `excel_compare.py` main app
- `build.ps1` build script
- `ExcelCompare.spec` PyInstaller spec
- `icons/` UI assets
- `icons/icon.ico` / `icons/icon.png` app icon

## License

MIT. See `LICENSE`.
