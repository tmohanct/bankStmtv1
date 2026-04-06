# Windows Setup

## Recommended machine prerequisites
- Windows 10 or Windows 11
- Internet connection for `pip install`
- `winget` if you want the setup script to auto-install Python or Tesseract
- Tesseract OCR only if you need `icici` statement support

## Fastest setup on a new machine
1. Copy or clone this repo to the new machine.
2. Open PowerShell in the repo root.
3. Run:
   ```powershell
   .\install_new_machine.bat
   ```
4. After setup finishes, run the parser with:
   ```powershell
   .\run_bank_parser.bat --bank sbi --file sbi.pdf
   ```

If the folder was copied from another machine, the installer will automatically rebuild a stale `.venv`.
If Python is missing and `winget` is available, the installer will try to install Python 3.11 automatically.

## Create a clean package to share with another machine
Run this on the source machine:

```powershell
.\build_fresh_machine_package.bat
```

That creates a zip inside `dist\` with:
- source code
- setup scripts
- `input\Rules.xlsx` if present

By default it excludes:
- `.venv`
- logs
- existing output files
- input PDFs

If you want to include PDFs in the package:

```powershell
.\build_fresh_machine_package.bat --include-input-pdfs
```

## What the setup script does
- Finds a usable Python 3.11+ install and ignores the Microsoft Store `python.exe` alias in `WindowsApps`
- Installs Python 3.11 with `winget` if Python is missing and `winget` is available
- Creates `.venv`
- Recreates `.venv` automatically if it was copied from another machine and still points to an old Python path
- Upgrades `pip`
- Installs all Python packages from `requirements.txt`
- Checks for Tesseract OCR and tries to install it with `winget`
- Saves `TESSERACT_CMD` for the current Windows user when Tesseract is found
- Verifies project Python files with `py_compile`

## If you do not need ICICI support
You can skip the OCR step:

```powershell
.\setup_windows.bat -SkipTesseract
```

## Manual setup commands
If you prefer doing it yourself:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

If `py` is not available yet, install Python 3.11 first:

```powershell
winget install --id Python.Python.3.11 -e --accept-package-agreements --accept-source-agreements
```

If you need ICICI support and Tesseract is not installed, install it manually or use `winget`:

```powershell
winget install --id UB-Mannheim.TesseractOCR -e --accept-package-agreements --accept-source-agreements
```

## Input and output
- Put PDFs in `input\`
- Keep `Rules.xlsx` in `input\`
- `output.xlsx` is written to `output\output.xlsx`
- Final workbooks are written to `output\`

## Run examples
Single file:

```powershell
.\run_bank_parser.bat --bank hdfc --file hdfc.pdf
```

Multiple files:

```powershell
.\run_bank_parser.bat --bank axis --file "axis.pdf;axis2.pdf"
```

SBI:

```powershell
.\run_bank_parser.bat --bank sbi --file sbi.pdf
```

ICICI:

```powershell
.\run_bank_parser.bat --bank icici --file icici.pdf
```

Kotak:

```powershell
.\run_bank_parser.bat --bank kotak --file BALASUBRAMANIAN.pdf
```

## Supported bank codes
- `axis`
- `bob`
- `bom`
- `canara`
- `cub`
- `federal`
- `hdfc`
- `icici`
- `idbi`
- `idfc`
- `indian`
- `indus`
- `iob`
- `kotak`
- `kvb`
- `pnb`
- `sbi`
- `southind`
- `unionbank`

## Troubleshooting
If `python` is not recognized:
- The Microsoft Store alias may be taking over `python.exe`
- Run `py -3 --version`; if that works, rerun `.\install_new_machine.bat`
- Otherwise install Python 3.11 from python.org with `Add python.exe to PATH`, or use:
  `winget install --id Python.Python.3.11 -e --accept-package-agreements --accept-source-agreements`

If ICICI fails on the new machine:
- Check that `tesseract.exe` exists
- Check that `TESSERACT_CMD` points to the correct path
- Restart PowerShell after Tesseract installation

If PowerShell blocks the setup script:
- Use `setup_windows.bat` instead of running the `.ps1` file directly
