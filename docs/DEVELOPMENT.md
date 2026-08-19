# Development, Testing, Packaging, and Troubleshooting

## 1. Requirements

- Windows 10/11 for supplied setup and batch scripts.
- Python 3.11+ according to `setup_windows.ps1`.
- Internet during installation.
- `winget` for automatic Python/Tesseract installation when needed.
- Tesseract for OCR layouts and OCR detection.

| Package | Range | Purpose |
|---|---|---|
| `pandas` | `>=2.2.0,<3.0.0` | DataFrames, rules, Excel assembly. |
| `openpyxl` | `>=3.1.0,<4.0.0` | `.xlsx` writing, formatting, charts. |
| `pdfplumber` | `>=0.11.0,<1.0.0` | PDF text, words, tables. |
| `pymupdf` | `>=1.27.1,<2.0.0` | PDF access, rendering, metadata, decryption. |
| `Pillow` | `>=12.0.0,<13.0.0` | OCR/chart images. |
| `pytesseract` | `>=0.3.13,<1.0.0` | Interface to external Tesseract. |

## 2. Automated Windows setup

```powershell
.\install_new_machine.bat
```

The setup:

1. Searches the Python launcher, PATH, and registered installations.
2. Rejects the Microsoft Store alias.
3. Requires Python 3.11+ and can install 3.11 with `winget`.
4. Checks `.venv` health and recreates a stale copied environment.
5. Creates `.venv`, upgrades pip, and installs `requirements.txt`.
6. Finds/installs Tesseract unless skipped and saves `TESSERACT_CMD` for the user.
7. Runs `py_compile` on project Python files.

```powershell
.\setup_windows.bat -SkipTesseract
.\setup_windows.bat -ForceRecreateVenv
```

## 3. Manual setup

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

```powershell
winget install --id UB-Mannheim.TesseractOCR -e --accept-package-agreements --accept-source-agreements
```

## 4. Repository layout

```text
.
├── run.py                         Root launcher/.venv handoff
├── run_bank_parser.bat, stmt.bat  Windows run wrappers
├── setup_windows.*                Environment installer
├── build_fresh_machine_package.*  Recommended zip builder
├── build_windows_package.ps1      Alternate package builder
├── input/                         Rules.xlsx and conventional PDFs
├── output/                        Generated workbooks/temp data
├── src/code/                      Active CLI, parsers, helpers, builder
├── src/parsers/                   Modular/shared parsers
├── src/transform/, export/, utils/ Scaffold pipeline components
├── src/logs/                      Active logs
└── tests/                         unittest suite
```

## 5. Module responsibilities

| Module | Responsibility |
|---|---|
| `run.py` | `.venv` detection/re-execution and active launch. |
| `src/code/run.py` | CLI, active parser map, inputs, passwords, progress, merge, output coordination. |
| `src/code/bank_detector.py` | Signatures, native extraction, OCR and filename fallbacks. |
| `src/code/utils.py` | Cleaning, dates, amounts, cheques, generic parsing, paths/decryption, reconciliation. |
| `src/code/parser_helpers.py` | Shared record, signed-balance, and date helpers. |
| `src/code/final_excel_builder.py` | Rules, analysis sheets, PDF status, styles, chart, naming. |
| `src/code/*_parser.py` | Active bank adapters/implementations. |
| `src/parsers/*_parser.py` | Class parsers and record parsers reused by adapters. |

## 6. Tests

```powershell
python -m unittest discover -s tests -p "test_*.py"
python -m unittest tests.test_cheque_normalization
python -m unittest tests.test_axis_layouts.AxisLayoutTests
```

Some regression tests require local representative PDFs and skip when absent. Synthetic unit tests still cover key layout logic.

Coverage includes Axis layouts/header reuse; bank regressions; cheque normalization; balance consistency; multi-page/wrapped details; rule merging; return/reject filtering; repeat amount ordering; monthly summaries/style; number formatting; PDF summary/status placement; edge-page totals; and negative-balance reporting.

Other checks:

```powershell
python -m compileall src tests
python run.py --help
```

After integration testing, inspect the terminal, log, intermediate workbook, final workbook, and beginning/middle/end source transactions.

## 7. Parser test design

Test header mapping, first/last row, multiple pages, headerless continuation pages, wrapped narration, debit/credit/balance consistency, summary exclusion, date normalization, leading-zero cheques, invalid cells, progress callbacks, old-layout regression, and detection signatures.

## 8. Packaging

```powershell
.\build_fresh_machine_package.bat
.\build_fresh_machine_package.bat --include-input-pdfs
```

Output: `dist/bankStmtv1_fresh_windows_YYMMDD_HHMMSS.zip`.

The Python packager includes setup/docs, `src/`, `input/`, and `tests/`, and creates placeholder input/output/log directories. It excludes `.git`, `.venv`, caches, `dist`, configured temp directories, bytecode, logs, output contents, Excel lock files, and input PDFs unless requested.

Review packages before sharing because PDFs and rules may contain sensitive data.

## 9. Troubleshooting

### Stale virtual environment

```powershell
.\setup_windows.bat -ForceRecreateVenv
```

### Python not recognized

Try `py -3 --version`, then install Python 3.11+ and rerun setup. The setup deliberately ignores the WindowsApps Store alias.

### Missing dependency

```powershell
.\install_new_machine.bat
```

The root launcher's current error text mentions `install_fresh_machine.bat`, but the actual maintained file is `install_new_machine.bat`.

### Input not found

Check spelling/extension, use root `input/`, quote spaces, and single-quote `$` filenames in PowerShell. A wrong non-`.pdf` suffix is not corrected.

### Auto-detection failed

Use explicit `--bank`. TMB always needs `--bank tmb` until a detector signature is added.

### Encrypted PDF failed

Verify `--pwd`. For different passwords, use per-file filename metadata or separate runs because one `--pwd` applies to every file.

### OCR failed

Run `tesseract --version`. Set `TESSERACT_CMD` for nonstandard installations and open a new terminal.

### Zero rows

Verify the bank/layout, read the log, check whether the PDF is scanned, and add a regression test before changing code. The process may still build an empty workbook; treat zero rows as a failed business result.

### Reconciliation mismatch

Review page boundaries, wrapped rows, summary rows, missing pages, debit/credit direction, OCR decimals, and what the printed totals include.

### Rules do not create a sheet

Check `Rules.xlsx`, first-sheet headers, nonblank `SheetName`, numeric `AMT` values, `Detail_Clean`, and rule messages in the log.

### Workbook locked

Close Excel. The final path can timestamp, but `output/output.xlsx` always requires the same writable path.

### PowerShell blocked

Use `.\setup_windows.bat`; it invokes the setup script with execution-policy bypass.

## 10. Security

- Do not commit PDFs, output workbooks, logs, or passwords.
- Restrict `input/`, `output/`, and `src/logs/`.
- Remove decrypted `_tmp_run_*` leftovers after abnormal termination.
- Inspect zips before sharing.
- Treat filename and shell-history passwords as exposed metadata.
- `PDF_Status` is heuristic, not forensic proof.
- Validate transactions against originals before decisions.

## 11. Known maintenance notes

- Active CLI and `src/main.py` scaffold coexist with different schemas/bank lists.
- TMB parsing is active but TMB detection is missing.
- Intermediate output is overwritten.
- Final money styling rounds displayed values to whole numbers.
- Timestamp collision handling has no same-second counter.
- Recommended Python and alternate PowerShell packagers both exist; prefer and test the recommended path.

