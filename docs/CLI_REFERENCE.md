# Command-Line Reference

## 1. Preferred command forms

Run from the repository root:

```powershell
python run.py --file statement.pdf
.\run_bank_parser.bat --file statement.pdf
.\stmt.bat --file statement.pdf
```

The `.bat` wrappers are Windows-specific. All forms ultimately run `src/code/run.py`.

## 2. Syntax

```text
run.py [-h] --pdf PDF [--bank BANK] [--pwd PWD] [--out OUT]
```

`--file` is an alias for `--pdf`.

## 3. Active parser options

### `-h`, `--help`

Shows usage, option descriptions, and the active bank list, then exits successfully.

```powershell
python run.py --help
```

### `--pdf <value>` / `--file <value>`

Required. Supplies one PDF or a semicolon-separated list.

Behavior:

- Quote a multi-file value; otherwise the shell may treat `;` specially.
- Whitespace around each item is removed and empty items are ignored.
- `.pdf` is added when a value has no suffix.
- An existing absolute or relative path is accepted.
- Otherwise the project searches `src/input/` and root `input/`.
- A filename may contain `$password` before `.pdf`.

Examples:

```powershell
python run.py --pdf axis
python run.py --file axis.pdf
python run.py --file "C:\Statements\July 2026.pdf"
python run.py --file "jan.pdf;feb.pdf;mar.pdf" --out quarter_1
```

For multiple PDFs, rows are merged in argument order. `Source` records the original filename and `Sno` is reassigned across the merged result.

### `--bank <code>`

Optional. Selects a parser explicitly. The value is case-insensitive and surrounding whitespace is ignored.

| Code | Bank |
|---|---|
| `axis` | Axis Bank |
| `bob` | Bank of Baroda |
| `boi` | Bank of India |
| `bom` | Bank of Maharashtra |
| `canara` | Canara Bank |
| `central` | Central Bank of India |
| `cub` | City Union Bank |
| `dbs` | DBS Bank |
| `federal` | Federal Bank |
| `hdfc` | HDFC Bank |
| `icici` | ICICI Bank |
| `idbi` | IDBI Bank |
| `idfc` | IDFC FIRST Bank |
| `indian` | Indian Bank |
| `indus` | IndusInd Bank |
| `iob` | Indian Overseas Bank |
| `kotak` | Kotak Mahindra Bank |
| `kvb` | Karur Vysya Bank |
| `pnb` | Punjab National Bank |
| `sbi` | State Bank of India |
| `southind` | South Indian Bank |
| `tmb` | Tamilnad Mercantile Bank |
| `unionbank` | Union Bank of India |

If omitted, the bank is auto-detected separately for each PDF. If provided for a multi-file run, the selected parser is applied to every PDF.

```powershell
python run.py --file statement.pdf --bank axis
python run.py --file statement.pdf --bank AXIS
python run.py --file "part1.pdf;part2.pdf" --bank iob
```

### `--pwd <password>`

Optional. Password used to unlock encrypted PDFs.

- `--pwd` overrides filename-embedded passwords.
- The same `--pwd` is used for every PDF in a multi-file run.
- Without it, each file may derive its own password from its filename.
- A wrong or missing password for an encrypted file is a runtime error.

```powershell
python run.py --file protected.pdf --bank icici --pwd "Secret123"
python run.py --file "a.pdf;b.pdf" --pwd "SharedPassword"
```

Passwords may be visible in shell history or process listings. Do not put real passwords in documentation or committed scripts.

### `--out <name>`

Optional. Controls the final workbook stem and log filename.

- Defaults to the first PDF filename without `.pdf` or embedded password metadata.
- Path components are discarded; only the final name is used.
- A trailing `.xlsx` is removed.
- Windows-invalid characters `< > : " / \ | ? *` become `_`.
- Leading/trailing spaces and periods are removed.
- An empty sanitized name is rejected with exit code 2.
- `output/output.xlsx` keeps its fixed name regardless of `--out`.

```powershell
python run.py --file statement.pdf --out customer_july
python run.py --file statement.pdf --out customer_july.xlsx
python run.py --file "jan.pdf;feb.pdf" --out "Customer - Jan-Feb"
```

If `output/customer_july.xlsx` exists, the new final workbook becomes `customer_july_YYMMDD_HHMMSS.xlsx`. `src/logs/customer_july.log` is not timestamped and is opened in append mode.

## 4. Password in filename

A filename can embed a per-file password after the first `$`:

```text
statement$Secret123.pdf
```

Either command can locate it:

```powershell
python run.py --file 'statement$Secret123.pdf' --bank icici
python run.py --file statement.pdf --bank icici
```

The plain form works only when exactly one `statement$*.pdf` match exists at the resolved location. In PowerShell, single-quote a literal `$` filename to prevent variable expansion. The default output stem is `statement`.

A filename password is visible in directory listings and the source path may be logged. `--pwd` avoids the filename but may expose the password through shell history.

## 5. Workflow examples

### Auto-detect one file

```powershell
.\run_bank_parser.bat --file statement.pdf
```

### Force a parser

```powershell
.\run_bank_parser.bat --file statement.pdf --bank federal
```

### Merge statement segments

```powershell
.\run_bank_parser.bat --file "statement_1.pdf;statement_2.pdf" --bank axis --out full_statement
```

### Merge different banks

Omit `--bank` so each PDF is detected independently:

```powershell
.\run_bank_parser.bat --file "axis.pdf;hdfc.pdf;icici.pdf" --out combined_accounts
```

### Encrypted input

```powershell
python run.py --file protected.pdf --bank kvb --pwd "password"
```

## 6. Runtime output

The CLI reports file start, selected/detected bank, row progress, a five-second status ticker, per-file row count, workbook-building status, and final paths.

Non-fatal warnings may include negative balances, reconciliation differences, PDF structural/metadata indicators, or unusable rules. On an exception, the CLI prints the error and log path.

## 7. Exit codes

| Code | Conditions |
|---:|---|
| `0` | Help displayed or parsing/workbook creation completed. |
| `1` | Missing dependency, authentication failure, parser/OCR/Excel error, or another runtime exception. |
| `2` | Invalid CLI syntax, unsupported bank, empty file list, missing input, or invalid `--out`. |

Batch wrappers return the Python process exit code.

## 8. Setup command parameters

`install_new_machine.bat` and `setup_windows.bat` forward options to `setup_windows.ps1`:

```powershell
.\install_new_machine.bat [-SkipTesseract] [-ForceRecreateVenv]
```

| Option | Effect |
|---|---|
| `-SkipTesseract` | Skips Tesseract detection/installation. Text-only layouts continue to work, but OCR layouts and OCR bank detection do not. |
| `-ForceRecreateVenv` | Removes `.venv`, creates it again, and reinstalls `requirements.txt`. |

```powershell
.\setup_windows.bat -SkipTesseract
.\setup_windows.bat -ForceRecreateVenv
.\setup_windows.bat -ForceRecreateVenv -SkipTesseract
```

## 9. Packaging command parameters

### Recommended package builder

```powershell
.\build_fresh_machine_package.bat [--include-input-pdfs]
```

| Option | Effect |
|---|---|
| `-h`, `--help` | Shows help when invoking `build_fresh_machine_package.py` directly. |
| `--include-input-pdfs` | Includes PDFs under `input/`; by default they are excluded. |

Output: `dist/bankStmtv1_fresh_windows_YYMMDD_HHMMSS.zip`.

### Alternate PowerShell package builder

```powershell
.\build_windows_package.ps1 [-OutputDir <directory>] [-IncludeInputPdfs] [-IncludeOutputFiles]
```

| Parameter | Default | Effect |
|---|---|---|
| `-OutputDir` | `dist` | Staging and zip output directory. |
| `-IncludeInputPdfs` | off | Copies root `input/*.pdf`. |
| `-IncludeOutputFiles` | off | Copies files from root `output/`. |

Prefer `build_fresh_machine_package.bat` for the maintained clean-package workflow.

## 10. Alternate scaffold CLI

`src/main.py` defines a different, smaller CLI:

```text
python src/main.py --bank {axis,boi,iob,kotak,southind,tmb,unionbank} --pdf <filename>
```

Both arguments are required and `--pdf` is one filename inside root `input/`. It does not expose `--file`, multiple PDFs, `--pwd`, `--out`, or active auto-detection. It is not the recommended runtime command.

