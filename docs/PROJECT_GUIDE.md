# Complete Project Guide

## 1. What this project does

This project reads one or more bank-statement PDF files, selects a bank-specific parser, converts transaction rows into a common schema, writes an intermediate Excel file, and builds a formatted analysis workbook.

The project is intended for statement analysis rather than general PDF-to-Excel conversion. Each bank prints dates, narrations, cheque numbers, debit/credit amounts, and balances differently. The parser layer handles those differences, while the rest of the pipeline works with a consistent transaction model.

A normal run performs the following work:

1. Parse and validate command-line arguments.
2. Resolve each requested PDF from a path or the project input folders.
3. Determine a password from `--pwd` or filename metadata.
4. Create a temporary decrypted copy when a PDF is encrypted.
5. Use `--bank`, or auto-detect a bank separately for each PDF.
6. Run the corresponding bank parser.
7. Warn about negative balances and reconcile extracted totals when the PDF exposes summary totals.
8. Merge all parsed rows and renumber `Sno` from 1.
9. Normalize the output to the common transaction columns.
10. Overwrite `output/output.xlsx` with the merged statement.
11. Load `input/Rules.xlsx` and build rule-based analysis sheets.
12. Build and style the final workbook, including PDF status checks and summary sheets.
13. Remove temporary decrypted files and report output/log paths.

## 2. Active entry points

### `run.py`

This is the preferred Python entry point. It:

- Locates the repository root from its own file location.
- Checks for `.venv\Scripts\python.exe`.
- Ignores a copied/stale virtual environment when `pyvenv.cfg` points to a missing base Python installation.
- Re-executes itself with the virtual-environment Python when the environment is healthy and the current interpreter is different.
- Adds `src/code` and `src` to `sys.path` and runs `src/code/run.py`.
- Prints dependency-repair instructions if an import fails.

The internal environment variable `BANKSTMT_SKIP_VENV_REEXEC=1` prevents a re-execution loop. It is set automatically and is not a normal user option.

### `run_bank_parser.bat` and `stmt.bat`

These Windows wrappers have equivalent behavior:

1. Use `.venv\Scripts\python.exe` when present.
2. Otherwise try `py -3`.
3. If the Windows Python launcher is unavailable, try `python`.
4. Forward all arguments to `src\code\run.py` and return its exit code.

### `src/code/run.py`

This is the active application CLI and orchestration layer. Direct execution is supported, although the root `run.py` or batch wrapper is more convenient because the root launcher handles a healthy virtual environment automatically.

### `src/main.py`

`src/main.py` is a separate, earlier modular pipeline scaffold. Its CLI requires `--bank` and `--pdf`, uses `src/parsers/parser_registry.py`, normalizes through `src/transform/`, and exports through `src/export/`. It supports fewer registered banks and is not the command used by the Windows wrappers or current full workbook flow.

The active CLI is defined in `src/code/run.py` and supports multiple files, auto-detection, passwords, and `--out`.

## 3. End-to-end runtime flow

### 3.1 Argument validation

The active CLI requires `--pdf`/`--file`. `--bank`, `--pwd`, and `--out` are optional. Invalid argparse syntax exits with code 2 and prints an example command.

An explicit bank value is trimmed and lowercased. If it is not in the active parser map, the application prints all supported codes and exits with code 2.

### 3.2 File list and path resolution

The file argument is split on semicolons. Empty parts are ignored. A missing suffix is changed to `.pdf`.

For each value, the resolver tries:

1. The supplied value as a path, relative to the current directory or absolute.
2. `src/input/<value>`.
3. `input/<value>` at the repository root.

At each location, the resolver can also recover a unique password-bearing filename such as `statement$secret.pdf` when the user requested `statement.pdf`. Ambiguous matches are not selected.

### 3.3 Password handling

The password for each file is selected in this order:

1. `--pwd`, when supplied.
2. Text after the first `$` in that PDF's filename stem.
3. No password.

Encrypted files are opened with PyMuPDF. If authentication succeeds, an unencrypted temporary copy is written under `output/_tmp_run_<milliseconds>/` and passed to parsers. The temporary directory is removed in a `finally` block.

### 3.4 Parser selection

If `--bank` is present, the same selected parser is used for every file. Otherwise, each PDF is independently auto-detected using the first two pages of extracted text, OCR of the first page if necessary, and finally the filename as a fallback. Detection is described in [PARSER_REFERENCE.md](PARSER_REFERENCE.md).

### 3.5 Transaction extraction

Every active parser returns a list of transaction dictionaries. Parsers use `pdfplumber` tables/words, PyMuPDF lines/words, Tesseract OCR, and layout-specific regular expressions or balance-difference logic. Wrapped narration lines are joined where supported.

### 3.6 Progress and sanity checks

The terminal displays file start/completion messages, `Processing row : N`, and a five-second ticker containing file number, filename, rows parsed so far, and elapsed seconds.

After each file, numeric balances below zero are counted. The terminal prints up to the first three and last three examples and asks the operator to cross-check the PDF. This warning does not stop the run.

### 3.7 Reconciliation

For each PDF, the application logs parsed transaction count, total debit, and total credit. It scans up to the first three and last three pages for `Total Debit`, `Total Credit`, and `Total Transaction(s)`.

When those values are found, parsed results are compared with a tolerance of 0.01 for money and exact equality for count. Mismatches are warnings and do not stop workbook creation. If a statement uses other wording or image-only totals, no expected totals may be found.

### 3.8 Merge and normalize

Rows are appended in command-line file order. Each row receives its original PDF filename in `Source`, and the merged list is renumbered sequentially in `Sno`.

`records_to_dataframe()` guarantees the common output columns and applies cheque-number sanitization. Missing columns are added with empty values.

### 3.9 Excel creation

The merged data is written to `output/output.xlsx`, sheet `Statement`. The final workbook is then assembled from the statement, rules, PDF information, and built-in analysis. See [INPUT_RULES_OUTPUTS.md](INPUT_RULES_OUTPUTS.md).

## 4. Common transaction data model

| Column | Meaning |
|---|---|
| `Sno` | Sequential row number across the merged run. |
| `Date` | Transaction date, normally normalized to `DD/MM/YYYY` internally and written as an Excel date when recognized. |
| `Details` | Human-readable narration/transaction description. |
| `Detail_Clean` | `Details` with non-alphanumeric characters removed; used for rule matching. |
| `Cheque No` | Sanitized cheque/instrument number, stored as text so leading zeroes survive. |
| `Debit` | Debit/withdrawal amount, or blank/zero when not applicable. |
| `Credit` | Credit/deposit amount, or blank/zero when not applicable. |
| `Balance` | Running account balance after the transaction when available. |
| `Source` | Original PDF filename. |

The final builder coerces `Debit` and `Credit` to numbers and replaces missing/non-numeric values with `0.0`. Monetary cells in most final sheets are rounded to whole values for display and formatted with Indian digit grouping (`#,##,##0`). The intermediate workbook retains values before final styling.

## 5. Shared normalization

### Dates

Shared helpers recognize common numeric and month-name forms, including `DD/MM/YYYY`, `DD-MM-YYYY`, two-digit years, ISO `YYYY-MM-DD`, and `DD-Mon-YYYY`. Individual parsers may support more. The final builder writes recognized values as real Excel dates using `yyyy-mm-dd` display format.

### Amounts

Shared amount parsing removes commas/spaces; ignores common `CR`, `DR`, `INR`, `RS`, and `MR`; supports parentheses or minus signs as negatives; accepts leading decimals; and returns no value for blank or unparseable text. Parsers may use running balance changes to decide debit versus credit.

### Cheque numbers

- Numeric values are retained, including leading zeroes.
- `123456.0` becomes `123456`.
- All-zero and non-numeric values are removed.
- Cheque/clearing hints permit numeric cheque values.
- UPI, IMPS, NEFT, RTGS, card, ATM, transfer, and similar hints suppress unrelated references.
- Known narration patterns can supply a missing cheque number.

### Formula-like text safety

Strings beginning with `=` are forced to Excel text cells before saving intermediate and final workbooks.

## 6. Logging and errors

Each run writes `src/logs/<output-stem>.log`. It records parser details, reconciliation, rule matches, workbook creation, warnings, and tracebacks. Reusing an output stem appends to the existing log.

| Code | Meaning |
|---:|---|
| `0` | Run completed and outputs were written. |
| `1` | Runtime failure such as parser, PDF, dependency, or workbook error. |
| `2` | CLI usage, unsupported bank, empty file list, missing input, or invalid output name. |

## 7. Architecture

### Active application

- `run.py` — root launcher and virtual-environment handoff.
- `src/code/run.py` — CLI, active parser map, orchestration, logging, decryption, merge, and output calls.
- `src/code/bank_detector.py` — bank scoring and OCR fallback.
- `src/code/utils.py` — cleaning, dates, amounts, cheque normalization, generic table parsing, input/decryption, reconciliation, and intermediate output.
- `src/code/final_excel_builder.py` — rules, analysis sheets, PDF status, styles, charts, and final naming.
- `src/code/*_parser.py` — active bank adapters and implementations.
- `src/parsers/*_parser.py` — modular parsers, some reused by active adapters.

### Modular scaffold

- `src/main.py` — alternate minimal pipeline.
- `src/parsers/base_parser.py`, `parser_registry.py`, and `detector.py` — class interface and smaller registry; scaffold auto-detection is unimplemented.
- `src/transform/` — alternate normalization/validation.
- `src/export/` and `src/utils/` — alternate export/helpers.

The active application uses `Sno`, `Date`, and `Details`; the scaffold targets fields such as `Txn_Date`, `Value_Date`, `Description`, `Bank`, and `Account_Number`. They are not interchangeable.

## 8. Reliability boundaries

- Extraction is layout-sensitive; bank template changes can require parser updates.
- Auto-detection selects the highest weighted match and has no confidence threshold beyond at least one match.
- Reconciliation requires recognizable summary text.
- Negative balances may be legitimate overdrafts.
- `PDF_Status` is heuristic and does not cryptographically validate signatures.
- Verify results against original statements before consequential decisions.

