# Bank Statement Parser — AI Project Context

## Purpose

This repository converts Indian bank-statement PDFs into Excel workbooks for review and analysis. It extracts transactions, normalizes the common fields, applies user-maintained matching rules, and produces analysis sheets such as cheque transactions, repeated amounts, return/reject entries, top debits/credits, and monthly totals.

The project is designed to support many banks. PDFs can be text-based or, for selected parsers and bank detection, OCR-assisted.

## Start Here: What Actually Runs

There are two implementations in the repository. They must not be confused.

1. **Active production implementation: `src/code/`**
   - Top-level `run.py` starts `src/code/run.py`.
   - `stmt.bat` and `run_bank_parser.bat` also start `src/code/run.py`.
   - This is the feature-complete path and is the path used for normal customer work.

2. **New modular implementation: `src/parsers/`, `src/transform/`, `src/export/`, and `src/main.py`**
   - This is a migration/scaffold toward a cleaner architecture.
   - It is not the default runtime path.
   - It has only a smaller parser registry and some unfinished pieces (notably the modular Axis parser and automatic detection).
   - Do not switch the launcher to this path without completing parity work and regression testing.

When changing live behavior, make the change in the active `src/code/` path unless the task explicitly concerns the migration.

## Runtime Flow

The active entry point is `src/code/run.py`.

```text
CLI / batch launcher
  -> resolve one or more PDFs
  -> decrypt password-protected files when necessary
  -> select bank parser (explicit --bank or automatic detection)
  -> parse each PDF into transaction dictionaries
  -> reconcile available PDF totals and warn on negative balances
  -> merge all transactions and add Source + final Sno
  -> output/output.xlsx (intermediate Statement sheet)
  -> final Excel analysis workbook in output/
  -> src/logs/<output-name>.log
```

### Command-line interface

Use the repository root launcher, normally through the virtual environment:

```powershell
python run.py --pdf "statement.pdf" --bank axis
python run.py --pdf "file1.pdf;file2.pdf" --out combined_output
python run.py --pdf "encrypted.pdf" --pwd mypassword
```

`--bank` is optional. Without it, the program attempts bank detection from text extracted from the first two PDF pages, then OCR, then the filename.

PDF names may embed a password using `$`, for example `CUSTOMER$secret.pdf`. An explicit `--pwd` overrides that embedded value. The source PDF itself is not modified; an unlocked temporary copy is created below `output/_tmp_run_*` and removed after processing.

## Inputs and Outputs

### Inputs

- `input/` contains source statement PDFs.
- The active pipeline expects `input/Rules.xlsx`.
- A statement can also be passed using a full path.

### Intermediate output

`output/output.xlsx` is overwritten on every run. It contains the merged `Statement` worksheet with these columns:

| Column | Meaning |
| --- | --- |
| `Sno` | Final transaction serial number across all input PDFs |
| `Date` | Normalized transaction date, normally `DD/MM/YYYY` |
| `Details` | Transaction narration/details |
| `Detail_Clean` | Alphanumeric-only key derived from `Details`, used for matching |
| `Cheque No` | Sanitized/extracted cheque number when applicable |
| `Debit` | Debit value or blank |
| `Credit` | Credit value or blank |
| `Balance` | Running balance |
| `Source` | Original source PDF filename |

### Final output

The final file is normally `output/<first-pdf-stem>.xlsx`, or `output/<--out value>.xlsx`. If that name already exists, the active builder adds a `YYMMDD_HHMMSS` timestamp.

It normally includes these sheets:

- `PDF_Status`: heuristic PDF integrity/modification indicators and first-PDF account summary.
- `Statement`: complete parsed statement.
- `Ret/Rej`: transactions whose details indicate cheque/electronic return, rejection, dishonour, or related charges.
- Rule-based sheets: matching transactions, grouped by `SheetName` in the rules workbook.
- `Cheque_Transactions`: rows with a usable cheque number.
- `Repeat_Credit_Amount` and `Repeat_Debit_Amount`: amounts occurring more than twice.
- `Top30_Debit` and `Top30_Credit`: 30 largest positive values.
- `month_dr_cr`: monthly debit, credit, net, end-of-month balance, counts, and averages.

`PDF_Status` is a warning/audit aid, not cryptographic proof that a PDF is authentic or unmodified. In particular, a detected digital-signature marker is not signature validation.

## Rules Workbook

The final workbook builder reads the first worksheet in `input/Rules.xlsx`.

The supported columns are case-insensitive aliases of:

| Recommended column | Purpose |
| --- | --- |
| `Order` | Sort order for output rule sheets |
| `Category` | `AMT` for an amount rule; any other value is a text rule |
| `subCategory` | Text to find in `Detail_Clean`, or the numeric amount for `AMT` |
| `SheetName` | Target analysis-sheet name |

Text matching is case-insensitive and matches against a compact alphanumeric version of narration. Amount rules match either debit or credit within a small tolerance. Multiple rules targeting the same `SheetName` are merged into one worksheet and duplicate statement rows are removed.

### Current configuration note

The repository currently contains `input/RulesAll.xlsx` and `input/RulesSam.xlsx`, but the active runner looks specifically for `input/Rules.xlsx`. If `Rules.xlsx` is absent, the active path logs a warning and still creates the normal non-rule sheets; it omits rule-derived sheets. The modular path instead treats the missing file as an error.

## Bank Selection and Parsers

The active `PARSERS` dictionary supports these bank codes:

```text
axis, bob, boi, bom, canara, central, cub, dbs, federal, hdfc,
icici, idbi, idfc, indian, indus, iob, kvb, kotak, pnb, sbi,
southind, tmb, unionbank
```

`src/code/bank_detector.py` uses weighted text signatures such as bank names and IFSC prefixes. It deliberately checks extracted document text before OCR because OCR is slower and less reliable.

Each parser returns transaction dictionaries in the common legacy structure. Parsers use the extraction strategy best suited to the PDF layout:

- `pdfplumber` table extraction for conventional tabular statements.
- PyMuPDF (`fitz`) word positions for difficult layouts or wrapped rows.
- Tesseract OCR in selected bank parsers and detection fallbacks.
- Bank-specific heuristics for dates, debit/credit conventions, balances, continuation lines, and cheque references.

Some active legacy parser modules delegate to reusable code under `src/parsers/`:

- `boi`
- `kotak`
- `southind`
- `tmb`
- `unionbank`

The remaining active parsers are implemented directly in `src/code/`.

## Shared Active-Path Behavior

`src/code/utils.py` contains important shared behavior:

- Parses amount text including commas and CR/DR markers.
- Normalizes several date formats.
- Detects statement headers dynamically for generic table parsers.
- Joins continuation narration rows where appropriate.
- Sanitizes cheque values so transaction IDs from UPI/IMPS/NEFT/RTGS are not incorrectly labelled as cheque numbers.
- Extracts cheque numbers from narration when explicit cheque columns are unreliable.
- Writes leading `=` strings as text so Excel does not treat statement content as formulas.
- Reconciles parsed transaction counts/debit/credit totals against summary values discoverable in the PDF and logs mismatches.

`src/code/final_excel_builder.py` is responsible for all final-sheet generation, spreadsheet styling, Indian number formatting, PDF-status checks, and month-wise chart formatting.

## New Modular Architecture (Migration Target)

The intended newer architecture is:

```text
src/parsers/<bank>_parser.py  -> bank-specific extraction
src/transform/                -> normalization and validation
src/export/                   -> workbook output
src/utils/                    -> shared helpers
```

Its normalized target schema is:

```text
Txn_Date, Value_Date, Description, Debit, Credit, Balance,
Currency, Bank, Account_Number, Reference, Source_Page
```

The modular registry currently contains only `axis`, `boi`, `iob`, `kotak`, `southind`, `tmb`, and `unionbank`. It should be treated as a work in progress:

- `src/parsers/axis_parser.py` does not yet extract transactions.
- `src/parsers/detector.py` intentionally raises `NotImplementedError` for automatic detection.
- `src/transform/validate.py` only checks required columns and rejects rows with both debit and credit populated.
- `src/export/final_excel_builder.py` currently creates only `Transactions` and `Metadata` sheets, not the complete analytical workbook.

## Testing

Tests live in `tests/` and are primarily `unittest` tests. They cover parser regressions, cheque normalization, bank detection, rule-sheet merging, monthly totals, return/reject detection, and Excel formatting.

Before changing a parser or final workbook behavior, run the relevant focused tests first, then the full suite if practical:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests
```

Parser changes should be validated using representative PDFs for the affected bank and compared for transaction count, debit total, credit total, balance sequence, and continuation narration handling.

## Safe Change Guidelines

1. Preserve the active output schema unless the final workbook is updated at the same time.
2. Keep bank-specific layout assumptions isolated to that bank's parser.
3. Add or update regression tests whenever a parser changes.
4. Do not silently discard rows merely to make reconciliation pass; log why a row is skipped.
5. Treat PDF extraction as imperfect. A successful run does not prove every row is correct.
6. Keep the `Source` column and final serial numbering correct when processing multiple PDFs.
7. Do not rely on `PDF_Status` as legal-grade tamper detection.
8. Avoid changing the default launcher to `src/main.py` until the modular path reaches feature parity.

## Known Maintenance Risks

- The coexistence of active legacy and newer modular paths can cause fixes to land in the wrong place.
- The default rule-file name does not currently match the rule workbooks visible in `input/`.
- PDF formats can change without notice; extraction uses layout-sensitive heuristics.
- OCR requires the external Tesseract application in addition to the Python dependency.
- Reconciliation is advisory: mismatches are logged, not used to stop generation.
- The active final builder is large and handles both data logic and presentation logic, so changes there need focused tests.

## Key Files

| File | Role |
| --- | --- |
| `run.py` | Root launcher; prefers the project virtual environment then starts the active runner |
| `src/code/run.py` | Active CLI orchestration |
| `src/code/bank_detector.py` | Automatic bank identification |
| `src/code/utils.py` | Shared parsing, normalization, file, Excel, and reconciliation helpers |
| `src/code/*_parser.py` | Active bank parsers and wrappers |
| `src/code/final_excel_builder.py` | Full analytical workbook builder |
| `src/main.py` | New modular pipeline entry point; not the default path |
| `src/parsers/` | New modular parser implementation/migration target |
| `src/transform/` | New normalization and validation layer |
| `src/export/` | New export layer |
| `tests/` | Regression and behavior tests |
| `input/` | PDFs and rule workbook |
| `output/` | Generated Excel workbooks |
| `src/logs/` | Runtime logs created by the active runner |

