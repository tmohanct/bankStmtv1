# Bank Statement Parser — AI Project Context

## Purpose

This repository converts Indian bank-statement PDFs into Excel workbooks for review and analysis. It extracts transactions, normalizes the common fields, applies user-maintained matching rules, and produces analysis sheets such as cheque transactions, repeated amounts, return/reject entries, top debits/credits, and monthly totals.

The project is designed to support many banks. PDFs can be text-based or, for selected parsers and bank detection, OCR-assisted.

## Start Here: What Actually Runs

All launchers use `src/main.py`. The root launcher and batch files prefer a healthy local virtual environment. `src/code/` contains compatibility imports, not a second implementation.

- Bank extraction and detection signatures: `src/parsers/<bank>_parser.py`.
- One discovered registry for all 24 banks: `src/parsers/parser_registry.py`.
- Normalization, overlap removal, validation: `src/transform/`.
- Intermediate/final Excel output and text safety: `src/export/`.
- Shared parsing/file helpers: `src/utils/`.

## Runtime Flow

The active entry point is `src/main.py`.

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
- `Ret_Rej`: cheque return/rejection transactions detected by `src/transform/cheque_returns.py`; excludes electronic returns, fees/charges, and explicit reversals. Supports cheque numbers and blank-narration fallback to `Detail_Clean`.
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
| `Order` | Optional sort order for output rule sheets; defaults to workbook row order |
| `Category` | Optional: `AMT` for an amount rule; defaults to text matching |
| `searchString` | Text to find in `Detail_Clean`, or the numeric amount for `AMT`; legacy `subCategory` is also accepted |
| `SheetName` | Target analysis-sheet name |

For text rules, only `searchString` and `SheetName` are required. For example, `RAMASAMY` and `KANCHANA` can both target the `RAMASAMY` sheet.

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

`src/parsers/detector.py` uses weighted text signatures such as bank names and IFSC prefixes. It deliberately checks extracted document text before OCR because OCR is slower and less reliable.

Each parser returns transaction dictionaries in the common legacy structure. Parsers use the extraction strategy best suited to the PDF layout:

- `pdfplumber` table extraction for conventional tabular statements.
- PyMuPDF (`fitz`) word positions for difficult layouts or wrapped rows.
- Tesseract OCR in selected bank parsers and detection fallbacks.
- Bank-specific heuristics for dates, debit/credit conventions, balances, continuation lines, and cheque references.

All bank parsers now live under `src/parsers/`. Old `src/code/` imports delegate to these modules.

## Shared Active-Path Behavior

Shared parsing helpers, normalization, validation and export modules provide:

- Parses amount text including commas and CR/DR markers.
- Normalizes several date formats.
- Detects statement headers dynamically for generic table parsers.
- Joins continuation narration rows where appropriate.
- Sanitizes cheque values so transaction IDs from UPI/IMPS/NEFT/RTGS are not incorrectly labelled as cheque numbers.
- Extracts cheque numbers from narration when explicit cheque columns are unreliable.
- Writes leading `=` strings as text so Excel does not treat statement content as formulas.
- Reconciles parsed transaction counts/debit/credit totals against summary values discoverable in the PDF and returns explicit passed/failed/unavailable results.

`src/export/final_excel_builder.py` assembles and styles final sheets, including Indian number formatting and PDF-status presentation. Rule matching and transaction summaries live in `src/transform/analysis.py`; monthly chart image rendering lives in `src/export/monthly_chart.py`. Tesseract discovery is shared through `src/utils/ocr.py`.

## Data contract and validation

The active record schema remains `Sno, Date, Details, Detail_Clean, Cheque No, Debit, Credit, Balance, Source`.
Internal account/source metadata is removed before export. Repeated rows within one PDF are preserved. Cross-PDF overlap removal requires a known, unmasked account identity and at least two consecutive exact matches including running balances. Ambiguous matches are retained.

Empty parses, invalid dates/amounts and mismatched printed totals stop the run before workbook output with exit code 1. Missing printed totals are reported as reconciliation unavailable, not a pass. Monetary cell values retain decimal precision; number formats may display whole units.

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
8. Keep old import compatibility wrappers thin; make behavior changes in the canonical modules.

## Known Maintenance Risks

- Compatibility wrappers under `src/code/` must continue delegating to the canonical implementation.
- The default rule-file name does not currently match the rule workbooks visible in `input/`.
- PDF formats can change without notice; extraction uses layout-sensitive heuristics.
- OCR requires the external Tesseract application in addition to the Python dependency.
- Reconciliation can only check summary values present in extractable PDF text; missing totals are reported explicitly.
- The active final builder is large and handles both data logic and presentation logic, so changes there need focused tests.

## Key Files

| File | Role |
| --- | --- |
| `run.py` | Root launcher; prefers the project virtual environment then starts the active runner |
| `src/main.py` | Active CLI orchestration |
| `src/parsers/detector.py` | Automatic bank identification |
| `src/utils/statement_utils.py` | Shared extraction, amount/date and input-file helpers |
| `src/transform/normalize.py` | Cheque normalization and conservative overlap removal |
| `src/transform/validate.py` | Record validation and reconciliation |
| `src/export/excel_writer.py` | Intermediate export |
| `src/parsers/*_parser.py` | Active bank parsers and wrappers |
| `src/transform/analysis.py` | Rule matching and analytical tables |
| `src/export/monthly_chart.py` | Monthly chart rendering |
| `src/export/final_excel_builder.py` | Workbook assembly and styling |
| `src/utils/ocr.py` | Shared Tesseract discovery |
| `src/code/` | Compatibility wrappers for old imports |
| `src/parsers/` | Canonical bank parsers and detection signatures |
| `src/transform/` | Normalization and validation |
| `src/export/` | Workbook exports |
| `tests/` | Regression and behavior tests |
| `input/` | PDFs and rule workbook |
| `output/` | Generated Excel workbooks |
| `src/logs/` | Runtime logs created by the active runner |
