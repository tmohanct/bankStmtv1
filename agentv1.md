# bankStmtv1 Agent Guide

This file is a project-specific operating manual for Codex or any other coding agent working in this repository.

Use this file as the primary project brief before making changes.
The goal is that an agent reading only this document can understand:

- what this project does
- which code path is actually used today
- where inputs and outputs belong
- how bank parsers are expected to behave
- how to safely test and extend the system
- what not to break while making changes

## 1. Project Mission

This repository is a Windows-friendly Python bank statement analysis tool.

Its job is to:

1. read one or more bank statement PDFs
2. detect or accept the bank type
3. parse statement transactions into a normalized tabular structure
4. write an intermediate workbook
5. generate a final Excel workbook with summary and analysis sheets

The project must remain bank-extensible.
Bank-specific parsing logic should stay isolated so new banks can be added without rewriting the full pipeline.

## 2. What The User Usually Wants

The user is usually asking for one of these things:

- test a specific PDF and see whether it parses correctly
- fix a parser for a new statement layout of an existing bank
- add support for a new bank
- improve workbook output sheets
- make the project easier to run on another Windows machine
- preserve old working formats while adding new ones

When the user says "test this PDF", do not stop at analysis.
Actually run the parser, inspect the result, and if parsing fails or is incomplete, patch the relevant parser and verify it.

## 3. Current Architecture: Important Reality

This repo has two architectures at the same time.

### 3.1 Active production-style path used today

The currently active and feature-complete path is the legacy-but-working flow under:

- `src/code/run.py`
- `src/code/*_parser.py`
- `src/code/utils.py`
- `src/code/final_excel_builder.py`
- `src/code/bank_detector.py`

This is the path that currently supports the broad multi-bank workflow, auto-detection, reconciliation, progress printing, and the richer final workbook.

If the user wants a real PDF tested today, this is usually the correct path to run first.

### 3.2 New modular scaffold

There is also a newer modular structure under:

- `src/main.py`
- `src/parsers/`
- `src/transform/`
- `src/export/`
- `src/utils/`

This is the target architecture for long-term cleanup and extension, but it is not yet the full runtime replacement.

Important current limitations of the modular path:

- bank coverage is partial
- auto-detection is not implemented there yet
- output workbook generation is simpler than the active legacy builder

### 3.3 Practical instruction for agents

Do not assume the newer modular scaffold is the full live pipeline.

If the task is:

- testing real PDFs
- fixing a currently supported bank parser
- preserving existing workbook behavior
- keeping the current CLI working

then inspect `src/code/` first.

If the task is:

- adding cleaner architecture
- creating a new modular parser
- improving transform/export separation

then inspect `src/` and decide whether the change also needs to be mirrored in `src/code/` so the user-facing workflow still works.

Do not break the current `src/code` flow while trying to improve the modular structure.

## 4. Repository Layout

### 4.1 Top-level folders

- `input/`
  Incoming PDFs and `Rules.xlsx`

- `output/`
  Intermediate and final Excel outputs

- `src/code/`
  Current active parser runner, legacy bank parsers, export builder, detection logic, parsing helpers

- `src/parsers/`
  New modular parser classes

- `src/transform/`
  New modular normalization and validation logic

- `src/export/`
  New modular workbook output code

- `src/utils/`
  New modular file/date/amount helpers

- `tests/`
  Regression tests and unit tests

- `dist/`
  Packaged distribution zips for fresh machines

### 4.2 Important root files

- `AGENTS.md`
  Short project guardrails already present in the repo

- `README.md`
  Minimal setup reference

- `SETUP_WINDOWS.md`
  Windows setup and run instructions

- `run_bank_parser.bat`
  Windows launcher for the active parser flow

- `install_new_machine.bat`
  Rebuilds environment and installs dependencies on a new Windows system

- `requirements.txt`
  Python dependencies

## 5. Input And Output Contract

These paths are important and should not be casually changed.

### 5.1 Inputs

- source PDFs belong in `input/`
- parsing rules workbook belongs in `input/Rules.xlsx`

### 5.2 Outputs

- intermediate workbook must be written to `output/output.xlsx`
- final workbook must be written to `output/<pdf_base_name>.xlsx`
- if that final workbook already exists, append timestamp format `YYMMDD_HHMMSS`

Examples:

- `output/VAIRAKANNU.xlsx`
- `output/VAIRAKANNU_260408_191318.xlsx`

### 5.3 Log files

The active runner writes logs to:

- `src/logs/<output_name>.log`

When debugging parsing issues, always inspect the log file if the run fails or row counts look suspicious.

## 6. Supported Banks

### 6.1 Banks supported by the active legacy runner in `src/code/run.py`

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

### 6.2 Banks currently registered in the newer modular parser registry

As of now, `src/parsers/parser_registry.py` registers:

- `axis`
- `iob`
- `kotak`
- `southind`
- `unionbank`

That list is smaller than the active legacy runner.
Do not confuse the two.

## 7. Active Runtime Flow

This section describes what happens in the main user-facing run path.

### 7.1 Main entry point

Primary current runner:

- `python .\src\code\run.py --file "sample.pdf"`

Common Windows wrapper:

- `.\run_bank_parser.bat --file "sample.pdf"`

### 7.2 CLI behavior

The active runner accepts:

- `--pdf` or `--file`
- `--bank`
- `--pwd`
- `--out`

Multiple PDFs can be passed as:

- `"file1.pdf;file2.pdf"`

### 7.3 Password behavior

Encrypted PDFs may use either:

- explicit `--pwd`
- or a password embedded in the file name using `$`

Example:

- `BALASUBRAMANIAN$35795624.pdf`

### 7.4 Processing steps

The active runner does the following:

1. parse CLI args
2. resolve PDF path from `input/` or an explicit file path
3. infer password from `--pwd` or file name suffix
4. create temporary decrypted copy if needed
5. auto-detect bank if `--bank` was omitted
6. call the bank-specific parser
7. merge records from all PDFs
8. assign final `Sno` values across the merged result
9. write `output/output.xlsx`
10. build the final workbook in `output/`
11. write a log file in `src/logs/`

### 7.5 Auto-detection

The active bank detector is:

- `src/code/bank_detector.py`

Detection uses:

- extracted text via `pdfplumber`
- extracted text via PyMuPDF (`fitz`)
- OCR fallback via Tesseract when needed
- filename heuristics as a last fallback

If auto-detection fails, the user can rerun with:

- `--bank <code>`

## 8. Output Row Schema In The Active Legacy Flow

The active legacy parser stack writes rows with these columns:

- `Sno`
- `Date`
- `Details`
- `Cheque No`
- `Debit`
- `Credit`
- `Balance`
- `Detail_Clean`
- `Source`

Important expectations:

- `Sno` is assigned after parsing and merge
- `Date` should be normalized to `YYYY-MM-DD` whenever possible
- `Details` is the human-readable transaction narration
- `Detail_Clean` is a normalized tokenized version used for matching/grouping
- `Cheque No` should be blank unless the parser is reasonably confident it is a cheque reference
- exactly one of `Debit` or `Credit` should normally be populated
- `Balance` may legitimately be negative, especially for OD/CC accounts
- `Source` is the source PDF file name

## 9. Final Workbook Expectations In The Active Flow

The active `src/code/final_excel_builder.py` creates a richer workbook than the modular scaffold.

Planned sheets include:

- `Statement`
- `Ret_Rej`
- `Cheque_Trans`
- `Repeat_Dt`
- `Repeat_Cr`
- `Top30_Dt`
- `Top30_Cr`
- `month_dr_cr`
- rule-driven sheets based on `search.xlsx`

### 9.1 Rule-driven sheets

`search.xlsx` is used to build categorized sheets. Search the string from search.xlsx column searchVal and lookup in Detail_clean column. MAtching transaction put in SheetName.

Based on tests and current implementation, the rules workflow expects columns such as:

- `searchVal`
- `SheetName`
- `Order` may also appear

Important behavior:

- multiple rules can target the same sheet
- duplicate sheet names should be merged rather than written as separate suffix sheets
- matching is driven by cleaned detail strings

### 9.2 Return and reject logic

The workbook includes a return/reject sheet for cheque return, reject, return charges, and related electronic return patterns.

When changing return/reject logic, verify:

- row inclusion is correct
- the sheet remains near the front of the workbook
- existing tests still pass

### 9.3 Month debit/credit sheet

The workbook includes a `month_dr_cr` analysis sheet with:

- monthly debit/credit totals
- net movement
- end-of-month balance
- debit and credit counts
- average debit and credit
- custom styling and chart/image presentation

If changing final workbook generation, do not accidentally remove this sheet or reduce its functionality unless the user explicitly asks.

## 10. Modular Target Architecture

The newer modular architecture intends to separate responsibilities clearly:

- bank-specific extraction in `src/parsers/<bank>_parser.py`
- normalization in `src/transform/`
- validation in `src/transform/validate.py`
- export in `src/export/`
- generic helpers in `src/utils/`

### 10.1 Modular parser contract

The base parser interface is in:

- `src/parsers/base_parser.py`

Parser class contract:

- each parser subclasses `BaseStatementParser`
- each parser defines `bank_code`
- each parser implements `parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame`

### 10.2 Modular normalized output target

The new transform layer normalizes to:

- `Txn_Date`
- `Value_Date`
- `Description`
- `Debit`
- `Credit`
- `Balance`
- `Currency`
- `Bank`
- `Account_Number`
- `Reference`
- `Source_Page`

### 10.3 Modular validation rule

The current modular validator enforces that a row cannot have both:

- `Debit`
- `Credit`

populated at the same time.

## 11. Parser Development Rules

These rules are important whenever adding or fixing a bank parser.

### 11.1 Isolation

Keep bank-specific parsing logic inside the bank parser file.

For modular work:

- `src/parsers/<bank>_parser.py`

For the active user-facing runner today:

- `src/code/<bank>_parser.py`

### 11.2 Preferred extraction order

Prefer this order unless the statement format forces something else:

1. `pdfplumber` text/table extraction
2. PyMuPDF (`fitz`) text extraction
3. OCR only when necessary

OCR should be the fallback, not the first approach, unless the bank is known to require it.

### 11.3 Multi-line transaction handling

Many statements split details across lines or rows.
Parsers should:

- detect a real transaction start robustly
- merge continuation lines into the previous transaction detail
- ignore page headers, brought-forward rows, totals, and end-of-report lines

### 11.4 Date handling

Normalize dates into `DD/MM/YYYY`.

Be prepared to support:

- `dd/mm/yyyy`
- `dd-mm-yyyy`
- `yyyy-mm-dd`
- `dd-MMM-yyyy`
- `dd-MMMM-yyyy`

If a new bank uses a different date style, add support carefully and verify it does not regress existing banks.

### 11.5 Amount handling

Amount parsers should handle:

- commas
- `CR` and `DR`
- negative signs
- brackets for negative values
- leading decimal values when OCR is noisy

### 11.6 Balance handling

Negative balances are not automatically a bug.
Overdraft and cash credit accounts may be negative for the whole statement.

Do not "fix" negative balances unless they are clearly parsed incorrectly.

### 11.7 Cheque numbers

Only populate `Cheque No` when the parser is reasonably confident the token is actually a cheque reference.
Do not blindly treat every numeric token as a cheque number.

### 11.8 Reconciliation

If the statement exposes totals like:

- total debit
- total credit
- total transactions

then compare parsed totals against them.
This project already has reconciliation helpers in the legacy path.

### 11.9 Preserve source details

Do not over-clean or over-normalize the human-readable transaction description.
The original detail text is useful for rules, review, and auditing.

## 12. Testing Expectations

Any meaningful parser or workbook change should be verified.

### 12.1 Preferred test style in this repo

Use `unittest`.
Do not assume `pytest` is installed on the machine.

A safe command pattern is:

```powershell
python -m unittest tests.test_cub_parser tests.test_bom_parser tests.test_kvb_parser
```

### 12.2 Real PDF regression tests

When possible, add or update regression tests that use real sample PDFs already stored in `input/`.

Good examples:

- `tests/test_cub_parser.py`
- `tests/test_bom_parser.py`
- `tests/test_unionbank_parser.py`
- `tests/test_iob_parser.py`

### 12.3 Unit-level tests

If a full-PDF test is too heavy or too fragile, add focused unit tests around:

- date normalization
- row start detection
- multiline merge behavior
- cheque extraction
- summary metric extraction
- workbook sheet generation

### 12.4 Minimum validation after parser changes

After changing a parser, try to verify:

1. row count
2. first row correctness
3. last row correctness
4. debit total
5. credit total
6. balance parsing sanity
7. output workbook exists

### 12.5 When the user gives a new PDF

The expected workflow is:

1. confirm the PDF is in `input/`
2. run the active parser path
3. inspect row count and output workbook
4. if rows are zero or incomplete, inspect extracted text or tables
5. patch the correct parser
6. add a regression test if the format is important
7. rerun the parser
8. rerun relevant tests

## 13. Recommended Commands

### 13.1 Run the active parser flow

```powershell
python .\src\code\run.py --file "sample.pdf"
```

Explicit bank:

```powershell
python .\src\code\run.py --bank cub --file "VAIRAKANNU.pdf"
```

Multiple files:

```powershell
python .\src\code\run.py --bank axis --file "axis.pdf;axis2.pdf"
```

### 13.2 Run through the Windows wrapper

```powershell
.\run_bank_parser.bat --bank sbi --file sbi.pdf
```

### 13.3 Run the modular scaffold

Use only when working specifically on the new modular path:

```powershell
python .\src\main.py --bank iob --pdf AKILANMANIVANNAN.pdf
```

### 13.4 Run tests

```powershell
python -m unittest
```

or a focused subset:

```powershell
python -m unittest tests.test_cub_parser
```

## 14. Known Project Quirks

### 14.1 The copied virtual environment may be stale

If this repo was copied from another Windows machine, `.venv` may point to a Python path that no longer exists.

Use:

- `.\install_new_machine.bat`

or:

- `.\setup_windows.bat`

to rebuild it.

### 14.2 The active runner import order matters

`src/code/run.py` intentionally manipulates `sys.path` so its legacy `utils.py` and parser modules resolve correctly.

Be careful when changing imports there.
It is easy to accidentally import `src/utils/` instead of `src/code/utils.py`.

### 14.3 The modular path is not yet the complete replacement

Do not delete or bypass legacy code just because modular equivalents exist.

### 14.4 Some tests depend on sample PDFs

Several regression tests are guarded by `skipUnless(sample_pdf.exists())`.
If a sample file is missing, that test may skip.

### 14.5 OCR is only required for some workflows

Tesseract is mainly relevant for OCR-based parsing and OCR-based bank detection fallback.
Do not require OCR for all banks if simple text extraction is enough.

## 15. Decision Rules For Codex

When making decisions in this repo, follow these rules.

### 15.1 If the user asks to fix a parser for a real bank statement

- start with the active `src/code` parser path
- run the actual PDF
- patch the specific bank parser
- add a regression test
- verify the workbook output

### 15.2 If the user asks to add a new bank for future architecture

- create the parser cleanly in `src/parsers/`
- register it in `src/parsers/parser_registry.py`
- add transform/export compatibility if needed
- decide whether the user also needs it wired into `src/code/run.py` immediately

### 15.3 If the user asks to improve workbook sheets

- inspect `src/code/final_excel_builder.py`
- inspect related tests in `tests/`
- preserve existing sheet order and major sheet names unless told otherwise

### 15.4 If the user asks to clean up architecture

- make incremental changes
- keep current outputs stable
- do not break the active CLI path

### 15.5 If you are unsure which path to modify

Ask yourself:

- is this for real current user-facing PDF parsing now
- or for long-term modular cleanup

If it affects current user runs, the answer usually includes `src/code/`.

## 16. Things Not To Break

- `input/Rules.xlsx` location
- `output/output.xlsx` intermediate output
- final workbook naming convention
- bank auto-detection in the active runner
- multiline detail merging behavior
- reconciliation checks
- rule-driven workbook sheets
- cheque number cleanup rules
- existing sample PDF regression coverage

## 17. Success Criteria

A good change in this repository usually satisfies all of the following:

1. the user can place a PDF in `input/`
2. the parser can process it through the expected runner
3. parsed rows are structurally correct
4. debit and credit values make sense
5. output workbooks are written to the correct paths
6. existing behavior for other banks is not broken
7. relevant tests pass

## 18. Recommended Working Style For An Agent

When working on this repo:

1. inspect the active path first
2. run the real input when available
3. avoid broad refactors unless requested
4. make the smallest change that fixes the real problem
5. add or update a regression test
6. report exactly what was verified

If you have to choose between a beautiful rewrite and a verified fix that matches the user's workflow, choose the verified fix first.

## 19. Short Summary

This is a bank statement PDF-to-Excel project with:

- a current working multi-bank runtime under `src/code/`
- a partial future modular architecture under `src/`
- fixed input/output locations
- bank-specific parsing responsibilities
- a strong need for regression testing with real PDFs

For real user work today, assume `src/code/run.py` is the operational path unless the user explicitly asks for modular-architecture work.
