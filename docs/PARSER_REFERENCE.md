# Parser and Bank Reference

## 1. Active bank map

The active `PARSERS` dictionary is in `src/code/run.py` and exposes 22 codes.

| Code | Bank | Main extraction approach |
|---|---|---|
| `axis` | Axis Bank | Shared configurable `pdfplumber` table parser with header aliases and fallback positions. |
| `bob` | Bank of Baroda | Custom tables with positioned-word fallback when tables find no rows. |
| `boi` | Bank of India | PyMuPDF positioned-line parser in `src/parsers/boi_parser.py`. |
| `bom` | Bank of Maharashtra | `pdfplumber` tables, serial tracking, and summary-count checks. |
| `canara` | Canara Bank | PyMuPDF line parser with running-balance classification and continuity warnings. |
| `central` | Central Bank of India | Custom `pdfplumber` table parser. |
| `cub` | City Union Bank | Custom `pdfplumber` tables with CUB date normalization. |
| `dbs` | DBS Bank | Shared configurable `pdfplumber` table parser. |
| `federal` | Federal Bank | Custom tables supporting multiple known layouts and CR/DR balance indicators. |
| `hdfc` | HDFC Bank | `pdfplumber` text lines using opening/running balances and summary-count checks. |
| `icici` | ICICI Bank | Multiple native-text layouts, then rendered-page Tesseract OCR fallback. |
| `idbi` | IDBI Bank | `pdfplumber` page text with serial continuity reporting. |
| `idfc` | IDFC FIRST Bank | Custom `pdfplumber` table parser. |
| `indian` | Indian Bank | PyMuPDF lines, known date orders, wrapped details, footer filtering, and balance classification. |
| `indus` | IndusInd Bank | `pdfplumber` lines with a layout-specific row expression and continuations. |
| `iob` | Indian Overseas Bank | `pdfplumber` tables supporting known IOB date/code layouts. |
| `kotak` | Kotak Mahindra Bank | `pdfplumber` table parser in `src/parsers/kotak_parser.py`. |
| `kvb` | Karur Vysya Bank | Detects tokenized text, regular native text, or Tesseract OCR layout. |
| `pnb` | Punjab National Bank | Custom `pdfplumber` table parser. |
| `sbi` | State Bank of India | Custom tables with month-name date handling. |
| `southind` | South Indian Bank | PyMuPDF positioned words/lines with multiple balance-continuity paths. |
| `tmb` | Tamilnad Mercantile Bank | `pdfplumber` tables with layout selection in `src/parsers/tmb_parser.py`. |
| `unionbank` | Union Bank of India | Custom `pdfplumber` tables in `src/parsers/unionbank_parser.py`. |

This describes known implementations, not every PDF a bank may issue.

## 2. Automatic bank detection

When `--bank` is omitted, `src/code/bank_detector.py` performs three stages.

### Stage 1: native text

It concatenates the first two pages using both `pdfplumber` and PyMuPDF, uppercases text, collapses whitespace, and scores weighted signatures. The highest `(score, bank_code)` wins; a score tie is therefore broken by the lexically larger code.

### Stage 2: OCR

If native text has no match and Tesseract is available, the first page is rendered at 1.5x and OCR text is scored.

Tesseract discovery checks:

1. `TESSERACT_CMD`.
2. `tesseract` on `PATH`.
3. `C:\Program Files\Tesseract-OCR\tesseract.exe`.
4. `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`.

### Stage 3: filename

If content and OCR fail, the filename stem is scored. A filename match is logged as a warning. If all stages fail, use `--bank <code>`.

### Signature examples

| Code | Examples |
|---|---|
| `axis` | `AXIS BANK`, `UTIB0` |
| `bob` | `BANK OF BARODA`, `BARB0` |
| `boi` | `BANK OF INDIA`, `BKID0` |
| `bom` | `BANK OF MAHARASHTRA`, `MAHB0` |
| `canara` | `CNRB0`, `STATEMENT FOR A/C` |
| `central` | `CENTRAL BANK OF INDIA`, `CBIN0` |
| `cub` | `CITY UNION BANK`, `CIUB0` |
| `dbs` | `DBS BANK`, `DBSCPIN` |
| `federal` | `FEDERAL BANK`, `FDRL0` |
| `hdfc` | `HDFC BANK`, `HDFC0` |
| `icici` | `ICICI BANK`, `ICIC0` |
| `idbi` | `IDBI BANK`, `IBKL0` |
| `idfc` | `IDFC FIRST BANK`, `IDFB0` |
| `indian` | `IDIB0`, `ACCOUNT STATEMENT`, `ACCOUNT ACTIVITY` |
| `indus` | `INDUSIND BANK`, `INDB0` |
| `iob` | `INDIAN OVERSEAS BANK`, `IOBA0` |
| `kvb` | `KARUR VYSYA BANK`, `KVBL0` |
| `kotak` | `KOTAK MAHINDRA BANK`, `KKBK` |
| `pnb` | `PUNJAB NATIONAL BANK`, `PUNB0` |
| `sbi` | `STATE BANK OF INDIA`, `SBIN0` |
| `southind` | `SOUTH INDIAN BANK`, `SIBL` |
| `unionbank` | `UNION BANK`, `UBIN` |

`tmb` has an active parser but no `BANK_SIGNATURES` entry. Automatic detection cannot select it; use `--bank tmb`.

## 3. Generic table parser

Axis and DBS use the shared configurable parser in `src/code/utils.py`:

1. Extract all tables with `pdfplumber`.
2. Normalize whitespace.
3. Detect a header by aliases, normally requiring three canonical matches.
4. Reuse the last header mapping on headerless pages.
5. Use configured fallback positions for missing mappings.
6. Treat recognized-date rows as new transactions.
7. Merge narration-only rows into the preceding transaction.
8. Skip opening/closing/transaction-total summaries.
9. Prefer dedicated debit/credit columns.
10. Otherwise use an amount with DR/CR, or its sign.
11. Normalize date, details, cheque number, amounts, and balance.

## 4. OCR requirements

`pytesseract` is only a Python bridge; the Tesseract desktop executable must be installed.

OCR affects:

- ICICI when native layouts return no rows.
- KVB when text layouts are not detected.
- Auto-detection when native PDF text is unusable.

OCR is slower and less reliable. Common errors include date punctuation, decimal placement, digit/letter substitution, and shifted columns. Review OCR-based outputs.

## 5. Active parser contract

The runner calls:

```python
records = parser_fn(str(readable_pdf_path), logger, progress_cb=callback)
```

An active parser must:

- Accept `pdf_path`, `logger`, and optional `progress_cb`.
- Return a list of transaction dictionaries.
- Avoid writing final outputs itself.
- Call `progress_cb(current_count)` as rows are added where practical.
- Produce common keys or let shared normalization add missing keys.
- Keep layout logic in its bank module.

Recommended record:

```python
{
    "Sno": 1,
    "Date": "31/01/2026",
    "Details": "NEFT PAYMENT",
    "Detail_Clean": "NEFTPAYMENT",
    "Cheque No": "",
    "Debit": 1000.0,
    "Credit": None,
    "Balance": 9000.0,
}
```

The runner sets `Source` and regenerates `Sno` across merged inputs.

## 6. Add a new bank

1. Create `src/parsers/<code>_parser.py` per repository instructions.
2. If necessary, expose a thin `src/code/<code>_parser.py` adapter with `parse()`.
3. Reuse shared helpers where appropriate.
4. Import it in `src/code/run.py` and add it to `PARSERS`.
5. Add weighted `BANK_SIGNATURES` if auto-detection should support it.
6. Add focused synthetic layout tests and representative PDF regression tests.
7. Test CLI help, explicit parsing, detection, output columns, balance continuity, multi-page rows, and final workbook creation.
8. Update documentation bank lists.

## 7. Add a layout to an existing bank

- Identify stable layout headers/signatures.
- Separate layout detection from row parsing.
- Normalize every layout to the same record contract.
- Handle running balance, continuations, headers, footers, and summaries explicitly.
- Log the selected path.
- Add tests for new and old layouts.

## 8. Modular registry versus active map

The scaffold registry in `src/parsers/parser_registry.py` contains only `axis`, `boi`, `iob`, `kotak`, `southind`, `tmb`, and `unionbank`. It belongs to `src/main.py`.

The user-facing CLI uses the 22-entry `PARSERS` map in `src/code/run.py`. Use that map when determining command support.

