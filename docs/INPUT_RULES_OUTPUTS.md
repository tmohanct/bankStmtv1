# Inputs, Rules, and Outputs

## 1. Directory contract

```text
bankStmtv1/
├── input/
│   ├── Rules.xlsx
│   └── statement PDFs
├── output/
│   ├── output.xlsx
│   └── final workbooks
├── src/logs/
│   └── run logs
└── run.py
```

The active runner creates `output/` and `src/logs/` when missing. A PDF may be outside `input/` when a valid relative or absolute path is supplied.

## 2. PDF inputs

### Names and resolution

- `statement` becomes `statement.pdf`.
- `statement.pdf` remains unchanged.
- A value with any suffix is left as supplied.
- Existing absolute and relative paths are accepted.
- Several inputs are separated with semicolons.

For each input, the active resolver checks:

1. The supplied path.
2. `src/input/<name>`.
3. Root `input/<name>`.

Root `input/` is the project convention; `src/input/` is a compatibility location.

### Multi-file behavior

- Files are parsed in argument order.
- With `--bank`, every file uses that parser.
- Without it, detection runs separately for each file.
- Password selection runs per file unless `--pwd` provides one shared override.
- Rows are concatenated, `Source` identifies the PDF, and `Sno` is regenerated.
- The output name defaults to the first PDF's plain stem.
- The account summary at the top of `PDF_Status` comes from the first PDF; the status table checks all PDFs.

## 3. Encrypted PDFs and temporary data

An encrypted PDF must authenticate before parsing. PyMuPDF writes a decrypted copy under:

```text
output/_tmp_run_<milliseconds>/
```

The source PDF is never overwritten. Cleanup runs whether parsing succeeds or fails, and later runs remove empty `_tmp_run_*` directories. An abruptly terminated process may leave sensitive decrypted data; remove it after confirming no parser is active.

## 4. `Rules.xlsx`

The final builder reads only the first sheet of `input/Rules.xlsx`. Use this two-column format for text rules:

| searchString | SheetName |
|---|---|
| RAMASAMY | RAMASAMY |
| KANCHANA | RAMASAMY |

These rules combine transactions matching either name into one `RAMASAMY` sheet. A transaction matching both appears only once. Rules are processed in workbook row order.

The supported fields are:

| Header | Purpose |
|---|---|
| `Order` | Optional numeric rule/sheet processing order. |
| `Category` | Optional rule type. `AMT` means amount matching; every other value means text matching. |
| `searchString` | Search keyword or amount. Legacy `subCategory` is also accepted. |
| `SheetName` | Destination worksheet for matches. |

Header matching is case-insensitive and trims whitespace. Accepted aliases are:

| Logical field | Accepted headers |
|---|---|
| Category | `category` |
| Search value | `searchString`, `subcategory`, `sub_category`, `sub category`, `name`, `keyword`, `search_name`, `searchname`, `match` |
| Destination | `sheetname`, `sheet_name`, `sheet` |
| Order | `sheet_order`, `sheetorder`, `sheet order`, `order` |

Search value and destination columns are required. If absent, no rules load. Category and order are optional: missing category defaults to text behavior; missing order uses workbook row order.

Rows with blank search values or `SheetName` are ignored. Some current workbook rows have blank destinations and therefore do not create sheets.

### Text rules

Any category other than `AMT` uses text matching. `FIN`, `Fin`, `fin`, `TEXT`, blank, and other labels all behave as text rules.

The rule and transaction `Detail_Clean` are converted to text, stripped of non-alphanumeric characters, and uppercased. A transaction matches when its normalized detail contains the normalized rule as a substring.

| Order | Category | searchString | SheetName |
|---:|---|---|---|
| 1 | FIN | Acme Finance | Acme |

This can match `NEFT / ACME-FINANCE / 123`. Matching is substring-based, not whole-word-based, so short rules can produce false positives.

### Amount rules

Only category `AMT` activates amount matching. Commas/spaces are removed before numeric parsing.

| Order | Category | searchString | SheetName |
|---:|---|---:|---|
| 10 | AMT | 25,000 | Amount_25000 |

A row matches when debit or credit differs from the target by no more than `0.005`. Debit matches are ordered before credit matches, then by date and `Sno`. Invalid amounts are skipped and logged.

### Several rules targeting one sheet

Destination names are grouped case-insensitively. Rules for `CustomerA` and `customera` merge into the first encountered name.

Merged sheets:

- Keep source order when numeric `Sno` exists.
- Remove duplicates by `Sno`.
- Otherwise remove duplicate rows.
- Are created only when at least one rule matches.

### Sheet-name safety

Excel-invalid `\ / * ? : [ ]` characters become `_`, names are limited to 31 characters, and blank names become `Sheet`. Collisions receive `_2`, `_3`, and so on while staying within 31 characters.

### Rules maintenance recommendations

- Keep `Order` numeric and unique where practical.
- Use `AMT` only for numeric amount rules.
- Use distinctive text keys.
- Map spelling variants to one `SheetName` when they represent the same entity.
- Do not rename `Rules.xlsx` without changing code.
- Close it in Excel if Windows locking prevents reads.
- Review matches; a rule is an analysis filter, not proof of identity.

## 5. Intermediate workbook

Every successful run writes `output/output.xlsx`.

- It is overwritten on the next successful run.
- It contains one `Statement` sheet.
- It uses the common transaction columns.
- Formula-like narration is stored as text.
- It has no rule sheets, status checks, summaries, or final styling.

Use it for a straightforward merged table, and copy it before the next run if it must be retained.

## 6. Final workbook naming

The final name is based on `--out` or the first input stem:

```text
output/<stem>.xlsx
```

If it exists:

```text
output/<stem>_YYMMDD_HHMMSS.xlsx
```

If the timestamped name also exists, `_1`, `_2`, etc. are added until an unused name is found.

## 7. Final workbook sheets

Planned order:

1. `PDF_Status`
2. `Statement`
3. `Ret/Rej`
4. Matching rule-generated sheets
5. `Cheque_Transactions`
6. `Repeat_Credit_Amount`
7. `Repeat_Debit_Amount`
8. `Top30_Debit`
9. `Top30_Credit`
10. `month_dr_cr`

### `PDF_Status`

The top block summarizes the first source PDF:

- Customer Name
- Bank Name
- Account Number
- Address
- Statement Date Between

The table below contains `PDF`, `Check`, `Status`, `Result`, and `Details` for every PDF, with PASS/WARNING/FAIL colors.

Checks include:

- Original-source SHA-256 fingerprint, access/password authentication and structure repair.
- Parser-reported saved revision count, followed by text and rendered-page comparisons of retained revisions, including subsequently reverted edits. Metadata-only saves are not content edits. Signature widgets and annotations are excluded from the rendered comparison and reported separately.
- Creation/modification dates with timezone handling, creator/producer software and XMP dates, software and edit history.
- Every page's text cover-ups, images drawn over earlier visible text, conflicting monetary amounts printed at substantially overlapping positions, invisible text, annotations and font inventory. Ordinary backgrounds drawn before text, duplicate same-amount painting and font variation alone are not edit evidence.
- Large raster images (one image covering at least half a page), including pages that also contain searchable text. Their contents require review; the audit does not authenticate image pixels or compare OCR against bank records.
- Parsed active objects, editable fields and XFA forms.
- Signature byte-range syntax, bounds, populated excluded signature bytes and bytes appended after a signed revision. This is structural inspection only; signature cryptography and bank certificate trust are **not** verified.

The overall `Result` distinguishes:

| Result | Meaning |
| --- | --- |
| Content changes detected after an earlier save | Retained PDF versions show a text, appearance or page-count change. This does not establish who changed it or whether the change was authorized. |
| Possible manual editing - review evidence | Suspicious features were found; page/object references and available examples are included. These can have legitimate explanations. |
| Inconclusive - inspection has limitations | Checks failed, coverage was incomplete, images/OCR limit assessment, or signatures require verification. |
| No evidence of manual editing found | Available checks completed without indicators. This is not proof of an untouched bank original. |
| Cannot assess | The file could not be opened/authenticated or contained no usable PDF pages. A FAIL here means access failure, not a confirmed edit. |

A failed check must not erase already detected content changes. Revision comparison is bounded to 64 candidate byte boundaries and 2,000 page pairs; reaching either limit is explicitly reported as incomplete. Comparison rendering is capped at 144 dpi and a 1,600-pixel longest side, so very small visual differences may escape it. Text comparison uses full extracted text; evidence examples are shortened for workbook readability.

Image-overlap detection uses drawing order and bounding boxes; transparency or clipping can produce benign overlaps. Full rewrites, flattened edits and image edits can leave no recoverable evidence. No unsigned-PDF heuristic can reliably answer "never edited"; obtain an independently trusted bank original or perform cryptographic signature and issuer validation when authenticity must be established.

### `Statement`

The complete normalized, merged dataset. Repeated rows within a PDF are preserved. Cross-PDF overlap removal requires the same known, unmasked account and at least two adjacent exact transaction matches including running balances. Missing identity or ambiguous single matches are retained. Internal identity fields are not exported.

### `Ret_Rej`

Cheque return/rejection transactions only, for both issued and deposited cheques and either debit or credit entries. Shared detection lives in `src/transform/cheque_returns.py` and recognizes cheque/clearing/CTS context with return, rejection, dishonour, bounce, unpaid, and abbreviated RTN/RETD/RET/REJ descriptions in either word order. Recognized reasons include insufficient funds, exceeds arrangement, stopped payment, signature issues, closed/frozen/blocked accounts, and stale or post-dated cheques.

A usable cheque number supports reason-only or return/rejection-led narrations. Electronic returns (NEFT, RTGS, IMPS, UPI, NACH, ACH, ECS), return fees/charges/taxes, and explicit reversals or negated returns are excluded. These rows remain in `Statement`. No amount threshold is used. `Detail_Clean` is consulted only if `Details` is blank.

Detection is based on narration evidence; unknown bank-specific abbreviations, numeric reason codes without descriptive context, or unreadable OCR may still need additional rules. `Rules.xlsx` does not control this built-in sheet.

### Rule-generated sheets

Filtered matches from `Rules.xlsx`, merged and deduplicated when several rules share a destination.

Text-search sheets also include a date-ordered schedule to the right of the original
table, separated by one blank column. Both tables have a merged sheet-name title
in row 1, headers in row 2, and corresponding transactions beginning in row 3.
Rows 1–2 are frozen. Its fixed columns are `Due No`, `Name`, `Actual Date`, `Day`,
`Freq`, `Cheque No`, `Debit`, `Credit`, `Paid Date`, and `Other Name`.

- `NAME` is the matching search string for that destination. If several strings
  match one transaction, the first rule in processing order supplies the name;
  the transaction appears once.
- Both tables sort together by actual date, then numeric `Sno`. The first due's
  `Freq` counts calendar days from the first preceding credit. Later dues count
  days from the previous debit; intervening credits do not reset this interval.
  Credit rows have blank frequency. Same-date intervals show zero. A first due
  without a preceding credit has blank frequency. Unreadable dates remain
  visible at the end with blank `Day` and `Freq`.
- `DUE NO` counts positive debit rows only; credit rows are blank.
- `CHEQUE NO` shows the cheque identifier (preserving leading zeros) or an
  evidenced transaction mode such as IMPS, UPI, RTGS or NEFT. Cheque transactions
  without an available identifier show `CHEQUE`; unidentified modes remain blank.
- `DR` and `CR` retain the transaction amounts and show totals below the table.
  Zero amounts display blank. `PAID DATE` and `OTHER NAME` are blank for manual entry.
- Fonts, sizes, date/amount widths and numeric display follow the other sheets.
  The title, teal header, alternating fills and vertical group borders follow
  the schedule reference. Left titles/headers use `B7DEE8`, right titles use
  `92CDDC`, and right headers use `31869B` with white text. Both tables share
  alternating `EAF3FB`/`F8F1E5` rows. Schedule credits use red text; source
  transactions and schedule debits use black text. Filtering/sorting covers
  both tables to keep the corresponding rows together.

Amount-only destinations retain their original table. If a destination combines
text and amount rules, amount-only matches retain a blank row in the schedule
so subsequent transactions stay aligned with the left table.
Schedules are regenerated on each run; manual entries from older output files
are not imported.

### `Cheque_Transactions`

Rows with nonblank sanitized cheque numbers, sorted numerically where possible.

### Repeat amount sheets

`Repeat_Credit_Amount` and `Repeat_Debit_Amount` contain positive amounts appearing more than twice in that column. Sort order is amount descending, cheque-present first, cheque number, date, then `Sno`. Repeated groups receive alternating colors.

### Top 30 sheets

`Top30_Debit` and `Top30_Credit` contain the 30 largest positive values, sorted by amount descending then `Sno`.

### `month_dr_cr`

| Column | Calculation |
|---|---|
| `Yr-Month` | `YY-Mon`, plus final `Total`. |
| `Dr` | Monthly debit sum. |
| `Cr` | Monthly credit sum. |
| `Net` | Credit minus debit. |
| `EOM Balance` | Last nonblank monthly balance after date/`Sno` sort. |
| `#.Of.Dr` | Count of debits strictly greater than 30. |
| `#.Of.Cr` | Count of credits strictly greater than 30. |
| `Avg.Dr` | Average debit using only values greater than 30. |
| `Avg.Cr` | Average credit using only values greater than 30. |

The sheet includes a threshold footnote and chart-oriented styling.

## 8. Final formatting

Normal transaction sheets receive Aptos font, styled headers, alternating row fills, frozen headers, filters, type-based alignment, Excel dates using `yyyy-mm-dd`, cheque text formatting, Indian money grouping, and formula-like text protection.

Final styling displays whole units in most sheets while retaining the exact stored monetary values, including paise/cents. Both exports sanitize illegal control characters and store leading `=` text as text.

## 9. Log file

`src/logs/<output-stem>_run_<run-id>.log` includes paths, selected bank, parsing details, counts/totals, reconciliation, rule matches, output paths, warnings, and tracebacks. Each execution also has one row in the SQLite `run_history` table. See [run history](RUN_HISTORY.md) for its location, columns and statistics queries.

Logs may contain sensitive paths and statement metadata; protect them like the PDFs and workbooks.
