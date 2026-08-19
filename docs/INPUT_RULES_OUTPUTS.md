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

The final builder reads only the first sheet of `input/Rules.xlsx`. The current workbook has one table, 283 rows including its header, and these 4 columns:

| Current header | Purpose |
|---|---|
| `Order` | Numeric rule/sheet processing order. |
| `Category` | Rule type. `AMT` means amount matching; every other value means text matching. |
| `subCategory` | Search keyword or amount. |
| `SheetName` | Destination worksheet for matches. |

Header matching is case-insensitive and trims whitespace. Accepted aliases are:

| Logical field | Accepted headers |
|---|---|
| Category | `category` |
| Search value | `subcategory`, `sub_category`, `sub category`, `name`, `keyword`, `search_name`, `searchname`, `match` |
| Destination | `sheetname`, `sheet_name`, `sheet` |
| Order | `sheet_order`, `sheetorder`, `sheet order`, `order` |

Search value and destination columns are required. If absent, no rules load. Category and order are optional: missing category defaults to text behavior; missing order uses workbook row order.

Rows with blank search values or `SheetName` are ignored. Some current workbook rows have blank destinations and therefore do not create sheets.

### Text rules

Any category other than `AMT` uses text matching. `FIN`, `Fin`, `fin`, `TEXT`, blank, and other labels all behave as text rules.

The rule and transaction `Detail_Clean` are converted to text, stripped of non-alphanumeric characters, and uppercased. A transaction matches when its normalized detail contains the normalized rule as a substring.

| Order | Category | subCategory | SheetName |
|---:|---|---|---|
| 1 | FIN | Acme Finance | Acme |

This can match `NEFT / ACME-FINANCE / 123`. Matching is substring-based, not whole-word-based, so short rules can produce false positives.

### Amount rules

Only category `AMT` activates amount matching. Commas/spaces are removed before numeric parsing.

| Order | Category | subCategory | SheetName |
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

The helper does not add a counter if an identical timestamped path already exists. Avoid two identical output builds in the same second.

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

- Overall modification status
- File access/password authentication
- Page count and encryption
- Parser-reported structure repair
- Incremental update markers (`/Prev`, `%%EOF`)
- Cross-reference/save revision indicators
- Creation/modification dates
- XMP metadata and history fields
- Creator/producer metadata
- Annotations and JavaScript
- Active/object markers: launch actions, embedded files, rich media, forms, XFA, and open actions
- Combined abnormal detection
- Digital-signature marker presence

These are heuristics. Signature markers are not cryptographically validated, and no tool can prove an unsigned PDF was never changed.

### `Statement`

The complete normalized, merged dataset.

### `Ret/Rej`

Rows whose normalized details contain known cheque return/rejection, dishonour, return-charge, or NEFT/RTGS/IMPS return patterns.

### Rule-generated sheets

Filtered matches from `Rules.xlsx`, merged and deduplicated when several rules share a destination.

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

The final styling step rounds numeric debit, credit, and balance cells to whole values in most sheets. If paise/cents matter, use the intermediate workbook or change the final rounding behavior.

## 9. Log file

`src/logs/<output-stem>.log` includes paths, selected bank, parsing details, counts/totals, reconciliation, rule matches, output paths, warnings, and tracebacks. It appends for repeated output stems.

Logs may contain sensitive paths and statement metadata; protect them like the PDFs and workbooks.

