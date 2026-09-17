# SQLite run history

Every pipeline invocation inserts one row into the single `run_history` table.
Running the same statement again always creates another row. Multiple PDFs in
one invocation share a row, with one object per input in `input_details_json`.
No transaction rows or PDF contents are stored in the database.

## Location and usage

On Windows the database is created automatically at:

```text
%LOCALAPPDATA%\bankStmtv1\bankstmt.db
```

This keeps the active database outside the OneDrive checkout. On other platforms
the default is `$XDG_DATA_HOME/bankStmtv1/bankstmt.db`, falling back to
`~/.local/share/bankStmtv1/bankstmt.db`. SQLite is included with Python; no server
or additional Python package is required.

Existing commands automatically record history:

```powershell
python run.py --pdf statement.pdf --bank axis
```

To record combined monetary totals, explicitly confirm that **every input has
the same currency**:

```powershell
python run.py --pdf statement.pdf --bank axis --currency INR
python run.py --pdf "jan.pdf;feb.pdf" --currency INR
```

`--currency` accepts INR, USD, EUR, GBP, SGD, AUD and CAD, all using two decimal
places. It labels run statistics; it does not convert amounts or change Excel
output. Without it, `currency`, `total_debit_minor`, and `total_credit_minor` are
NULL. Per-PDF totals are still available in JSON as decimal strings. Do not pass
the flag for a batch with mixed or unknown currencies.

To choose another database location for the current PowerShell session:

```powershell
$env:BANKSTMT_DB_PATH = 'C:\BankStmtData\bankstmt.db'
```

The program creates the parent directory if needed. Keep the live database on a
local, writable disk. The CLI prints the database location, run ID and final
status. Existing Excel naming and collision handling remain unchanged; log files
are now `src/logs/<output-stem>_run_<run-id>.log`.

## Columns

| Columns | Meaning |
|---|---|
| `run_id` | INTEGER primary key; always a new row per invocation. |
| `started_at`, `finished_at` | ISO 8601 UTC timestamps; finish is NULL while running. |
| `duration_ms` | Elapsed monotonic time, saved at completion. |
| `status` | `running`, `success`, `failed`, or `interrupted`. |
| `app_version` | SHA-256 fingerprint of Python source under `src/`, including local edits. |
| `rules_sha256` | Fingerprint of `input/Rules.xlsx`, or NULL when absent. |
| `input_count` | Number of requested PDFs. |
| `processed_count` | PDFs whose parsing, validation and reconciliation completed successfully; an unavailable printed summary is allowed. |
| `failed_count` | PDFs that failed; an export failure can have zero failed PDFs. |
| `skipped_count` | PDFs not attempted when the run stopped. Interrupted inputs have their own JSON status. |
| `transaction_count` | Merged transaction count after deduplication; NULL if merging was not reached. |
| `duplicates_removed` | Number of rows removed by the existing overlap/deduplication logic. |
| `total_debit_minor`, `total_credit_minor` | INTEGER totals after merging, only for an explicitly confirmed currency. INR values are paise: 10049 means Rs. 100.49. |
| `currency` | The optional currency declared by the caller. |
| `ocr_file_count` | PDFs for which OCR was invoked during detection or transaction parsing, even if the OCR attempt failed. Does not include Excel PDF_Status inspection. |
| `reconciliation_status` | `passed` if all passed, `failed` if any failed, `unavailable` if all lacked printed totals, `partial` for mixed/completed and uncompleted checks, or `not_run`. |
| `warning_count` | Number of WARNING-level log records; not the number of affected transactions or PDF_Status findings. |
| `input_details_json` | Per-PDF information described below. |
| `output_path` | Final workbook path; NULL until export returns successfully. |
| `log_path` | This execution's log path, if logging was initialized. |
| `error_stage`, `error_message` | Failure stage and redacted explanation. |

Per-PDF objects contain safe filename, status, start/finish times, source SHA-256,
file size, page count, encryption flag, bank code, parser source fingerprint,
account last four digits and printed statement period when available. After
validation they also include transaction count, debit/credit decimal totals,
first/last transaction dates, negative-balance count, OCR usage and reconciliation
status. Unavailable metadata is omitted or NULL; transaction dates are not
presented as the printed statement period. Optional inspection errors are recorded
as `metadata_error` and do not prevent parsing.

No PDF password is stored. Known explicit and filename-derived passwords are
redacted from database filenames, paths and errors. Existing text logs and shell
history may still contain sensitive source paths; database redaction does not
sanitize those separate records.

## Failure and interruption behavior

The initial `running` row is committed before argument/input validation. File
completion checkpoints update that same row. Missing files, unsupported banks,
invalid arguments, decryption errors, validation failures and export errors are
recorded. A run becomes `success` only after the final workbook export returns.

Ctrl+C records `interrupted` and returns exit code 130. A forced process kill or
power loss can leave `running` with `finished_at = NULL`; this means the run was
not finalized, not that it is necessarily still active. Later runs never rewrite
older rows or mark other live processes interrupted.

`--help` creates no row. Failures before the pipeline imports, or inability to
open/create the database, cannot be recorded. If database initialization fails,
the command prints an error and stops before parsing. If a later database write
fails, the command reports failure; the row may remain `running`. A workbook
already written before a final database failure can still exist.

## Example statistics queries

Run these SQL queries against the database using a SQLite client:

```sql
-- Run counts, including repeated statements.
SELECT status, COUNT(*) AS runs
FROM run_history
GROUP BY status;

-- Daily operational volume in UTC.
SELECT substr(started_at, 1, 10) AS day,
       COUNT(*) AS runs,
       SUM(input_count) AS requested_pdfs,
       SUM(processed_count) AS parsed_pdfs,
       SUM(failed_count) AS failed_pdfs,
       ROUND(AVG(duration_ms) / 1000.0, 2) AS average_seconds
FROM run_history
GROUP BY day
ORDER BY day DESC;

-- Recent failures, including those that generated no workbook.
SELECT run_id, started_at, error_stage, error_message
FROM run_history
WHERE status = 'failed'
ORDER BY run_id DESC;

-- Unique source PDFs among inputs whose hashes were collected.
SELECT COUNT(DISTINCT json_extract(input.value, '$.sha256')) AS unique_pdfs
FROM run_history AS run, json_each(run.input_details_json) AS input;

-- Per-bank processing attempts, including reruns.
SELECT json_extract(input.value, '$.bank') AS bank,
       json_extract(input.value, '$.status') AS status,
       COUNT(*) AS attempts
FROM run_history AS run, json_each(run.input_details_json) AS input
GROUP BY bank, status;
```

JSON queries require a SQLite build with JSON functions. Counts of successful
transactions or monetary totals across runs describe processing volume, not unique
financial activity: reruns and overlapping statements can count the same activity
again. Group monetary comparisons by currency and only use successful runs.

## Code and tests

`src/storage/run_history.py` owns persistence. `src/main.py` supplies lifecycle
events and pipeline statistics. Shared OCR observation lives in
`src/utils/ocr.py`; parsers signal actual OCR invocations without accessing the DB.
Writes use parameters and short transactions with a busy timeout. No connection
or write lock is held while parsing a PDF.

```powershell
python -B -m unittest discover -s tests -p test_run_history.py
```

Tests use temporary database paths and do not add records to your real history.
