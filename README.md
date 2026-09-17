# bankStmtv1

[![Python tests](https://github.com/tmohanct/bankStmtv1/actions/workflows/tests.yml/badge.svg)](https://github.com/tmohanct/bankStmtv1/actions/workflows/tests.yml)

Use [SETUP_WINDOWS.md](SETUP_WINDOWS.md) to install and run this project on another Windows machine.

For a fresh Windows machine, run `.\install_new_machine.bat`.

To create a clean shareable zip package for another machine, run `.\build_fresh_machine_package.bat`.

The canonical pipeline is `src/main.py`, with bank parsers in `src/parsers/`, normalization and validation in `src/transform/`, exports in `src/export/`, and shared helpers in `src/utils/`. `src/code/` retains compatibility imports.

Run all regressions with `python -m unittest discover -s tests -p "test_*.py"`. Empty/invalid parses and mismatched printed totals stop before workbook generation. Final workbooks preserve monetary precision.

Every run is recorded in a single SQLite `run_history` table, including reruns and failures. The default Windows database is `%LOCALAPPDATA%\bankStmtv1\bankstmt.db`; set `BANKSTMT_DB_PATH` to override it. Add `--currency INR` when all input PDFs are in INR to record combined monetary totals. See [run history](docs/RUN_HISTORY.md) for columns and statistics queries.
