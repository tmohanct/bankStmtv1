# Bank Statement Analyzer Documentation

This is the documentation index for `bankStmtv1`, a Python tool that converts supported Indian bank-statement PDFs into normalized Excel transactions and a multi-sheet analysis workbook.

## Start here

- [Complete project guide](docs/PROJECT_GUIDE.md) — purpose, end-to-end execution flow, data model, architecture, logging, errors, and reliability boundaries.
- [Command-line reference](docs/CLI_REFERENCE.md) — every active CLI option, command examples, precedence rules, exit codes, setup options, and packaging options.
- [Inputs, rules, and outputs](docs/INPUT_RULES_OUTPUTS.md) — PDF resolution, encrypted inputs, the actual `Rules.xlsx` schema, matching behavior, output files, and every final workbook sheet.
- [Parser and bank reference](docs/PARSER_REFERENCE.md) — all 22 active bank codes, parser techniques, auto-detection, OCR, parser contracts, and extension instructions.
- [Development and troubleshooting](docs/DEVELOPMENT.md) — setup, dependencies, module ownership, tests, packaging, diagnostics, security, and known maintenance notes.
- [Windows setup](SETUP_WINDOWS.md) — shorter installation guide for a new Windows machine.

## Quick command

```powershell
.\install_new_machine.bat
.\run_bank_parser.bat --file statement.pdf --bank axis
```

Use `python run.py --help` to display the live option list. `--file` and `--pdf` are aliases, and `--bank` may be omitted to auto-detect most banks.

## Active bank codes

`axis`, `bob`, `boi`, `bom`, `canara`, `central`, `cub`, `dbs`, `federal`, `hdfc`, `icici`, `idbi`, `idfc`, `indian`, `indus`, `iob`, `kotak`, `kvb`, `pnb`, `sbi`, `southind`, `tmb`, and `unionbank`.

TMB currently requires explicit `--bank tmb` because it has an active parser but no auto-detection signature.

