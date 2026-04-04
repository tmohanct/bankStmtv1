# AGENTS

## Project Scope
- This repository is a Python bank statement analysis tool scaffold.
- The first supported bank is Axis Bank (`axis`).
- Parser architecture should stay modular so additional banks can be added without changing the pipeline.

## Input And Output
- Put source PDFs in `input/`.
- Keep parsing rules in `input/Rules.xlsx`.
- Write intermediate output to `output/output.xlsx`.
- Write final workbook to `output/<pdf_base_name>.xlsx`.
- If `output/<pdf_base_name>.xlsx` exists, append timestamp format `YYMMDD_HHMMSS`.

## Code Organization
- Keep bank-specific logic only inside `src/parsers/<bank>_parser.py`.
- Keep normalization and validation logic in `src/transform/`.
- Keep export logic in `src/export/`.
- Keep shared helpers in `src/utils/`.
