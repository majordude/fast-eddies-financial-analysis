# Fast Eddies Financial Spending and Income Analysis

![Check](https://github.com/majordude/fast-eddies-financial-analysis/actions/workflows/check.yml/badge.svg)

Turns bank and credit card statement PDFs into an Excel workbook: income vs expenses, spending by category, transfers,
charts, and proof that every month ties to the statements. Statements are redacted first; the analyzer reads only the
redacted copies. Nothing is uploaded anywhere.

Start with README_FIRST.txt (what to install on Windows 11 and how). Day-to-day use is in HOWTO.txt, and how the
numbers are checked is in HOW_IT_WORKS.txt.

No statements, workbooks or personal rules are stored in this repository. Your own rules go in personal_config.py,
which git ignores (see personal_config.example.py).

## License

MIT License (see LICENSE). Copyright (c) 2026 majordude.

## Disclaimer

This is a bookkeeping tool, not financial, tax or legal advice. It is provided as is, with no warranty. Check its
output against your own statements before relying on it.
