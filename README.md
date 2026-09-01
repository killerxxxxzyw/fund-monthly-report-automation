# Fund Monthly Report Automation

A sanitized portfolio project based on a real asset-management reporting workflow. It converts valuation statements and subscription/redemption/dividend files into a structured monthly report, performs consistency checks, and generates a safe browser-automation change plan.

> **Privacy note:** all portfolio names, fund names, amounts and sample files in this repository are synthetic. No employer data, document links, login state, cookies or production browser paths are included.

## What it demonstrates

- Batch Excel ingestion with `pandas` and `openpyxl`
- Flexible valuation-field extraction without hard-coded source row numbers
- Underlying-fund market-value aggregation
- Subscription / redemption / dividend classification and summarization
- Structural-risk detection for new or fully redeemed positions
- Report workbook generation with multiple review sheets
- Safe `dry-run` workflow before any online-sheet automation
- Environment-variable based handling of document URLs and local browser paths

## Workflow

```text
valuation files ─────┐
                     ├─> extraction / validation ─> monthly_report.xlsx
transaction files ──┘                               │
                                                   └─> dry-run change plan
```

## Quick start

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
python main.py --sample
python playwright_writer.py --dry-run
python -m unittest discover -s tests -v
```

The first command copies the synthetic workbooks from `sample_data/` into `input/` and produces:

```text
output/monthly_report.xlsx
```

The dry-run produces:

```text
output/playwright_change_plan.txt
```

## Project structure

```text
.
├── main.py
├── config.py
├── valuation_extractor.py
├── transaction_extractor.py
├── report_builder.py
├── playwright_writer.py
├── utils.py
├── sample_data/          # synthetic demonstration files only
├── input/                # ignored by Git except .gitkeep
├── output/               # ignored by Git except .gitkeep
├── tests/
├── .env.example
├── .gitignore
└── requirements.txt
```

## Safety design

The public repository intentionally excludes production selectors and login state. `playwright_writer.py` defaults to dry-run. A write attempt requires both `--write --yes` and a locally configured `KDOCS_DOC_URL`; the script refuses to proceed when unsafe plan items remain.

## Notes for recruiters / reviewers

This repository focuses on the engineering pattern rather than exposing a proprietary spreadsheet template. The original workflow automated recurring FOF-style reporting tasks: extracting valuation metrics, reconciling underlying positions and transactions, generating review-ready Excel outputs, and preparing online-sheet updates with safety checks.
