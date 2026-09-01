# Workflow details

1. `main.py` scans `input/` and separates valuation files from transaction files.
2. `valuation_extractor.py` locates valuation headers and key accounting fields by label, then extracts underlying positions using a configurable account-code prefix.
3. `transaction_extractor.py` identifies subscription, redemption and dividend files from their headers and aggregates cash flows by portfolio and underlying fund.
4. `report_builder.py` merges the two streams and writes review sheets plus an online-sheet paste layout.
5. `playwright_writer.py --dry-run` reads the generated report and creates a change plan. It flags row-count or structural mismatches before any browser write can be attempted.

All demonstration data is synthetic.
