# Test results

Verified in the build environment on 2026-09-01.

## Automated checks executed

```bash
python main.py --sample
python playwright_writer.py --dry-run
python -m unittest discover -s tests -v
```

Results:

- Synthetic input files detected: 6
- Valuation files detected: 3
- Transaction files detected: 3
- `output/monthly_report.xlsx` generated successfully
- `playwright_writer.py --dry-run` generated a change plan successfully
- Dry-run safety validation: passed for the synthetic template
- Unit tests: 4/4 passed
- Sensitive-string scan for original company/product names, original document URL, and original local user path: no matches in public source files

## Environment used for the smoke test

- Python 3.13.5
- pandas 2.2.3
- openpyxl 3.1.5

`xlrd` is listed in `requirements.txt` for legacy `.xls` input support, but the included sample dataset uses `.xlsx`, so the `.xls` branch was not exercised in this environment.

## Browser automation boundary

The data pipeline and dry-run planning path were executed end to end. A real online-sheet write was **not** executed because the public package intentionally contains no real document URL, authenticated browser profile, or production account. The retained Playwright writer includes read-before-write comparison, block-write retry logic, safety validation, and persistent-profile support, but a real write must be tested locally against a disposable document copy first.
