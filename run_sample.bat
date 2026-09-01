@echo off
python main.py --sample
python playwright_writer.py --dry-run
python -m unittest discover -s tests -v
pause
