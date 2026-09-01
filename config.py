# -*- coding: utf-8 -*-
"""Project configuration for the public demo version.

All names and sample data in this repository are synthetic. Real document URLs,
login state, cookies, local browser paths, and business data are intentionally
excluded from source control.
"""
from __future__ import annotations

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "monthly_report.xlsx"

UNDERLYING_CODE_PREFIX = "1108.26.01."

FIELD_NAV = "基金单位净值"
FIELD_PAID_IN_CAPITAL = "实收资本"
FIELD_NET_ASSET = "基金资产净值"
FIELD_BANK_DEPOSIT = "银行存款"
FIELD_OTC_MONEY_FUND = "场外货币基金"

PURCHASE_HEADERS = {"确认份额", "成交金额"}
REDEMPTION_HEADERS = {"赎回份额", "赎回金额（含费）"}
DIVIDEND_HEADERS = {"到账金额", "红利截止日"}

# Synthetic portfolio/fund order used by the sample reporting template.
PORTFOLIO_FUND_ORDER = {
    "Portfolio_A": [
        ("Manager Alpha", "Alpha Market Neutral Fund"),
        ("Manager Beta", "Beta Quant Hedge Fund"),
        ("Manager Gamma", "Gamma Multi Strategy Fund"),
    ],
    "Portfolio_B": [
        ("Manager Delta", "Delta Absolute Return Fund"),
        ("Manager Epsilon", "Epsilon Equity Neutral Fund"),
    ],
    "Portfolio_C": [
        ("Manager Zeta", "Zeta Quant Selection Fund"),
        ("Manager Eta", "Eta Macro Strategy Fund"),
        ("Manager Theta", "Theta Enhanced Index Fund"),
    ],
}

PORTFOLIO_ORDER = list(PORTFOLIO_FUND_ORDER)

# Online-sheet automation configuration. These are deliberately non-secret
# placeholders. Put real values in environment variables or a local .env file
# that is never committed.
DOC_URL = os.getenv("KDOCS_DOC_URL", "")
TARGET_SHEET_KEYWORD = os.getenv("KDOCS_TARGET_SHEET", "Monthly Report")
CHROME_PATH = os.getenv("CHROME_PATH", "")
LOGIN_STATE_DIR = BASE_DIR / "browser_login_state"

# Synthetic demo cell mapping. A real deployment should replace these values
# locally after validating the target template.
SUMMARY_CELL_MAP = {
    "Portfolio_A": {"净值": "C6", "产品成本规模": "E5", "产品市值规模": "E6", "剩余资金": "E10", "剩余资金备注": "F10"},
    "Portfolio_B": {"净值": "C16", "产品成本规模": "E15", "产品市值规模": "E16", "剩余资金": "E20", "剩余资金备注": "F20"},
    "Portfolio_C": {"净值": "C26", "产品成本规模": "E25", "产品市值规模": "E26", "剩余资金": "E30", "剩余资金备注": "F30"},
}

MARKET_VALUE_START_CELLS = {
    "Portfolio_A": "D7",
    "Portfolio_B": "D17",
    "Portfolio_C": "D27",
}

EXPECTED_MARKET_VALUE_ROWS = {
    portfolio: len(funds) for portfolio, funds in PORTFOLIO_FUND_ORDER.items()
}
