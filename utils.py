# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from pathlib import Path
from typing import Any


def to_number(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return 0.0
    try:
        return float(text)
    except (TypeError, ValueError):
        return 0.0


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().replace("（", "(").replace("）", ")")
    return re.sub(r"\s+", "", text)


def clean_product_name(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    if "_" in text and re.match(r"^[A-Za-z0-9]+_", text):
        # Keep synthetic Portfolio_A names intact; strip only leading security codes.
        if not text.startswith("Portfolio_"):
            text = text.split("_", 1)[1]
    return text.replace("（CW）", "").replace("(CW)", "").strip()


def clean_portfolio_name(value: Any) -> str:
    text = "" if value is None else str(value).strip()
    if text.startswith("Portfolio_"):
        return text
    if "_" in text:
        return text.split("_", 1)[1].strip()
    return text


def extract_portfolio_from_text(text: str) -> str:
    if not text:
        return ""
    match = re.search(r"Portfolio_[A-Z]", str(text), flags=re.IGNORECASE)
    if match:
        value = match.group(0)
        return "Portfolio_" + value[-1].upper()
    match = re.search(r"示例组合\s*([A-Z])", str(text), flags=re.IGNORECASE)
    if match:
        return f"Portfolio_{match.group(1).upper()}"
    return ""


def iter_input_files(input_dir: Path):
    for path in sorted(input_dir.iterdir()):
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() in {".xls", ".xlsx"}:
            yield path
