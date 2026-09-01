# -*- coding: utf-8 -*-
"""Extract and summarize subscription, redemption and dividend files."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd

from config import DIVIDEND_HEADERS, PURCHASE_HEADERS, REDEMPTION_HEADERS
from utils import clean_portfolio_name, clean_product_name, to_number


def _read_table(file_path: Path) -> pd.DataFrame:
    if file_path.suffix.lower() == ".xls":
        return pd.read_excel(file_path, dtype=object, engine="xlrd")
    return pd.read_excel(file_path, dtype=object, engine="openpyxl")


def _classify_transaction(headers: List[str]) -> str:
    header_set = {str(header).strip() for header in headers}
    if PURCHASE_HEADERS.issubset(header_set):
        return "申购"
    if REDEMPTION_HEADERS.issubset(header_set):
        return "赎回"
    if DIVIDEND_HEADERS.issubset(header_set):
        return "分红"
    return "未知"


def format_date(value: Any) -> str:
    if value is None or str(value).strip() == "" or str(value).lower() == "nan":
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y%m%d")
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y%m%d", "%Y.%m.%d"):
        try:
            candidate = text[:8] if fmt == "%Y%m%d" else text[:10]
            return datetime.strptime(candidate, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return text


def amount_to_wan_text(amount: float) -> str:
    wan = to_number(amount) / 10000
    return f"{int(round(wan))}万" if abs(wan - round(wan)) < 1e-6 else f"{wan:.2f}万"


def make_progress_text(txn_type: str, date_text: str, amount: float) -> str:
    prefix = date_text if date_text else ""
    return f"{prefix}{txn_type}{amount_to_wan_text(amount)}"


def extract_transaction_file(file_path: Path) -> Tuple[str, List[Dict[str, Any]]]:
    df = _read_table(file_path)
    df.columns = [str(col).strip() for col in df.columns]
    txn_type = _classify_transaction(list(df.columns))
    if txn_type == "未知":
        return txn_type, [{"来源文件": file_path.name, "错误": "无法识别交易文件类型，请检查表头"}]

    rows: List[Dict[str, Any]] = []
    for _, row in df.iterrows():
        portfolio = clean_portfolio_name(row.get("产品"))
        fund = clean_product_name(row.get("资产管理产品"))
        if not portfolio and not fund:
            continue
        if txn_type == "申购":
            amount = to_number(row.get("成交金额"))
            share = to_number(row.get("确认份额"))
            date_value = row.get("确认日期")
        elif txn_type == "赎回":
            amount = to_number(row.get("赎回金额（含费）"))
            share = to_number(row.get("赎回份额"))
            date_value = row.get("确认日期")
        else:
            amount = to_number(row.get("到账金额"))
            share = 0.0
            date_value = row.get("到账日期")
        date_text = format_date(date_value)
        rows.append({
            "交易类型": txn_type,
            "组合产品": portfolio,
            "底层产品名称": fund,
            "日期": date_text,
            "金额": amount,
            "金额_万元": amount / 10000,
            "份额": share,
            "在线表格文字": make_progress_text(txn_type, date_text, amount),
            "来源文件": file_path.name,
        })
    return txn_type, rows


def extract_all_transactions(files: List[Path]):
    details: List[Dict[str, Any]] = []
    for file_path in files:
        try:
            _, rows = extract_transaction_file(file_path)
            details.extend(rows)
        except Exception as exc:
            details.append({"交易类型": "错误", "来源文件": file_path.name, "错误": str(exc)})
    return details


def _join_text(old: str, new: str) -> str:
    if not new:
        return old or ""
    return new if not old else f"{old}/{new}"


def summarize_transactions(details: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    grouped: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for row in details:
        if row.get("错误"):
            continue
        key = (row.get("组合产品", ""), row.get("底层产品名称", ""))
        item = grouped.setdefault(key, {
            "组合产品": key[0], "底层产品名称": key[1],
            "申购金额": 0.0, "赎回金额": 0.0, "分红金额": 0.0,
            "申购文字_追加": "", "赎回文字_追加": "", "分红文字_追加": "",
        })
        amount = to_number(row.get("金额"))
        text = row.get("在线表格文字", "")
        txn_type = row.get("交易类型")
        if txn_type == "申购":
            item["申购金额"] += amount
            item["申购文字_追加"] = _join_text(item["申购文字_追加"], text)
        elif txn_type == "赎回":
            item["赎回金额"] += amount
            item["赎回文字_追加"] = _join_text(item["赎回文字_追加"], text)
        elif txn_type == "分红":
            item["分红金额"] += amount
            item["分红文字_追加"] = _join_text(item["分红文字_追加"], text)
    result = []
    for item in grouped.values():
        item["净申购"] = item["申购金额"] - item["赎回金额"]
        item["已投金额调整"] = item["净申购"]
        item["已投金额调整_万元"] = item["净申购"] / 10000
        result.append(item)
    return sorted(result, key=lambda x: (x["组合产品"], x["底层产品名称"]))


def summarize_transactions_by_portfolio(details: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    grouped: Dict[str, Dict[str, Any]] = {}
    for row in details:
        if row.get("错误"):
            continue
        portfolio = row.get("组合产品", "")
        if not portfolio:
            continue
        item = grouped.setdefault(portfolio, {
            "组合产品": portfolio,
            "申购合计": 0.0, "赎回合计": 0.0, "分红合计": 0.0,
            "申购文字_追加": "", "赎回文字_追加": "", "分红文字_追加": "",
        })
        amount = to_number(row.get("金额"))
        text = row.get("在线表格文字", "")
        txn_type = row.get("交易类型")
        if txn_type == "申购":
            item["申购合计"] += amount
            item["申购文字_追加"] = _join_text(item["申购文字_追加"], text)
        elif txn_type == "赎回":
            item["赎回合计"] += amount
            item["赎回文字_追加"] = _join_text(item["赎回文字_追加"], text)
        elif txn_type == "分红":
            item["分红合计"] += amount
            item["分红文字_追加"] = _join_text(item["分红文字_追加"], text)
    for item in grouped.values():
        item["净申购合计"] = item["申购合计"] - item["赎回合计"]
        item["已投金额调整"] = item["净申购合计"]
        item["已投金额调整_万元"] = item["净申购合计"] / 10000
    return sorted(grouped.values(), key=lambda x: x["组合产品"])
