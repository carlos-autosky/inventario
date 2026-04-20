
from __future__ import annotations
import re
import unicodedata
from datetime import datetime
import pandas as pd


def strip_accents(value: str) -> str:
    value = unicodedata.normalize("NFKD", str(value))
    return "".join(ch for ch in value if not unicodedata.combining(ch))


def normalize_token(value: str) -> str:
    value = strip_accents(str(value)).lower().strip()
    value = re.sub(r"\s+", " ", value)
    return value


def safe_num(series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def parse_dates(series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def fmt_num(value) -> str:
    try:
        return f"{float(value):,.2f}"
    except Exception:
        return "0.00"


def fmt_pct(value) -> str:
    try:
        return f"{float(value):,.2f}%"
    except Exception:
        return "0.00%"


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
