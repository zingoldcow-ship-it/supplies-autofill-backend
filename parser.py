from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd


@dataclass
class ParsedItem:
    name: str
    spec: str
    qty: float
    unit_list: float
    unit_sale: float
    product_code: str
    site: str = "아이스크림몰"


def _to_number(x) -> float:
    """Convert '28,400원' / '1,234' / 1234 / None -> float."""
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "":
        return 0.0
    # keep digits, dot, minus
    s = re.sub(r"[^\d\.\-]", "", s)
    if s in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def split_name_spec(raw_name: str) -> Tuple[str, str]:
    """Try to split spec from name using (), [], or ' / ' patterns."""
    if not raw_name:
        return "", ""
    name = str(raw_name).strip()
    # parentheses
    m = re.search(r"^(.*?)[\s]*\(([^)]{1,80})\)[\s]*$", name)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # brackets
    m = re.search(r"^(.*?)[\s]*\[([^\]]{1,80})\][\s]*$", name)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # slash with spaces
    if " / " in name:
        left, right = name.split(" / ", 1)
        # treat right side as spec if it's short-ish
        if len(right.strip()) <= 30:
            return left.strip(), right.strip()
    return name, ""


def detect_header_row(df: pd.DataFrame) -> int:
    """Find a likely header row index by scanning for '상품명' and '수량'."""
    for i in range(min(30, len(df))):
        row = df.iloc[i].astype(str).fillna("")
        row_join = " ".join(row.tolist())
        if ("상품명" in row_join) and ("수량" in row_join):
            return i
    return 0


def normalize_columns(cols: List[str]) -> List[str]:
    out = []
    for c in cols:
        s = str(c).strip()
        s = re.sub(r"\s+", " ", s)
        out.append(s)
    return out


def parse_icecream_excel(upload_bytes: bytes) -> pd.DataFrame:
    """
    Parse an IcecreamMall cart/estimate Excel into a normalized dataframe with columns:
    품목, 규격, 수량, 단가(정가), 단가(할인), 금액(정가), 최종금액, 상품코드, 사이트
    """
    # read first sheet raw (no header)
    raw = pd.read_excel(upload_bytes, header=None, engine="openpyxl")
    header_i = detect_header_row(raw)
    df = pd.read_excel(upload_bytes, header=header_i, engine="openpyxl")
    df.columns = normalize_columns(list(df.columns))

    # candidate columns
    def pick(*cands):
        for c in cands:
            if c in df.columns:
                return c
        # fuzzy contains
        for col in df.columns:
            for c in cands:
                if c and (c in str(col)):
                    return col
        return None

    col_name = pick("상품명", "품명", "상품")
    col_qty = pick("수량", "구매수량", "주문수량")
    col_unit_list = pick("1개당 금액", "정가", "판매가", "상품가격", "단가")
    col_unit_sale = pick("할인적용금액", "할인가", "할인금액", "할인적용")
    col_code = pick("상품코드", "상품 코드", "코드", "상품번호")

    # If sale unit missing, fall back to list unit.
    if col_unit_sale is None:
        col_unit_sale = col_unit_list

    # build rows
    items = []
    for _, r in df.iterrows():
        nm = "" if col_name is None else r.get(col_name)
        if pd.isna(nm) or str(nm).strip() == "":
            continue

        qty = _to_number(r.get(col_qty)) if col_qty else 0.0
        unit_list = _to_number(r.get(col_unit_list)) if col_unit_list else 0.0
        unit_sale = _to_number(r.get(col_unit_sale)) if col_unit_sale else unit_list
        code = "" if col_code is None else ("" if pd.isna(r.get(col_code)) else str(r.get(col_code)).strip())

        name, spec = split_name_spec(str(nm))
        items.append(
            {
                "품목": name,
                "규격": spec,
                "수량": qty,
                "단가(정가)": unit_list,
                "단가(할인)": unit_sale,
                "금액(정가)": qty * unit_list,
                "최종금액": qty * unit_sale,
                "상품코드": code,
                "사이트": "아이스크림몰",
            }
        )

    out = pd.DataFrame(items)
    # basic type cleanup
    if not out.empty:
        out["수량"] = out["수량"].astype(float)
        for c in ["단가(정가)", "단가(할인)", "금액(정가)", "최종금액"]:
            out[c] = out[c].astype(float)
    return out
