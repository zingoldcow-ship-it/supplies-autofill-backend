from __future__ import annotations
import io
import re
import pandas as pd
import numpy as np

RE_NUM = re.compile(r"[^0-9\-]+")

def _to_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (int, np.integer)):
        return int(x)
    s = str(x).strip()
    if s == "" or s.lower() == "none":
        return None
    s = RE_NUM.sub("", s)
    if s in ("", "-"):
        return None
    try:
        return int(float(s))
    except Exception:
        return None

def _find_header_row(raw: pd.DataFrame, max_scan: int = 30) -> int:
    # Find a row that likely contains header names.
    for r in range(min(max_scan, len(raw))):
        row = raw.iloc[r].astype(str).fillna("").str.replace("\n", " ").str.strip()
        joined = " | ".join(row.tolist())
        if ("상품" in joined and "수량" in joined) or ("상품명" in joined):
            return r
    return 0

def _pick_col(cols, keywords):
    for k in keywords:
        for c in cols:
            if k in c:
                return c
    return None

def _extract_code_from_text(name: str):
    # Extract product code from patterns like [11033697] or (11033697) or trailing digits
    if not name:
        return None, name
    code = None
    m = re.search(r"\[(\d{6,})\]", name)
    if m:
        code = m.group(1)
        name = re.sub(r"\s*\[\d{6,}\]\s*", " ", name).strip()
        return code, name
    m = re.search(r"\((\d{6,})\)", name)
    if m:
        code = m.group(1)
        name = re.sub(r"\s*\(\d{6,}\)\s*", " ", name).strip()
        return code, name
    m = re.search(r"(\d{6,})\s*$", name)
    if m:
        code = m.group(1)
        # don't remove if the whole name is just digits
        if len(name) > len(code) + 1:
            name = re.sub(r"\s*\d{6,}\s*$", "", name).strip()
    return code, name

def _extract_spec_from_option_line(text: str):
    # Examples: "사이즈별 : 25mm" / "옵션: 3호" / "색상: 빨강"
    if not text:
        return None
    t = str(text).strip()
    m = re.match(r"^\s*([^\:]{1,12})\s*[:：]\s*(.+?)\s*$", t)
    if not m:
        return None
    key = m.group(1).strip()
    val = m.group(2).strip()
    # Keep only the value as "규격" (teacher wanted spec found from the row)
    if val:
        return val
    return None

def parse_icecream_excel(xlsx_bytes: bytes) -> pd.DataFrame:
    # Read without header to detect header row
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), header=None)
    header_row = _find_header_row(raw)
    df = pd.read_excel(io.BytesIO(xlsx_bytes), header=header_row)
    df.columns = [str(c).strip().replace("\n", " ") for c in df.columns]

    cols = list(df.columns)

    col_name = _pick_col(cols, ["상품명", "상품"])
    col_qty = _pick_col(cols, ["수량", "개수"])
    col_unit_regular = _pick_col(cols, ["정가", "판매가", "상품가격", "단가(정가)", "단가"])
    col_unit_discount = _pick_col(cols, ["할인가", "할인적용", "구매가", "단가(할인)"])
    col_total_discount = _pick_col(cols, ["최종금액", "결제금액", "합계", "금액(할인)", "할인금액"])
    col_code = _pick_col(cols, ["상품코드", "상품번호", "상품ID", "상품코드(옵션)"])

    if col_name is None:
        raise ValueError("상품명 컬럼을 찾지 못했습니다. (엑셀 헤더가 예상과 다릅니다)")

    # Build normalized rows
    out_rows = []
    last_idx = None

    for _, row in df.iterrows():
        name = row.get(col_name)
        name = "" if (name is None or (isinstance(name, float) and np.isnan(name))) else str(name).strip()
        if name == "":
            continue

        qty = _to_int(row.get(col_qty)) if col_qty else None
        unit_r = _to_int(row.get(col_unit_regular)) if col_unit_regular else None
        unit_d = _to_int(row.get(col_unit_discount)) if col_unit_discount else None
        tot_d = _to_int(row.get(col_total_discount)) if col_total_discount else None
        code = row.get(col_code) if col_code else None
        code = None if (code is None or (isinstance(code, float) and np.isnan(code))) else str(code).strip()
        if code and not re.fullmatch(r"\d{4,}", code):
            # keep only digits if mixed
            digits = re.sub(r"\D", "", code)
            code = digits if digits else None

        # Option/spec line handling: when qty and prices are missing, treat as spec for previous item
        if (qty is None) and (unit_r is None) and (unit_d is None) and (tot_d is None):
            spec = _extract_spec_from_option_line(name)
            if spec and last_idx is not None:
                out_rows[last_idx]["규격"] = spec
            continue

        # Regular product line
        # Extract code embedded in name if code column absent
        embedded_code, cleaned_name = _extract_code_from_text(name)
        if not code and embedded_code:
            code = embedded_code
        name = cleaned_name

        # If unit_d missing but tot_d and qty exist -> compute
        if unit_d is None and tot_d is not None and qty:
            unit_d = int(round(tot_d / qty))
        if unit_r is None:
            unit_r = unit_d

        # If tot_d missing -> compute
        if tot_d is None and qty and unit_d is not None:
            tot_d = qty * unit_d

        out = {
            "품목": name,
            "규격": None,                # may be filled by option line below
            "수량": qty,
            "단가(정가)": unit_r,
            "단가(할인)": unit_d,
            "금액(정가)": (qty * unit_r) if (qty and unit_r is not None) else None,
            "최종금액": tot_d,
            "상품코드": code,
            "사이트": "아이스크림몰",
        }
        out_rows.append(out)
        last_idx = len(out_rows) - 1

    out_df = pd.DataFrame(out_rows, columns=["품목","규격","수량","단가(정가)","단가(할인)","금액(정가)","최종금액","상품코드","사이트"])
    # Clean NaNs to None for nicer display
    out_df = out_df.where(pd.notnull(out_df), None)
    return out_df
