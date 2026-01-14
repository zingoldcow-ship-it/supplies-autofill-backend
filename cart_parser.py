from __future__ import annotations
from dataclasses import dataclass
from typing import List, Optional
import re
from openpyxl import load_workbook

@dataclass
class CartItem:
    name_raw: str
    qty: int
    unit_price_list: int   # 정가(1개당 금액)
    unit_price_sale: int   # 할인적용금액(할인 후 단가)
    product_code: str = "" # 장바구니 엑셀에는 보통 없음(비워둠)

def _to_int_price(val) -> int:
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip()
    s = s.replace(",", "").replace("원", "").strip()
    m = re.search(r"(\d+)", s)
    return int(m.group(1)) if m else 0

def _to_int_qty(val) -> int:
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val)
    s = str(val).strip().replace(",", "")
    return int(s) if s.isdigit() else 0

def _split_name_and_spec(name: str) -> tuple[str, str]:
    """
    규격이 상품명 안에 같이 적힌 경우가 많아서,
    괄호/대괄호 뒤쪽을 '규격'으로 분리해보되,
    실패하면 규격은 빈칸으로 둡니다.
    """
    if not name:
        return "", ""
    # 예: "에어 더블클립 (19mm)" -> ("에어 더블클립", "19mm")
    m = re.match(r"^(.*?)\s*[\(\[]\s*(.+?)\s*[\)\]]\s*$", name)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return name.strip(), ""

def parse_iscreammall_cart_xlsx(path: str) -> List[CartItem]:
    """
    아이스크림몰 '견적서/장바구니' 엑셀을 읽어 상품 목록을 추출합니다.
    기대 헤더(예시): 순번, 상품명, 1개당 금액, 할인적용금액, 수량
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    header_row = None
    col_map = {}
    for r in range(1, min(ws.max_row, 100) + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(isinstance(v, str) and v.strip() == "순번" for v in row_vals):
            header_row = r
            # map headers -> col index
            for c, v in enumerate(row_vals, start=1):
                if isinstance(v, str):
                    key = v.strip()
                    col_map[key] = c
            break

    if header_row is None:
        raise ValueError("상품정보 헤더(순번)가 있는 행을 찾지 못했습니다. 원본 장바구니/견적서 엑셀인지 확인해주세요.")

    def col(name: str) -> int:
        if name not in col_map:
            raise ValueError(f"필수 열 '{name}'을(를) 찾지 못했습니다. 현재 헤더: {list(col_map.keys())}")
        return col_map[name]

    c_name = col("상품명")
    c_list = col("1개당 금액")
    c_sale = col("할인적용금액")
    c_qty  = col("수량")

    items: List[CartItem] = []
    # 데이터는 헤더 다음 행부터
    for r in range(header_row + 1, ws.max_row + 1):
        name_val = ws.cell(r, c_name).value
        qty_val  = ws.cell(r, c_qty).value
        list_val = ws.cell(r, c_list).value
        sale_val = ws.cell(r, c_sale).value

        # 빈 줄 스킵
        if name_val is None and qty_val is None and list_val is None and sale_val is None:
            continue

        name = str(name_val).strip() if name_val else ""
        qty = _to_int_qty(qty_val)
        if not name or qty <= 0:
            continue

        unit_list = _to_int_price(list_val)
        unit_sale = _to_int_price(sale_val)
        items.append(CartItem(name_raw=name, qty=qty, unit_price_list=unit_list, unit_price_sale=unit_sale))

    if not items:
        raise ValueError("추출된 상품이 없습니다. 장바구니에 상품이 담긴 엑셀인지 확인해주세요.")
    return items
