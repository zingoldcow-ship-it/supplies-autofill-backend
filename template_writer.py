from __future__ import annotations

from io import BytesIO
from typing import Optional

import openpyxl
import pandas as pd


def fill_template(template_bytes: bytes, items_df: pd.DataFrame) -> bytes:
    """
    Fill the provided Icecream estimate template:
    - Write rows starting at row 12
    - Columns:
      A 순번, B 판매자(빈칸), C 상품명(품목+규격), D 1개당 금액(정가),
      E 업체 할인금액(정가-할인), F 쿠폰 할인금액(0),
      G 할인적용금액(할인), H 수량, I 합계금액(수량*할인)
    Keep styles by copying from the first data row.
    """
    wb = openpyxl.load_workbook(BytesIO(template_bytes))
    ws = wb[wb.sheetnames[0]]

    start_row = 12
    # find last existing data row (by col A numeric)
    r = start_row
    while ws.cell(r, 1).value not in (None, ""):
        r += 1
    existing_rows = r - start_row

    n = len(items_df)
    if n <= 0:
        out = BytesIO()
        wb.save(out)
        return out.getvalue()

    # Ensure enough rows: insert if template doesn't have space
    # We'll insert rows below the first data row style row if needed.
    # Determine current max row to keep footer lines intact.
    if n > existing_rows and existing_rows > 0:
        insert_at = start_row + existing_rows
        ws.insert_rows(insert_at, amount=(n - existing_rows))

    # style source row: start_row (may be empty but has formatting)
    style_row = start_row

    def copy_style(src_cell, dst_cell):
        dst_cell._style = src_cell._style
        dst_cell.font = src_cell.font
        dst_cell.border = src_cell.border
        dst_cell.fill = src_cell.fill
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = src_cell.protection
        dst_cell.alignment = src_cell.alignment

    # Write rows
    for idx, row in enumerate(items_df.itertuples(index=False), start=1):
        rr = start_row + (idx - 1)

        # copy styles across A..I from style_row
        for col in range(1, 10):
            copy_style(ws.cell(style_row, col), ws.cell(rr, col))

        # compose display name
        display_name = str(getattr(row, "품목", "")).strip()
        spec = str(getattr(row, "규격", "")).strip()
        if spec:
            display_name = f"{display_name} ({spec})"

        qty = float(getattr(row, "수량", 0.0) or 0.0)
        unit_list = float(getattr(row, "단가(정가)", 0.0) or 0.0)
        unit_sale = float(getattr(row, "단가(할인)", 0.0) or 0.0)

        ws.cell(rr, 1).value = idx                         # A
        ws.cell(rr, 2).value = ""                          # B
        ws.cell(rr, 3).value = display_name                # C
        ws.cell(rr, 4).value = unit_list                   # D
        ws.cell(rr, 5).value = max(unit_list - unit_sale, 0.0)  # E
        ws.cell(rr, 6).value = 0                           # F
        ws.cell(rr, 7).value = unit_sale                   # G
        ws.cell(rr, 8).value = qty                         # H
        ws.cell(rr, 9).value = qty * unit_sale             # I

    # Optionally update the "합계 금액" at top (cell A3 has text)
    # We'll compute total of column I.
    total = float(items_df["최종금액"].sum()) if "최종금액" in items_df.columns else 0.0
    # Write to L9 maybe? In template, L9 is final total (text like 28,400원)
    # We'll put numeric in L9 and format as currency-like with 원 suffix via number_format.
    try:
        ws["L9"].value = total
        ws["L9"].number_format = '#,##0"원"'
        ws["H7"].value = total
        ws["H7"].number_format = '#,##0"원"'
        ws["H9"].value = total
        ws["H9"].number_format = '#,##0"원"'
        ws["L7"].value = total
        ws["L7"].number_format = '#,##0"원"'
    except Exception:
        pass

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
