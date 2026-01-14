from __future__ import annotations

from typing import List
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from cart_parser import CartItem


COLUMNS = [
    ("품명", 38),
    ("규격", 22),
    ("수량", 10),
    ("단가(정가)", 14),
    ("단가(할인)", 14),
    ("금액(정가)", 14),
    ("최종금액", 14),
    ("상품코드", 16),
    ("사이트", 14),
]


def _apply_table_style(ws, header_row: int, start_col: int, end_col: int, end_row: int):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    header_fill = PatternFill("solid", fgColor="E8F0FE")
    header_font = Font(bold=True)

    for c in range(start_col, end_col + 1):
        cell = ws.cell(header_row, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border

    for r in range(header_row + 1, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(r, c)
            cell.border = border
            if c in (1, 2, 8, 9):
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")


def build_output_workbook(items: List[CartItem]) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "변환결과"

    # set column widths + header
    for i, (name, width) in enumerate(COLUMNS, start=1):
        ws.cell(1, i).value = name
        ws.column_dimensions[get_column_letter(i)].width = width

    # data rows
    for r_idx, it in enumerate(items, start=2):
        ws.cell(r_idx, 1).value = it.name
        ws.cell(r_idx, 2).value = it.spec
        ws.cell(r_idx, 3).value = it.qty
        ws.cell(r_idx, 4).value = it.unit_price_list
        ws.cell(r_idx, 5).value = it.unit_price_sale
        # formulas (E=할인단가)
        ws.cell(r_idx, 6).value = f"=C{r_idx}*D{r_idx}"  # 금액(정가)
        ws.cell(r_idx, 7).value = f"=C{r_idx}*E{r_idx}"  # 최종금액(할인)
        ws.cell(r_idx, 8).value = it.product_code or ""
        ws.cell(r_idx, 9).value = "아이스크림몰"

    end_row = 1 + len(items)

    # number formats
    for r in range(2, end_row + 1):
        ws.cell(r, 3).number_format = "0"
        for c in (4, 5, 6, 7):
            ws.cell(r, c).number_format = "#,##0"

    # totals row
    total_row = end_row + 1
    ws.cell(total_row, 1).value = "합계"
    ws.cell(total_row, 1).font = Font(bold=True)
    ws.cell(total_row, 6).value = f"=SUM(F2:F{end_row})"
    ws.cell(total_row, 7).value = f"=SUM(G2:G{end_row})"
    ws.cell(total_row, 6).font = Font(bold=True)
    ws.cell(total_row, 7).font = Font(bold=True)
    ws.cell(total_row, 6).number_format = "#,##0"
    ws.cell(total_row, 7).number_format = "#,##0"

    _apply_table_style(ws, 1, 1, len(COLUMNS), total_row)
    ws.freeze_panes = "A2"

    # sheet 2: 원본정리 (정가/할인가 확인용)
    ws2 = wb.create_sheet("가격정보")
    for i, h in enumerate(["품명", "규격", "수량", "정가(단가)", "할인가(단가)", "정가(금액)", "할인가(금액)"], start=1):
        ws2.cell(1, i).value = h
        ws2.column_dimensions[get_column_letter(i)].width = 26 if i in (1,) else 18

    for r_idx, it in enumerate(items, start=2):
        ws2.cell(r_idx, 1).value = it.name
        ws2.cell(r_idx, 2).value = it.spec
        ws2.cell(r_idx, 3).value = it.qty
        ws2.cell(r_idx, 4).value = it.unit_price_list
        ws2.cell(r_idx, 5).value = it.unit_price_sale
        ws2.cell(r_idx, 6).value = f"=C{r_idx}*D{r_idx}"
        ws2.cell(r_idx, 7).value = f"=C{r_idx}*E{r_idx}"
        ws2.cell(r_idx, 3).number_format = "0"
        for c in (4, 5, 6, 7):
            ws2.cell(r_idx, c).number_format = "#,##0"

    total_row2 = 1 + len(items) + 1
    ws2.cell(total_row2, 1).value = "합계"
    ws2.cell(total_row2, 1).font = Font(bold=True)
    ws2.cell(total_row2, 6).value = f"=SUM(F2:F{total_row2-1})"
    ws2.cell(total_row2, 7).value = f"=SUM(G2:G{total_row2-1})"
    ws2.cell(total_row2, 6).font = Font(bold=True)
    ws2.cell(total_row2, 7).font = Font(bold=True)
    ws2.cell(total_row2, 6).number_format = "#,##0"
    ws2.cell(total_row2, 7).number_format = "#,##0"

    _apply_table_style(ws2, 1, 1, 7, total_row2)
    ws2.freeze_panes = "A2"

    return wb


def workbook_to_bytes(wb: Workbook) -> bytes:
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.read()
