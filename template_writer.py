from __future__ import annotations

from copy import copy
from io import BytesIO
from typing import Tuple

import openpyxl
import pandas as pd
from openpyxl.cell.cell import MergedCell


def _copy_cell_style(src_cell, dst_cell) -> None:
    """Copy openpyxl styles safely (copy(), not assign by reference)."""
    try:
        if not getattr(src_cell, "has_style", False):
            return
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)
    except Exception:
        pass


def _top_left_of_merge(ws, row: int, col: int) -> Tuple[int, int]:
    for rng in ws.merged_cells.ranges:
        min_col, min_row, max_col, max_row = rng.bounds
        if min_row <= row <= max_row and min_col <= col <= max_col:
            return (min_row, min_col)
    return (row, col)


def _safe_cell(ws, row: int, col: int):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        tl_row, tl_col = _top_left_of_merge(ws, row, col)
        return ws.cell(row=tl_row, column=tl_col)
    return cell


def fill_template(template_bytes: bytes, items_df: pd.DataFrame) -> bytes:
    """Fill the provided template workbook with items (robust to merged cells)."""
    wb = openpyxl.load_workbook(BytesIO(template_bytes))
    ws = wb.active

    start_row = 12
    style_row = 12

    def _get(row_obj, name: str, default=None):
        return getattr(row_obj, name, default)

    for idx, row in enumerate(items_df.itertuples(index=False), start=1):
        rr = start_row + (idx - 1)

        # copy formatting A..I
        for col in range(1, 10):
            _copy_cell_style(_safe_cell(ws, style_row, col), _safe_cell(ws, rr, col))

        name = _get(row, "품명", "") or _get(row, "품목", "") or _get(row, "상품명", "")
        spec = _get(row, "규격", "")
        title = f"{name} ({spec})" if spec else f"{name}"

        unit_list = _get(row, "단가(정가)", None)
        unit_disc = _get(row, "단가(할인)", None)
        qty = _get(row, "수량", None)

        _safe_cell(ws, rr, 1).value = idx          # A
        _safe_cell(ws, rr, 2).value = ""           # B
        _safe_cell(ws, rr, 3).value = title        # C
        _safe_cell(ws, rr, 4).value = unit_list    # D

        if unit_list is not None and unit_disc is not None:
            try:
                _safe_cell(ws, rr, 5).value = float(unit_list) - float(unit_disc)  # E
            except Exception:
                _safe_cell(ws, rr, 5).value = None
        else:
            _safe_cell(ws, rr, 5).value = None

        _safe_cell(ws, rr, 6).value = 0            # F
        _safe_cell(ws, rr, 7).value = unit_disc    # G
        _safe_cell(ws, rr, 8).value = qty          # H

        if qty is not None and unit_disc is not None:
            _safe_cell(ws, rr, 9).value = f"=H{rr}*G{rr}"  # I
        else:
            _safe_cell(ws, rr, 9).value = None

    # update totals (best-effort)
    try:
        last_rr = start_row + len(items_df) - 1
        if last_rr >= start_row:
            total_formula = f"=SUM(I{start_row}:I{last_rr})"
            for addr in ("L9", "H7", "H9", "L7"):
                try:
                    cell = ws[addr]
                    if isinstance(cell, MergedCell):
                        cell = _safe_cell(ws, cell.row, cell.column)
                    cell.value = total_formula
                except Exception:
                    pass
    except Exception:
        pass

    out = BytesIO()
    wb.save(out)
    return out.getvalue()
