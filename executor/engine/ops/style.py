"""格式：按条件着色、字体颜色、加粗、清除底色"""
from __future__ import annotations

from openpyxl.styles import PatternFill

from .. import cond as C
from . import font_with, op, row_getter, solid_fill


@op("format", "按条件设置格式（着色）", required=["columns"],
    optional=["where", "fill", "font_color", "bold", "clear_fill", "include_header", "number_format"],
    aliases=["color"])
def format_cells(ctx, step):
    view = ctx.view(step)
    cols = view.cols(step["columns"])
    fill = solid_fill(step["fill"]) if step.get("fill") else None
    getter = row_getter(ctx, view)
    rows = view.data_rows()
    n = 0
    for r in ([view.header_row] if step.get("include_header") else []) + rows:
        if r != view.header_row and "where" in step and not C.evaluate(step["where"], getter(r)):
            continue
        for c in cols:
            cell = view.ws.cell(row=r, column=c)
            if fill is not None:
                cell.fill = fill
            elif step.get("clear_fill"):
                cell.fill = PatternFill(fill_type=None)
            if step.get("font_color") or "bold" in step:
                cell.font = font_with(cell, step.get("font_color"), step.get("bold"))
            if step.get("number_format"):
                cell.number_format = step["number_format"]
        n += 1
    return {"rows": n, "detail": f"{len(cols)} 列 × {n} 行设置格式"}
