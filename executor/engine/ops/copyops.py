"""区块复制与整表复制（可含公式、格式、列宽、合并单元格）"""
from __future__ import annotations

import copy

from openpyxl.cell.cell import MergedCell

from .. import refs
from ..book import copy_style
from ..errors import StepError
from . import op

ALL_CONTENT = ["values", "formulas", "styles", "widths", "merges"]


def _content(step):
    c = step.get("content", "all")
    if c == "all":
        return set(ALL_CONTENT)
    if isinstance(c, str):
        c = [c]
    bad = [x for x in c if x not in ALL_CONTENT]
    if bad:
        raise StepError(f"content 只能包含：{'、'.join(ALL_CONTENT)}，不认识：{'、'.join(bad)}")
    return set(c)


def _copy_block(ctx, src_fh, src_sheet, box, dst_view_ws, dst_c, dst_r, content, fallback_mode):
    """把 src 的 box=(c1,r1,c2,r2) 复制到目标左上角 (dst_c, dst_r)。返回 (单元格数, 转成值的公式数)"""
    c1, r1, c2, r2 = box
    src_f = src_fh.formula_book()[src_sheet]
    src_v = src_fh.value_book()[src_sheet]
    same_book = src_fh is ctx.main
    main_sheets = ctx.main.wb.sheetnames
    dc, dr = dst_c - c1, dst_r - r1
    n, fb = 0, 0
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            s = src_f._cells.get((r, c))
            if s is None or isinstance(s, MergedCell):
                continue
            t = dst_view_ws.cell(row=r + dr, column=c + dc)
            v = s.value
            if isinstance(v, str) and v.startswith("="):
                if "formulas" in content:
                    bad = refs.has_external_ref(v) or (not same_book and any(x not in main_sheets for x in refs.explicit_sheets(v)))
                    if bad:
                        if fallback_mode == "error":
                            raise StepError(f"{src_sheet}!{s.coordinate} 的公式引用了主表中不存在的工作表或其他文件")
                        sv = src_v._cells.get((r, c)) if not same_book else None
                        t.value = sv.value if sv is not None else ctx.main.cached.get(id(s))
                        fb += 1
                    else:
                        t.value = refs.translate_formula(v, dr, dc)
                        ctx.main.written_formulas.add(id(t))
                elif "values" in content:
                    sv = src_v._cells.get((r, c)) if not same_book else None
                    t.value = sv.value if sv is not None else ctx.main.cached.get(id(s))
            elif "values" in content:
                t.value = v
            if "styles" in content and s.has_style:
                copy_style(s, t, same_book)
            n += 1
    if "widths" in content:
        for c in range(c1, c2 + 1):
            d = src_f.column_dimensions.get(refs.col_letter(c))
            if d is not None and d.width:
                dst_view_ws.column_dimensions[refs.col_letter(c + dc)].width = d.width
    if "merges" in content:
        for mr in list(src_f.merged_cells.ranges):
            if mr.min_col >= c1 and mr.max_col <= c2 and mr.min_row >= r1 and mr.max_row <= r2:
                dst_view_ws.merge_cells(start_row=mr.min_row + dr, start_column=mr.min_col + dc,
                                        end_row=mr.max_row + dr, end_column=mr.max_col + dc)
    return n, fb


@op("copy_range", "区块复制（可含公式与格式）", required=["from", "to"], optional=["content", "formula_fallback"])
def copy_range(ctx, step):
    """from: {"file": "ref", "sheet": "汇总", "range": "AF1:AI30"}；to: {"sheet": "...", "cell": "AF1"}"""
    src = step["from"]
    if not isinstance(src, dict) or "range" not in src:
        raise StepError('from 应写成 {"file": "ref", "sheet": "工作表名", "range": "AF1:AI30"}')
    fh = ctx.file(src.get("file", "main"))
    sheet = fh.resolve_sheet(src.get("sheet"))
    box = refs.parse_range(src["range"])
    dst = step["to"]
    if isinstance(dst, str):
        dst = {"cell": dst}
    view = ctx.main.view(dst.get("sheet", step.get("sheet")), step.get("header_row"))
    dc, dr = refs.parse_cell(dst["cell"])
    n, fb = _copy_block(ctx, fh, sheet, box, view.ws, dc, dr, _content(step), step.get("formula_fallback", "value"))
    view.invalidate()
    if fb:
        ctx.warn(f"{fb} 个公式引用了主表中不存在的工作表或其他文件，已改为复制计算结果")
    end = refs.col_letter(dc + box[2] - box[0]) + str(dr + box[3] - box[1])
    return {"cells": n, "detail": f"从「{fh.label}」{sheet}!{src['range']} 复制到 {view.name}!{dst['cell']}:{end}，共 {n} 个单元格"}


@op("copy_sheet", "复制整个工作表到主表", required=["from"], optional=["name", "content", "replace", "formula_fallback"])
def copy_sheet(ctx, step):
    """from: {"file": "ref", "sheet": "单价表"}；name: 新表名（默认同名）。复制进来后，公式就能引用它"""
    src = step["from"]
    if isinstance(src, str):
        src = {"file": src}
    fh = ctx.file(src.get("file", "ref"))
    sheet = fh.resolve_sheet(src.get("sheet"))
    name = step.get("name") or sheet
    wb = ctx.main.wb
    if name in wb.sheetnames:
        if not step.get("replace"):
            raise StepError(f"主表里已有工作表「{name}」，可换个 name，或加 \"replace\": true 覆盖")
        del wb[name]
    ws = wb.create_sheet(name)
    src_ws = fh.formula_book()[sheet]
    if not src_ws._cells:
        return {"detail": f"「{sheet}」是空表"}
    max_r = max(r for r, _ in src_ws._cells)
    max_c = max(c for _, c in src_ws._cells)
    n, fb = _copy_block(ctx, fh, sheet, (1, 1, max_c, max_r), ws, 1, 1, _content(step), step.get("formula_fallback", "value"))
    for r, d in src_ws.row_dimensions.items():
        if d.height:
            ws.row_dimensions[r].height = d.height
    if src_ws.freeze_panes:
        ws.freeze_panes = src_ws.freeze_panes
    if fb:
        ctx.warn(f"复制工作表时有 {fb} 个公式引用了不存在的工作表，已改为计算结果")
    return {"cells": n, "detail": f"把「{fh.label}」的「{sheet}」复制为主表的「{name}」（{max_r} 行 × {max_c} 列）"}
