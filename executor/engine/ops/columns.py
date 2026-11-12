"""列操作：删除列、插入列、改表头、写单元格、数字格式、列宽"""
from __future__ import annotations

from .. import refs
from ..book import check_tables, copy_style, extend_autofilter, restructure
from ..errors import StepError
from . import compile_content, op, row_getter, write_content
from .. import cond as C


@op("delete_columns", "删除列", required=["columns"])
def delete_columns(ctx, step):
    view = ctx.view(step)
    cols = sorted(set(view.cols(step["columns"])))
    names = [view.header_of(c) or refs.col_letter(c) for c in cols]
    check_tables(view.ws, at_col=cols)
    restructure(ctx.main, view.name, refs.IndexMap(deleted=cols), None)
    return {"detail": f"删除 {len(cols)} 列：{'、'.join(names)}"}


def _insert_position(view, step):
    keys = [k for k in ("before", "after", "at", "end") if k in step]
    if len(keys) != 1:
        raise StepError("插入位置要用 before / after / at / end 其中一个")
    k = keys[0]
    if k == "before":
        return view.col(step["before"])
    if k == "after":
        return view.col(step["after"]) + 1
    if k == "at":
        return refs.col_index(str(step["at"])) if not isinstance(step["at"], int) else int(step["at"])
    return max(view.max_col(), max(view.headers().values() or [0])) + 1


@op("insert_column", "插入列", required=["header", "before|after|at|end"],
    optional=["value", "formula", "formula_from", "from_col", "where", "style_from", "width"])
def insert_column(ctx, step):
    view = ctx.view(step)
    k = _insert_position(view, step)
    if k <= view.max_col():
        check_tables(view.ws, at_col=[k])
        restructure(ctx.main, view.name, refs.IndexMap(insert_at=k, count=1), None)
    else:
        extend_autofilter(view, k)
    ws = view.ws
    hr = view.header_row
    # 格式：默认取左边一列（与 Excel 插入列时一致）
    style_from = step.get("style_from", "left")
    src_col = None
    if style_from == "left" and k > 1:
        src_col = k - 1
    elif style_from == "right":
        src_col = k + 1
    elif style_from not in (None, "none", "left", "right"):
        src_col = view.col(style_from)
    rows = view.data_rows()
    if src_col:
        for r in [hr] + rows:
            s = ws._cells.get((r, src_col))
            if s is not None and s.has_style:
                copy_style(s, ws.cell(row=r, column=k), True)
        sd = ws.column_dimensions.get(refs.col_letter(src_col))
        if sd is not None and sd.width:
            ws.column_dimensions[refs.col_letter(k)].width = sd.width
    if step.get("width"):
        ws.column_dimensions[refs.col_letter(k)].width = float(step["width"])
    view.set_value(hr, k, step["header"])
    view.invalidate()
    content = compile_content(ctx, step, view, k)
    n = 0
    if content is not None:
        getter = row_getter(ctx, view)
        for r in rows:
            if "where" in step and not C.evaluate(step["where"], getter(r)):
                continue
            write_content(view, r, k, content(r))
            n += 1
    return {"rows": n, "detail": f"在第 {refs.col_letter(k)} 列插入「{step['header']}」" + (f"，写入 {n} 行" if content is not None else "")}


@op("rename_header", "修改表头", required=["column", "to"])
def rename_header(ctx, step):
    view = ctx.view(step)
    c = view.col(step["column"])
    old = view.header_of(c)
    view.set_value(view.header_row, c, step["to"])
    view.invalidate()
    return {"detail": f"{refs.col_letter(c)} 列表头「{old}」→「{step['to']}」"}


@op("set_cells", "写入指定单元格", required=["cells"])
def set_cells(ctx, step):
    """cells: {"AC1": "备注", "B2": "=SUM(B3:B9)"} 或 [{"cell": "AC1", "value": "备注"}]"""
    view = ctx.view(step)
    items = step["cells"]
    if isinstance(items, dict):
        items = [{"cell": k, "value": v} for k, v in items.items()]
    n = 0
    for it in items:
        c, r = refs.parse_cell(it["cell"])
        write_content(view, r, c, it.get("value"))
        n += 1
    view.invalidate()
    return {"cells": n, "detail": f"写入 {n} 个单元格"}


@op("number_format", "设置数字格式", required=["columns", "format"], optional=["where", "include_header"])
def number_format(ctx, step):
    view = ctx.view(step)
    cols = view.cols(step["columns"])
    getter = row_getter(ctx, view)
    n = 0
    for r in view.data_rows():
        if "where" in step and not C.evaluate(step["where"], getter(r)):
            continue
        for c in cols:
            view.ws.cell(row=r, column=c).number_format = step["format"]
        n += 1
    return {"rows": n, "detail": f"{len(cols)} 列 × {n} 行设为「{step['format']}」"}


@op("set_width", "设置列宽", required=["columns", "width"])
def set_width(ctx, step):
    view = ctx.view(step)
    cols = view.cols(step["columns"])
    for c in cols:
        view.ws.column_dimensions[refs.col_letter(c)].width = float(step["width"])
    return {"detail": f"{len(cols)} 列宽度设为 {step['width']}"}
