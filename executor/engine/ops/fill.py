"""填充与文本处理：整列/按条件写值或公式、下拉填充公式、查找替换"""
from __future__ import annotations

import re

from .. import cond as C
from .. import refs
from ..book import copy_style, is_blank
from ..errors import StepError
from . import compile_content, op, row_getter, write_content


@op("fill", "填充（固定值/公式/条件写入）", required=["column", "value|formula|formula_from|from_col"],
    optional=["where", "rows", "only_empty"], aliases=["write_if"])
def fill(ctx, step):
    """在某列写入内容。
    value: 固定值；formula: 公式模板（可用 {row}、[@列名]、[列名]）；
    formula_from: 复制某单元格的公式并平移；from_col: 取同一行另一列的值。
    where: 只处理满足条件的行；only_empty: 只填空单元格。"""
    view = ctx.view(step)
    c = view.col(step["column"])
    content = compile_content(ctx, step, view, c)
    getter = row_getter(ctx, view)
    n = 0
    for r in view.data_rows(step.get("rows")):
        if "where" in step and not C.evaluate(step["where"], getter(r)):
            continue
        if step.get("only_empty"):
            cell = view.cell(r, c)
            if cell is not None and not is_blank(cell.value):
                continue
        if write_content(view, r, c, content(r)):
            n += 1
    return {"rows": n, "detail": f"「{view.header_of(c) or refs.col_letter(c)}」写入 {n} 行"}


@op("fill_down", "向下填充公式", required=["columns"], optional=["from_row", "to_row", "with_style"])
def fill_down(ctx, step):
    """把 from_row（默认第一行数据）那一行的公式/值，像 Excel 下拉一样填到最后一行数据"""
    view = ctx.view(step)
    cols = view.cols(step["columns"])
    rows = view.data_rows()
    if not rows:
        return {"detail": "没有数据行"}
    src_r = int(step.get("from_row", rows[0]))
    end_r = int(step.get("to_row", rows[-1]))
    ws = view.ws
    n = 0
    for c in cols:
        src = ws._cells.get((src_r, c))
        if src is None:
            continue
        v = src.value
        for r in range(src_r + 1, end_r + 1):
            if isinstance(v, str) and v.startswith("="):
                view.set_formula(r, c, refs.translate_formula(v, r - src_r, 0))
            else:
                view.set_value(r, c, v)
            if step.get("with_style", True) and src.has_style:
                copy_style(src, ws.cell(row=r, column=c), True)
            n += 1
    return {"cells": n, "detail": f"{len(cols)} 列从第 {src_r} 行向下填充到第 {end_r} 行"}


@op("replace_text", "查找替换", required=["columns", "find"],
    optional=["replace", "match_case", "whole_cell", "regex", "where"])
def replace_text(ctx, step):
    view = ctx.view(step)
    cols = view.cols(step["columns"])
    find, rep = str(step["find"]), str(step.get("replace", ""))
    flags = 0 if step.get("match_case") else re.I
    if step.get("regex"):
        try:
            pat = re.compile(find, flags)
        except re.error as e:
            raise StepError(f"正则表达式有误：{e}")
    else:
        pat = re.compile(re.escape(find), flags)
    getter = row_getter(ctx, view)
    n = 0
    for r in view.data_rows():
        if "where" in step and not C.evaluate(step["where"], getter(r)):
            continue
        for c in cols:
            cell = view.cell(r, c)
            if cell is None or cell.data_type == "f" or not isinstance(cell.value, str):
                continue
            v = cell.value
            if step.get("whole_cell"):
                if pat.fullmatch(v):
                    view.set_value(r, c, rep)
                    n += 1
            else:
                nv = pat.sub(rep, v)
                if nv != v:
                    view.set_value(r, c, nv)
                    n += 1
    return {"cells": n, "detail": f"替换 {n} 个单元格"}
