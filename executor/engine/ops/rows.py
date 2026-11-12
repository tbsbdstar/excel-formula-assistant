"""行操作：删除行、保留行、插入行、去重、排序"""
from __future__ import annotations

import datetime as dt

from .. import cond as C
from .. import refs
from ..book import check_tables, is_blank, norm_header, restructure
from ..errors import StepError
from . import col_cache, op, row_getter


def _delete(ctx, view, rows):
    if not rows:
        return
    check_tables(view.ws, at_rows=rows)
    restructure(ctx.main, view.name, None, refs.IndexMap(deleted=rows))


@op("delete_rows", "删除行", required=["where|rows|blank"])
def delete_rows(ctx, step):
    """where: 条件；rows: "5:10"；blank: true 删除整行为空的行"""
    view = ctx.view(step)
    targets = []
    if "rows" in step:
        targets = view.data_rows(step["rows"])
    elif step.get("blank"):
        cols = list(view.headers().values())
        targets = [r for r in view.data_rows() if all(is_blank(view.get(r, c)) for c in cols)]
    else:
        getter = row_getter(ctx, view)
        targets = [r for r in view.data_rows() if C.evaluate(step["where"], getter(r))]
    _delete(ctx, view, targets)
    return {"rows": len(targets), "detail": f"删除 {len(targets)} 行" if targets else "没有符合条件的行"}


@op("keep_rows", "只保留符合条件的行", required=["where"])
def keep_rows(ctx, step):
    view = ctx.view(step)
    getter = row_getter(ctx, view)
    targets = [r for r in view.data_rows() if not C.evaluate(step["where"], getter(r))]
    _delete(ctx, view, targets)
    return {"rows": len(targets), "detail": f"删除不符合条件的 {len(targets)} 行"}


@op("insert_rows", "插入空行", required=["before_row|at_end"], optional=["count"])
def insert_rows(ctx, step):
    view = ctx.view(step)
    n = int(step.get("count", 1))
    if n < 1:
        raise StepError("count 至少为 1")
    if step.get("at_end"):
        at = view.last_row() + 1
        return {"rows": n, "detail": f"在第 {at} 行起预留 {n} 行（末尾无需移动）"}
    at = int(step["before_row"])
    check_tables(view.ws, at_rows=[at])
    restructure(ctx.main, view.name, None, refs.IndexMap(insert_at=at, count=n))
    return {"rows": n, "detail": f"在第 {at} 行前插入 {n} 行"}


@op("dedupe_rows", "删除重复行", optional=["keys", "keep"])
def dedupe_rows(ctx, step):
    view = ctx.view(step)
    keys = view.cols(step["keys"]) if step.get("keys") else list(view.headers().values())
    keep = step.get("keep", "first")
    if keep not in ("first", "last"):
        raise StepError("keep 只能是 first（保留第一次出现）或 last（保留最后一次）")
    rows = view.data_rows()
    order = rows if keep == "first" else list(reversed(rows))
    seen, dup = set(), []
    for r in order:
        k = tuple(_key_text(view.get(r, c)) for c in keys)
        if k in seen:
            dup.append(r)
        else:
            seen.add(k)
    _delete(ctx, view, dup)
    return {"rows": len(dup), "detail": f"按 {len(keys)} 列判断重复，删除 {len(dup)} 行"}


def _key_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    return str(v).strip()


def _sort_key(v, desc):
    # 与 Excel 一致：数字 < 文本 < 逻辑值；空值始终排最后
    if is_blank(v):
        return (1, 0, 0)
    if isinstance(v, bool):
        rank, val = 3, int(v)
    elif isinstance(v, (int, float)):
        rank, val = 1, float(v)
    elif isinstance(v, (dt.datetime, dt.date)):
        rank, val = 1, (v if isinstance(v, dt.datetime) else dt.datetime(v.year, v.month, v.day)).timestamp() / 86400
    else:
        rank, val = 2, str(v).lower()
    if desc:
        rank = -rank
        val = -val if isinstance(val, float) else _Rev(val)
    return (0, rank, val)


class _Rev:
    __slots__ = ("v",)

    def __init__(self, v):
        self.v = v

    def __lt__(self, o):
        return self.v > o.v

    def __eq__(self, o):
        return self.v == o.v


@op("sort_rows", "排序", required=["by"])
def sort_rows(ctx, step):
    """by: [{"column": "金额", "order": "desc"}, {"column": "日期"}]"""
    view = ctx.view(step)
    ws = view.ws
    by = step["by"] if isinstance(step["by"], list) else [step["by"]]
    keys = []
    for b in by:
        if isinstance(b, (str, dict)) and not (isinstance(b, dict) and "column" in b):
            b = {"column": b}
        order = b.get("order", "asc")
        if order not in ("asc", "desc"):
            raise StepError("order 只能是 asc（升序）或 desc（降序）")
        keys.append((view.col(b["column"]), order == "desc"))
    rows = view.data_rows()
    if not rows:
        return {"detail": "没有数据行"}
    first, last = rows[0], rows[-1]
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= last and mr.max_row >= first:
            raise StepError(f"排序区域内有合并单元格（{mr.coord}），请先取消合并")
    check_tables(ws, at_rows=rows)
    decorated = []
    for r in rows:
        decorated.append((tuple(_sort_key(view.get(r, c), d) for c, d in keys), r))
    decorated.sort(key=lambda x: x[0])  # 稳定排序
    new_of = {old: first + i for i, (_, old) in enumerate(decorated)}
    moved = {}
    for (r, c), cell in list(ws._cells.items()):
        if first <= r <= last:
            moved[(r, c)] = cell
            del ws._cells[(r, c)]
    for (r, c), cell in moved.items():
        nr = new_of[r]
        cell.row = nr
        if cell.data_type == "f" and isinstance(cell.value, str) and nr != r:
            cell.value = refs.translate_formula(cell.value, nr - r, 0)
            ctx.main.cached.pop(id(cell), None)  # 公式已变，原缓存结果不再可靠
        ws._cells[(nr, c)] = cell
    heights = {r: ws.row_dimensions[r] for r in rows if r in ws.row_dimensions}
    for r, d in heights.items():
        del ws.row_dimensions[r]
    for r, d in heights.items():
        d.index = new_of[r]
        ws.row_dimensions[new_of[r]] = d
    view.invalidate()
    desc = "、".join(f"{norm_header(view.header_of(c)) or refs.col_letter(c)}{'降序' if d else '升序'}" for c, d in keys)
    return {"rows": len(rows), "detail": f"按 {desc} 排序 {len(rows)} 行"}
