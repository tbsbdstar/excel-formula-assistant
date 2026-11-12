"""多文件合并（追加行）与分组汇总"""
from __future__ import annotations

from pathlib import Path

from openpyxl.styles import Font, PatternFill

from .. import cond as C
from .. import refs
from ..book import copy_style, extend_autofilter, grow_rows, is_blank, norm_header
from ..errors import StepError
from . import col_cache, op, row_getter
from .rows import _sort_key

HEAD_FILL = PatternFill("solid", start_color="FFDDEBF7", end_color="FFDDEBF7")


def _qsheet(name: str) -> str:
    return "'" + str(name).replace("'", "''") + "'"


def _ensure_col(view, header, created=None):
    """主表没有这一列时加在最后（并纳入筛选范围），返回列号；created 列表用于告知调用方是新建的"""
    h = view.headers()
    name = norm_header(header)
    if name in h:
        return h[name]
    c = max(view.max_col(), max(h.values() or [0])) + 1
    extend_autofilter(view, c)
    view.set_value(view.header_row, c, header)
    view.invalidate()
    if created is not None:
        created.append(c)
    return c


# ================================================================ 追加行（多文件合并）
@op("append_rows", "追加行（多文件合并）", required=["from"],
    optional=["add_missing_columns", "source_column", "where", "extend_formulas", "skip_blank"])
def append_rows(ctx, step):
    """把一个或多个文件（“多个文件”）的数据行追加到主表末尾，按表头名称对应列。
    from: {"file": "stores", "sheet": "明细", "header_row": 1}"""
    src = step["from"]
    if isinstance(src, str):
        src = {"file": src}
    if not isinstance(src, dict) or not src.get("file"):
        raise StepError('from 应写成 {"file": "参照表代号", "sheet": "工作表（可选）"}')
    view = ctx.view(step)
    ws = view.ws
    sources = ctx.files_of(src["file"])
    if any(fh is ctx.main for fh in sources):
        raise StepError("不能把主表追加到它自己")
    first_data = view.header_row + 1
    old_last = view.last_row()
    template_row = old_last if old_last > view.header_row else None
    # 模板行（追加前的最后一行数据）：新行沿用它的格式；它里面的公式在来源没有对应值时向下延续
    template = {}
    if template_row:
        for (r, c), cell in list(ws._cells.items()):
            if r == template_row:
                template[c] = cell
    extend = step.get("extend_formulas", True)
    skip_blank = step.get("skip_blank", True)
    add_missing = step.get("add_missing_columns", False)
    src_col_name = step.get("source_column")
    created = []
    src_col = _ensure_col(view, src_col_name, created) if src_col_name else None
    if src_col and created and template_row:
        # 新建的来源列：主表原有的行也标上主表自己的文件名
        for r in range(first_data, old_last + 1):
            if any(not is_blank(ws._cells[(r, c)].value) for c in range(1, src_col) if (r, c) in ws._cells):
                view.set_value(r, src_col, Path(ctx.main.path).stem)
    next_row = old_last + 1
    per_file, skipped_cols = [], set()
    for fh in sources:
        sv = fh.view(src.get("sheet"), src.get("header_row"))
        mapping = []  # (来源列, 主表列)
        for name, sc in sv.headers().items():
            th = view.headers()
            if name in th:
                mapping.append((sc, th[name]))
            elif add_missing:
                mapping.append((sc, _ensure_col(view, name)))
            else:
                skipped_cols.add(name)
        if src_col and src_col_name in sv.headers():
            mapping = [(s, t) for s, t in mapping if t != src_col]
        getter = row_getter(ctx, sv, col_cache(sv))
        n = 0
        for r in sv.data_rows():
            vals = [(t, sv.get(r, s, ctx.warn), s) for s, t in mapping]
            if skip_blank and all(is_blank(v) for _, v, _ in vals):
                continue
            if "where" in step and not C.evaluate(step["where"], getter(r)):
                continue
            filled = set()
            for t, v, s in vals:
                tc = template.get(t)
                if tc is not None:
                    copy_style(tc, ws.cell(row=next_row, column=t), True)
                if is_blank(v):
                    continue
                cell = ws.cell(row=next_row, column=t)
                view.set_value(next_row, t, v)
                if isinstance(v, str) and v.startswith("="):
                    cell.data_type = "s"   # 来源里的文字，不要当成公式
                srcc = sv.cell(r, s)
                if srcc is not None and srcc.number_format not in (None, "General") and (tc is None or tc.number_format in (None, "General")):
                    cell.number_format = srcc.number_format
                filled.add(t)
            if extend and template_row:
                for t, tc in template.items():
                    if t in filled or t == src_col or tc.data_type != "f" or not isinstance(tc.value, str):
                        continue
                    copy_style(tc, ws.cell(row=next_row, column=t), True)
                    view.set_formula(next_row, t, refs.translate_formula(tc.value, next_row - template_row, 0))
            if src_col:
                view.set_value(next_row, src_col, Path(fh.path).stem)
            next_row += 1
            n += 1
        per_file.append((Path(fh.path).name, n))
    added = next_row - old_last - 1
    if added and old_last > view.header_row:
        grow_rows(ctx.main, view.name, first_data, old_last, added)
    view.invalidate()
    if skipped_cols:
        ctx.warn("来源中有主表没有的列，未合并：" + "、".join(sorted(skipped_cols)) + "（需要的话可勾选“自动添加主表没有的列”）")
    first_cells = [ws._cells.get((old_last, c)) for c in range(1, 4)]
    if any(c is not None and isinstance(c.value, str) and any(k in c.value for k in ("合计", "总计")) for c in first_cells):
        ctx.warn(f"主表追加前的最后一行（第 {old_last} 行）看起来是合计行，新数据排在了它下面。建议在追加之前先用“删除行”去掉合计行")
    detail = f"从 {len(per_file)} 个文件追加 {added} 行" + ("：" + "、".join(f"{n}（{k} 行）" for n, k in per_file) if len(per_file) <= 8 else "")
    return {"rows": added, "detail": detail}


# ================================================================ 分组汇总
FUNCS = {"sum": "合计", "count": "记录数", "avg": "平均", "max": "最大", "min": "最小", "count_distinct": "不重复个数"}


@op("group_summary", "分组汇总", required=["group_by", "metrics"],
    optional=["sheet_name", "replace", "mode", "total_row", "sort", "where"])
def group_summary(ctx, step):
    """按一列或多列分组，生成汇总工作表。
    metrics: [{"column": "金额", "func": "sum", "title": "金额合计"}, {"func": "count"}]
    mode: formula（默认，写 SUMIFS 等公式，数据变化后打开即更新）/ value（写计算结果）"""
    view = ctx.view(step)
    mode = step.get("mode", "formula")
    if mode not in ("formula", "value"):
        raise StepError("mode 只能是 formula（写公式）或 value（写结果值）")
    if mode == "formula" and "where" in step:
        raise StepError("按条件筛选（where）只能配合“写入结果值”使用；要用公式，请先用“删除行”或“只保留符合条件的行”")
    gcols = view.cols(step["group_by"], "分组列")
    metrics = step["metrics"]
    if not isinstance(metrics, list) or not metrics:
        raise StepError("metrics 应是非空列表，例如 [{\"column\": \"金额\", \"func\": \"sum\"}]")
    mdefs = []
    for i, m in enumerate(metrics, 1):
        if not isinstance(m, dict) or m.get("func") not in FUNCS:
            raise StepError(f"第 {i} 个汇总指标的 func 应是：{'、'.join(FUNCS)}")
        f = m["func"]
        if f == "count_distinct" and mode == "formula":
            raise StepError("不重复个数只能用“写入结果值”方式（公式在数据量大时会非常慢）")
        c = None
        if f != "count":
            if not m.get("column"):
                raise StepError(f"第 {i} 个汇总指标缺少要汇总的列")
            c = view.col(m["column"], "汇总列")
        title = m.get("title") or (f"{norm_header(m['column'])}{FUNCS[f]}" if c else FUNCS[f])
        mdefs.append({"func": f, "col": c, "title": title})
    sheet_name = str(step.get("sheet_name") or "分组汇总").strip()[:31]
    for ch in '[]:*?/\\':
        sheet_name = sheet_name.replace(ch, "_")
    wb = ctx.main.wb
    if sheet_name == view.name:
        raise StepError("汇总表不能和数据表同名")
    idx = None
    if sheet_name in wb.sheetnames:
        if not step.get("replace"):
            raise StepError(f"主表里已有工作表「{sheet_name}」。换个名字，或勾选“已有同名工作表时替换”")
        idx = wb.sheetnames.index(sheet_name)
        wb.remove(wb[sheet_name])

    rows = view.data_rows()
    getter = row_getter(ctx, view, col_cache(view))
    groups = {}
    for r in rows:
        if "where" in step and not C.evaluate(step["where"], getter(r)):
            continue
        key = tuple(_norm_key(view.get(r, c, ctx.warn)) for c in gcols)
        if all(k is None for k in key) and all(is_blank(view.get(r, c)) for c in [m["col"] for m in mdefs if m["col"]]):
            continue
        groups.setdefault(key, []).append(r)
    values = {k: [_agg(view, rs, m, ctx) for m in mdefs] for k, rs in groups.items()}
    keys = list(groups)
    if step.get("sort") == "metric_desc":
        keys.sort(key=lambda k: (values[k][0] is None, -(values[k][0] or 0)))
    else:
        keys.sort(key=lambda k: tuple(_sort_key(v, False) for v in k))

    out = wb.create_sheet(sheet_name, idx)
    nk = len(gcols)
    heads = [view.header_of(c) for c in gcols] + [m["title"] for m in mdefs]
    for j, h in enumerate(heads, 1):
        cell = out.cell(row=1, column=j, value=h)
        cell.font = Font(bold=True)
        cell.fill = HEAD_FILL
        out.column_dimensions[refs.col_letter(j)].width = max(10, min(40, len(str(h)) * 2 + 4))
    first, last = (rows[0], rows[-1]) if rows else (view.header_row + 1, view.header_row + 1)
    q = _qsheet(view.name)
    rng = lambda c: f"{q}!${refs.col_letter(c)}${first}:${refs.col_letter(c)}${last}"
    sample_fmt = {}
    for m in mdefs:
        if m["col"] and rows:
            cell = view.cell(rows[0], m["col"])
            sample_fmt[m["col"]] = cell.number_format if cell is not None else None
    for i, k in enumerate(keys, 2):
        crits = []
        for j, (v, gc) in enumerate(zip(k, gcols), 1):
            cell = out.cell(row=i, column=j, value=v if v is not None else "（空白）")
            src = view.cell(groups[k][0], gc)
            if src is not None and v is not None:
                cell.number_format = src.number_format
            crits.append((rng(gc), f"${refs.col_letter(j)}{i}" if v is not None else '""'))
        for j, m in enumerate(mdefs, nk + 1):
            cell = out.cell(row=i, column=j)
            if mode == "value":
                cell.value = values[k][j - nk - 1]
            else:
                cell.value = _formula(m, rng, crits)
                ctx.main.written_formulas.add(id(cell))
            if m["func"] in ("count", "count_distinct"):
                cell.number_format = "0"
            elif sample_fmt.get(m["col"]):
                cell.number_format = sample_fmt[m["col"]]
    end = len(keys) + 1
    if step.get("total_row", True) and keys:
        tr = end + 1
        out.cell(row=tr, column=1, value="合计").font = Font(bold=True)
        for j, m in enumerate(mdefs, nk + 1):
            L = refs.col_letter(j)
            cell = out.cell(row=tr, column=j)
            cell.font = Font(bold=True)
            if mode == "value":
                all_rows = [r for rs in groups.values() for r in rs]
                cell.value = _agg(view, all_rows, m, ctx)
            elif m["func"] in ("sum", "count"):
                cell.value = f"=SUM({L}2:{L}{end})"
                ctx.main.written_formulas.add(id(cell))
            elif m["func"] == "avg":
                cell.value = f'=IFERROR(AVERAGE({rng(m["col"])}),"")'
                ctx.main.written_formulas.add(id(cell))
            elif m["func"] in ("max", "min"):
                cell.value = f'={m["func"].upper()}({rng(m["col"])})'
                ctx.main.written_formulas.add(id(cell))
            if m["func"] in ("count", "count_distinct"):
                cell.number_format = "0"
            elif sample_fmt.get(m["col"]):
                cell.number_format = sample_fmt[m["col"]]
    out.freeze_panes = "A2"
    if not keys:
        ctx.warn("没有可汇总的数据行，汇总表只有表头")
    return {"rows": len(keys), "detail": f"生成工作表「{sheet_name}」：按 {'、'.join(heads[:nk])} 分 {len(keys)} 组，"
                                         f"{len(mdefs)} 个指标（{'公式' if mode == 'formula' else '结果值'}）"}


def _norm_key(v):
    if is_blank(v):
        return None
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _formula(m, rng, crits):
    f = m["func"]
    if f == "count":
        return "=COUNTIFS(" + ",".join(f"{a},{b}" for a, b in crits) + ")"
    val = rng(m["col"])
    if f == "sum":
        return f"=SUMIFS({val}," + ",".join(f"{a},{b}" for a, b in crits) + ")"
    if f == "avg":
        return f'=IFERROR(AVERAGEIFS({val},' + ",".join(f"{a},{b}" for a, b in crits) + '),"")'
    # INDEX(…,0) 让 Excel 按数组计算：直接写进文件的 AGGREGATE(14,6,区域/(条件),1) 打开后是 #VALUE!（实测）
    k = 14 if f == "max" else 15
    mask = "*".join(f"({a}={b})" for a, b in crits)
    return f'=IFERROR(AGGREGATE({k},6,INDEX({val}/({mask}),0),1),"")'


def _agg(view, rows, m, ctx):
    f = m["func"]
    if f == "count":
        return len(rows)
    vals = [view.get(r, m["col"], ctx.warn) for r in rows]
    if f == "count_distinct":
        return len({_norm_key(v) for v in vals if not is_blank(v)})
    nums = [x for x in (C._num(v) for v in vals) if x is not None]
    if not nums:
        return None
    if f == "sum":
        return round(sum(nums), 10)
    if f == "avg":
        return sum(nums) / len(nums)
    return max(nums) if f == "max" else min(nums)
