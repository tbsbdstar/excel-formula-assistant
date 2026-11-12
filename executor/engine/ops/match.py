"""跨表匹配：建立匹配（单键/多键）、按匹配复制值或公式、比对差异"""
from __future__ import annotations

import datetime as dt

from openpyxl.styles import Font

from .. import cond as C
from .. import refs
from ..book import copy_style, extend_autofilter, is_blank
from ..errors import StepError
from . import col_cache, op, parse_color, row_getter, solid_fill, write_content


def key_text(v, ignore_case=False, trim=True):
    if v is None:
        s = ""
    elif isinstance(v, bool):
        s = "TRUE" if v else "FALSE"
    elif isinstance(v, float) and v.is_integer():
        s = str(int(v))
    elif isinstance(v, (int, float)):
        s = repr(v)
    elif isinstance(v, dt.datetime):
        s = v.date().isoformat() if (v.hour, v.minute, v.second) == (0, 0, 0) else v.isoformat(sep=" ")
    elif isinstance(v, dt.date):
        s = v.isoformat()
    else:
        s = str(v)
    if trim:
        s = s.strip()
    return s.casefold() if ignore_case else s


class MatchIndex:
    def __init__(self, mid, ref_view, main_keys, ref_keys, opts):
        self.id = mid
        self.ref_view = ref_view
        self.main_keys = main_keys
        self.ref_keys = ref_keys
        self.trim = opts.get("trim", True)
        self.ignore_case = opts.get("ignore_case", False)
        self.index = {}
        self.dups = 0
        self.ref_rows = []

    def key(self, values):
        return tuple(key_text(v, self.ignore_case, self.trim) for v in values)

    def build(self, ctx, duplicates, ref_where):
        rv = self.ref_view
        cols = [rv.col(k, "参照表键列") for k in self.ref_keys]
        getter = row_getter(ctx, rv)
        for r in rv.data_rows():
            if ref_where and not C.evaluate(ref_where, getter(r)):
                continue
            k = self.key([rv.get(r, c) for c in cols])
            if all(x == "" for x in k):
                continue
            self.ref_rows.append(r)
            if k in self.index:
                self.dups += 1
                if duplicates == "error":
                    raise StepError(f"参照表中键 {'+'.join(k)} 重复（第 {self.index[k]} 行和第 {r} 行），请检查数据或改用 duplicates: first/last")
                if duplicates == "last":
                    self.index[k] = r
                continue
            self.index[k] = r

    def main_key(self, view, r, colof):
        return self.key([view.get(r, colof(k)) for k in self.main_keys])


@op("match", "建立匹配（单键/多键）", required=["id", "ref", "main_keys"],
    optional=["ref_keys", "ref_sheet", "ref_header_row", "trim", "ignore_case", "duplicates", "ref_where"], writes=False)
def match(ctx, step):
    """main_keys: ["详情ID", "采购单", "条码"]；ref_keys 不写时与 main_keys 同名"""
    view = ctx.view(step)
    ref_fh = ctx.file(step["ref"])
    ref_view = ref_fh.view(step.get("ref_sheet"), step.get("ref_header_row"))
    main_keys = step["main_keys"] if isinstance(step["main_keys"], list) else [step["main_keys"]]
    ref_keys = step.get("ref_keys", main_keys)
    ref_keys = ref_keys if isinstance(ref_keys, list) else [ref_keys]
    if len(ref_keys) != len(main_keys):
        raise StepError(f"main_keys 有 {len(main_keys)} 列，ref_keys 有 {len(ref_keys)} 列，数量要一致")
    duplicates = step.get("duplicates", "first")
    if duplicates not in ("first", "last", "error"):
        raise StepError("duplicates 只能是 first / last / error")
    mi = MatchIndex(step["id"], ref_view, main_keys, ref_keys, step)
    mi.build(ctx, duplicates, step.get("ref_where"))
    colof = col_cache(view)
    for k in main_keys:
        colof(k)  # 提前检查主表键列
    hit, miss, blank = 0, 0, 0
    used = set()
    for r in view.data_rows():
        k = mi.main_key(view, r, colof)
        if all(x == "" for x in k):
            blank += 1
            continue
        rr = mi.index.get(k)
        if rr is None:
            miss += 1
        else:
            hit += 1
            used.add(rr)
    rk = [ref_view.col(x) for x in ref_keys]
    ref_only = sum(1 for r in mi.ref_rows if r not in used and mi.index.get(mi.key([ref_view.get(r, c) for c in rk])) == r)
    ctx.matches[step["id"]] = mi
    stats = {"主表匹配成功": hit, "仅主表有": miss, "仅参照表有": ref_only, "主表键为空": blank, "参照表重复键": mi.dups}
    ctx.report.match_stats[step["id"]] = stats
    detail = f"按 {'+'.join(map(str, main_keys))} 匹配「{ref_fh.label}」：成功 {hit} 行，仅主表有 {miss} 行，仅参照表有 {ref_only} 行"
    if mi.dups:
        detail += f"，参照表有 {mi.dups} 个重复键（取{'第一' if duplicates == 'first' else '最后'}次出现）"
        ctx.warn(f"匹配「{step['id']}」：参照表有 {mi.dups} 个重复键")
    return {"rows": hit, "detail": detail}


def _pairs(step):
    if "columns" in step:
        out = []
        for p in step["columns"]:
            if isinstance(p, dict):
                out.append((p["from"], p.get("to", p["from"])))
            else:
                out.append((p, p))
        return out
    fr, to = step.get("from"), step.get("to")
    fr = fr if isinstance(fr, list) else [fr]
    to = to if isinstance(to, list) else [to]
    if len(fr) != len(to):
        raise StepError("from 和 to 的列数要一致")
    return list(zip(fr, to))


def get_match(ctx, step):
    mid = step["match"]
    if mid not in ctx.matches:
        raise StepError(f"找不到匹配「{mid}」，请在这一步之前先添加“建立匹配”步骤（id 为 {mid}）")
    return ctx.matches[mid]


@op("copy_matched", "按匹配复制（值或公式）", required=["match", "columns|from"],
    optional=["to", "mode", "where", "skip_empty", "only_empty", "fill", "color_columns", "not_found",
              "copy_format", "create_missing"])
def copy_matched(ctx, step):
    view = ctx.view(step)
    mi = get_match(ctx, step)
    rv = mi.ref_view
    mode = step.get("mode", "value")
    if mode not in ("value", "formula"):
        raise StepError("mode 只能是 value（复制值）或 formula（复制公式）")
    pairs = []
    for fr, to in _pairs(step):
        sc = rv.col(fr, "参照表列")
        try:
            tc = view.col(to)
        except StepError:
            if not step.get("create_missing"):
                raise StepError(f"主表里没有列「{to}」。可先用“插入列”添加，或在这一步加 \"create_missing\": true 自动加在最后")
            tc = max(view.max_col(), max(view.headers().values() or [0])) + 1
            extend_autofilter(view, tc)
            view.set_value(view.header_row, tc, to)
            view.invalidate()
        pairs.append((sc, tc))
    color_cols = view.cols(step["color_columns"]) if step.get("color_columns") else [tc for _, tc in pairs]
    fill = solid_fill(step["fill"]) if step.get("fill") else None
    nf = step.get("not_found")
    nf_col = view.col(nf["column"]) if nf else None
    skip_empty = step.get("skip_empty", True)
    colof = col_cache(view)
    getter = row_getter(ctx, view, colof)
    ref_colof = col_cache(rv)
    ref_ws_f = rv.fws if mode == "formula" else None
    main_sheets = ctx.main.wb.sheetnames
    same_book = rv.fh is ctx.main
    copied, rows_hit, fallback = 0, 0, 0
    for r in view.data_rows():
        rr = mi.index.get(mi.main_key(view, r, colof))
        if rr is None:
            if nf_col:
                write_content(view, r, nf_col, nf.get("value", "未找到"))
            continue
        get_ref = (lambda rr_: (lambda ref: rv.get(rr_, ref_colof(ref), ctx.warn)))(rr)
        if "where" in step and not C.evaluate(step["where"], getter(r), get_ref):
            continue
        rows_hit += 1
        for sc, tc in pairs:
            if step.get("only_empty"):
                cur = view.cell(r, tc)
                if cur is not None and not is_blank(cur.value):
                    continue
            val = rv.get(rr, sc, ctx.warn)
            if mode == "formula":
                src = ref_ws_f._cells.get((rr, sc))
                f = src.value if src is not None else None
                if isinstance(f, str) and f.startswith("="):
                    bad = refs.has_external_ref(f) or (not same_book and any(s not in main_sheets for s in refs.explicit_sheets(f)))
                    if bad:
                        fallback += 1
                    else:
                        val = refs.translate_formula(f, r - rr, tc - sc)
                elif f is not None:
                    val = f
            if is_blank(val) and skip_empty:
                continue
            write_content(view, r, tc, val)
            if step.get("copy_format"):
                srcc = rv.cell(rr, sc) if mode == "value" else ref_ws_f._cells.get((rr, sc))
                if srcc is not None:
                    view.ws.cell(row=r, column=tc).number_format = srcc.number_format
            copied += 1
        if fill is not None:
            for c in color_cols:
                view.ws.cell(row=r, column=c).fill = fill
    if fallback:
        ctx.warn(f"{fallback} 个公式引用了主表中不存在的工作表或其他文件，已改为复制计算结果")
    return {"rows": rows_hit, "cells": copied,
            "detail": f"{rows_hit} 行匹配且满足条件，复制 {copied} 个单元格（{'公式' if mode == 'formula' else '值'}）" + (f"并着色 {len(color_cols)} 列" if fill else "")}


@op("compare", "比对两表差异", required=["match"],
    optional=["status_column", "labels", "compare_columns", "tolerance", "diff_column", "highlight", "report_sheet", "export_csv"])
def compare(ctx, step):
    view = ctx.view(step)
    mi = get_match(ctx, step)
    rv = mi.ref_view
    labels = {"same": "一致", "diff": "有差异", "main_only": "仅主表有"}
    labels.update(step.get("labels") or {})
    tol = float(step.get("tolerance", 0) or 0)
    colof = col_cache(view)
    # 要逐列比较的列
    spec = step.get("compare_columns", "same_name")
    pairs = []
    if spec == "same_name":
        keys = {str(k) for k in mi.main_keys}
        for name, c in view.headers().items():
            if name in keys:
                continue
            if name in rv.headers():
                pairs.append((name, c, rv.headers()[name]))
    else:
        for p in spec:
            a, b = (p.get("main"), p.get("ref", p.get("main"))) if isinstance(p, dict) else (p, p)
            pairs.append((str(a), view.col(a), rv.col(b, "参照表列")))

    def ensure_col(header):
        try:
            return view.col(header)
        except StepError:
            c = max(view.max_col(), max(view.headers().values() or [0])) + 1
            extend_autofilter(view, c)
            view.set_value(view.header_row, c, header)
            view.invalidate()
            return c

    status_c = ensure_col(step.get("status_column") or "比对结果")
    diff_name = step.get("diff_column", "差异列")
    diff_c = ensure_col(diff_name) if diff_name else None
    hl = step.get("highlight", "FFEB9C")
    hl_fill = solid_fill(hl) if hl else None
    n_same = n_diff = n_main_only = 0
    used = set()
    details = []
    for r in view.data_rows():
        k = mi.main_key(view, r, colof)
        if all(x == "" for x in k):
            continue
        rr = mi.index.get(k)
        if rr is None:
            view.set_value(r, status_c, labels["main_only"])
            n_main_only += 1
            continue
        used.add(rr)
        diffs = []
        for name, mc, rc in pairs:
            a, b = view.get(r, mc, ctx.warn), rv.get(rr, rc, ctx.warn)
            if _same(a, b, tol):
                continue
            diffs.append(name)
            details.append([r, "+".join(k), name, _show(a), _show(b)])
            if hl_fill is not None:
                view.ws.cell(row=r, column=mc).fill = hl_fill
        if diffs:
            n_diff += 1
            view.set_value(r, status_c, labels["diff"])
        else:
            n_same += 1
            view.set_value(r, status_c, labels["same"])
        if diff_c:
            view.set_value(r, diff_c, "、".join(diffs) if diffs else None)
    ref_key_cols = [rv.col(k) for k in mi.ref_keys]
    ref_only = [r for r in mi.ref_rows if r not in used and mi.index.get(mi.key([rv.get(r, c) for c in ref_key_cols])) == r]
    stats = {"一致": n_same, "有差异": n_diff, "仅主表有": n_main_only, "仅参照表有": len(ref_only), "差异单元格": len(details)}
    ctx.report.match_stats[f"{step['match']}（比对）"] = stats
    sheet_name = step.get("report_sheet", "比对报告")
    if sheet_name:
        _write_report_sheet(ctx, sheet_name, stats, details, ref_only, rv, mi)
    if step.get("export_csv"):
        ctx.extra_csv.append(("比对差异明细", [["主表行号", "键", "列", "主表值", "参照表值"]] + details))
        ctx.extra_csv.append(("仅参照表有", [list(rv.headers().keys())] + [[_show(rv.get(r, c)) for c in rv.headers().values()] for r in ref_only]))
    return {"rows": n_same + n_diff + n_main_only,
            "detail": f"一致 {n_same} 行、有差异 {n_diff} 行（{len(details)} 个单元格）、仅主表有 {n_main_only} 行、仅参照表有 {len(ref_only)} 行"}


def _same(a, b, tol):
    if is_blank(a) and is_blank(b):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(a - b) <= tol + 1e-9
    return C._eq(a, b, False)


def _show(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _write_report_sheet(ctx, name, stats, details, ref_only, rv, mi):
    wb = ctx.main.wb
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    bold = Font(bold=True)
    head = solid_fill("E8F3EC")
    ws.append(["比对汇总"])
    ws["A1"].font = Font(bold=True, size=13)
    for k, v in stats.items():
        ws.append([k, v])
    ws.append([])
    row = ws.max_row + 1
    ws.cell(row=row, column=1, value=f"差异明细（共 {len(details)} 处{'，只列前 5000 处' if len(details) > 5000 else ''}）").font = bold
    ws.append(["主表行号", "键", "列", "主表值", "参照表值"])
    for c in range(1, 6):
        ws.cell(row=ws.max_row, column=c).fill = head
    for d in details[:5000]:
        ws.append(d)
    ws.append([])
    ws.cell(row=ws.max_row + 1, column=1, value=f"仅参照表有（共 {len(ref_only)} 行{'，只列前 5000 行' if len(ref_only) > 5000 else ''}）").font = bold
    hdr = list(rv.headers().items())[:30]
    ws.append(["参照表行号"] + [h for h, _ in hdr])
    for c in range(1, len(hdr) + 2):
        ws.cell(row=ws.max_row, column=c).fill = head
    for r in ref_only[:5000]:
        ws.append([r] + [_show(rv.get(r, c)) for _, c in hdr])
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
