"""校验：执行过程中检查数据，结果写进校验报告；可设置不通过时中止（不输出文件）"""
from __future__ import annotations

from .. import cond as C
from .. import refs
from ..book import is_blank
from ..errors import StepError
from . import col_cache, op, row_getter
from .match import key_text

CHECK_TYPES = ("required_columns", "not_empty", "unique", "row_count", "headers_match", "condition")


@op("validate", "校验", required=["checks"], writes=False)
def validate(ctx, step):
    """checks: [
         {"type": "required_columns", "columns": ["A", "B"]},
         {"type": "not_empty", "columns": ["工号"]},
         {"type": "unique", "columns": ["工号"]},
         {"type": "row_count", "min": 1, "max": 100000},
         {"type": "headers_match", "file": "ref", "columns": "A:AB"},
         {"type": "condition", "where": {...}, "expect": "none", "message": "不应有负库存"}
       ]  每项可加 "on_fail": "stop" 让不通过时中止"""
    view = ctx.view(step)
    checks = step["checks"]
    if not isinstance(checks, list) or not checks:
        raise StepError("checks 应是非空列表")
    passed, failed = 0, 0
    for i, ck in enumerate(checks, 1):
        t = ck.get("type")
        if t not in CHECK_TYPES:
            raise StepError(f"第 {i} 项校验类型「{t}」不支持，可用：{'、'.join(CHECK_TYPES)}")
        ok, msg = _run(ctx, view, ck)
        ctx.report.checks.append({"ok": ok, "name": ck.get("message") or _label(ck), "detail": msg})
        if ok:
            passed += 1
        else:
            failed += 1
            if ck.get("on_fail") == "stop":
                raise StepError(f"校验不通过，已中止：{ck.get('message') or _label(ck)}（{msg}）")
    return {"detail": f"{passed} 项通过，{failed} 项未通过", "failed": failed}


def _label(ck):
    t = ck["type"]
    cols = ck.get("columns")
    cols_txt = "、".join(map(str, cols)) if isinstance(cols, list) else str(cols or "")
    return {
        "required_columns": f"必须有列：{cols_txt}",
        "not_empty": f"不能为空：{cols_txt}",
        "unique": f"不能重复：{cols_txt}",
        "row_count": f"数据行数在 {ck.get('min', 0)}～{ck.get('max', '不限')} 之间",
        "headers_match": f"表头与「{ck.get('file', 'ref')}」一致：{cols_txt}",
        "condition": "条件检查",
    }[t]


def _run(ctx, view, ck):
    t = ck["type"]
    rows = view.data_rows()
    if t == "required_columns":
        missing = []
        for c in ck.get("columns", []):
            try:
                view.col(c)
            except StepError:
                missing.append(str(c))
        return (not missing, "缺少：" + "、".join(missing) if missing else "全部存在")
    if t == "not_empty":
        bad = []
        for cref in ck.get("columns", []):
            c = view.col(cref)
            n = sum(1 for r in rows if is_blank(view.get(r, c)))
            if n:
                first = next(r for r in rows if is_blank(view.get(r, c)))
                bad.append(f"「{cref}」有 {n} 个空值（如第 {first} 行）")
        return (not bad, "；".join(bad) or "无空值")
    if t == "unique":
        cols = [view.col(c) for c in ck.get("columns", [])]
        seen, dup = {}, []
        for r in rows:
            k = tuple(key_text(view.get(r, c)) for c in cols)
            if all(x == "" for x in k):
                continue
            if k in seen:
                dup.append((seen[k], r))
            else:
                seen[k] = r
        return (not dup, f"{len(dup)} 处重复（如第 {dup[0][0]} 行与第 {dup[0][1]} 行）" if dup else "无重复")
    if t == "row_count":
        n = len(rows)
        lo, hi = ck.get("min", 0), ck.get("max")
        ok = n >= lo and (hi is None or n <= hi)
        return (ok, f"实际 {n} 行")
    if t == "headers_match":
        other = ctx.file(ck.get("file", "ref")).view(ck.get("sheet"), ck.get("header_row"))
        spec = ck.get("columns")
        if isinstance(spec, str) and ":" in spec:
            a, b = spec.split(":", 1)
            cols = range(refs.col_index(a), refs.col_index(b) + 1)
        else:
            cols = range(1, max(view.max_col(), other.max_col()) + 1)
        diff = [f"{refs.col_letter(c)}列「{view.header_of(c)}」≠「{other.header_of(c)}」"
                for c in cols if view.header_of(c) != other.header_of(c)]
        return (not diff, "；".join(diff[:10]) + (f" 等 {len(diff)} 处" if len(diff) > 10 else "") if diff else "一致")
    # condition
    getter = row_getter(ctx, view, col_cache(view))
    hits = [r for r in rows if C.evaluate(ck["where"], getter(r))]
    expect = ck.get("expect", "none")
    if expect == "none":
        return (not hits, f"{len(hits)} 行符合（如第 {hits[0]} 行）" if hits else "没有符合的行")
    if expect == "all":
        return (len(hits) == len(rows), f"{len(rows) - len(hits)} 行不符合")
    return (bool(hits), f"{len(hits)} 行符合")
