"""原子操作注册表与公用工具。

新增操作：在本目录（或工具包的 extensions/ 目录）新建 .py 文件，用 @op 装饰器注册一个函数：

    from engine.ops import op

    @op("my_op", "我的操作", required=["column"], optional=["value"])
    def my_op(ctx, step):
        view = ctx.view(step)            # 目标工作表视图
        col = view.col(step["column"])   # 按表头定位列
        ...
        return {"rows": 处理行数, "detail": "给报告看的一句话"}
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List

from openpyxl.styles import Font, PatternFill

from .. import refs
from ..errors import ConfigError, StepError

COMMON_KEYS = {"op", "title", "file", "sheet", "header_row", "enabled", "note"}


@dataclass
class OpDef:
    name: str
    label: str
    fn: Callable
    required: List[str] = field(default_factory=list)   # "a|b" 表示二选一
    optional: List[str] = field(default_factory=list)
    writes: bool = True                                  # 是否修改主表


REGISTRY: Dict[str, OpDef] = {}


def op(name, label, required=(), optional=(), aliases=(), writes=True):
    def deco(fn):
        d = OpDef(name, label, fn, list(required), list(optional), writes)
        REGISTRY[name] = d
        for a in aliases:
            REGISTRY[a] = d
        return fn
    return deco


def validate_step(step: dict, idx: int):
    where = f"第 {idx} 步"
    if not isinstance(step, dict):
        raise ConfigError(f"{where}应是一个对象")
    name = step.get("op")
    if name not in REGISTRY:
        raise ConfigError(f"{where}的操作「{name}」不存在。可用操作：{'、'.join(sorted(k for k in REGISTRY))}")
    d = REGISTRY[name]
    for req in d.required:
        if not any(k in step for k in req.split("|")):
            alts = " 或 ".join(req.split("|"))
            raise ConfigError(f"{where}（{d.label}）缺少参数 {alts}")
    allowed = COMMON_KEYS | {k for r in d.required for k in r.split("|")} | set(d.optional)
    unknown = [k for k in step if k not in allowed]
    if unknown:
        raise ConfigError(f"{where}（{d.label}）有不认识的参数：{'、'.join(unknown)}。可用参数：{'、'.join(sorted(allowed - COMMON_KEYS))}")
    from .. import cond
    for key in ("where", "ref_where"):
        if key in step:
            cond.validate(step[key], f"{where}的 {key}")
    for key in ("fill", "font_color"):
        if key in step and step[key] not in (None, ""):
            parse_color(step[key], f"{where}的 {key}")


# ---------------------------------------------------------------- 颜色
def parse_color(v, where="颜色") -> str:
    s = str(v).strip().lstrip("#").upper()
    named = {"红": "FFC7CE", "黄": "FFEB9C", "绿": "C6EFCE", "蓝": "BDD7EE", "橙": "FCE4D6", "灰": "D9D9D9"}
    s = named.get(s, s)
    if re.fullmatch(r"[0-9A-F]{6}", s):
        return "FF" + s
    if re.fullmatch(r"[0-9A-F]{8}", s):
        return s
    raise ConfigError(f"{where}「{v}」不是合法的颜色，应为 6 位十六进制，如 FCE4D6")


def solid_fill(color: str) -> PatternFill:
    c = parse_color(color)
    return PatternFill(fill_type="solid", start_color=c, end_color=c)


def font_with(cell, color=None, bold=None):
    f = cell.font
    return Font(name=f.name, size=f.size, bold=f.bold if bold is None else bool(bold), italic=f.italic,
                underline=f.underline, strike=f.strike, vertAlign=f.vertAlign,
                color=parse_color(color) if color else f.color)


# ---------------------------------------------------------------- 写入内容（固定值 / 公式模板 / 复制公式 / 取另一列）
_TPL_ROW = re.compile(r"\[@([^\[\]]+)\]")
_TPL_RNG = re.compile(r"\[([^@\[\]][^\[\]]*)\]")


def compile_content(ctx, step, view, target_col: int):
    """返回 fn(row) → 要写入的值或公式；None 表示步骤没有给内容"""
    keys = [k for k in ("value", "formula", "formula_from", "from_col") if k in step]
    if len(keys) > 1:
        raise StepError(f"value / formula / formula_from / from_col 只能选一个，现在给了：{'、'.join(keys)}")
    if not keys:
        return None
    k = keys[0]
    if k == "value":
        v = step["value"]
        return lambda r: v
    if k == "from_col":
        src = view.col(step["from_col"])
        return lambda r: view.get(r, src, ctx.warn)
    if k == "formula":
        tpl = str(step["formula"]).strip()
        if not tpl.startswith("="):
            tpl = "=" + tpl
        rows = view.data_rows()
        first, last = (rows[0], rows[-1]) if rows else (view.header_row + 1, view.header_row + 1)

        def rng(m):
            c = view.col(m.group(1).strip())
            L = refs.col_letter(c)
            return f"${L}${first}:${L}${last}"

        tpl = _TPL_RNG.sub(rng, tpl)
        tpl = _TPL_ROW.sub(lambda m: refs.col_letter(view.col(m.group(1).strip())) + "\x00", tpl)
        parts = tpl.replace("{row}", "\x00").split("\x00")
        return lambda r: str(r).join(parts)
    # formula_from：复制某个单元格的公式，并像 Excel 粘贴一样平移相对引用
    spec = step["formula_from"]
    if not isinstance(spec, dict) or "cell" not in spec:
        raise StepError('formula_from 应写成 {"file": "ref", "sheet": "工作表名", "cell": "P2"}')
    fh = ctx.file(spec.get("file", step.get("file", "main")))
    sheet = fh.resolve_sheet(spec.get("sheet", step.get("sheet") if fh is ctx.main else None))
    sc, sr = refs.parse_cell(spec["cell"])
    src_ws = fh.formula_book()[sheet]
    cell = src_ws._cells.get((sr, sc))
    src = cell.value if cell is not None else None
    if not (isinstance(src, str) and src.startswith("=")):
        return lambda r: src
    if refs.has_external_ref(src):
        raise StepError(f"{spec['cell']} 的公式引用了其他文件，执行器不支持跨文件公式")
    if fh is not ctx.main:
        missing = [s for s in refs.explicit_sheets(src) if s not in ctx.main.wb.sheetnames]
        if missing:
            raise StepError(f"{spec['cell']} 的公式引用了主表中不存在的工作表：{'、'.join(missing)}。可先用 copy_sheet 把它复制进主表")
    return lambda r: refs.translate_formula(src, r - sr, target_col - sc)


def write_content(view, r, c, v):
    if isinstance(v, str) and v.startswith("="):
        return view.set_formula(r, c, v)
    return view.set_value(r, c, v)


def col_cache(view):
    """按列名定位并缓存（一步之内表结构不变）"""
    cache = {}

    def colof(ref):
        key = ref if isinstance(ref, (str, int)) else json.dumps(ref, sort_keys=True, ensure_ascii=False)
        if key not in cache:
            cache[key] = view.col(ref)
        return cache[key]

    return colof


def row_getter(ctx, view, colof=None):
    """返回 for_row(r) → get(列) 函数，供条件判断取当前行的值"""
    colof = colof or col_cache(view)

    def for_row(r):
        return lambda ref: view.get(r, colof(ref), ctx.warn)

    return for_row


def load_all():
    """加载内置操作模块"""
    from . import columns, rows, fill, style, match, copyops, check, merge  # noqa: F401
