"""单元格引用工具：列字母换算、区域解析、公式中引用的平移与结构调整。

openpyxl 插入/删除行列时不会像 Excel 一样调整公式里的引用，
这里用正则扫描公式（跳过字符串常量），按映射函数改写引用。
"""
from __future__ import annotations

import bisect
import re
from typing import Callable, Optional

from openpyxl.utils import column_index_from_string, get_column_letter

MAX_COL = 16384
MAX_ROW = 1048576

# 公式里的引用：可选工作表前缀 + A1 / A1:B2 / A:B / 1:2
_SHEET = r"(?:'(?:[^']|'')+'|[^\s'!(),;=+\-*/^&<>\"{}:\[\]]+)!"
REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.$\]一-鿿])"
    r"(?P<sheet>" + _SHEET + r")?"
    r"(?P<ref>"
    r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?"
    r"|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"
    r"|\$?\d+:\$?\d+"
    r")"
    r"(?![A-Za-z0-9_(.!一-鿿])"
)
_CELL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_COL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^(\$?)(\d+)$")


def col_letter(idx: int) -> str:
    return get_column_letter(idx)


def col_index(letters: str) -> int:
    return column_index_from_string(letters.upper())


def unquote_sheet(s: str) -> str:
    s = s[:-1] if s.endswith("!") else s
    if s.startswith("'") and s.endswith("'"):
        s = s[1:-1].replace("''", "'")
    return s


def quote_sheet(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


def _split_strings(formula: str):
    """切分为 (是否字符串常量, 片段)"""
    parts, buf, i, n = [], "", 0, len(formula)
    while i < n:
        ch = formula[i]
        if ch == '"':
            if buf:
                parts.append((False, buf))
                buf = ""
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            parts.append((True, formula[i:j + 1]))
            i = j + 1
        else:
            buf += ch
            i += 1
    if buf:
        parts.append((False, buf))
    return parts


def rewrite_refs(formula: str, fn: Callable[[Optional[str], str], Optional[str]]) -> str:
    """对公式中每个引用调用 fn(工作表名或 None, 引用文本)，返回新引用文本；返回 None 表示不改。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    out = []
    for is_str, seg in _split_strings(formula):
        if is_str:
            out.append(seg)
            continue

        def repl(m):
            sheet_raw = m.group("sheet")
            sheet = unquote_sheet(sheet_raw) if sheet_raw else None
            new = fn(sheet, m.group("ref"))
            if new is None:
                return m.group(0)
            if new == "#REF!":
                return "#REF!"
            return (sheet_raw or "") + new

        out.append(REF_RE.sub(repl, seg))
    return "".join(out)


# ---------------------------------------------------------------- 引用解析
class Ref:
    """解析后的引用：单元格、区域、整列、整行"""

    def __init__(self, text: str):
        self.kind = None
        parts = text.split(":")
        if len(parts) == 1:
            m = _CELL_RE.match(parts[0])
            if not m:
                raise ValueError(text)
            self.kind = "cell"
            self.a = [m.group(1) == "$", col_index(m.group(2)), m.group(3) == "$", int(m.group(4))]
            self.b = None
            return
        p, q = parts
        m1, m2 = _CELL_RE.match(p), _CELL_RE.match(q)
        if m1 and m2:
            self.kind = "area"
            self.a = [m1.group(1) == "$", col_index(m1.group(2)), m1.group(3) == "$", int(m1.group(4))]
            self.b = [m2.group(1) == "$", col_index(m2.group(2)), m2.group(3) == "$", int(m2.group(4))]
            return
        c1, c2 = _COL_RE.match(p), _COL_RE.match(q)
        if c1 and c2:
            self.kind = "cols"
            self.a = [c1.group(1) == "$", col_index(c1.group(2)), False, None]
            self.b = [c2.group(1) == "$", col_index(c2.group(2)), False, None]
            return
        r1, r2 = _ROW_RE.match(p), _ROW_RE.match(q)
        if r1 and r2:
            self.kind = "rows"
            self.a = [False, None, r1.group(1) == "$", int(r1.group(2))]
            self.b = [False, None, r2.group(1) == "$", int(r2.group(2))]
            return
        raise ValueError(text)

    @staticmethod
    def _fmt(p, with_col=True, with_row=True):
        s = ""
        if with_col:
            s += ("$" if p[0] else "") + col_letter(p[1])
        if with_row:
            s += ("$" if p[2] else "") + str(p[3])
        return s

    def text(self) -> str:
        if self.kind == "cell":
            return self._fmt(self.a)
        if self.kind == "area":
            return self._fmt(self.a) + ":" + self._fmt(self.b)
        if self.kind == "cols":
            return self._fmt(self.a, True, False) + ":" + self._fmt(self.b, True, False)
        return self._fmt(self.a, False, True) + ":" + self._fmt(self.b, False, True)


# ---------------------------------------------------------------- 结构变化映射
class IndexMap:
    """插入/删除行或列后，旧序号 → 新序号（被删除返回 None）"""

    def __init__(self, insert_at: int = None, count: int = 0, deleted=None):
        self.insert_at, self.count = insert_at, count
        self.deleted = sorted(set(deleted or []))
        self._dset = set(self.deleted)

    def map(self, x: int) -> Optional[int]:
        if self.insert_at is not None:
            return x + self.count if x >= self.insert_at else x
        if x in self._dset:
            return None
        return x - bisect.bisect_left(self.deleted, x)

    def map_span(self, lo: int, hi: int):
        """区域端点映射：插入时自然扩展；删除时向内收缩，全部被删返回 None"""
        if self.insert_at is not None:
            return self.map(lo), self.map(hi)
        a = lo
        while a <= hi and a in self._dset:
            a += 1
        b = hi
        while b >= a and b in self._dset:
            b -= 1
        if a > b:
            return None
        return self.map(a), self.map(b)

    def map_boundary(self, b: int) -> int:
        """分隔位置（如冻结窗格的第一个未冻结列/行）：前面的列/行被插入或删除时随之移动，自身不会被删除"""
        if self.insert_at is not None:
            return b + self.count if self.insert_at < b else b
        return b - bisect.bisect_left(self.deleted, b)

    def is_identity(self):
        return (self.insert_at is None and not self.deleted) or (self.insert_at is not None and self.count == 0)


class GrowMap(IndexMap):
    """在数据末尾追加行：单元格都不动，但“以最后一行数据结尾”的区域向下延伸，
    使汇总公式、筛选、条件格式等自动包含新追加的行（像 Excel 表格自动扩展那样）"""

    def __init__(self, last: int, count: int):
        super().__init__()
        self.last, self.grow = last, count

    def map(self, x):
        return x

    def map_span(self, lo, hi):
        if lo < hi == self.last:
            return lo, hi + self.grow
        return lo, hi

    def map_boundary(self, b):
        return b

    def is_identity(self):
        return self.grow == 0


def shift_ref_text(text: str, cmap: Optional[IndexMap], rmap: Optional[IndexMap]) -> str:
    """按行/列映射调整一个引用；无法表示时返回 #REF!"""
    try:
        r = Ref(text)
    except ValueError:
        return text
    if r.kind == "cell":
        c = cmap.map(r.a[1]) if cmap else r.a[1]
        w = rmap.map(r.a[3]) if rmap else r.a[3]
        if c is None or w is None:
            return "#REF!"
        r.a[1], r.a[3] = c, w
        return r.text()
    if r.kind in ("area", "cols"):
        if cmap:
            span = cmap.map_span(min(r.a[1], r.b[1]), max(r.a[1], r.b[1]))
            if span is None:
                return "#REF!"
            r.a[1], r.b[1] = span
    if r.kind in ("area", "rows"):
        if rmap:
            span = rmap.map_span(min(r.a[3], r.b[3]), max(r.a[3], r.b[3]))
            if span is None:
                return "#REF!"
            r.a[3], r.b[3] = span
    return r.text()


def shift_formula(formula: str, host_sheet: str, target_sheet: str,
                  cmap: Optional[IndexMap], rmap: Optional[IndexMap]) -> str:
    """target_sheet 发生结构变化后，调整位于 host_sheet 的公式里指向 target_sheet 的引用"""

    def fn(sheet, ref):
        name = sheet if sheet is not None else host_sheet
        if name != target_sheet:
            return None
        return shift_ref_text(ref, cmap, rmap)

    return rewrite_refs(formula, fn)


def shift_range_list(text: str, cmap, rmap) -> Optional[str]:
    """调整 'A1:B5 D3' 这类空格分隔的区域列表；全部失效返回 None"""
    out = []
    for part in str(text).split():
        new = shift_ref_text(part, cmap, rmap)
        if new != "#REF!":
            out.append(new)
    return " ".join(out) if out else None


# ---------------------------------------------------------------- 复制公式时的相对平移
def translate_formula(formula: str, d_row: int, d_col: int) -> str:
    """像在 Excel 里复制粘贴公式一样，平移所有相对引用（带 $ 的不动）。越界的引用变为 #REF!"""
    if not d_row and not d_col:
        return formula

    def fn(sheet, ref):
        try:
            r = Ref(ref)
        except ValueError:
            return None
        for p in (r.a, r.b):
            if p is None:
                continue
            if p[1] is not None and not p[0]:
                p[1] += d_col
                if not 1 <= p[1] <= MAX_COL:
                    return "#REF!"
            if p[3] is not None and not p[2]:
                p[3] += d_row
                if not 1 <= p[3] <= MAX_ROW:
                    return "#REF!"
        return r.text()

    return rewrite_refs(formula, fn)


def referenced_sheets(formula: str, host_sheet: str):
    """公式引用到的工作表名集合（不含外部工作簿引用）"""
    names = set()
    if not isinstance(formula, str) or not formula.startswith("="):
        return names

    def fn(sheet, ref):
        names.add(sheet if sheet is not None else host_sheet)
        return None

    rewrite_refs(formula, fn)
    return names


def explicit_sheets(formula: str):
    """公式里明确写了表名的引用所指的工作表（不带表名的引用指当前所在表，不算在内）"""
    names = set()
    if not isinstance(formula, str) or not formula.startswith("="):
        return names

    def fn(sheet, ref):
        if sheet is not None:
            names.add(sheet)
        return None

    rewrite_refs(formula, fn)
    return names


def has_external_ref(formula: str) -> bool:
    return isinstance(formula, str) and re.search(r"\[[^\]]+\][^!]*!", formula) is not None


def parse_range(text: str):
    """'A1:C5' → (min_col, min_row, max_col, max_row)；'B3' → 单格"""
    t = text.replace("$", "").strip().upper()
    if ":" in t:
        a, b = t.split(":", 1)
    else:
        a = b = t
    ma, mb = _CELL_RE.match(a), _CELL_RE.match(b)
    if not ma or not mb:
        raise ValueError(f"区域格式不正确：{text}（应类似 A1:C10）")
    c1, r1 = col_index(ma.group(2)), int(ma.group(4))
    c2, r2 = col_index(mb.group(2)), int(mb.group(4))
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


def parse_cell(text: str):
    m = _CELL_RE.match(text.replace("$", "").strip().upper())
    if not m:
        raise ValueError(f"单元格地址不正确：{text}（应类似 B3）")
    return col_index(m.group(2)), int(m.group(4))
