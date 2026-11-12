"""极简公式计算：只算“同一张表里单元格之间的四则运算”这类最常见的公式，如 =G5*H5、=H5-I5、=ROUND(K5*1.13,2)。

执行器不负责计算公式（那是 Excel/WPS 的事），但条件、排序、分组、匹配需要读到公式的值：
- 文件里已有的公式，Excel 保存过计算结果，可以直接用；
- 本次运行新写入（或数据被改过）的公式还没有结果。能用这里算出来的，就用实时计算的值；
  其他公式（含跨表引用、文字拼接、查找函数等）返回“算不了”，由调用方退回到文件里的结果或按空值处理。
"""
from __future__ import annotations

import datetime as dt
import re
from decimal import ROUND_HALF_UP, Decimal

_TOKEN = re.compile(r"\s*(?:(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)|(\$?[A-Za-z]{1,3}\$?\d+)(?![A-Za-z0-9_(!])|([A-Za-z][A-Za-z0-9.]*)\s*\(|([-+*/^(),%]))")
_FUNCS = {"ROUND", "ABS", "MAX", "MIN", "INT"}


class Unsupported(Exception):
    pass


class _ExcelError(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.code = code


def _tokens(formula: str):
    s = formula[1:] if formula.startswith("=") else formula
    pos, out = 0, []
    while pos < len(s):
        m = _TOKEN.match(s, pos)
        if not m or m.end() == pos:
            if s[pos:].strip() == "":
                break
            raise Unsupported(s[pos:])
        num, ref, fn, op = m.groups()
        if num is not None:
            out.append(("num", float(num)))
        elif ref is not None:
            out.append(("ref", ref.replace("$", "").upper()))
        elif fn is not None:
            if fn.upper() not in _FUNCS:
                raise Unsupported(fn)
            out.append(("fn", fn.upper()))
        else:
            out.append(("op", op))
        pos = m.end()
    return out


def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (dt.date, dt.datetime, dt.time)):
        raise Unsupported("date")          # 日期运算交给 Excel
    if isinstance(v, str):
        t = v.strip()
        if t.startswith("#"):
            raise _ExcelError(t)
        if t == "":
            return 0.0
        try:
            return float(t.replace(",", ""))
        except ValueError:
            raise _ExcelError("#VALUE!")
    raise Unsupported(type(v).__name__)


class _Parser:
    def __init__(self, toks, getter):
        self.t, self.i, self.get = toks, 0, getter

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expr(self):
        v = self.term()
        while self.peek() in (("op", "+"), ("op", "-")):
            op = self.take()[1]
            r = self.term()
            v = v + r if op == "+" else v - r
        return v

    def term(self):
        v = self.power()
        while self.peek() in (("op", "*"), ("op", "/")):
            op = self.take()[1]
            r = self.power()
            if op == "*":
                v = v * r
            else:
                if r == 0:
                    raise _ExcelError("#DIV/0!")
                v = v / r
        return v

    def power(self):
        v = self.unary()
        while self.peek() == ("op", "^"):
            self.take()
            v = v ** self.unary()
        return v

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return -self.unary()
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        v = self.atom()
        if self.peek() == ("op", "%"):
            self.take()
            v = v / 100
        return v

    def atom(self):
        kind, val = self.take()
        if kind == "num":
            return val
        if kind == "ref":
            m = re.match(r"([A-Z]+)(\d+)$", val)
            return _num(self.get(_col_index(m.group(1)), int(m.group(2))))
        if kind == "op" and val == "(":
            v = self.expr()
            if self.take() != ("op", ")"):
                raise Unsupported("paren")
            return v
        if kind == "fn":
            args = []
            if self.peek() != ("op", ")"):
                args.append(self.expr())
                while self.peek() == ("op", ","):
                    self.take()
                    args.append(self.expr())
            if self.take() != ("op", ")"):
                raise Unsupported("paren")
            return _call(val, args)
        raise Unsupported(str(val))


def _call(fn, args):
    if fn == "ROUND" and len(args) == 2:
        # 与 Excel 一致：先按 15 位有效数字取值（15.254999… → 15.255），再四舍五入（远离 0）
        q = Decimal(1).scaleb(-int(args[1]))
        d = Decimal(format(args[0], ".15g")).copy_abs().quantize(q, rounding=ROUND_HALF_UP)
        return float(d if args[0] >= 0 else -d)
    if fn == "ABS" and len(args) == 1:
        return abs(args[0])
    if fn == "INT" and len(args) == 1:
        import math
        return float(math.floor(args[0]))
    if fn in ("MAX", "MIN") and args:
        return max(args) if fn == "MAX" else min(args)
    raise Unsupported(fn)


def evaluate(formula, getter):
    """返回 (能否计算, 结果)。getter(列号, 行号) → 单元格当前值（可递归计算其他公式）。
    结果是数字，或 Excel 错误值文字（如 "#DIV/0!"）。"""
    if not isinstance(formula, str) or not formula.startswith("="):
        return False, None
    try:
        toks = _tokens(formula)
        if not toks:
            return False, None
        p = _Parser(toks, getter)
        v = p.expr()
        if p.i != len(toks):
            return False, None
        if isinstance(v, complex):
            return True, "#NUM!"
        return True, (int(v) if float(v).is_integer() and abs(v) < 1e15 else round(v, 10))
    except _ExcelError as e:
        return True, e.code
    except (Unsupported, OverflowError, ValueError, ZeroDivisionError, RecursionError):
        return False, None
