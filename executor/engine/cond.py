"""条件判断：用于“满足条件的行才处理”。

写法：
  {"col": "未分配数", "op": ">=", "value": 10}
  {"col": "实发", "op": "=", "col2": "应发"}             两列比较
  {"col": "备注", "op": "empty"}
  {"col": "名称", "op": "contains", "value": "测试"}
  {"col": "状态", "op": "in", "value": ["已发货", "已签收"]}
  {"ref_col": "入库时间", "op": "not_empty"}             匹配到的参照表行（仅匹配类操作可用）
  {"all": [条件, 条件]} / {"any": [...]} / {"not": 条件}
"""
from __future__ import annotations

import datetime as dt
import re

from .book import is_blank
from .errors import ConfigError, StepError

OPS = {
    "=", "==", "!=", "<>", ">", ">=", "<", "<=", "contains", "not_contains", "starts_with", "ends_with",
    "empty", "not_empty", "in", "not_in", "regex", "is_number", "is_text", "is_error",
}
OP_ALIASES = {"等于": "=", "不等于": "!=", "大于": ">", "大于等于": ">=", "小于": "<", "小于等于": "<=",
              "包含": "contains", "不包含": "not_contains", "开头是": "starts_with", "结尾是": "ends_with",
              "为空": "empty", "不为空": "not_empty", "属于": "in", "不属于": "not_in"}
ERROR_VALUES = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!", "#SPILL!", "#CALC!"}


def validate(cond, where="条件"):
    if cond is None:
        return
    if not isinstance(cond, dict):
        raise ConfigError(f"{where}应是一个对象，例如 {{\"col\": \"数量\", \"op\": \">=\", \"value\": 10}}")
    for k in ("all", "any"):
        if k in cond:
            if not isinstance(cond[k], list) or not cond[k]:
                raise ConfigError(f"{where}的 {k} 应是非空列表")
            for i, c in enumerate(cond[k], 1):
                validate(c, f"{where}.{k}[{i}]")
            return
    if "not" in cond:
        validate(cond["not"], f"{where}.not")
        return
    if "col" not in cond and "ref_col" not in cond:
        raise ConfigError(f"{where}缺少 col（要判断的列）")
    op = OP_ALIASES.get(cond.get("op"), cond.get("op"))
    if op not in OPS:
        raise ConfigError(f"{where}的 op「{cond.get('op')}」不支持，可用：{'、'.join(sorted(OPS))}")
    if op not in ("empty", "not_empty", "is_number", "is_text", "is_error") and "value" not in cond and "col2" not in cond and "ref_col2" not in cond:
        raise ConfigError(f"{where}缺少 value（比较的值）或 col2（比较的另一列）")
    if op in ("in", "not_in") and not isinstance(cond.get("value"), list):
        raise ConfigError(f"{where}使用 in/not_in 时 value 应是列表")


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if s.endswith("%"):
            try:
                return float(s[:-1]) / 100
            except ValueError:
                return None
        try:
            return float(s)
        except ValueError:
            return None
    return None


def _date(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    if isinstance(v, str):
        m = re.fullmatch(r"\s*(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*", v)
        if m:
            y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            hh, mi, ss = int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0)
            try:
                return dt.datetime(y, mo, d, hh, mi, ss)
            except ValueError:
                return None
    return None


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def compare(a, op, b, ignore_case=False):
    if op in ("=", "=="):
        return _eq(a, b, ignore_case)
    if op in ("!=", "<>"):
        return not _eq(a, b, ignore_case)
    if op in (">", ">=", "<", "<="):
        x, y = _date(a), _date(b)
        if x is None or y is None:
            x, y = _num(a), _num(b)
        if x is None or y is None:
            return False
        return {">": x > y, ">=": x >= y, "<": x < y, "<=": x <= y}[op]
    ta, tb = _text(a), _text(b)
    if ignore_case:
        ta, tb = ta.lower(), tb.lower()
    if op == "contains":
        return tb in ta
    if op == "not_contains":
        return tb not in ta
    if op == "starts_with":
        return ta.startswith(tb)
    if op == "ends_with":
        return ta.endswith(tb)
    if op == "regex":
        return re.search(str(b), _text(a), re.I if ignore_case else 0) is not None
    raise StepError(f"不支持的比较方式：{op}")


def _eq(a, b, ignore_case):
    if is_blank(a) and is_blank(b):
        return True
    x, y = _num(a), _num(b)
    if x is not None and y is not None and not (isinstance(a, str) and isinstance(b, str)):
        return abs(x - y) < 1e-9
    da, db = _date(a), _date(b)
    if da is not None and db is not None:
        return da == db
    ta, tb = _text(a).strip(), _text(b).strip()
    return ta.lower() == tb.lower() if ignore_case else ta == tb


def evaluate(cond, get, get_ref=None):
    """get(列) → 当前行该列的值；get_ref(列) → 匹配到的参照表行该列的值"""
    if cond is None:
        return True
    if "all" in cond:
        return all(evaluate(c, get, get_ref) for c in cond["all"])
    if "any" in cond:
        return any(evaluate(c, get, get_ref) for c in cond["any"])
    if "not" in cond:
        return not evaluate(cond["not"], get, get_ref)
    if "ref_col" in cond:
        if get_ref is None:
            raise StepError("条件里的 ref_col 只能用在匹配类操作中")
        a = get_ref(cond["ref_col"])
    else:
        a = get(cond["col"])
    op = OP_ALIASES.get(cond.get("op"), cond.get("op"))
    ic = bool(cond.get("ignore_case"))
    if op == "empty":
        return is_blank(a)
    if op == "not_empty":
        return not is_blank(a)
    if op == "is_number":
        return _num(a) is not None and not isinstance(a, str)
    if op == "is_text":
        return isinstance(a, str) and a.strip() != ""
    if op == "is_error":
        return isinstance(a, str) and a.strip() in ERROR_VALUES
    if "col2" in cond:
        b = get(cond["col2"])
    elif "ref_col2" in cond:
        if get_ref is None:
            raise StepError("条件里的 ref_col2 只能用在匹配类操作中")
        b = get_ref(cond["ref_col2"])
    else:
        b = cond.get("value")
    if op in ("in", "not_in"):
        hit = any(_eq(a, x, ic) for x in b)
        return hit if op == "in" else not hit
    return compare(a, op, b, ic)


def columns_used(cond, acc=None):
    """条件里用到的主表列（用于提前校验）"""
    acc = acc if acc is not None else []
    if not cond:
        return acc
    for k in ("all", "any"):
        if k in cond:
            for c in cond[k]:
                columns_used(c, acc)
            return acc
    if "not" in cond:
        return columns_used(cond["not"], acc)
    if "col" in cond:
        acc.append(cond["col"])
    if "col2" in cond:
        acc.append(cond["col2"])
    return acc
