"""公式检查：
1. 语法检查（不需要 Excel）：括号/引号是否配对、是否含 #REF!、是否用了 Excel 2016 没有的函数
2. 可选重算校验：本机有 Excel 或 WPS 时，后台打开结果文件重算并保存，再扫描计算结果中的错误值
"""
from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.tokenizer import Tokenizer, TokenizerError

from .excel_functions import EXCEL2016, NEWER_ONLY

FUNC_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_.]*)\s*\(")
ERRORS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!", "#N/A", "#SPILL!", "#CALC!")
RECALC_PS1 = Path(__file__).with_name("recalc.ps1")


def _strip_strings(f: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', '""', f)


def syntax_check(wb, only_ids=None, limit=50):
    """返回 (检查的公式数, 问题列表[(位置, 说明)])"""
    checked, problems = 0, []
    seen_bad_funcs = set()
    for ws in wb.worksheets:
        for cell in ws._cells.values():
            if cell.data_type != "f" or not isinstance(cell.value, str):
                continue
            if only_ids is not None and id(cell) not in only_ids:
                continue
            checked += 1
            f = cell.value
            where = f"{ws.title}!{cell.coordinate}"
            body = _strip_strings(f)
            if "#REF!" in body:
                problems.append((where, "公式里有失效引用 #REF!（引用的行/列被删除了）"))
                continue
            try:
                Tokenizer(f)
            except (TokenizerError, Exception) as e:  # noqa: BLE001
                problems.append((where, f"公式语法有误：{e}"))
                continue
            if body.count("(") != body.count(")"):
                problems.append((where, "括号不配对"))
                continue
            for name in FUNC_RE.findall(body):
                u = name.upper()
                if u.startswith("_XLFN.") or u.startswith("_XLWS."):
                    u = u.split(".", 1)[1]
                if u in EXCEL2016 or u in seen_bad_funcs:
                    continue
                seen_bad_funcs.add(u)
                alt = NEWER_ONLY.get(u)
                problems.append((where, f"函数 {u} " + (f"是新版 Excel 才有的，Excel 2016/WPS 旧版打不开（可改用 {alt}）" if alt else "不在 Excel 2016 函数列表中")))
            if len(problems) >= limit:
                return checked, problems
    return checked, problems


def formula_positions(wb):
    pos = {}
    for ws in wb.worksheets:
        s = {k for k, c in ws._cells.items() if c.data_type == "f"}
        if s:
            pos[ws.title] = s
    return pos


def recalc(path: str, timeout=240):
    """用本机 Excel/WPS 重算并保存。返回 (状态, 说明)：ok / noapp / error"""
    if os.name != "nt":
        return "noapp", "非 Windows 系统"
    try:
        p = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(RECALC_PS1), str(Path(path).resolve())],
            capture_output=True, timeout=timeout, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except FileNotFoundError:
        return "noapp", "找不到 PowerShell"
    except subprocess.TimeoutExpired:
        return "error", f"重算超过 {timeout} 秒未完成，已跳过（可能有弹窗卡住了 Excel/WPS）"
    out = (p.stdout or b"").decode("utf-8", "ignore") + (p.stderr or b"").decode("gbk", "ignore")
    line = next((x.strip() for x in out.splitlines() if x.strip().startswith(("OK", "NOAPP", "ERR"))), out.strip()[:200])
    if line.startswith("OK"):
        app = line[2:].strip()
        return "ok", {"Excel.Application": "Excel", "KET.Application": "WPS 表格", "ET.Application": "WPS 表格"}.get(app, app)
    if line.startswith("NOAPP"):
        return "noapp", "本机没有安装 Excel 或 WPS"
    return "error", line.replace("ERR", "", 1).strip() or "重算失败"


def scan_values(path: str, positions, limit=50):
    """读取重算后的结果，找出公式单元格里的错误值。返回 (错误数, #N/A 数, 示例)"""
    wb = load_workbook(path, data_only=True, read_only=True)
    n_err, n_na, examples = 0, 0, []
    try:
        for title, coords in positions.items():
            if title not in wb.sheetnames:
                continue
            for row in wb[title].iter_rows():
                for c in row:
                    rc = (getattr(c, "row", None), getattr(c, "column", None))
                    if rc not in coords:
                        continue
                    v = c.value
                    if isinstance(v, str) and v in ERRORS:
                        if v == "#N/A":
                            n_na += 1
                        else:
                            n_err += 1
                        if len(examples) < limit:
                            examples.append((f"{title}!{c.coordinate}", v))
    finally:
        wb.close()
    return n_err, n_na, examples
