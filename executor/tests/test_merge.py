"""阶段 3：多文件合并（追加行）与分组汇总"""
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

from engine.config import validate  # noqa: E402
from engine.errors import ConfigError  # noqa: E402
from engine.runner import run  # noqa: E402


def make_book(path, rows, headers=("门店", "商品", "数量", "单价", "金额"), total=False, summary=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(list(headers))
    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)
    for i, r in enumerate(rows, 2):
        ws.append(list(r))
        if "金额" in headers:
            q, p, a = (headers.index(x) + 1 for x in ("数量", "单价", "金额"))
            L = lambda c: chr(64 + c)
            ws.cell(row=i, column=a, value=f"={L(q)}{i}*{L(p)}{i}")
            ws.cell(row=i, column=p).number_format = "0.00"
    last = len(rows) + 1
    if total:
        ws.append(["合计"])
    ws.auto_filter.ref = f"A1:{chr(64 + len(headers))}{last}"
    if summary:
        s = wb.create_sheet("汇总")
        s["A1"] = "金额合计"
        s["B1"] = f"=SUM(明细!E2:E{last})"
        s["A2"] = "第一行"
        s["B2"] = "=明细!$E$2"
    wb.save(path)


class MergeTest(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="efa_m_"))
        self.out = self.dir / "out"
        self.main = self.dir / "总表.xlsx"
        make_book(self.main, [("总部", "A", 1, 10), ("总部", "B", 2, 20)])
        self.s1 = self.dir / "门店1.xlsx"
        self.s2 = self.dir / "门店2.xlsx"
        # 来源列顺序不同、少“金额”、多一列“备注”
        make_book(self.s1, [("1店", "A", 3, 10, "x"), ("1店", "C", 1, 5, None), (None, None, None, None, None), ("1店", "D", 4, 1, "y")],
                  headers=("门店", "商品", "数量", "单价", "备注"), summary=False)
        make_book(self.s2, [("B", "2店", 20, 5, "z")], headers=("商品", "门店", "单价", "数量", "备注"), summary=False)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def cfg(self, steps):
        return {"schema": "excel-workflow/1", "name": "合并",
                "files": [{"id": "main", "role": "main"}, {"id": "stores", "role": "ref", "label": "门店表", "multiple": True}],
                "options": {"recalc": "off"}, "steps": steps}

    def go(self, steps, paths=None):
        rep = run(self.cfg(steps), paths or {"main": str(self.main), "stores": [str(self.s1), str(self.s2)]}, self.out, progress=lambda m: None)
        self.assertTrue(rep.ok, rep.render())
        return rep, load_workbook(rep.output)

    def test_append_rows_by_header(self):
        rep, wb = self.go([{"op": "append_rows", "from": {"file": "stores"}, "source_column": "来源"}])
        ws = wb["明细"]
        self.assertEqual([c.value for c in ws[1]], ["门店", "商品", "数量", "单价", "金额", "来源"])
        got = [[ws.cell(row=r, column=c).value for c in range(1, 7)] for r in range(2, 8)]
        # 原有 2 行 + 门店1 的 3 行（空行跳过）+ 门店2 的 1 行；按表头对应；金额公式向下延续；来源写文件名
        self.assertEqual(got[2], ["1店", "A", 3, 10, "=C4*D4", "门店1"])
        self.assertEqual(got[4], ["1店", "D", 4, 1, "=C6*D6", "门店1"])
        self.assertEqual(got[5], ["2店", "B", 5, 20, "=C7*D7", "门店2"])
        self.assertIsNone(ws["A8"].value)
        self.assertEqual(ws["D5"].number_format, "0.00", "新行沿用模板行格式")
        self.assertEqual(wb["汇总"]["B1"].value, "=SUM(明细!E2:E7)", "汇总公式自动延伸到追加的行")
        self.assertEqual(wb["汇总"]["B2"].value, "=明细!$E$2", "单个单元格引用不变")
        self.assertEqual(ws.auto_filter.ref, "A1:F7")
        self.assertTrue(any("备注" in w for s in rep.steps for w in s["warnings"]), "来源多出的列要提示")
        self.assertIn("2 个文件追加 4 行", rep.steps[0]["detail"])

    def test_append_add_missing_and_where(self):
        rep, wb = self.go([{"op": "append_rows", "from": {"file": "stores"}, "add_missing_columns": True,
                            "where": {"col": "数量", "op": ">=", "value": 3}, "extend_formulas": False}])
        ws = wb["明细"]
        self.assertEqual(ws["F1"].value, "备注")
        self.assertEqual([ws.cell(row=r, column=2).value for r in range(2, 7)], ["A", "B", "A", "D", "B"])
        self.assertEqual(ws["F3"].value, None)
        self.assertEqual(ws["F4"].value, "x")
        self.assertIsNone(ws["E4"].value, "不延续公式时，来源没有的列留空")

    def test_main_with_total_row_warns(self):
        make_book(self.main, [("总部", "A", 1, 10)], total=True)
        rep, wb = self.go([{"op": "append_rows", "from": {"file": "stores"}}])
        self.assertTrue(any("合计行" in w for s in rep.steps for w in s["warnings"]))

    def test_multiple_only_for_append(self):
        cfg = self.cfg([{"op": "match", "id": "m", "ref": "stores", "main_keys": ["商品"]}])
        with self.assertRaises(ConfigError):
            validate(cfg)
        cfg = self.cfg([])
        cfg["files"][0]["multiple"] = True
        cfg["steps"] = [{"op": "set_cells", "cells": {"A1": 1}}]
        with self.assertRaises(ConfigError):
            validate(cfg)

    def test_group_summary_formula(self):
        rep, wb = self.go([
            {"op": "append_rows", "from": {"file": "stores"}},
            {"op": "group_summary", "group_by": ["门店"], "metrics": [
                {"column": "金额", "func": "sum"}, {"func": "count"}, {"column": "单价", "func": "max", "title": "最高单价"}]},
        ])
        s = wb["分组汇总"]
        self.assertEqual([c.value for c in s[1]], ["门店", "金额合计", "记录数", "最高单价"])
        self.assertEqual([s.cell(row=r, column=1).value for r in range(2, 5)], ["1店", "2店", "总部"])
        self.assertEqual(s["B2"].value, "=SUMIFS('明细'!$E$2:$E$7,'明细'!$A$2:$A$7,$A2)")
        self.assertEqual(s["C3"].value, "=COUNTIFS('明细'!$A$2:$A$7,$A3)")
        self.assertEqual(s["D4"].value, "=IFERROR(AGGREGATE(14,6,INDEX('明细'!$D$2:$D$7/(('明细'!$A$2:$A$7=$A4)),0),1),\"\")")
        self.assertEqual(s["A5"].value, "合计")
        self.assertEqual(s["B5"].value, "=SUM(B2:B4)")
        self.assertEqual(rep.formula["problems"], [], "汇总公式都应符合 Excel 2016")

    def test_group_summary_values_and_errors(self):
        rep, wb = self.go([
            {"op": "append_rows", "from": {"file": "stores"}},
            {"op": "group_summary", "group_by": ["门店", "商品"], "mode": "value", "sort": "metric_desc", "sheet_name": "按店按品",
             "where": {"col": "数量", "op": ">", "value": 1},
             "metrics": [{"column": "数量", "func": "sum"}, {"column": "商品", "func": "count_distinct"}, {"column": "单价", "func": "avg"}]},
        ])
        s = wb["按店按品"]
        rows = [[c.value for c in r] for r in s.iter_rows(min_row=2)]
        self.assertEqual(rows[0][:3], ["2店", "B", 5])           # 数量合计最大的排第一
        self.assertEqual(rows[-1][0], "合计")
        self.assertEqual(rows[-1][2], 14)                        # 2+3+4+5（数量>1）
        bad = run(self.cfg([{"op": "group_summary", "group_by": ["门店"], "metrics": [{"column": "商品", "func": "count_distinct"}]}]),
                  {"main": str(self.main), "stores": [str(self.s1)]}, self.out, progress=lambda m: None)
        self.assertFalse(bad.ok)
        self.assertIn("结果值", bad.error)


class MiniExprTest(unittest.TestCase):
    def test_evaluate(self):
        from engine import miniexpr
        cells = {(1, 1): 3, (2, 1): 4.5, (3, 1): None, (4, 1): "abc", (5, 1): "12"}
        g = lambda c, r: cells.get((c, r))
        ev = lambda f: miniexpr.evaluate(f, g)
        self.assertEqual(ev("=A1*B1"), (True, 13.5))
        self.assertEqual(ev("=ROUND(A1*B1*1.13,2)"), (True, 15.26))
        self.assertEqual(ev("=-A1+C1-(B1-1)/2"), (True, -4.75))      # 空单元格按 0
        self.assertEqual(ev("=A1/C1"), (True, "#DIV/0!"))
        self.assertEqual(ev("=A1*D1"), (True, "#VALUE!"))
        self.assertEqual(ev("=E1+1"), (True, 13))
        self.assertEqual(ev("=10%*A1"), (True, 0.3))
        for f in ("=SUM(A1:B1)", "=Sheet2!A1*2", "=A1&B1", "=IF(A1>1,1,0)", "=VLOOKUP(A1,B:C,2,0)", "='x'!A1"):
            self.assertEqual(ev(f)[0], False, f)


class FreshFormulaTest(unittest.TestCase):
    """本次新写入的简单公式，在后面的条件、排序、汇总里也能读到值"""

    def test_condition_and_sort_on_new_formula(self):
        d = Path(tempfile.mkdtemp(prefix="efa_f_"))
        try:
            p = d / "t.xlsx"
            make_book(p, [("A", "x", 2, 10), ("B", "y", 5, 3), ("C", "z", 1, 50)], summary=False)
            cfg = {"schema": "excel-workflow/1", "name": "t", "files": [{"id": "main", "role": "main"}], "options": {"recalc": "off"}, "steps": [
                {"op": "insert_column", "header": "差", "end": True, "formula": "=[@金额]-[@单价]*2"},
                {"op": "format", "columns": ["差"], "where": {"col": "差", "op": "<", "value": 0}, "fill": "红"},
                {"op": "sort_rows", "by": [{"column": "差", "order": "desc"}]},
                {"op": "group_summary", "group_by": ["门店"], "mode": "value", "metrics": [{"column": "差", "func": "sum"}]},
            ]}
            rep = run(cfg, {"main": str(p)}, d / "out", progress=lambda m: None)
            self.assertTrue(rep.ok, rep.render())
            self.assertFalse(any("没有计算结果" in w for s in rep.steps for w in s["warnings"]))
            ws = load_workbook(rep.output)["明细"]
            # 差 = 金额-单价*2：A=20-20=0，B=15-6=9，C=50-100=-50 → 降序 B、A、C
            self.assertEqual([ws.cell(row=r, column=1).value for r in (2, 3, 4)], ["B", "A", "C"])
            self.assertTrue(ws["F4"].fill.start_color.rgb.endswith("FFC7CE"))
            self.assertFalse(str(ws["F2"].fill.start_color.rgb).endswith("FFC7CE"))
            s = load_workbook(rep.output)["分组汇总"]
            self.assertEqual([[c.value for c in r] for r in s.iter_rows(min_row=2)], [["A", 0], ["B", 9], ["C", -50], ["合计", -41]])
        finally:
            shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
