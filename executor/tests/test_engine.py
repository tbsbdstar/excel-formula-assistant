"""执行器自动化测试：python -m unittest discover -s executor/tests -v"""
import datetime as dt
import shutil
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402
from openpyxl.styles import Font, PatternFill  # noqa: E402

from engine import refs  # noqa: E402
from engine.config import validate  # noqa: E402
from engine.errors import ConfigError  # noqa: E402
from engine.runner import run  # noqa: E402

DEPTS = ["销售部", "市场部", "测试", "技术部"]


def make_main(path, n=20):
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    ws.append(["编号", "名称", "部门", "数量", "单价", "金额", "备注"])
    for c in range(1, 8):
        ws.cell(row=1, column=c).font = Font(bold=True)
        ws.cell(row=1, column=c).fill = PatternFill("solid", start_color="FFDDEBF7", end_color="FFDDEBF7")
    for i in range(n):
        r = i + 2
        ws.append([f"A{i:03d}", f"产品{i}", DEPTS[i % 4], i + 1, 10.0 + i, f"=D{r}*E{r}", None])
        ws.cell(row=r, column=5).number_format = "0.00"
    ws.column_dimensions["E"].width = 21
    ws.conditional_formatting.add(f"F2:F{n + 1}", CellIsRule(operator="greaterThan", formula=["300"], fill=PatternFill("solid", start_color="FFFFC7CE", end_color="FFFFC7CE")))
    s = wb.create_sheet("汇总")
    s["A1"] = "汇总"
    s.merge_cells("A1:B1")
    s["A2"] = "金额合计"
    s["B2"] = f"=SUM('明细'!F2:F{n + 1})"
    s["A3"] = "销售部数量"
    s["B3"] = f'=SUMIF(明细!C2:C{n + 1},"销售部",明细!$D$2:$D${n + 1})'
    s["A4"] = "第一行金额"
    s["B4"] = "=明细!$F$2"
    wb.save(path)


def make_ref(path, n=12):
    wb = Workbook()
    ws = wb.active
    ws.title = "上期"
    ws.append(["编号", "部门", "上期金额", "入库时间", "公式列"])
    for i in range(n):
        r = i + 2
        ws.append([f"A{i:03d}", DEPTS[i % 4], 100 + i, dt.datetime(2024, 1, 1 + i) if i % 3 else None, f"=C{r}*2"])
        ws.cell(row=r, column=3).fill = PatternFill("solid", start_color="FFFCE4D6", end_color="FFFCE4D6")
    ws.append(["B999", "市场部", 5, None, "=C14*2"])
    blk = wb.create_sheet("块")
    blk["A1"] = "小计"
    blk["A1"].font = Font(bold=True, color="FFFF0000")
    blk["B1"] = "=SUM(B2:B3)"
    blk["B2"] = 1
    blk["B3"] = 2
    blk.merge_cells("A2:A3")
    wb.save(path)


def base_cfg(steps, **kw):
    cfg = {"schema": "excel-workflow/1", "name": "测试",
           "files": [{"id": "main", "role": "main", "label": "主表"}, {"id": "ref", "role": "ref", "label": "参照表"}],
           "options": {"recalc": "off"}, "steps": steps}
    cfg.update(kw)
    return cfg


class RefsTest(unittest.TestCase):
    def test_shift_delete_column(self):
        m = refs.IndexMap(deleted=[2])
        self.assertEqual(refs.shift_formula("=D2*E2", "S", "S", m, None), "=C2*D2")
        self.assertEqual(refs.shift_formula("=SUM(A1:C9)", "S", "S", m, None), "=SUM(A1:B9)")
        self.assertEqual(refs.shift_formula("=B5", "S", "S", m, None), "=#REF!")
        self.assertEqual(refs.shift_formula('=SUMIF(S!C2:C9,"B2",S!$D$2:$D$9)', "X", "S", m, None), '=SUMIF(S!B2:B9,"B2",S!$C$2:$C$9)')
        self.assertEqual(refs.shift_formula("=D2", "Other", "S", m, None), "=D2")
        self.assertEqual(refs.shift_formula("='S'!D2+LOG10(D2)", "Other", "S", m, None), "='S'!C2+LOG10(D2)")

    def test_shift_insert_and_rows(self):
        ins = refs.IndexMap(insert_at=3, count=1)
        self.assertEqual(refs.shift_formula("=SUM(B2:D2)+C$1+$A1", "S", "S", ins, None), "=SUM(B2:E2)+D$1+$A1")
        self.assertEqual(refs.shift_formula("=SUM(C:C)", "S", "S", ins, None), "=SUM(D:D)")
        rm = refs.IndexMap(deleted=[3, 4])
        self.assertEqual(refs.shift_formula("=SUM(F2:F21)", "S", "S", None, rm), "=SUM(F2:F19)")
        self.assertEqual(refs.shift_formula("=SUM(F3:F4)", "S", "S", None, rm), "=SUM(#REF!)")

    def test_translate(self):
        self.assertEqual(refs.translate_formula("=C2*2+$C$2+Sheet2!B2", 3, 1), "=D5*2+$C$2+Sheet2!C5")
        self.assertEqual(refs.translate_formula('=IF(A1="A1",1,0)', 1, 0), '=IF(A2="A1",1,0)')
        self.assertEqual(refs.translate_formula("=A1", -5, 0), "=#REF!")


class EngineTest(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="efa_"))
        self.main = self.dir / "本期.xlsx"
        self.ref = self.dir / "上期.xlsx"
        make_main(self.main)
        make_ref(self.ref)
        self.out = self.dir / "out"

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def go(self, steps, **kw):
        rep = run(base_cfg(steps, **kw), {"main": str(self.main), "ref": str(self.ref)}, self.out, progress=lambda m: None)
        return rep

    def ok(self, steps, **kw):
        rep = self.go(steps, **kw)
        self.assertTrue(rep.ok, rep.render())
        return rep, load_workbook(rep.output)

    def test_delete_column_keeps_formulas(self):
        rep, wb = self.ok([{"op": "delete_columns", "columns": ["名称"]}])
        ws = wb["明细"]
        self.assertEqual([c.value for c in ws[1]][:6], ["编号", "部门", "数量", "单价", "金额", "备注"])
        self.assertEqual(ws["E2"].value, "=C2*D2")
        self.assertEqual(wb["汇总"]["B2"].value, "=SUM('明细'!E2:E21)")
        self.assertEqual(wb["汇总"]["B3"].value, '=SUMIF(明细!B2:B21,"销售部",明细!$C$2:$C$21)')
        self.assertEqual(wb["汇总"]["B4"].value, "=明细!$E$2")
        self.assertAlmostEqual(ws.column_dimensions["D"].width, 21)
        self.assertEqual(str(next(iter(ws.conditional_formatting)).sqref), "E2:E21")
        self.assertEqual(ws["D2"].number_format, "0.00")
        self.assertEqual(load_workbook(self.main)["明细"]["B1"].value, "名称", "原文件不能被修改")

    def test_freeze_panes_and_autofilter_follow_columns(self):
        wb0 = load_workbook(self.main)
        ws0 = wb0["明细"]
        ws0.freeze_panes = "C2"          # 冻结第 1 行和 A、B 两列
        ws0.auto_filter.ref = "A1:G21"
        wb0.save(self.main)
        # 删掉冻结区里的 A 列：冻结列数减 1，冻结行不变（以前会整个丢掉冻结）
        rep, wb = self.ok([{"op": "delete_columns", "columns": ["编号"]}])
        self.assertEqual(wb["明细"].freeze_panes, "B2")
        # 冻结区内插入：冻结列数加 1；冻结区外插入/删除：不变
        rep, wb = self.ok([{"op": "insert_column", "header": "新", "before": "名称"}, {"op": "delete_columns", "columns": ["备注"]}])
        self.assertEqual(wb["明细"].freeze_panes, "D2")
        # 只冻结行（A2）时删除 A 列仍然冻结第 1 行
        wb0 = load_workbook(self.main)
        wb0["明细"].freeze_panes = "A2"
        wb0.save(self.main)
        rep, wb = self.ok([{"op": "delete_columns", "columns": ["编号"]}, {"op": "delete_rows", "rows": "2:3"}])
        self.assertEqual(wb["明细"].freeze_panes, "A2")
        # 末尾新增的列自动纳入筛选范围
        rep, wb = self.ok([
            {"op": "insert_column", "header": "备注2", "end": True},
            {"op": "match", "id": "m", "ref": "ref", "main_keys": ["编号"]},
            {"op": "copy_matched", "match": "m", "columns": ["入库时间"], "create_missing": True},
        ])
        self.assertEqual(wb["明细"].auto_filter.ref, "A1:I21")

    def test_insert_fill_and_format(self):
        rep, wb = self.ok([
            {"op": "insert_column", "header": "折扣", "before": "金额", "value": 0.9},
            {"op": "insert_column", "header": "折后金额", "end": True, "formula": "=[@金额]*[@折扣]"},
            {"op": "format", "columns": ["数量", "金额"], "where": {"col": "数量", "op": ">=", "value": 10}, "fill": "FCE4D6"},
            {"op": "fill", "column": "备注", "where": {"all": [{"col": "部门", "op": "=", "value": "测试"}, {"col": "数量", "op": ">", "value": 5}]}, "value": "逾期"},
        ])
        ws = wb["明细"]
        self.assertEqual(ws["F1"].value, "折扣")
        self.assertEqual(ws["G2"].value, "=D2*E2")
        self.assertEqual(ws["F1"].font.b, True, "新列表头应继承左侧格式")
        self.assertEqual(ws["I1"].value, "折后金额")
        self.assertEqual(ws["I5"].value, "=G5*F5")
        self.assertEqual(wb["汇总"]["B2"].value, "=SUM('明细'!G2:G21)")
        self.assertEqual(ws["D11"].fill.start_color.rgb, "FFFCE4D6")
        self.assertNotEqual(ws["D10"].fill.start_color.rgb, "FFFCE4D6")
        self.assertEqual(ws["H8"].value, "逾期")  # 第 7 条数据：测试部，数量 7
        self.assertIsNone(ws["H4"].value)        # 第 3 条：测试部，数量 3

    def test_match_copy_and_compare(self):
        rep, wb = self.ok([
            {"op": "insert_column", "header": "入库时间", "end": True},
            {"op": "insert_column", "header": "上期两倍", "end": True},
            {"op": "match", "id": "m", "ref": "ref", "main_keys": ["编号", "部门"]},
            {"op": "copy_matched", "match": "m", "columns": ["入库时间"], "where": {"col": "数量", "op": ">=", "value": 3},
             "fill": "FCE4D6", "color_columns": "A:C"},
            {"op": "copy_matched", "match": "m", "columns": [{"from": "公式列", "to": "上期两倍"}], "mode": "formula"},
            {"op": "compare", "match": "m", "compare_columns": [{"main": "金额", "ref": "上期金额"}]},
        ])
        ws = wb["明细"]
        h = {c.value: c.column_letter for c in ws[1] if c.value}
        st = rep.match_stats["m"]
        self.assertEqual(st["主表匹配成功"], 12)
        self.assertEqual(st["仅主表有"], 8)
        self.assertEqual(st["仅参照表有"], 1)
        self.assertEqual(ws[h["入库时间"] + "6"].value, dt.datetime(2024, 1, 5))   # 第 5 条（i=4）
        self.assertIsNone(ws[h["入库时间"] + "3"].value)                            # 数量 2 < 3
        self.assertIsNone(ws[h["入库时间"] + "5"].value)                            # 参照表为空，不写
        self.assertEqual(ws["A6"].fill.start_color.rgb, "FFFCE4D6")
        # 参照表 E4 的 =C4*2 复制到主表 I4：列向右平移 4 列 → =G4*2
        self.assertEqual(h["上期两倍"], "I")
        self.assertEqual(ws["I4"].value, "=G4*2")
        self.assertEqual(ws[h["比对结果"] + "2"].value, "有差异")
        self.assertEqual(ws[h["比对结果"] + "21"].value, "仅主表有")
        self.assertIn("比对报告", wb.sheetnames)

    def test_formula_copy_translates_rows(self):
        rep, wb = self.ok([
            {"op": "delete_rows", "where": {"col": "部门", "op": "=", "value": "测试"}},
            {"op": "insert_column", "header": "上期两倍", "end": True},
            {"op": "match", "id": "m", "ref": "ref", "main_keys": "编号"},
            {"op": "copy_matched", "match": "m", "columns": [{"from": "公式列", "to": "上期两倍"}], "mode": "formula"},
        ])
        ws = wb["明细"]
        # 删掉 5 行“测试”后，编号 A003（参照表第 5 行，公式 =C5*2）落在主表第 4 行；
        # 从参照表 E 列复制到主表 H 列：行 -1、列 +3 → =F4*2
        self.assertEqual(ws["A4"].value, "A003")
        self.assertEqual(ws["H4"].value, "=F4*2")
        self.assertEqual(wb["汇总"]["B2"].value, "=SUM('明细'!F2:F16)")
        self.assertEqual(ws["F4"].value, "=D4*E4")

    def test_sort_dedupe_replace(self):
        rep, wb = self.ok([
            {"op": "sort_rows", "by": [{"column": "数量", "order": "desc"}]},
            {"op": "replace_text", "columns": ["名称"], "find": "产品", "replace": "商品"},
            {"op": "dedupe_rows", "keys": ["部门"]},
        ])
        ws = wb["明细"]
        self.assertEqual(ws["D2"].value, 20)
        self.assertEqual(ws["F2"].value, "=D2*E2")
        self.assertEqual(ws["B2"].value, "商品19")
        self.assertEqual(ws.max_row, 5)

    def test_copy_range_and_sheet(self):
        rep, wb = self.ok([
            {"op": "copy_range", "from": {"file": "ref", "sheet": "块", "range": "A1:B3"}, "to": {"cell": "J1"}},
            {"op": "copy_sheet", "from": {"file": "ref", "sheet": "上期"}, "name": "上期数据"},
        ])
        ws = wb["明细"]
        self.assertEqual(ws["K1"].value, "=SUM(K2:K3)")
        self.assertEqual(ws["J1"].font.color.rgb, "FFFF0000")
        self.assertIn("J2:J3", [str(r) for r in ws.merged_cells.ranges])
        self.assertEqual(wb["上期数据"]["E2"].value, "=C2*2")

    def test_validate_and_failure_rollback(self):
        rep, wb = self.ok([{"op": "validate", "checks": [
            {"type": "unique", "columns": ["编号"]},
            {"type": "not_empty", "columns": ["备注"]},
            {"type": "condition", "where": {"col": "数量", "op": "<", "value": 0}, "expect": "none"},
        ]}])
        self.assertEqual([c["ok"] for c in rep.checks], [True, False, True])
        bad = self.go([{"op": "fill", "column": "不存在的列", "value": 1}])
        self.assertFalse(bad.ok)
        self.assertIn("找不到列「不存在的列」", bad.error)
        self.assertIsNone(bad.output)
        self.assertEqual(len(list(self.out.glob("*.xlsx"))), 1, "失败时不应生成结果文件")
        stop = self.go([{"op": "validate", "checks": [{"type": "not_empty", "columns": ["备注"], "on_fail": "stop"}]}])
        self.assertFalse(stop.ok)

    def test_config_errors(self):
        with self.assertRaises(ConfigError):
            validate(base_cfg([{"op": "nope"}]))
        with self.assertRaises(ConfigError):
            validate(base_cfg([{"op": "fill", "column": "A"}]))
        with self.assertRaises(ConfigError):
            validate(base_cfg([{"op": "copy_matched", "match": "x", "columns": ["A"]}]))
        with self.assertRaises(ConfigError):
            validate(base_cfg([{"op": "format", "columns": "A", "fill": "红色啊"}]))
        with self.assertRaises(ConfigError):
            validate(base_cfg([{"op": "fill", "column": "A", "value": 1, "file": "ref"}]))

    def test_written_formula_syntax_check(self):
        rep, wb = self.ok([{"op": "insert_column", "header": "新", "end": True, "formula": "=TEXTJOIN(\",\",TRUE,A{row}:B{row})"}])
        self.assertTrue(any("TEXTJOIN" in m for _, m in rep.formula["problems"]))


if __name__ == "__main__":
    unittest.main()
