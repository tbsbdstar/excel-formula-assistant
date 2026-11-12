"""手动运行的性能与重算测试（会启动本机 Excel/WPS）：python tests/perf_and_recalc.py [行数]"""
import random
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8")

from openpyxl import Workbook  # noqa: E402

from engine.runner import run  # noqa: E402

N = int(sys.argv[1]) if len(sys.argv) > 1 else 10000
d = Path(tempfile.mkdtemp(prefix="efa_perf_"))
rnd = random.Random(1)

t = time.time()
wb = Workbook()
ws = wb.active
ws.title = "明细"
ws.append(["详情ID", "采购单", "条码", "名称", "数量", "单价", "未分配数", "金额", "备注", "出货单创建日"])
for i in range(N):
    r = i + 2
    ws.append([f"D{i}", f"P{i % 997}", f"690{i:07d}", f"商品{i}", rnd.randint(0, 50), rnd.randint(1, 500),
               rnd.randint(0, 30), f"=E{r}*F{r}", None, "2024-01-01"])
wb.save(d / "main.xlsx")
ref = Workbook()
rs = ref.active
rs.append(["详情ID", "采购单", "条码", "下单时间", "入库时间", "折扣金额"])
for i in range(0, N, 2):
    r = rs.max_row + 1
    rs.append([f"D{i}", f"P{i % 997}", f"690{i:07d}", "2024-02-01", "2024-02-05", f"=H{r}*0.9"])
ref.save(d / "ref.xlsx")
print(f"生成测试文件 {N} 行：{time.time() - t:.1f}s")

cfg = {
    "schema": "excel-workflow/1", "name": "性能测试",
    "files": [{"id": "main", "role": "main"}, {"id": "ref", "role": "ref"}],
    "options": {"recalc": sys.argv[2] if len(sys.argv) > 2 else "auto"},
    "steps": [
        {"op": "delete_columns", "columns": ["出货单创建日"]},
        {"op": "insert_column", "header": "淘宝数量", "after": "数量", "value": 0},
        {"op": "format", "columns": ["未分配数"], "where": {"col": "未分配数", "op": ">=", "value": 10}, "fill": "FCE4D6"},
        {"op": "match", "id": "m", "ref": "ref", "main_keys": ["详情ID", "采购单", "条码"]},
        {"op": "copy_matched", "match": "m", "columns": ["下单时间", "入库时间"], "create_missing": True,
         "where": {"col": "未分配数", "op": ">=", "value": 10}, "fill": "FCE4D6", "color_columns": "A:D"},
        {"op": "fill", "column": "备注", "where": {"all": [{"col": "未分配数", "op": ">=", "value": 10},
                                                          {"col": "未分配数", "op": "=", "col2": "数量"}]}, "value": "逾期"},
        {"op": "insert_column", "header": "含税金额", "end": True, "formula": "=ROUND([@金额]*1.13,2)"},
        {"op": "validate", "checks": [{"type": "unique", "columns": ["详情ID"]}]},
    ],
}
t = time.time()
rep = run(cfg, {"main": str(d / "main.xlsx"), "ref": str(d / "ref.xlsx")}, d / "out", progress=lambda m: print(f"  {time.time() - t:6.1f}s {m}"))
print(f"总用时 {time.time() - t:.1f}s，成功={rep.ok}")
print(rep.render())
