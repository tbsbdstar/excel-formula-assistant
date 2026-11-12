"""网页配置器的操作目录（js/wf-catalog.js）必须与执行器实际注册的操作、参数一致。

执行器新增/修改操作参数后忘了同步配置器，这里会失败。
"""
import json
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]          # 项目根目录
EXEC = ROOT / "executor"
sys.path.insert(0, str(EXEC))
sys.path.insert(0, str(ROOT))

from engine import cond  # noqa: E402
from engine.ops import COMMON_KEYS, REGISTRY, load_all  # noqa: E402


def load_catalog():
    text = (ROOT / "js" / "wf-catalog.js").read_text(encoding="utf-8")
    start = text.index("WF_CATALOG = ") + len("WF_CATALOG = ")
    end = text.rindex("};") + 1
    return json.loads(text[start:end])


class CatalogMatchesExecutor(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        load_all()
        cls.cat = load_catalog()
        cls.ops = {d["op"]: d for d in cls.cat["ops"]}
        cls.defs = {}
        for name, d in REGISTRY.items():
            cls.defs.setdefault(d.name, d)

    def test_same_operations(self):
        self.assertEqual(set(self.ops), set(self.defs), "配置器与执行器的操作列表不一致")

    def test_aliases(self):
        for name, d in REGISTRY.items():
            if name != d.name:
                self.assertIn(name, self.ops[d.name].get("aliases", []), f"别名 {name} 未写进配置器目录")

    def test_parameters(self):
        groups = {g["id"] for g in self.cat["groups"]}
        for name, d in self.defs.items():
            c = self.ops[name]
            self.assertIn(c["group"], groups)
            keys = set(c.get("hidden", []))
            req = set()
            for f in c["fields"]:
                ks = f.get("keys", [f["key"]])
                keys.update(ks)
                if f.get("required"):
                    req.add(frozenset(f.get("requiredKeys", ks)))
            allowed = {k for r in d.required for k in r.split("|")} | set(d.optional)
            self.assertEqual(keys, allowed, f"{name}：配置器字段与执行器参数不一致")
            self.assertEqual(req, {frozenset(r.split("|")) for r in d.required}, f"{name}：必填参数不一致")
            self.assertEqual(c.get("writes", True), d.writes, f"{name}：是否修改主表不一致")
            self.assertFalse(keys & COMMON_KEYS, f"{name}：通用字段不应出现在操作字段里")

    def test_condition_operators(self):
        js = (ROOT / "js" / "wf-model.js").read_text(encoding="utf-8")
        ui_ops = set(re.findall(r"\{ v: '([^']+)'", js))
        self.assertEqual(ui_ops, cond.OPS - {"==", "<>"}, "条件判断方式不一致")

    def test_examples_in_sync(self):
        import build
        self.assertEqual((ROOT / "js" / "wf-examples.js").read_text(encoding="utf-8"), build.examples_source(),
                         "js/wf-examples.js 与 executor/templates 不同步，请运行 python build.py")


if __name__ == "__main__":
    unittest.main()
