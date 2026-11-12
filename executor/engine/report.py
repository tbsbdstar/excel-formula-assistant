"""校验报告（给人看的 .txt）与执行日志"""
from __future__ import annotations

import datetime as dt
from pathlib import Path


class Report:
    def __init__(self, cfg):
        self.cfg = cfg
        self.started = dt.datetime.now()
        self.finished = None
        self.ok = False
        self.files = []          # (角色, 标签, 路径, 工作表, 表头行, 数据行数)
        self.steps = []          # dict(no, label, title, status, seconds, detail, warnings)
        self.warnings = []       # 全局警告
        self.match_stats = {}
        self.checks = []
        self.formula = {}        # 公式检查结果
        self.output = None
        self.extra_outputs = []
        self.error = None
        self.log_lines = []

    def log(self, msg):
        self.log_lines.append(f"{dt.datetime.now():%H:%M:%S} {msg}")

    def render(self) -> str:
        L = []
        name = self.cfg.get("name") or "未命名工作流"
        dur = ((self.finished or dt.datetime.now()) - self.started).total_seconds()
        L.append("Excel 工作流执行报告")
        L.append("=" * 40)
        L.append(f"工作流：{name}")
        if self.cfg.get("_path"):
            L.append(f"配置文件：{self.cfg['_path']}")
        L.append(f"执行时间：{self.started:%Y-%m-%d %H:%M:%S}（用时 {dur:.1f} 秒）")
        L.append(f"结果：{'成功 ✔' if self.ok else '失败 ✘（没有生成结果文件，原文件未改动）'}")
        if self.output:
            L.append(f"输出文件：{self.output}")
        for p in self.extra_outputs:
            L.append(f"附加文件：{p}")
        if self.error:
            L.append("")
            L.append(f"失败原因：{self.error}")
        L.append("")
        L.append("【输入文件】")
        for role, label, path, sheet, hr, rows in self.files:
            L.append(f"  {'主表' if role == 'main' else '参照表'}「{label}」：{path}")
            L.append(f"      工作表「{sheet}」，表头第 {hr} 行，数据 {rows} 行")
        L.append("")
        L.append("【处理步骤】")
        for s in self.steps:
            mark = {"ok": "✔", "fail": "✘", "skip": "－"}[s["status"]]
            title = f"{s['label']}" + (f"（{s['title']}）" if s.get("title") else "")
            L.append(f"  {mark} 第 {s['no']} 步  {title}  {s['seconds']:.2f} 秒")
            if s.get("detail"):
                L.append(f"        {s['detail']}")
            for w in s.get("warnings", []):
                L.append(f"        ⚠ {w}")
        if self.match_stats:
            L.append("")
            L.append("【匹配统计】")
            for mid, st in self.match_stats.items():
                L.append(f"  {mid}：" + "，".join(f"{k} {v}" for k, v in st.items()))
        if self.checks:
            L.append("")
            L.append("【数据校验】")
            for c in self.checks:
                L.append(f"  {'✔' if c['ok'] else '✘'} {c['name']}：{c['detail']}")
        if self.formula:
            f = self.formula
            L.append("")
            L.append("【公式检查】")
            L.append(f"  本次写入公式 {f.get('written', 0)} 个，语法检查 {f.get('checked', 0)} 个，发现问题 {len(f.get('problems', []))} 处")
            for where, msg in f.get("problems", [])[:30]:
                L.append(f"    ✘ {where}：{msg}")
            rc = f.get("recalc")
            if rc:
                L.append(f"  重算校验：{rc}")
            if "value_errors" in f:
                L.append(f"  重算后错误值：{f['value_errors']} 处；#N/A（通常是查找不到）{f['value_na']} 处")
                for where, v in f.get("value_examples", [])[:30]:
                    L.append(f"    {'⚠' if v == '#N/A' else '✘'} {where}：{v}")
            L.append("  说明：公式的最终结果在用 Excel/WPS 打开文件时计算。")
        if self.warnings:
            L.append("")
            L.append("【注意事项】")
            for w in self.warnings:
                L.append(f"  ⚠ {w}")
        L.append("")
        return "\n".join(L)

    def write(self, out_dir: Path, stem: str):
        out_dir.mkdir(parents=True, exist_ok=True)
        rp = out_dir / f"{stem}_校验报告.txt"
        rp.write_text(self.render(), encoding="utf-8-sig")
        lp = out_dir / f"{stem}_执行日志.log"
        lp.write_text("\n".join(self.log_lines) + "\n", encoding="utf-8-sig")
        return rp, lp
