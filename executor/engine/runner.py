"""执行工作流：打开文件 → 逐步执行 → 公式检查 → 安全保存 → 写报告。
任一步失败：不生成结果文件，输入文件始终不改动。"""
from __future__ import annotations

import csv
import datetime as dt
import importlib.util
import os
import sys
import time
import traceback
from pathlib import Path

from . import config as cfgmod
from . import scan
from .book import FileHandle
from .errors import ConfigError, StepError
from .ops import REGISTRY, load_all
from .report import Report


class Context:
    def __init__(self, files, main, report, multi=None):
        self.files = files
        self.multi = multi or {}
        self.main = main
        self.report = report
        self.matches = {}
        self.extra_csv = []
        self._step_warnings = None

    def file(self, fid):
        if fid in self.multi:
            raise StepError(f"「{fid}」是“多个文件”，只能在“追加行（多文件合并）”里使用")
        if fid not in self.files:
            raise StepError(f"配置里没有 id 为「{fid}」的文件")
        return self.files[fid]

    def files_of(self, fid):
        """单个文件返回 [文件]；“多个文件”返回全部匹配到的文件"""
        if fid in self.multi:
            return self.multi[fid]
        return [self.file(fid)]

    def view(self, step):
        fh = self.file(step.get("file", self.main.id))
        return fh.view(step.get("sheet"), step.get("header_row"))

    def warn(self, msg):
        target = self._step_warnings if self._step_warnings is not None else self.report.warnings
        if msg not in target:
            target.append(msg)


def load_extensions(folder):
    """加载 extensions/ 目录里的自定义操作（.py 文件）。只放可信来源的文件"""
    loaded = []
    folder = Path(folder)
    if not folder.is_dir():
        return loaded
    for p in sorted(folder.glob("*.py")):
        spec = importlib.util.spec_from_file_location(f"extensions.{p.stem}", p)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        loaded.append(p.name)
    return loaded


def output_name(cfg, main_path):
    pattern = (cfg.get("output") or {}).get("file_name") or "{main}_处理结果_{time}.xlsx"
    now = dt.datetime.now()
    name = pattern.format(main=Path(main_path).stem, name=cfg.get("name") or "工作流",
                          time=now.strftime("%Y%m%d_%H%M%S"), date=now.strftime("%Y%m%d"))
    name = "".join("_" if ch in '\\/:*?"<>|' else ch for ch in name)
    return name if name.lower().endswith(".xlsx") else name + ".xlsx"


def run(cfg: dict, paths: dict, out_dir, recalc=None, progress=print) -> Report:
    """paths: {文件id: 路径（“多个文件”为路径列表）}；recalc: None 按配置；'off'/'auto'/'required' 覆盖"""
    load_all()
    report = Report(cfg)
    out_dir = Path(out_dir)
    stem = None
    try:
        cfgmod.validate(cfg)
        main_id = cfgmod.main_id(cfg)
        files, multi = {}, {}
        for spec in cfg["files"]:
            p = paths.get(spec["id"])
            label = spec.get("label") or spec["id"]
            plist = [x for x in (p if isinstance(p, (list, tuple)) else [p]) if x]
            if not plist:
                raise StepError(f"没有提供「{label}」的文件")
            if len(plist) > 1 and not spec.get("multiple"):
                raise StepError(f"「{label}」只能对应一个文件，现在给了 {len(plist)} 个")
            for x in plist:
                if not Path(x).is_file():
                    raise StepError(f"找不到文件：{x}")
            if spec.get("multiple"):
                multi[spec["id"]] = [FileHandle(dict(spec, label=f"{label}：{Path(x).name}"), str(x), writable=False) for x in plist]
            else:
                files[spec["id"]] = FileHandle(spec, str(plist[0]), writable=spec["id"] == main_id)
        main = files[main_id]
        stem = Path(output_name(cfg, main.path)).stem
        progress(f"正在读取文件…")
        for fh in list(files.values()) + [h for hs in multi.values() for h in hs]:
            t0 = time.time()
            fh.open()
            v = fh.view()
            report.files.append((fh.role, fh.label, fh.path, v.name, v.header_row, len(v.data_rows())))
            report.log(f"读取 {fh.path} 用时 {time.time() - t0:.1f}s")
            if fh.writable and fh.dropped_features:
                report.warnings.append(f"主表中的{'、'.join(fh.dropped_features)}在结果文件中不会保留（执行器不支持这些内容）")
            if fh.no_cached_values:
                report.warnings.append(f"「{fh.label}」里的公式没有保存计算结果（文件可能不是用 Excel/WPS 保存的），条件和匹配中用到公式列时会按空值处理。用 Excel/WPS 打开后保存一次即可")
        ctx = Context(files, main, report, multi)
        steps = [s for s in cfg["steps"] if s.get("enabled", True) is not False]
        for i, step in enumerate(cfg["steps"], 1):
            d = REGISTRY[step["op"]]
            if step.get("enabled", True) is False:
                report.steps.append({"no": i, "label": d.label, "title": step.get("title"), "status": "skip",
                                     "seconds": 0, "detail": "已停用", "warnings": []})
                continue
            progress(f"  第 {i}/{len(cfg['steps'])} 步：{step.get('title') or d.label}")
            ctx._step_warnings = []
            t0 = time.time()
            entry = {"no": i, "label": d.label, "title": step.get("title"), "status": "ok", "warnings": ctx._step_warnings}
            report.steps.append(entry)
            try:
                res = d.fn(ctx, step) or {}
            except StepError as e:
                entry.update(status="fail", seconds=time.time() - t0, detail=str(e))
                raise StepError(f"第 {i} 步（{step.get('title') or d.label}）：{e}")
            except Exception as e:  # noqa: BLE001
                report.log(traceback.format_exc())
                entry.update(status="fail", seconds=time.time() - t0, detail=f"程序出错：{e}")
                raise StepError(f"第 {i} 步（{step.get('title') or d.label}）程序出错：{e}（详情见执行日志）")
            entry.update(seconds=time.time() - t0, detail=res.get("detail", ""))
            report.log(f"第 {i} 步 {d.name} {entry['seconds']:.2f}s {entry['detail']}")
        ctx._step_warnings = None

        # 公式语法检查（只检查本次写入的公式）
        progress("正在检查公式…")
        checked, problems = scan.syntax_check(main.wb, only_ids=main.written_formulas)
        report.formula = {"written": len(main.written_formulas), "checked": checked, "problems": problems}

        # 安全保存：先写临时文件，全部成功后再改名
        out_dir.mkdir(parents=True, exist_ok=True)
        final = out_dir / f"{stem}.xlsx"
        tmp = out_dir / f"~{stem}.tmp.xlsx"
        positions = scan.formula_positions(main.wb)
        progress("正在保存结果…")
        main.wb.save(tmp)
        mode = recalc or (cfg.get("options") or {}).get("recalc", "auto")
        if mode != "off" and positions:
            progress("正在用 Excel/WPS 重算公式（可能需要几十秒）…")
            status, info = scan.recalc(str(tmp))
            if status == "ok":
                n_err, n_na, ex = scan.scan_values(str(tmp), positions)
                report.formula.update(recalc=f"已用 {info} 重算", value_errors=n_err, value_na=n_na, value_examples=ex)
            elif status == "noapp":
                if mode == "required":
                    raise StepError("配置要求重算校验，但本机没有 Excel 或 WPS")
                report.formula["recalc"] = f"未执行（{info}），只做了语法检查"
            else:
                if mode == "required":
                    raise StepError(f"重算校验失败：{info}")
                report.formula["recalc"] = f"未完成（{info}），只做了语法检查"
        if final.exists():
            final.unlink()
        os.replace(tmp, final)
        report.output = str(final)
        for name, rows in ctx.extra_csv:
            p = out_dir / f"{stem}_{name}.csv"
            with open(p, "w", newline="", encoding="utf-8-sig") as fp:
                csv.writer(fp).writerows(rows)
            report.extra_outputs.append(str(p))
        report.ok = True
    except (ConfigError, StepError) as e:
        report.error = str(e)
    except Exception as e:  # noqa: BLE001
        report.error = f"程序出错：{e}"
        report.log(traceback.format_exc())
    finally:
        report.finished = dt.datetime.now()
        try:
            for p in out_dir.glob("~*.tmp.xlsx"):
                p.unlink()
        except OSError:
            pass
        if stem is None:
            stem = f"工作流_{dt.datetime.now():%Y%m%d_%H%M%S}"
        report.report_path, report.log_path = report.write(out_dir, stem)
    return report
