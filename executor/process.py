"""Excel 工作流执行器入口。

双击 run.bat 运行（交互模式）：选择工作流 → 自动在 input 目录找文件（找不到就提示拖入）→ 执行 → 结果在 output 目录。
也可以把文件直接拖到 run.bat 图标上。

命令行（高级）：
  process.py --config configs/xx.json --file main=路径 --file ref=路径 [--yes] [--no-recalc] [--out 目录]
  process.py --check configs/xx.json       只检查配置格式
  process.py --list-ops                     列出所有可用操作
"""
from __future__ import annotations

import argparse
import fnmatch
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
for p in (ROOT / "libs", ROOT):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

for _s in (sys.stdout, sys.stderr, sys.stdin):
    try:
        _s.reconfigure(encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass

try:
    import openpyxl  # noqa: F401
except ImportError:
    print("缺少 openpyxl 组件：请确认 libs 文件夹完整，或用工具包自带的 run.bat 启动。")
    sys.exit(3)

from engine import __version__, config as cfgmod  # noqa: E402
from engine.errors import ConfigError  # noqa: E402
from engine.ops import REGISTRY, load_all  # noqa: E402
from engine.runner import load_extensions, run  # noqa: E402

CONFIG_DIR = ROOT / "configs"
INPUT_DIR = ROOT / "input"
OUTPUT_DIR = ROOT / "output"
SAMPLE_DIR = ROOT / "templates" / "示例数据"
EXCEL_EXT = (".xlsx", ".xlsm")


class Abort(Exception):
    pass


def ask(prompt):
    try:
        return input(prompt)
    except EOFError:
        raise Abort()


def clean_path(s):
    s = s.strip().strip('"').strip("'").strip()
    return s


def choose(items, title, fmt=str):
    print(title)
    for i, it in enumerate(items, 1):
        print(f"  {i}. {fmt(it)}")
    while True:
        a = ask("请输入序号后回车：").strip()
        if a.isdigit() and 1 <= int(a) <= len(items):
            return items[int(a) - 1]
        print("  序号不对，请重新输入。")


def pick_example():
    """configs 为空时：说明配置放哪里，并允许直接运行内置示例（用示例数据）。返回示例配置路径或 None"""
    print("configs 文件夹里还没有工作流配置。")
    print("工作流配置（.json）写明“处理哪些表格、按顺序做哪些操作”。")
    if (ROOT / "工作流配置器.html").exists():
        print("用本文件夹里的「工作流配置器.html」（双击用浏览器打开）配置并保存，再把 .json 放进：")
    else:
        print("用网页“工作流配置器”配置并保存，再把 .json 放进：")
    print(f"  {CONFIG_DIR}")
    tpl = sorted((ROOT / "templates").glob("*.json"))
    if not tpl:
        return None
    print("\n现在可以先运行内置示例体验一下（使用自带的示例数据；结果保存到 output 文件夹）：")
    for i, t in enumerate(tpl, 1):
        try:
            c = cfgmod.load(t)
            name, desc = c.get("name") or t.stem, c.get("description") or ""
        except ConfigError:
            name, desc = t.stem, ""
        print(f"  {i}. {name}")
        if desc:
            print(f"     {desc}")
    while True:
        a = ask("\n输入序号后回车运行示例；直接回车退出：").strip()
        if not a:
            return None
        if a.isdigit() and 1 <= int(a) <= len(tpl):
            return tpl[int(a) - 1]
        print("  序号不对，请重新输入。")


def pick_config():
    cfgs = sorted(p for p in CONFIG_DIR.glob("*.json"))
    if not cfgs:
        return None
    if len(cfgs) == 1:
        return cfgs[0]

    def label(p):
        try:
            c = cfgmod.load(p)
            return f"{c.get('name') or p.stem}（{p.name}）"
        except ConfigError:
            return f"{p.name}（配置有误）"

    return choose(cfgs, "\n请选择要运行的工作流：", label)


def split_paths(s):
    """拖进 cmd 窗口的多个文件：路径之间用空格分隔，带空格的路径会加引号"""
    out, cur, quoted = [], "", False
    for ch in s:
        if ch == '"':
            quoted = not quoted
        elif ch == " " and not quoted:
            if cur:
                out.append(cur)
            cur = ""
        else:
            cur += ch
    if cur:
        out.append(cur)
    return [clean_path(x) for x in out if clean_path(x)]


def find_multi(spec, pool, inputs, used):
    """“多个文件”：符合规则的全部文件；一个都没有时提示拖入（可一次拖多个）"""
    label = spec.get("label") or spec["id"]
    pattern = spec.get("match")
    cands = []
    for src in (pool, inputs):
        for p in src:
            if str(p) in used or (pattern and not fnmatch.fnmatch(p.name.lower(), pattern.lower())):
                continue
            if str(p) not in {str(x) for x in cands}:
                cands.append(p)
    while not cands:
        hint = f"（文件名应符合：{pattern}）" if pattern else ""
        s = ask(f"\n没有找到「{label}」的文件。请把这些文件拖进本窗口（可以一次选中多个一起拖），然后按回车{hint}：\n> ")
        got = [Path(x) for x in split_paths(s)]
        bad = [p for p in got if not p.is_file() or p.suffix.lower() not in EXCEL_EXT]
        if bad:
            print("  这些不是可用的 .xlsx 文件：" + "、".join(p.name for p in bad))
        cands = [p for p in got if p not in bad]
    cands.sort(key=lambda p: p.name)
    print(f"  「{label}」→ {len(cands)} 个文件：" + "、".join(p.name for p in cands[:10]) + (" …" if len(cands) > 10 else ""))
    for p in cands:
        used.add(str(p))
    return [str(p) for p in cands]


def find_files(cfg, dropped, input_dir=INPUT_DIR):
    """为每个文件角色找到实际文件：拖入的文件 > input 目录按规则匹配 > 提示拖入"""
    result = {}
    pool = [Path(p) for p in dropped if Path(p).suffix.lower() in EXCEL_EXT]
    inputs = sorted((p for p in input_dir.glob("*") if p.suffix.lower() in EXCEL_EXT and not p.name.startswith("~$")),
                    key=lambda p: p.stat().st_mtime, reverse=True) if input_dir.is_dir() else []
    used = set()
    # 先定主表和单个文件，最后再收集“多个文件”，避免主表也被算进去
    order = sorted(cfg["files"], key=lambda f: (f.get("role") != "main", bool(f.get("multiple"))))
    for spec in order:
        if spec.get("multiple"):
            result[spec["id"]] = find_multi(spec, pool, inputs, used)
            continue
        label = spec.get("label") or spec["id"]
        pattern = spec.get("match")
        cands = []
        for src in (pool, inputs):
            for p in src:
                if str(p) in used:
                    continue
                if pattern and not fnmatch.fnmatch(p.name.lower(), pattern.lower()):
                    continue
                cands.append(p)
            if cands:
                break
        if not pattern and len(cands) > 1 and len(pool) == len(cfg["files"]):
            cands = [p for p in pool if str(p) not in used][:1]
        chosen = None
        same_as_multi = pattern and any(f.get("multiple") and (f.get("match") or "").lower() == pattern.lower() for f in cfg["files"])
        if len(cands) == 1:
            chosen = cands[0]
            print(f"  「{label}」→ {chosen.name}")
        elif len(cands) > 1 and same_as_multi:
            # 合并场景：主表与“多个文件”用同一规则，取第一个作主表，其余都会被追加，不用选
            chosen = sorted(cands, key=lambda p: p.name)[0]
            print(f"  「{label}」→ {chosen.name}（与“多个文件”规则相同，取第一个作主表，其余追加进来）")
        elif len(cands) > 1:
            chosen = choose(cands, f"\n「{label}」有多个候选文件：", lambda p: p.name)
        while chosen is None:
            hint = f"（文件名应符合：{pattern}）" if pattern else ""
            s = clean_path(ask(f"\n请把「{label}」文件拖进本窗口，然后按回车{hint}：\n> "))
            if not s:
                continue
            p = Path(s)
            if not p.is_file():
                print("  找不到这个文件，请重新拖入。")
                continue
            if p.suffix.lower() not in EXCEL_EXT:
                print("  只支持 .xlsx 文件。旧版 .xls 请先在 Excel/WPS 里另存为 .xlsx。")
                continue
            chosen = p
        used.add(str(chosen))
        result[spec["id"]] = str(chosen)
    return result


def interactive(dropped):
    print(f"Excel 工作流执行器 v{__version__}")
    print("=" * 40)
    cfg_path = pick_config()
    example = False
    if not cfg_path:
        cfg_path = pick_example()
        if not cfg_path:
            return 1
        example = True
    try:
        cfg = cfgmod.load(cfg_path)
    except ConfigError as e:
        print(f"\n配置文件有误：{e}")
        return 1
    print(f"\n工作流：{cfg.get('name') or cfg_path.stem}")
    if cfg.get("description"):
        print(f"说明：{cfg['description']}")
    print(f"共 {len(cfg['steps'])} 个步骤。\n正在查找文件…")
    paths = find_files(cfg, dropped, SAMPLE_DIR if example and SAMPLE_DIR.is_dir() else INPUT_DIR)
    print("\n开始执行（原文件不会被修改，结果保存到 output 文件夹）")
    rep = run(cfg, paths, OUTPUT_DIR, progress=print)
    print()
    print(rep.render())
    print(f"校验报告：{rep.report_path}")
    if example:
        print("\n以上是示例。处理自己的表格时：把工作流配置（.json）放进 configs 文件夹、要处理的表格放进 input 文件夹，再双击 run.bat。")
        print(f"示例配置在 {cfg_path.parent}，可以复制到 configs 后按需修改（写法见“配置格式说明.md”）。")
    if rep.ok and os.name == "nt":
        try:
            if ask("直接按回车结束；输入 o 再回车可打开输出文件夹：").strip().lower() == "o":
                os.startfile(str(OUTPUT_DIR))
        except Abort:
            pass
    return 0 if rep.ok else 2


def main(argv=None):
    load_all()
    exts = load_extensions(ROOT / "extensions")
    ap = argparse.ArgumentParser(description="Excel 工作流执行器")
    ap.add_argument("files", nargs="*", help="拖到 run.bat 上的文件")
    ap.add_argument("--config")
    ap.add_argument("--file", action="append", default=[], help="文件id=路径")
    ap.add_argument("--out", default=str(OUTPUT_DIR))
    ap.add_argument("--no-recalc", action="store_true")
    ap.add_argument("--yes", action="store_true", help="不交互")
    ap.add_argument("--check")
    ap.add_argument("--list-ops", action="store_true")
    a = ap.parse_args(argv)
    if exts:
        print("已加载扩展操作：" + "、".join(exts))
    if a.list_ops:
        seen = set()
        for name, d in sorted(REGISTRY.items()):
            if d.name in seen:
                continue
            seen.add(d.name)
            print(f"{d.name:16} {d.label}    必填：{'，'.join(d.required) or '无'}    可选：{'，'.join(d.optional) or '无'}")
        return 0
    if a.check:
        try:
            c = cfgmod.load(a.check)
            print(f"配置正确：{c.get('name')}，{len(c['steps'])} 个步骤")
            return 0
        except ConfigError as e:
            print(f"配置有误：{e}")
            return 1
    if a.config:
        cfg = cfgmod.load(a.config)
        paths = {}
        for x in a.file:
            k, v = x.split("=", 1)
            paths.setdefault(k, []).append(v)
        multi = {f["id"] for f in cfg["files"] if f.get("multiple")}
        paths = {k: (v if k in multi else v[0]) for k, v in paths.items()}
        rep = run(cfg, paths, Path(a.out), recalc="off" if a.no_recalc else None,
                  progress=(lambda m: None) if a.yes else print)
        print(rep.render())
        return 0 if rep.ok else 2
    return interactive(a.files)


if __name__ == "__main__":
    try:
        code = main()
    except (KeyboardInterrupt, Abort):
        print("\n已取消。")
        code = 130
    except ConfigError as e:
        print(f"\n配置文件有误：{e}")
        code = 1
    sys.exit(code)
