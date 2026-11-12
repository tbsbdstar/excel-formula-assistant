"""读取并检查工作流配置（JSON）"""
from __future__ import annotations

import json
import re
from pathlib import Path

from .errors import ConfigError
from .ops import REGISTRY, load_all, validate_step

SCHEMA = "excel-workflow/1"


def load(path) -> dict:
    p = Path(path)
    try:
        text = p.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        raise ConfigError(f"找不到配置文件：{p}")
    except UnicodeDecodeError:
        text = p.read_text(encoding="gb18030")
    try:
        cfg = json.loads(text)
    except json.JSONDecodeError as e:
        raise ConfigError(f"配置文件不是合法的 JSON（第 {e.lineno} 行第 {e.colno} 列）：{e.msg}")
    validate(cfg)
    cfg["_path"] = str(p)
    return cfg


def validate(cfg: dict):
    load_all()
    if not isinstance(cfg, dict):
        raise ConfigError("配置文件内容应是一个 JSON 对象")
    schema = cfg.get("schema", SCHEMA)
    if not str(schema).startswith("excel-workflow/"):
        raise ConfigError(f"不是 Excel 工作流配置（schema={schema}）")
    if schema != SCHEMA:
        raise ConfigError(f"配置版本 {schema} 比执行器新，请更新执行器（当前支持 {SCHEMA}）")
    files = cfg.get("files")
    if not isinstance(files, list) or not files:
        raise ConfigError("缺少 files（要处理的文件），至少要有一个主表")
    ids = set()
    mains = 0
    for i, f in enumerate(files, 1):
        if not isinstance(f, dict) or not f.get("id"):
            raise ConfigError(f"files 第 {i} 项缺少 id")
        if not re.fullmatch(r"[A-Za-z0-9_\-一-鿿]+", str(f["id"])):
            raise ConfigError(f"files 第 {i} 项的 id「{f['id']}」只能包含字母、数字、下划线或中文")
        if f["id"] in ids:
            raise ConfigError(f"files 里 id「{f['id']}」重复")
        ids.add(f["id"])
        role = f.get("role", "ref")
        if role not in ("main", "ref"):
            raise ConfigError(f"files「{f['id']}」的 role 只能是 main（主表，输出结果）或 ref（参照表，只读）")
        mains += role == "main"
        hr = f.get("header_row", 1)
        if not isinstance(hr, int) or hr < 1:
            raise ConfigError(f"files「{f['id']}」的 header_row 应是 ≥1 的整数")
        if f.get("multiple") and role != "ref":
            raise ConfigError(f"files「{f['id']}」：只有参照表可以设置 multiple（多个文件）")
    if mains != 1:
        raise ConfigError(f"files 里必须恰好有一个 role 为 main 的主表（现在有 {mains} 个）")
    multi = {f["id"] for f in files if f.get("multiple")}
    steps = cfg.get("steps")
    if not isinstance(steps, list) or not steps:
        raise ConfigError("缺少 steps（处理步骤）")
    match_ids = set()
    for i, s in enumerate(steps, 1):
        validate_step(s, i)
        if s.get("enabled") is False:
            continue
        for key in ("file",):
            if key in s and s[key] not in ids:
                raise ConfigError(f"第 {i} 步的 file「{s[key]}」不在 files 里")
        if s["op"] == "match":
            if s["ref"] not in ids:
                raise ConfigError(f"第 {i} 步匹配的参照表「{s['ref']}」不在 files 里")
            match_ids.add(s["id"])
        if "match" in s and s["op"] != "match" and s["match"] not in match_ids:
            raise ConfigError(f"第 {i} 步用到的匹配「{s['match']}」在它之前没有定义（先加一个 id 为 {s['match']} 的“建立匹配”步骤）")
        if REGISTRY[s["op"]].writes and s.get("file", "main") != _main_id(files):
            raise ConfigError(f"第 {i} 步要修改「{s['file']}」，但只有主表可以被修改（参照表只读）")
        for fid in _file_refs(s):
            if fid in multi and s["op"] != "append_rows":
                raise ConfigError(f"第 {i} 步用到的「{fid}」是“多个文件”，只能在“追加行（多文件合并）”里使用")
    opts = cfg.get("options") or {}
    if opts.get("recalc", "auto") not in ("auto", "off", "required"):
        raise ConfigError("options.recalc 只能是 auto（有 Excel/WPS 就重算校验）、off、required")


def _file_refs(s):
    """一个步骤引用到的所有文件 id"""
    out = []
    if s.get("file"):
        out.append(s["file"])
    if s.get("op") == "match" and s.get("ref"):
        out.append(s["ref"])
    for key in ("from", "formula_from"):
        v = s.get(key)
        if isinstance(v, dict) and v.get("file"):
            out.append(v["file"])
    for ck in s.get("checks") or []:
        if isinstance(ck, dict) and ck.get("type") == "headers_match":
            out.append(ck.get("file", "ref"))
    return out


def _main_id(files):
    return next(f["id"] for f in files if f.get("role") == "main")


def main_id(cfg):
    return _main_id(cfg["files"])
