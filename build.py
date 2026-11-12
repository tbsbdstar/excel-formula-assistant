"""把页面及其引用的 CSS / JS / 组件库打包成独立的单个 html 文件。

用法：python build.py
输出：
  dist/Excel智能函数助手.html   公式助手（V1）
  dist/Excel工作流配置器.html   工作流配置器（V2.0）
双击即可使用，可直接通过微信、邮件发送。

另外会把 executor/templates/*.json（执行器自带的示例工作流）同步到 js/wf-examples.js，供配置器“从示例开始”使用。
"""
import base64
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT_DIR = ROOT / "dist"
PAGES = [
    ("index.html", "Excel智能函数助手.html"),
    ("workflow.html", "Excel工作流配置器.html"),
]
EXAMPLES_JS = ROOT / "js" / "wf-examples.js"


def read(rel: str) -> str:
    path = (ROOT / rel).resolve()
    if ROOT not in path.parents:
        sys.exit(f"引用的文件不在项目目录内：{rel}")
    return path.read_text(encoding="utf-8")


def needs_base64(code: str) -> bool:
    """内联后可能破坏 HTML 解析的脚本，改用 base64 嵌入"""
    low = code.lower()
    return "</script" in low or ("<!--" in code and "<script" in low)


def inline_script(match: re.Match) -> str:
    src = match.group(1)
    code = read(src)
    if needs_base64(code):
        b64 = base64.b64encode(code.encode("utf-8")).decode("ascii")
        return (
            f'<script>/* {src} */(function(){{var b=atob("{b64}"),u=new Uint8Array(b.length);'
            "for(var i=0;i<b.length;i++)u[i]=b.charCodeAt(i);"
            "(0,eval)(new TextDecoder().decode(u));})();</script>"
        )
    return f"<script>/* {src} */\n{code}\n</script>"


def inline_style(match: re.Match) -> str:
    href = match.group(1)
    return f"<style>/* {href} */\n{read(href)}\n</style>"


def examples_source() -> str:
    """executor/templates/*.json → js/wf-examples.js 的内容"""
    items = []
    for p in sorted((ROOT / "executor" / "templates").glob("*.json")):
        items.append(json.loads(p.read_text(encoding="utf-8-sig")))
    body = json.dumps(items, ensure_ascii=False, indent=1)
    return ("/* 由 build.py 根据 executor/templates/*.json 自动生成，请勿手改（改示例请改 executor/templates）。 */\n"
            "(window.EFA = window.EFA || {}).WF_EXAMPLES = " + body + ";\n")


def sync_examples() -> bool:
    text = examples_source()
    old = EXAMPLES_JS.read_text(encoding="utf-8") if EXAMPLES_JS.exists() else ""
    if old != text:
        EXAMPLES_JS.write_text(text, encoding="utf-8")
        return True
    return False


def build_page(src_name: str, out_name: str) -> Path:
    html = (ROOT / src_name).read_text(encoding="utf-8")
    # 页面之间的链接改成打包后的文件名
    for s, o in PAGES:
        html = html.replace(f'href="{s}"', f'href="{o}"')
    # 先检查页面里只有本脚本能处理的本地引用
    without = re.sub(r'<link rel="stylesheet" href="[^"]+">|<script src="[^"]+"></script>', "", html)
    leftovers = [x for x in re.findall(r'(?:src|href)="(?!data:|https?:|#)([^"]+)"', without) if x not in {o for _, o in PAGES}]
    if leftovers:
        sys.exit(f"{src_name} 里有本脚本无法内联的本地引用：{leftovers}")
    html = re.sub(r'<link rel="stylesheet" href="([^"]+)">', inline_style, html)
    html = re.sub(r'<script src="([^"]+)"></script>', inline_script, html)
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / out_name
    out.write_text(html, encoding="utf-8")
    return out


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if sync_examples():
        print(f"已更新 {EXAMPLES_JS.relative_to(ROOT)}")
    for src, out_name in PAGES:
        out = build_page(src, out_name)
        print(f"已生成 {out.relative_to(ROOT)}（{out.stat().st_size / 1024 / 1024:.1f} MB）")


if __name__ == "__main__":
    main()
