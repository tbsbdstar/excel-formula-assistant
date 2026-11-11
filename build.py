"""把 index.html 及其引用的 CSS / JS / 组件库打包成一个独立的 html 文件。

用法：python build.py
输出：dist/Excel智能函数助手.html（双击即可使用，可直接通过微信、邮件发送）
"""
import base64
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "index.html"
OUT_DIR = ROOT / "dist"
OUT = OUT_DIR / "Excel智能函数助手.html"


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


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    html = SRC.read_text(encoding="utf-8")
    # 先检查页面里只有本脚本能处理的本地引用
    without = re.sub(r'<link rel="stylesheet" href="[^"]+">|<script src="[^"]+"></script>', "", html)
    leftovers = re.findall(r'(?:src|href)="(?!data:|https?:|#)([^"]+)"', without)
    if leftovers:
        sys.exit(f"页面里有本脚本无法内联的本地引用：{leftovers}")
    html = re.sub(r'<link rel="stylesheet" href="([^"]+)">', inline_style, html)
    html = re.sub(r'<script src="([^"]+)"></script>', inline_script, html)
    OUT_DIR.mkdir(exist_ok=True)
    OUT.write_text(html, encoding="utf-8")
    print(f"已生成 {OUT.relative_to(ROOT)}（{OUT.stat().st_size / 1024 / 1024:.1f} MB）")


if __name__ == "__main__":
    main()
