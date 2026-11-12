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


# ---------------------------------------------------------------- 操作手册：Markdown → 单个 html
MANUAL_MD = ROOT / "操作手册.md"


def _inline(text: str) -> str:
    s = html_escape(text)
    s = re.sub(r"`([^`]+)`", r"<code>\1</code>", s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r'<a href="\2">\1</a>', s)
    return s


def html_escape(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def md_to_html(md: str):
    """够用的 Markdown 转换（标题、段落、列表含嵌套、表格、引用、粗体、行内代码、链接）。返回 (正文 html, 目录项)"""
    lines = md.splitlines()
    out, para, stack, toc = [], [], [], []
    list_re = re.compile(r"^( *)(\d+\.|-)\s+(.*)$")

    def flush_para():
        if para:
            out.append("<p>" + _inline(" ".join(para)) + "</p>")
            para.clear()

    def close_lists(to=-1):
        while stack and stack[-1][0] > to:
            out.append(f"</li></{stack.pop()[1]}>")

    i, n_h2 = 0, 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            flush_para()
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1
            if not (j < len(lines) and list_re.match(lines[j])):
                close_lists()
            i += 1
            continue
        m = list_re.match(line)
        if line.startswith("#"):
            flush_para()
            close_lists()
            level = len(line) - len(line.lstrip("#"))
            text = line[level:].strip()
            if level == 1:
                out.append(f"<h1>{_inline(text)}</h1>")
            elif level == 2:
                n_h2 += 1
                toc.append((f"s{n_h2}", text))
                out.append(f'<h2 id="s{n_h2}">{_inline(text)}</h2>')
            else:
                out.append(f"<h{level}>{_inline(text)}</h{level}>")
        elif line.startswith("|"):
            flush_para()
            close_lists()
            rows = []
            while i < len(lines) and lines[i].startswith("|"):
                rows.append([c.strip() for c in lines[i].strip().strip("|").split("|")])
                i += 1
            i -= 1
            body = [r for r in rows[1:] if not all(re.fullmatch(r":?-{3,}:?", c) for c in r)]
            out.append('<div class="tbl"><table><thead><tr>' + "".join(f"<th>{_inline(c)}</th>" for c in rows[0]) + "</tr></thead><tbody>" +
                       "".join("<tr>" + "".join(f"<td>{_inline(c)}</td>" for c in r) + "</tr>" for r in body) + "</tbody></table></div>")
        elif line.startswith(">"):
            flush_para()
            close_lists()
            quote = []
            while i < len(lines) and lines[i].startswith(">"):
                quote.append(lines[i].lstrip(">").strip())
                i += 1
            i -= 1
            out.append("<blockquote><p>" + _inline(" ".join(quote)) + "</p></blockquote>")
        elif m:
            flush_para()
            indent, tag = len(m.group(1)), ("ol" if m.group(2)[0].isdigit() else "ul")
            if not stack or indent > stack[-1][0]:
                out.append(f"<{tag}><li>")
                stack.append([indent, tag])
            else:
                close_lists(indent)
                if stack and stack[-1][1] != tag:
                    out.append(f"</li></{stack.pop()[1]}><{tag}><li>")
                    stack.append([indent, tag])
                elif stack:
                    out.append("</li><li>")
                else:
                    out.append(f"<{tag}><li>")
                    stack.append([indent, tag])
            out.append(_inline(m.group(3)))
        elif stack and line.startswith(" "):
            out.append(" " + _inline(line.strip()))
        else:
            close_lists()
            para.append(line.strip())
        i += 1
    flush_para()
    close_lists()
    return "\n".join(out), toc


MANUAL_CSS = """
:root{--bg:#f4f6f5;--surface:#fff;--text:#1c2420;--text-2:#56625b;--border:#e2e7e4;--accent:#107c41;--accent-soft:#e6f3ec;--code:#f3f6f4;--th:#f8faf9}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#121614;--surface:#1a1f1c;--text:#e6ebe8;--text-2:#a9b4ae;--border:#2c3430;--accent:#2fa463;--accent-soft:#1d3326;--code:#141916;--th:#1f2522}}
*{box-sizing:border-box}html,body{margin:0}
body{background:var(--bg);color:var(--text);font:15px/1.75 -apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif}
.wrap{max-width:900px;margin:0 auto;padding:28px 20px 60px}
main{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:8px 36px 36px}
h1{font-size:26px;margin:24px 0 4px}h1+p{color:var(--text-2);margin-top:0}
h2{font-size:21px;margin:40px 0 12px;padding-bottom:6px;border-bottom:2px solid var(--accent-soft);scroll-margin-top:16px}
h3{font-size:17px;margin:28px 0 8px;color:var(--accent)}
p{margin:8px 0}ul,ol{padding-left:24px;margin:8px 0}li{margin:4px 0}li>ul,li>ol{margin:4px 0}
code{font:13px Consolas,"Cascadia Mono",monospace;background:var(--code);border:1px solid var(--border);border-radius:4px;padding:1px 5px}
.tbl{overflow-x:auto;margin:12px 0}table{border-collapse:collapse;width:100%;font-size:14px}
th,td{border:1px solid var(--border);padding:7px 10px;text-align:left;vertical-align:top}th{background:var(--th);white-space:nowrap}
blockquote{margin:12px 0;padding:8px 14px;border-left:4px solid var(--accent);background:var(--accent-soft);border-radius:0 8px 8px 0}
blockquote p{margin:0}
nav{background:var(--surface);border:1px solid var(--border);border-radius:14px;padding:14px 20px;margin-bottom:16px}
nav{background:var(--bg);margin:16px 0 8px}
nav b{display:block;margin-bottom:6px}nav ol{margin:0;padding:0;list-style:none;columns:2;column-gap:28px}nav li{margin:2px 0}nav a{color:var(--accent);text-decoration:none}nav a:hover{text-decoration:underline}
@media (max-width:640px){.wrap{padding:16px}main{padding:4px 16px 24px}nav ol{columns:1}h1{font-size:22px}}
@media print{body{background:#fff}nav{display:none}main{border:none;padding:0}h2{break-after:avoid}table{break-inside:avoid}}
"""


def build_manual(out_path: Path) -> Path:
    body, toc = md_to_html(MANUAL_MD.read_text(encoding="utf-8"))
    nav = "<nav><b>目录</b><ol>" + "".join(f'<li><a href="#{i}">{html_escape(t)}</a></li>' for i, t in toc) + "</ol></nav>"
    # 目录放在标题和版本说明之后
    k = body.find("</p>")
    body = body[:k + 4] + nav + body[k + 4:] if k >= 0 else nav + body
    page = ("<!doctype html>\n<html lang=\"zh-CN\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">"
            "<title>Excel 智能工具操作手册</title><style>" + MANUAL_CSS + "</style></head><body><div class=\"wrap\">" +
            "<main>" + body + "</main></div></body></html>\n")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(page, encoding="utf-8")
    return out_path


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
    m = build_manual(OUT_DIR / "操作手册.html")
    print(f"已生成 {m.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
