"""组装绿色便携工具包：嵌入版 Python + openpyxl + 执行器 → dist/Excel工作流工具/ 及 zip

用法（开发机上运行一次即可，需要能访问 python.org 和 PyPI）：
    python build_toolkit.py
"""
from __future__ import annotations

import datetime as dt
import random
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from engine import __version__  # noqa: E402

PY_VER = "3.12.10"
PY_URL = f"https://www.python.org/ftp/python/{PY_VER}/python-{PY_VER}-embed-amd64.zip"
LIBS = ["openpyxl==3.1.5", "et-xmlfile==2.0.0"]
CACHE = HERE / ".cache"
DIST = HERE / "dist"
NAME = "Excel工作流工具"
OUT = DIST / NAME


def log(msg):
    print(msg, flush=True)


def fetch_python():
    CACHE.mkdir(exist_ok=True)
    z = CACHE / Path(PY_URL).name
    if not z.exists():
        log(f"下载嵌入版 Python {PY_VER} …")
        tmp = z.with_suffix(".part")
        urllib.request.urlretrieve(PY_URL, tmp)
        tmp.replace(z)
    return z


def install_libs(target: Path):
    log("安装依赖库 " + "、".join(LIBS) + " …")
    cmd = [sys.executable, "-m", "pip", "install", "--target", str(target), "--no-deps", "--no-compile",
           "--only-binary=:all:", "--disable-pip-version-check", "--quiet", *LIBS]
    subprocess.run(cmd, check=True)
    for p in target.glob("*.dist-info"):
        # 保留许可证文件，删掉其余元数据
        for f in p.iterdir():
            if not f.name.upper().startswith("LICENSE") and f.is_file():
                f.unlink()
    for p in target.rglob("__pycache__"):
        shutil.rmtree(p, ignore_errors=True)


def make_samples(folder: Path):
    """生成虚构的示例数据，供体验 templates 里的示例工作流"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    folder.mkdir(parents=True, exist_ok=True)
    rnd = random.Random(7)
    customers = ["华东商贸", "南方百货", "测试客户", "北方连锁", "西部超市"]
    status = ["已发货", "待发货", "已签收"]
    head_fill = PatternFill("solid", start_color="FFDDEBF7", end_color="FFDDEBF7")

    def header(ws, names):
        ws.append(names)
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = head_fill

    wb = Workbook()
    ws = wb.active
    ws.title = "订单明细"
    header(ws, ["订单号", "商品编码", "客户", "数量", "未分配数", "单价", "金额", "状态", "内部备注"])
    rows = []
    for i in range(60):
        r = i + 2
        qty = rnd.randint(1, 40)
        unalloc = qty if i % 9 == 0 else rnd.randint(0, qty)
        price = rnd.choice([19.9, 59, 128, 399])
        st = rnd.choice(status)
        rows.append((f"SO{2400 + i // 2}", f"SKU{100 + i % 7}", round(qty * price, 2), st))
        ws.append([rows[-1][0], rows[-1][1], customers[i % 5], qty, unalloc, price, f"=D{r}*F{r}", st, "内部"])
    ws.append(list(rows[3][:2]) + ["华东商贸", 5, 0, 59, f"=D{ws.max_row + 1}*F{ws.max_row + 1}", "已发货", "重复行"])
    ws.column_dimensions["A"].width = 12
    wb.save(folder / "本期订单.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "订单明细"
    header(ws, ["订单号", "商品编码", "金额", "状态", "下单时间", "入库时间"])
    for i, (so, sku, amount, st) in enumerate(rows[:45]):
        # 大部分与本期一致，少数金额/状态不同，用来演示比对
        ws.append([so, sku, amount if i % 6 else round(amount * 0.9, 2), st if i % 8 else "待发货",
                   dt.datetime(2026, 10, 1) + dt.timedelta(days=i % 20), dt.datetime(2026, 10, 5) + dt.timedelta(days=i % 20)])
        ws.cell(row=ws.max_row, column=5).number_format = "yyyy-mm-dd"
        ws.cell(row=ws.max_row, column=6).number_format = "yyyy-mm-dd"
    ws.append(["SO9999", "SKU999", 100, "已签收", dt.datetime(2026, 9, 1), None])
    wb.save(folder / "上期订单.xlsx")
    # 有 Excel/WPS 时重算一次，让示例文件带上公式计算结果
    try:
        from engine.scan import recalc
        for f in folder.glob("*.xlsx"):
            recalc(str(f))
    except Exception:  # noqa: BLE001
        pass


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    pyzip = fetch_python()
    # 只清空内容、保留文件夹本身：有命令行窗口停在这个文件夹里时，删文件夹会被系统拒绝
    OUT.mkdir(parents=True, exist_ok=True)
    for child in OUT.iterdir():
        try:
            if child.is_dir():
                shutil.rmtree(child)
            else:
                child.unlink()
        except PermissionError as e:
            sys.exit(f"无法清理旧的打包结果：{child}\n请先关闭正在使用它的窗口（例如 run.bat 的黑色窗口），再重新打包。\n{e}")
    log("解压 Python …")
    with zipfile.ZipFile(pyzip) as z:
        z.extractall(OUT / "python")
    pth = next((OUT / "python").glob("python*._pth"))
    lines = pth.read_text(encoding="utf-8").splitlines()
    for extra in ("..", "..\\libs"):
        if extra not in lines:
            lines.insert(2, extra)
    pth.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    install_libs(OUT / "libs")
    log("复制执行器 …")
    shutil.copytree(HERE / "engine", OUT / "engine", ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    shutil.copy2(HERE / "process.py", OUT / "process.py")
    (OUT / "run.bat").write_bytes((HERE / "run.bat").read_text(encoding="ascii").replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))
    shutil.copytree(HERE / "templates", OUT / "templates")
    make_samples(OUT / "templates" / "示例数据")
    for d in ("configs", "input", "output", "extensions"):
        (OUT / d).mkdir(exist_ok=True)
    log("打包网页配置器 …")
    sys.path.insert(0, str(HERE.parent))
    import build as webbuild  # 项目根目录的 build.py
    webbuild.sync_examples()
    page = webbuild.build_page("workflow.html", "Excel工作流配置器.html")
    html = page.read_text(encoding="utf-8").replace('href="Excel智能函数助手.html"', 'href="#" hidden')
    (OUT / "工作流配置器.html").write_text(html, encoding="utf-8")
    (OUT / "configs" / "把工作流配置文件放在这里.txt").write_text(
        "用本工具包里的「工作流配置器.html」（双击用浏览器打开）配置工作流，保存出的 .json 配置文件放在这个文件夹。\r\n"
        "想先体验：这里没有配置时直接双击 run.bat，窗口里会列出内置示例，输入序号即可用示例数据运行。\r\n",
        encoding="utf-8-sig")
    (OUT / "input" / "把要处理的文件放在这里.txt").write_text("把本次要处理的 Excel 文件（.xlsx）放在这个文件夹，也可以运行时直接拖进窗口。\r\n", encoding="utf-8-sig")
    (OUT / "extensions" / "说明.txt").write_text(
        "高级用法：在这里放 .py 文件可以增加自定义操作（写法见“配置格式说明.md”最后一节）。\r\n"
        "这里的文件是程序代码，只放来自可信来源的文件。\r\n", encoding="utf-8-sig")
    for doc in ("使用手册.txt", "配置格式说明.md"):
        src = HERE / "docs" / doc
        if src.exists():
            shutil.copy2(src, OUT / doc)
    (OUT / "版本.txt").write_text(
        f"Excel 工作流执行器 v{__version__}\r\n打包时间 {dt.datetime.now():%Y-%m-%d %H:%M}\r\n"
        f"内置 Python {PY_VER}（嵌入版）、{'、'.join(LIBS)}\r\n支持 Windows 10/11（64 位）\r\n", encoding="utf-8-sig")
    zpath = DIST / f"{NAME}_v{__version__}.zip"
    if zpath.exists():
        zpath.unlink()
    log("压缩 …")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for p in sorted(OUT.rglob("*")):
            if "__pycache__" in p.parts:
                continue
            z.write(p, Path(NAME) / p.relative_to(OUT))
    size = sum(p.stat().st_size for p in OUT.rglob("*") if p.is_file()) / 1024 / 1024
    log(f"完成：{OUT}（解压后 {size:.1f}MB）")
    log(f"压缩包：{zpath}（{zpath.stat().st_size / 1024 / 1024:.1f}MB）")


if __name__ == "__main__":
    main()
