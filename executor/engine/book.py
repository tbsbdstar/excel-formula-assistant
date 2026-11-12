"""工作簿与工作表视图：读取文件、按表头定位列、读取单元格当前值、写入、结构调整（插入/删除行列）。"""
from __future__ import annotations

import copy
import datetime as dt
import re
import zipfile
from typing import Dict, List, Optional

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.cell_range import CellRange, MultiCellRange
from openpyxl.worksheet.dimensions import ColumnDimension

from . import refs
from .errors import StepError


def norm_header(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v).replace("\n", " ")).strip()


def is_blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


class FileHandle:
    """一个输入文件。主表可写；参照表只读（按需加载公式版）"""

    def __init__(self, spec: dict, path: str, writable: bool):
        self.id = spec["id"]
        self.label = spec.get("label") or spec["id"]
        self.role = spec.get("role", "ref")
        self.spec = spec
        self.path = path
        self.writable = writable
        self.wb = None            # 公式版（主表写入用 / 参照表复制公式和格式用）
        self.wb_values = None     # 计算结果版（参照表取值用）
        self.cached = {}          # 主表：id(公式单元格) → 文件里缓存的计算结果
        self.written_formulas = set()
        self.views: Dict[str, "SheetView"] = {}
        self.dropped_features: List[str] = []
        self.no_cached_values = False

    # ------------------------------------------------------------ 加载
    def open(self):
        if not self.path.lower().endswith((".xlsx", ".xlsm")):
            raise StepError(f"「{self.label}」不是 .xlsx 文件，请先在 Excel/WPS 中另存为 .xlsx：{self.path}")
        self._inspect_package()
        try:
            if self.writable:
                self.wb = load_workbook(self.path, keep_vba=False, keep_links=True)
                self._load_cached_values()
            else:
                self.wb_values = load_workbook(self.path, data_only=True)
        except PermissionError:
            raise StepError(f"无法读取「{self.label}」，文件可能正被其他程序独占：{self.path}")
        except zipfile.BadZipFile:
            raise StepError(f"「{self.label}」不是有效的 .xlsx 文件（可能已损坏或加密）：{self.path}")

    def _inspect_package(self):
        try:
            with zipfile.ZipFile(self.path) as z:
                names = z.namelist()
        except zipfile.BadZipFile:
            raise StepError(f"「{self.label}」不是有效的 .xlsx 文件（可能已损坏或加密）：{self.path}")
        except PermissionError:
            raise StepError(f"无法读取「{self.label}」，文件可能正被其他程序独占：{self.path}")
        if any(n.startswith("xl/charts/") for n in names):
            self.dropped_features.append("图表")
        if any(n.startswith("xl/media/") for n in names):
            self.dropped_features.append("图片")
        if any(n.startswith("xl/pivotTables/") for n in names):
            self.dropped_features.append("数据透视表（可能失效）")
        if any(n.startswith("xl/slicers/") for n in names):
            self.dropped_features.append("切片器")

    def _load_cached_values(self):
        """读取公式单元格在文件里保存的计算结果（Excel/WPS 最后一次保存时算出的值）"""
        need = {}
        for ws in self.wb.worksheets:
            coords = {k for k, c in ws._cells.items() if c.data_type == "f"}
            if coords:
                need[ws.title] = coords
        if not need:
            return
        wbv = load_workbook(self.path, data_only=True, read_only=True)
        total, filled = 0, 0
        try:
            for title, coords in need.items():
                wsv = wbv[title]
                ws = self.wb[title]
                total += len(coords)
                for row in wsv.iter_rows():
                    for c in row:
                        rc = getattr(c, "row", None), getattr(c, "column", None)
                        if rc in coords:
                            self.cached[id(ws._cells[rc])] = c.value
                            filled += c.value is not None
        finally:
            wbv.close()
        # 文件不是 Excel/WPS 保存的（如程序生成）时，公式没有保存计算结果
        self.no_cached_values = total > 0 and filled == 0

    def formula_book(self):
        """参照表的公式版（复制公式/格式时才加载）"""
        if self.wb is None:
            self.wb = load_workbook(self.path, keep_links=True)
        return self.wb

    def value_book(self):
        return self.wb_values if self.wb_values is not None else self.wb

    def sheet_names(self):
        return self.value_book().sheetnames

    # ------------------------------------------------------------ 视图
    def view(self, sheet=None, header_row=None) -> "SheetView":
        name = self.resolve_sheet(sheet)
        hr = int(header_row or self.spec.get("header_row") or 1)
        key = f"{name}|{hr}"
        if key not in self.views:
            self.views[key] = SheetView(self, name, hr)
        return self.views[key]

    def resolve_sheet(self, sheet=None) -> str:
        names = self.sheet_names()
        want = sheet if sheet not in (None, "") else self.spec.get("sheet")
        if want in (None, ""):
            return names[0]
        if isinstance(want, int):
            if 1 <= want <= len(names):
                return names[want - 1]
            raise StepError(f"「{self.label}」没有第 {want} 个工作表（共 {len(names)} 个）")
        if want in names:
            return want
        for n in names:
            if n.strip() == str(want).strip():
                return n
        raise StepError(f"「{self.label}」里找不到工作表「{want}」，现有：{'、'.join(names)}")

    def invalidate(self, sheet_name=None):
        for k, v in self.views.items():
            if sheet_name is None or v.name == sheet_name:
                v.invalidate()


class SheetView:
    """某个工作表 + 表头行的视图。表头/数据行在结构变化后自动重算"""

    def __init__(self, fh: FileHandle, name: str, header_row: int):
        self.fh = fh
        self.name = name
        self.header_row = header_row
        self._headers = None
        self._last_row = None

    @property
    def ws(self):
        return (self.fh.wb if self.fh.writable else self.fh.value_book())[self.name]

    @property
    def fws(self):
        """公式版工作表（参照表会按需加载）"""
        return self.fh.formula_book()[self.name]

    def invalidate(self):
        self._headers = None
        self._last_row = None

    # ------------------------------------------------------------ 表头与列定位
    def headers(self) -> Dict[str, int]:
        if self._headers is None:
            h = {}
            ws = self.ws
            for (r, c), cell in ws._cells.items():
                if r != self.header_row:
                    continue
                name = norm_header(cell.value)
                if name and name not in h:
                    h[name] = c
            self._headers = dict(sorted(h.items(), key=lambda kv: kv[1]))
        return self._headers

    def header_of(self, col: int) -> str:
        cell = self.ws._cells.get((self.header_row, col))
        return norm_header(cell.value) if cell is not None else ""

    def max_col(self) -> int:
        cols = [c for (r, c), cell in self.ws._cells.items() if not is_blank(cell.value)]
        return max(cols) if cols else 0

    def col(self, ref, what="列") -> int:
        """列定位：字符串=表头名（找不到且像列字母时按字母）；{"letter":"L"}；{"index":12}"""
        h = self.headers()
        if isinstance(ref, dict):
            if "letter" in ref:
                return refs.col_index(str(ref["letter"]))
            if "index" in ref:
                return int(ref["index"])
            if "name" in ref:
                ref = ref["name"]
            else:
                raise StepError(f"{what}写法不正确：{ref}")
        if isinstance(ref, int):
            return ref
        name = norm_header(ref)
        if name in h:
            return h[name]
        compact = name.replace(" ", "")
        for k, v in h.items():
            if k.replace(" ", "") == compact:
                return v
        if re.fullmatch(r"[A-Za-z]{1,3}", name):
            return refs.col_index(name)
        m = re.fullmatch(r"([A-Za-z]{1,3})列", name)
        if m:
            return refs.col_index(m.group(1))
        msg = f"在「{self.fh.label}」的「{self.name}」第 {self.header_row} 行表头里找不到{what}「{name}」。"
        other = self._row_with_header(name)
        if other:
            msg += (f"这个名字出现在第 {other} 行，可能这次文件的表头位置变了（配置里写的是第 {self.header_row} 行）。"
                    f"请确认文件格式；如果以后都是第 {other} 行，在配置器里把这个文件的「表头行」改成 {other}。")
        else:
            msg += f"现有表头：{'、'.join(list(h)[:40])}{' …' if len(h) > 40 else ''}"
        raise StepError(msg)

    def _row_with_header(self, name):
        """在前 20 行里找写着这个列名的行（用于提示表头位置变化）"""
        compact = name.replace(" ", "")
        for r in range(1, min(self.ws.max_row, 20) + 1):
            if r == self.header_row:
                continue
            for c in range(1, min(self.ws.max_column, 200) + 1):
                v = self.ws._cells.get((r, c))
                if v is not None and isinstance(v.value, str) and norm_header(v.value).replace(" ", "") == compact:
                    return r
        return None

    def cols(self, spec, what="列") -> List[int]:
        """多列：列表；'U:AB' 字母区间；'表头1:表头2' 表头区间；{"from":..,"to":..}"""
        if spec is None:
            raise StepError(f"缺少{what}")
        if isinstance(spec, list):
            out = []
            for s in spec:
                out.extend(self.cols(s, what) if isinstance(s, str) and ":" in s else [self.col(s, what)])
            return out
        if isinstance(spec, dict) and "from" in spec:
            a, b = self.col(spec["from"], what), self.col(spec["to"], what)
            return list(range(min(a, b), max(a, b) + 1))
        if isinstance(spec, str) and ":" in spec:
            a, b = [x.strip() for x in spec.split(":", 1)]
            if re.fullmatch(r"[A-Za-z]{1,3}", a) and re.fullmatch(r"[A-Za-z]{1,3}", b) and a not in self.headers() and b not in self.headers():
                ca, cb = refs.col_index(a), refs.col_index(b)
            else:
                ca, cb = self.col(a, what), self.col(b, what)
            return list(range(min(ca, cb), max(ca, cb) + 1))
        return [self.col(spec, what)]

    # ------------------------------------------------------------ 数据行
    def last_row(self) -> int:
        if self._last_row is None:
            rows = [r for (r, c), cell in self.ws._cells.items()
                    if r > self.header_row and not isinstance(cell, MergedCell) and not is_blank(cell.value)]
            self._last_row = max(rows) if rows else self.header_row
        return self._last_row

    def data_rows(self, spec=None) -> List[int]:
        start, end = self.header_row + 1, self.last_row()
        if spec:
            if isinstance(spec, dict):
                start = int(spec.get("from", start))
                end = int(spec.get("to", end))
            elif isinstance(spec, str) and ":" in spec:
                a, b = spec.split(":", 1)
                start, end = int(a), int(b)
        return list(range(start, end + 1))

    # ------------------------------------------------------------ 取值与写入
    def cell(self, r: int, c: int):
        return self.ws._cells.get((r, c))

    def get(self, r: int, c: int, warn=None):
        """单元格当前值；公式单元格取文件里保存的计算结果"""
        cell = self.ws._cells.get((r, c))
        if cell is None or isinstance(cell, MergedCell):
            return None
        if self.fh.writable and cell.data_type == "f":
            key = id(cell)
            if key in self.fh.cached:
                return self.fh.cached[key]
            if warn:
                warn("条件或匹配用到了本次新写入的公式单元格，它们还没有计算结果，按空值处理")
            return None
        return cell.value

    def set_value(self, r: int, c: int, value):
        ws = self.ws
        cell = ws._cells.get((r, c))
        if isinstance(cell, MergedCell):
            return False
        cell = ws.cell(row=r, column=c)
        self.fh.cached.pop(id(cell), None)
        self.fh.written_formulas.discard(id(cell))
        cell.value = value
        if r > (self._last_row or 0) and not is_blank(value):
            self._last_row = None
        return True

    def set_formula(self, r: int, c: int, formula: str):
        if not isinstance(formula, str) or not formula.startswith("="):
            return self.set_value(r, c, formula)
        ws = self.ws
        cell = ws._cells.get((r, c))
        if isinstance(cell, MergedCell):
            return False
        cell = ws.cell(row=r, column=c)
        self.fh.cached.pop(id(cell), None)
        cell.value = formula
        self.fh.written_formulas.add(id(cell))
        return True


# ================================================================ 结构调整
def _remap_cells(ws, cmap: Optional[refs.IndexMap], rmap: Optional[refs.IndexMap]):
    new = {}
    for (r, c), cell in ws._cells.items():
        nr = rmap.map(r) if rmap else r
        nc = cmap.map(c) if cmap else c
        if nr is None or nc is None:
            continue
        cell.row, cell.column = nr, nc
        new[(nr, nc)] = cell
    ws._cells = new
    for cell in ws._cells.values():
        link = getattr(cell, "_hyperlink", None)
        if link is not None:
            link.ref = cell.coordinate


def _remap_dimensions(ws, cmap, rmap):
    if cmap:
        dims = []
        for key, d in list(ws.column_dimensions.items()):
            lo, hi = (d.min or refs.col_index(key)), (d.max or refs.col_index(key))
            for c in range(lo, min(hi, refs.MAX_COL) + 1):
                nc = cmap.map(c)
                if nc is not None:
                    dims.append((nc, d))
        ws.column_dimensions.clear()
        for nc, d in dims:
            letter = refs.col_letter(nc)
            nd = ColumnDimension(ws, index=letter, width=d.width, bestFit=d.bestFit, hidden=d.hidden,
                                 outlineLevel=d.outlineLevel, collapsed=d.collapsed, customWidth=d.customWidth)
            nd._style = copy.copy(d._style)
            ws.column_dimensions[letter] = nd
    if rmap:
        items = list(ws.row_dimensions.items())
        ws.row_dimensions.clear()
        for r, d in items:
            nr = rmap.map(r)
            if nr is None:
                continue
            d.index = nr
            ws.row_dimensions[nr] = d


def _remap_merged(ws, cmap, rmap):
    ranges = []
    for cr in list(ws.merged_cells.ranges):
        txt = refs.shift_ref_text(cr.coord, cmap, rmap)
        if txt == "#REF!" or ":" not in txt:
            continue
        ranges.append(txt)
    ws.merged_cells = MultiCellRange()
    for txt in ranges:
        cr = CellRange(txt)
        if cr.min_row == cr.max_row and cr.min_col == cr.max_col:
            continue
        # 左上角必须是普通单元格
        tl = ws._cells.get((cr.min_row, cr.min_col))
        if tl is None or isinstance(tl, MergedCell):
            ws._cells[(cr.min_row, cr.min_col)] = Cell(ws, row=cr.min_row, column=cr.min_col)
        ws.merged_cells.add(cr)


def _remap_formatting(ws, target, cmap, rmap):
    old = ws.conditional_formatting
    new = ConditionalFormattingList()
    for cf in old:
        sq = refs.shift_range_list(str(cf.sqref), cmap, rmap)
        if not sq:
            continue
        for rule in cf.rules:
            if getattr(rule, "formula", None):
                rule.formula = [refs.shift_formula("=" + f, ws.title, target, cmap, rmap)[1:] for f in rule.formula]
            new.add(sq, rule)
    ws.conditional_formatting = new
    dvs = getattr(ws, "data_validations", None)
    if dvs is not None:
        keep = []
        for dv in dvs.dataValidation:
            sq = refs.shift_range_list(str(dv.sqref), cmap, rmap)
            if not sq:
                continue
            dv.sqref = MultiCellRange(sq)
            for attr in ("formula1", "formula2"):
                f = getattr(dv, attr, None)
                if f:
                    setattr(dv, attr, refs.shift_formula("=" + f, ws.title, target, cmap, rmap)[1:])
            keep.append(dv)
        dvs.dataValidation = keep
    if ws.auto_filter and ws.auto_filter.ref:
        ws.auto_filter.ref = refs.shift_range_list(ws.auto_filter.ref, cmap, rmap)
    if ws.freeze_panes:
        # 冻结窗格记录的是“分隔位置”（左上角第一个未冻结的单元格），不是单元格引用：
        # 冻结区内的列/行被删除时冻结范围缩小，冻结区内插入时扩大，与 Excel 一致
        try:
            c, r = refs.parse_cell(ws.freeze_panes)
        except Exception:  # noqa: BLE001
            c = r = None
        if c is not None:
            nc = cmap.map_boundary(c) if cmap else c
            nr = rmap.map_boundary(r) if rmap else r
            ws.freeze_panes = None if (nc, nr) == (1, 1) else f"{refs.col_letter(nc)}{nr}"
    for t in getattr(ws, "tables", {}).values():
        t.ref = refs.shift_range_list(t.ref, cmap, rmap) or t.ref
        if getattr(t, "autoFilter", None) is not None and t.autoFilter.ref:
            t.autoFilter.ref = t.ref


def check_tables(ws, cmap=None, rmap=None, at_col=None, at_rows=None):
    """插入/删除落在 Excel 表格（Table）内部时拒绝，避免损坏文件"""
    for name, t in getattr(ws, "tables", {}).items():
        min_c, min_r, max_c, max_r = refs.parse_range(t.ref)
        if at_col is not None and min_c <= max(at_col) and min(at_col) <= max_c:
            raise StepError(f"工作表「{ws.title}」的列落在 Excel 表格「{name}」({t.ref}) 内，暂不支持在表格中插入或删除列，请先在 Excel 中把它“转换为区域”")
        if at_rows is not None and min_r <= max(at_rows) and min(at_rows) <= max_r:
            raise StepError(f"工作表「{ws.title}」的行落在 Excel 表格「{name}」({t.ref}) 内，暂不支持，请先在 Excel 中把它“转换为区域”")


def restructure(fh: FileHandle, sheet_name: str, cmap: Optional[refs.IndexMap], rmap: Optional[refs.IndexMap]):
    """对工作表做行/列结构变化，并像 Excel 一样同步调整所有引用"""
    wb = fh.wb
    ws = wb[sheet_name]
    _remap_cells(ws, cmap, rmap)
    _remap_dimensions(ws, cmap, rmap)
    _remap_merged(ws, cmap, rmap)
    _remap_formatting(ws, sheet_name, cmap, rmap)
    # 所有工作表中的公式
    for other in wb.worksheets:
        for cell in other._cells.values():
            if cell.data_type == "f" and isinstance(cell.value, str):
                v = cell.value
                if other.title != sheet_name and "!" not in v:
                    continue
                nv = refs.shift_formula(v, other.title, sheet_name, cmap, rmap)
                if nv != v:
                    cell.value = nv
    # 定义的名称
    try:
        for name, dn in list(wb.defined_names.items()):
            if dn.attr_text:
                dn.attr_text = refs.shift_formula("=" + dn.attr_text, "", sheet_name, cmap, rmap)[1:]
    except AttributeError:
        pass
    fh.invalidate(sheet_name)


def extend_autofilter(view, col: int):
    """在表格最右边新加一列时，把它纳入原有的筛选范围（否则新列不能筛选）"""
    ws = view.ws
    ref = ws.auto_filter.ref if ws.auto_filter else None
    if not ref:
        return
    try:
        c1, r1, c2, r2 = refs.parse_range(ref)
    except Exception:  # noqa: BLE001
        return
    if r1 == view.header_row and col == c2 + 1:
        ws.auto_filter.ref = f"{refs.col_letter(c1)}{r1}:{refs.col_letter(col)}{r2}"


def copy_style(src_cell, dst_cell, same_book: bool):
    if src_cell is None or not getattr(src_cell, "has_style", False):
        return
    if same_book:
        dst_cell._style = copy.copy(src_cell._style)
    else:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)
