/* 工作簿：读取文件、识别表头与列类型、维护计算引擎、应用/撤销计算 */
(function (EFA) {
  'use strict';
  const U = EFA.U, F = EFA.Formula;

  const MAX_DISPLAY_ROWS = 200;
  const ENGINE_CONFIG = {
    licenseKey: 'gpl-v3',
    useArrayArithmetic: true,
    evaluateNullToZero: true,
    useColumnIndex: true,
    maxRows: 1048576,
    maxColumns: 16384,
    dateFormats: ['YYYY-MM-DD', 'YYYY/MM/DD', 'YYYY.MM.DD', 'MM/DD/YYYY'],
    timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  };
  const TOTAL_ROW_RE = /^(合计|总计|小计|总和|汇总|合 计|总 计|total|sum)$/i;

  /* ======================= 读取文件 ======================= */

  function decodeText(buf) {
    const bytes = new Uint8Array(buf);
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return new TextDecoder('utf-8').decode(bytes.subarray(3));
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (e) {
      return new TextDecoder('gb18030').decode(bytes); // 中文 Windows 下 Excel 存的 CSV 通常是 GBK
    }
  }

  async function readFile(file) {
    const name = file.name || 'data.xlsx';
    const ext = (name.split('.').pop() || '').toLowerCase();
    if (!['xlsx', 'xlsm', 'xls', 'csv', 'txt', 'tsv'].includes(ext)) {
      throw new Error('不支持的文件类型 .' + ext + '，请上传 .xlsx / .xls / .csv 文件');
    }
    const buf = await file.arrayBuffer();
    return fromBuffer(buf, name, ext);
  }

  function fromBuffer(buf, name, ext) {
    let xwb, kind;
    if (ext === 'csv' || ext === 'txt' || ext === 'tsv') {
      const text = decodeText(buf);
      xwb = XLSX.read(text, { type: 'string', raw: true, FS: ext === 'tsv' ? '\t' : undefined });
      kind = 'csv';
    } else {
      try {
        xwb = XLSX.read(buf, { type: 'array', cellDates: false, cellNF: true, cellFormula: true, cellText: true });
      } catch (e) {
        throw new Error('文件无法解析，可能已损坏或被加密：' + (e.message || e));
      }
      kind = ext === 'xls' ? 'xls' : 'xlsx';
    }
    if (!xwb.SheetNames || !xwb.SheetNames.length) throw new Error('文件里没有工作表');

    const wb = {
      fileName: name,
      kind,
      buffer: kind === 'xlsx' ? buf : null,
      sheets: [],
      outputs: [],      // 新建的结果表（汇总、分组）
      ops: [],          // 已确认的计算
      pending: null,    // 待确认的计算
      activeIndex: 0,
      warnings: [],
      date1904: !!(xwb.Workbook && xwb.Workbook.WBProps && xwb.Workbook.WBProps.date1904),
    };
    if (wb.date1904) wb.warnings.push('该文件使用 1904 日期系统，日期类计算结果可能有偏差');

    xwb.SheetNames.forEach(sn => {
      const ws = xwb.Sheets[sn];
      if (!ws || ws['!type'] === 'chart') return;
      wb.sheets.push(buildSheet(ws, sn, kind));
    });
    if (!wb.sheets.length) throw new Error('文件里没有可用的数据表');
    // 默认选中第一个有数据的表
    const idx = wb.sheets.findIndex(s => s.dataEnd >= s.dataStart);
    wb.activeIndex = idx >= 0 ? idx : 0;
    buildEngine(wb);
    return wb;
  }

  function cellValue(cell) {
    if (!cell) return null;
    switch (cell.t) {
      case 'n': return typeof cell.v === 'number' ? cell.v : null;
      case 's': case 'str': return cell.v == null ? null : String(cell.v);
      case 'b': return !!cell.v;
      case 'e': return cell.w || '#N/A';
      case 'd': return cell.v instanceof Date ? U.partsToSerial(cell.v.getFullYear(), cell.v.getMonth() + 1, cell.v.getDate()) : null;
      default: return null;
    }
  }

  function buildSheet(ws, name, kind) {
    const sheet = {
      name, values: [], display: [], fmts: [], formulas: {}, merges: ws['!merges'] || [],
      lastRow: -1, lastCol: -1, origLastCol: -1, headerRow: 0, dataStart: 1, dataEnd: 0, totalRow: null,
      columns: [], warnings: [], cols: ws['!cols'] || null,
    };
    const ref = ws['!ref'];
    if (!ref) return finalizeSheet(sheet, kind);
    const range = XLSX.utils.decode_range(ref);
    for (let r = 0; r <= range.e.r; r++) {
      const row = [];
      let disp = r < MAX_DISPLAY_ROWS + 20 ? [] : null;
      const fmtRow = [];
      for (let c = 0; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        const v = cellValue(cell);
        row.push(v);
        if (cell && cell.z && cell.t === 'n') fmtRow[c] = cell.z;
        if (disp) disp.push(cell ? (cell.w != null ? cell.w : (v == null ? '' : String(v))) : '');
        if (cell && cell.f && kind !== 'xlsx') sheet.formulas[r + ',' + c] = cell.f;
        if (v !== null && v !== '') {
          if (r > sheet.lastRow) sheet.lastRow = r;
          if (c > sheet.lastCol) sheet.lastCol = c;
        }
      }
      sheet.values.push(row);
      sheet.fmts.push(fmtRow);
      if (disp) sheet.display.push(disp);
    }
    if (kind === 'csv') coerceCsv(sheet);
    return finalizeSheet(sheet, kind);
  }

  /* CSV 全部按文本读入后，按列判断是否为数字/日期，避免“001”这类编码丢失前导零 */
  function coerceCsv(sheet) {
    const header = detectHeader(sheet);
    for (let c = 0; c <= sheet.lastCol; c++) {
      let nonEmpty = 0, numLike = 0, dateLike = 0, pctLike = 0;
      for (let r = header + 1; r <= sheet.lastRow; r++) {
        const v = sheet.values[r][c];
        if (v === null || v === '') continue;
        const s = String(v).trim();
        nonEmpty++;
        if (/^-?(0|[1-9]\d{0,2}(,\d{3})+|[1-9]\d*)(\.\d+)?$/.test(s) || /^-?0\.\d+$/.test(s)) numLike++;
        else if (/^-?\d+(\.\d+)?%$/.test(s)) pctLike++;
        else if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}( \d{1,2}:\d{2}(:\d{2})?)?$/.test(s)) dateLike++;
      }
      if (!nonEmpty) continue;
      const mode = numLike === nonEmpty ? 'num' : pctLike === nonEmpty ? 'pct' : dateLike === nonEmpty ? 'date' : null;
      if (!mode) continue;
      for (let r = header + 1; r <= sheet.lastRow; r++) {
        const v = sheet.values[r][c];
        if (v === null || v === '') continue;
        const s = String(v).trim();
        if (mode === 'num') sheet.values[r][c] = Number(s.replace(/,/g, ''));
        else if (mode === 'pct') { sheet.values[r][c] = Number(s.slice(0, -1)) / 100; sheet.fmts[r][c] = '0.00%'; }
        else {
          const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
          let serial = U.partsToSerial(+m[1], +m[2], +m[3]);
          if (m[4]) serial += ((+m[4]) * 3600 + (+m[5]) * 60 + (+(m[6] || 0))) / 86400;
          sheet.values[r][c] = serial;
          sheet.fmts[r][c] = m[4] ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd';
        }
      }
    }
  }

  /** 自动识别表头行：前 15 行中第一行“大部分是文字、且宽度与下方数据相当”的行 */
  function detectHeader(sheet) {
    const maxScan = Math.min(sheet.lastRow, 14);
    const width = r => {
      const row = sheet.values[r] || [];
      let n = 0;
      for (let c = 0; c <= sheet.lastCol; c++) if (row[c] !== null && row[c] !== '' && row[c] !== undefined) n++;
      return n;
    };
    let firstNonEmpty = -1;
    for (let r = 0; r <= maxScan; r++) {
      const row = sheet.values[r] || [];
      let nonEmpty = 0, texts = 0;
      for (let c = 0; c <= sheet.lastCol; c++) {
        const v = row[c];
        if (v === null || v === '' || v === undefined) continue;
        nonEmpty++;
        if (typeof v === 'string' && isNaN(Number(v))) texts++;
      }
      if (!nonEmpty) continue;
      if (firstNonEmpty < 0) firstNonEmpty = r;
      let below = 0;
      for (let k = 1; k <= 3 && r + k <= sheet.lastRow; k++) below = Math.max(below, width(r + k));
      if (texts >= Math.ceil(nonEmpty * 0.6) && nonEmpty >= Math.ceil(below * 0.6)) return r;
    }
    return Math.max(0, firstNonEmpty);
  }

  function finalizeSheet(sheet, kind) {
    sheet.origLastCol = sheet.lastCol;
    if (sheet.lastRow < 0) {
      sheet.headerRow = 0; sheet.dataStart = 1; sheet.dataEnd = 0; sheet.columns = [];
      return sheet;
    }
    setHeaderRow(sheet, detectHeader(sheet));
    return sheet;
  }

  /** 设置表头行并重新分析列 */
  function setHeaderRow(sheet, headerRow) {
    sheet.headerRow = headerRow;
    sheet.dataStart = headerRow + 1;
    sheet.dataEnd = sheet.lastRow;
    sheet.totalRow = null;
    sheet.warnings = [];
    // 末行是“合计”行时排除
    if (sheet.dataEnd > sheet.dataStart) {
      const last = sheet.values[sheet.dataEnd] || [];
      const isTotal = last.some(v => typeof v === 'string' && TOTAL_ROW_RE.test(v.trim()));
      if (isTotal) {
        sheet.totalRow = sheet.dataEnd;
        sheet.dataEnd -= 1;
        sheet.warnings.push('第 ' + (sheet.totalRow + 1) + ' 行是合计行，计算时已自动排除');
      }
    }
    // 表头行有合并单元格时提示
    if (sheet.merges.some(m => m.s.r <= headerRow && m.e.r >= headerRow && m.e.r > m.s.r)) {
      sheet.warnings.push('表头含跨行合并单元格，列名可能识别不准，可手动调整表头行');
    }
    analyzeColumns(sheet);
  }

  const NAME_RE = /姓名|名字|员工|人员|销售员|业务员|负责人|客户名|学生|联系人|收货人|经办人|姓 名/;
  const ID_RE = /编号|工号|单号|订单号|学号|编码|ID|id|序号|代码/;

  function analyzeColumns(sheet) {
    const header = sheet.values[sheet.headerRow] || [];
    const taken = new Set();
    sheet.columns = [];
    const lastCol = sheet.lastCol;
    for (let c = 0; c <= lastCol; c++) {
      const letter = U.colLetter(c);
      let name = header[c];
      name = name == null ? '' : String(typeof name === 'number' && isDateFmt(sheet.fmts[sheet.headerRow] && sheet.fmts[sheet.headerRow][c]) ? U.formatSerialDate(name) : name);
      name = name.replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
      const rawHeader = name;
      if (!name) name = letter + '列';
      if (taken.has(name)) name = name + '(' + letter + ')';
      taken.add(name);
      const col = { index: c, letter, name, rawHeader, added: c > sheet.origLastCol };
      computeColumnStats(sheet, col);
      sheet.columns.push(col);
    }
    // 去掉首尾完全空白、且无表头的列
    sheet.columns = sheet.columns.filter(col => !(col.type === 'empty' && !col.rawHeader));
  }

  function isDateFmt(z) {
    if (!z || typeof z !== 'string') return false;
    try { return XLSX.SSF.is_date(z); } catch (e) { return /y|d/i.test(z) && !/[#0]/.test(z); }
  }

  /** 统计一列的类型、样例、去重值 */
  function computeColumnStats(sheet, col) {
    const c = col.index;
    let nonEmpty = 0, num = 0, date = 0, str = 0, numStr = 0, bool = 0, err = 0;
    const distinct = new Map();
    let distinctOverflow = false;
    const samples = [];
    const nums = [];
    let dateFmt = null;
    const step = Math.max(1, Math.floor((sheet.dataEnd - sheet.dataStart + 1) / 5000));
    for (let r = sheet.dataStart; r <= sheet.dataEnd; r += 1) {
      const v = sheet.values[r] ? sheet.values[r][c] : null;
      if (v === null || v === '' || v === undefined) continue;
      nonEmpty++;
      const sampled = (r - sheet.dataStart) % step === 0;
      if (typeof v === 'number') {
        const z = sheet.fmts[r] && sheet.fmts[r][c];
        if (isDateFmt(z)) { date++; dateFmt = dateFmt || z; } else num++;
        if (sampled && nums.length < 5000) nums.push(v);
      } else if (typeof v === 'boolean') bool++;
      else if (typeof v === 'string') {
        if (/^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(v)) err++;
        else if (v.trim() !== '' && !isNaN(Number(v))) numStr++;
        else str++;
      }
      {
        if (!distinctOverflow) {
          distinct.set(v, (distinct.get(v) || 0) + 1);
          if (distinct.size > 3000) distinctOverflow = true;
        }
      }
      if (samples.length < 5) samples.push(displayValue(sheet, r, c));
    }
    let type = 'empty';
    if (nonEmpty) {
      if (date && date + num === nonEmpty && date >= num) type = 'date';
      else if (num + date === nonEmpty) type = 'number';
      else if (numStr === nonEmpty) type = 'textnum';
      else if (bool === nonEmpty) type = 'bool';
      else if (str + numStr + err === nonEmpty && str > 0) type = 'text';
      else if (num + date > nonEmpty * 0.8) type = 'number';
      else type = 'mixed';
    }
    nums.sort((a, b) => a - b);
    Object.assign(col, {
      type, nonEmpty, samples, dateFmt,
      distinct: distinctOverflow ? null : distinct,
      distinctCount: distinctOverflow ? Infinity : distinct.size,
      numStats: nums.length ? { min: nums[0], max: nums[nums.length - 1], median: nums[Math.floor(nums.length / 2)], count: nums.length } : null,
      isName: NAME_RE.test(col.name),
      isId: ID_RE.test(col.name),
      isIdCard: /身份证/.test(col.name) || (type === 'text' || type === 'textnum') && samples.length > 0 && samples.every(s => /^\d{17}[\dXx]$/.test(String(s))),
      isPhone: /手机|电话|联系方式/.test(col.name) || samples.length > 0 && samples.every(s => /^1[3-9]\d{9}$/.test(String(s))),
      isBirth: /出生|生日/.test(col.name),
      isHire: /入职|参加工作|加入|进入公司|入司/.test(col.name),
    });
    col.isCategory = (type === 'text' || type === 'textnum') && col.distinctCount <= 50 && col.distinctCount < nonEmpty * 0.7;
    if (type === 'textnum') {
      col.warning = '「' + col.name + '」里的数字是以文本形式存储的，Excel 中直接求和会得到 0';
    }
  }

  function displayValue(sheet, r, c) {
    const d = sheet.display[r];
    if (d && d[c] !== undefined && d[c] !== '') return d[c];
    const v = sheet.values[r] ? sheet.values[r][c] : null;
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      const z = sheet.fmts[r] && sheet.fmts[r][c];
      if (isDateFmt(z)) return U.formatSerialDate(v, /h/i.test(z));
      if (z && /%/.test(z)) return U.formatNumber(v, 'percent');
      return U.formatNumber(v);
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }

  /* ======================= 计算引擎 ======================= */

  function toEngineValue(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'string') {
      if (/^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(v)) return v;
      return "'" + v;
    }
    return v;
  }

  function buildEngine(wb) {
    if (wb.hf) { try { wb.hf.destroy(); } catch (e) { /* ignore */ } }
    const sheets = {};
    for (const s of wb.sheets) {
      sheets[s.name] = s.values.map(row => {
        const out = new Array(Math.min(row.length, s.origLastCol + 1));
        for (let c = 0; c < out.length; c++) out[c] = toEngineValue(row[c]);
        return out;
      });
      if (!sheets[s.name].length) sheets[s.name] = [[null]];
    }
    wb.hf = HyperFormula.buildFromSheets(sheets, ENGINE_CONFIG);
    wb.outputs = [];
  }

  function sheetId(wb, name) { return wb.hf.getSheetId(name); }

  function resolver(wb, sourceSheet) {
    return function (sheetName, colName) {
      const find = (sh, name) => {
        const n = String(name).trim();
        let col = sh.columns.find(c => c.name === n);
        if (!col) col = sh.columns.find(c => c.name.toLowerCase() === n.toLowerCase());
        if (!col) col = sh.columns.find(c => c.rawHeader && c.rawHeader.replace(/\s/g, '') === n.replace(/\s/g, ''));
        if (!col && /^[A-Z]{1,3}列?$/.test(n)) col = sh.columns.find(c => c.letter === n.replace('列', ''));
        return col ? { sheet: sh, col } : null;
      };
      if (sheetName) {
        const sh = wb.sheets.find(s => s.name === sheetName) || wb.sheets.find(s => s.name.replace(/\s/g, '') === sheetName.replace(/\s/g, ''));
        return sh ? find(sh, colName) : null;
      }
      return find(sourceSheet, colName);
    };
  }

  /* ---------- 结果值的读取与分析 ---------- */

  function readValue(wb, sid, r, c) {
    const v = wb.hf.getCellValue({ sheet: sid, row: r, col: c });
    if (v && typeof v === 'object' && 'type' in v && 'value' in v) return { error: v.value, message: v.message, type: v.type };
    return v;
  }

  function isErrorValue(v) { return v && typeof v === 'object' && 'error' in v; }

  function formatValue(v, fmt) {
    if (v === null || v === undefined) return '';
    if (isErrorValue(v)) return v.error;
    if (typeof v === 'number') {
      if (fmt === 'date') return U.formatSerialDate(v);
      if (fmt === 'month') return U.formatSerialMonth(v);
      if (fmt === 'year') return String(U.serialToParts(v).y) + '年';
      return U.formatNumber(v, fmt);
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }

  /** 汇总错误情况，给出通俗解释 */
  function summarizeErrors(values) {
    const counts = {};
    let total = 0;
    for (const v of values) {
      if (isErrorValue(v)) { counts[v.error] = (counts[v.error] || 0) + 1; total++; }
    }
    const tips = {
      '#N/A': '没找到匹配的值', '#DIV/0!': '除数为 0 或没有符合条件的数据', '#VALUE!': '参与计算的值类型不对（如文本参与了运算）',
      '#NUM!': '数值无效（如日期顺序颠倒）', '#REF!': '引用无效', '#NAME?': '函数名无法识别', '#ERROR!': '公式语法错误',
    };
    const parts = Object.keys(counts).map(k => counts[k] + ' 个 ' + k + '（' + (tips[k] || '计算出错') + '）');
    return { total, counts, text: parts.join('，') };
  }

  /* ======================= 计划 → 计算 ======================= */

  /**
   * 计划（模板与 AI 共用）：
   * { kind: 'column'|'scalar'|'group', sheet: 源工作表名, groupBy: [{column, unit:'value'|'month'|'year'}],
   *   outputs: [{ title, formula(列引用语法), format }], explanation, interpretation, source }
   */
  function normalizePlan(wb, plan) {
    const p = Object.assign({}, plan);
    p.sheet = p.sheet || wb.sheets[wb.activeIndex].name;
    p.outputs = (p.outputs || []).filter(o => o && o.formula).map(o => ({
      title: String(o.title || '计算结果').trim().slice(0, 60) || '计算结果',
      formula: String(o.formula).trim(),
      format: o.format || 'auto',
    }));
    if (!p.outputs.length) throw new Error('没有生成任何公式');
    p.groupBy = (p.groupBy || []).filter(g => g && g.column).map(g => ({ column: g.column, unit: g.unit || 'value' }));
    if (p.kind === 'group' && !p.groupBy.length) throw new Error('分组汇总缺少分组依据列');
    if (!['column', 'scalar', 'group'].includes(p.kind)) throw new Error('无法识别的结果类型');
    return p;
  }

  /** 应用计划为“待确认”状态：写入引擎并算出结果，可确认或撤销 */
  function applyPending(wb, plan) {
    discardPending(wb);
    const p = normalizePlan(wb, plan);
    const src = wb.sheets.find(s => s.name === p.sheet);
    if (!src) throw new Error('找不到工作表「' + p.sheet + '」');
    const resolve = resolver(wb, src);

    // 先展开并校验第一条公式，避免把错误公式写进引擎
    const op = { id: 'op' + Date.now() + Math.floor(Math.random() * 1000), plan: p, cells: [], kind: p.kind, sheet: src.name, warnings: [] };
    if (p.kind === 'column') buildColumnOp(wb, op, src, resolve);
    else if (p.kind === 'scalar') buildScalarOp(wb, op, src, resolve);
    else buildGroupOp(wb, op, src, resolve);

    // 语法与函数校验
    const check = new Set();
    for (const o of op.outputs) {
      const fc = F.checkFunctions(o.sampleFormula);
      fc.problems.forEach(x => check.add(x));
      o.functions = fc.used;
      if (!wb.hf.validateFormula(F.toEngine(o.sampleFormula))) check.add('公式语法有误：' + o.sampleFormula);
    }
    if (check.size) {
      const e = new Error(Array.from(check).join('；'));
      e.validation = true;
      throw e;
    }

    // 写入引擎（一次批量操作，便于撤销）
    const undoBefore = wb.hf.isThereSomethingToUndo();
    wb.hf.batch(() => {
      if (op.newSheet && !wb.hf.doesSheetExist(op.newSheet)) wb.hf.addSheet(op.newSheet);
      for (const blk of op.blocks) {
        const sid = sheetId(wb, blk.sheet);
        wb.hf.setCellContents({ sheet: sid, row: blk.row, col: blk.col }, blk.data.map(r => r.map(x => (typeof x === 'string' && x.startsWith('=')) ? F.toEngine(x) : toEngineValue(x))));
      }
    });
    op.undoable = true;
    op.hadUndoBefore = undoBefore;

    // 读取结果
    for (const o of op.outputs) {
      const sid = sheetId(wb, o.sheet);
      o.values = o.rows.map(r => readValue(wb, sid, r, o.col));
      o.errors = summarizeErrors(o.values);
      if (o.errors.counts['#ERROR!'] || o.errors.counts['#NAME?']) {
        const sample = o.values.find(isErrorValue);
        wb.hf.undo();
        const e = new Error('公式无法计算：' + (sample && sample.message ? sample.message : sample.error));
        e.validation = true;
        throw e;
      }
      if (o.format === 'auto') o.format = inferFormat(o, p);
    }
    wb.pending = op;
    return op;
  }

  function inferFormat(o, p) {
    const f = o.sampleFormula.toUpperCase();
    if (/^=\s*(DATE|EDATE|EOMONTH|TODAY|WORKDAY|DATEVALUE)\(/.test(f)) return 'date';
    return 'general';
  }

  function nextOutputSheetName(wb, base) {
    const taken = new Set(wb.sheets.map(s => s.name).concat(wb.outputs.map(o => o.name)));
    return U.uniqueName(U.sanitizeSheetName(base), taken);
  }

  const ROW_SENTINEL = '\u0001ROW\u0001';

  function buildColumnOp(wb, op, src, resolve) {
    const p = op.plan;
    if (src.dataEnd < src.dataStart) throw new Error('工作表「' + src.name + '」没有数据行');
    let col = src.lastCol + 1;
    const takenNames = new Set(src.columns.map(c => c.name));
    op.outputs = [];
    op.blocks = [];
    const nRows = src.dataEnd - src.dataStart + 1;
    for (const o of p.outputs) {
      const title = U.uniqueName(o.title, takenNames);
      takenNames.add(title);
      const tpl = F.expand(o.formula, { resolve, hostSheet: src.name, row: ROW_SENTINEL, allowRow: true });
      const formulas = [];
      for (let i = 0; i < nRows; i++) formulas.push(tpl.split(ROW_SENTINEL).join(String(src.dataStart + i + 1)));
      op.outputs.push({
        title, format: o.format, sheet: src.name, col, headerRow: src.headerRow,
        rows: Array.from({ length: nRows }, (_, i) => src.dataStart + i),
        formulas, sampleFormula: formulas[0], dsl: o.formula,
      });
      op.blocks.push({ sheet: src.name, row: src.headerRow, col, data: [[title]] });
      op.blocks.push({ sheet: src.name, row: src.dataStart, col, data: formulas.map(f => [f]) });
      col++;
    }
  }

  function buildScalarOp(wb, op, src, resolve) {
    const p = op.plan;
    let out = wb.outputs.find(o => o.kind === 'scalar');
    const sheetName = out ? out.name : nextOutputSheetName(wb, '计算结果');
    op.newSheet = out ? null : sheetName;
    op.outputs = [];
    op.blocks = [];
    let row = out ? out.nextRow : 1;
    if (!out) op.blocks.push({ sheet: sheetName, row: 0, col: 0, data: [['计算项', '结果', '需求描述']] });
    for (const o of p.outputs) {
      const f = F.expand(o.formula, { resolve, hostSheet: sheetName, row: null, allowRow: false });
      op.outputs.push({ title: o.title, format: o.format, sheet: sheetName, col: 1, rows: [row], formulas: [f], sampleFormula: f, dsl: o.formula });
      op.blocks.push({ sheet: sheetName, row, col: 0, data: [[o.title, f, p.request || '']] });
      row++;
    }
    op.scalarSheet = { name: sheetName, isNew: !out, nextRow: row };
  }

  /** 计算分组值（去重、保持出现顺序；按月/年时为该月/年第一天并升序） */
  function groupKeys(src, groupBy, resolve) {
    const cols = groupBy.map(g => {
      const hit = resolve(null, g.column);
      if (!hit) throw new Error('找不到分组列「' + g.column + '」');
      return { col: hit.col, unit: g.unit };
    });
    const seen = new Map();
    for (let r = src.dataStart; r <= src.dataEnd; r++) {
      const key = [];
      let skip = false;
      for (const { col, unit } of cols) {
        let v = src.values[r] ? src.values[r][col.index] : null;
        if (v === null || v === '' || v === undefined) { skip = true; break; }
        if (unit === 'month' || unit === 'year') {
          if (typeof v !== 'number') { skip = true; break; }
          const pp = U.serialToParts(v);
          v = unit === 'month' ? U.partsToSerial(pp.y, pp.m, 1) : U.partsToSerial(pp.y, 1, 1);
        }
        key.push(v);
      }
      if (skip) continue;
      const k = JSON.stringify(key);
      if (!seen.has(k)) seen.set(k, key);
    }
    let keys = Array.from(seen.values());
    if (cols.some(c => c.unit !== 'value')) {
      keys.sort((a, b) => {
        for (let i = 0; i < a.length; i++) {
          if (a[i] === b[i]) continue;
          if (typeof a[i] === 'number' && typeof b[i] === 'number') return a[i] - b[i];
          return String(a[i]).localeCompare(String(b[i]), 'zh');
        }
        return 0;
      });
    }
    if (keys.length > 5000) throw new Error('分组数量过多（' + keys.length + ' 组），请换一个分组列');
    return { cols, keys };
  }

  function buildGroupOp(wb, op, src, resolve) {
    const p = op.plan;
    const { cols, keys } = groupKeys(src, p.groupBy, resolve);
    if (!keys.length) throw new Error('分组列没有数据');
    const base = p.sheetTitle || ('按' + cols.map(c => (c.unit === 'month' ? '月' : c.unit === 'year' ? '年' : c.col.name)).join('、') + '汇总');
    const sheetName = nextOutputSheetName(wb, base);
    op.newSheet = sheetName;
    op.outputs = [];
    op.blocks = [];
    const header = cols.map(c => c.col.name + (c.unit === 'month' ? '(月份)' : c.unit === 'year' ? '(年份)' : '')).concat(p.outputs.map(o => o.title));
    op.blocks.push({ sheet: sheetName, row: 0, col: 0, data: [header] });
    op.blocks.push({ sheet: sheetName, row: 1, col: 0, data: keys });
    op.keyCols = cols.map((c, i) => ({ name: header[i], unit: c.unit, col: i, srcFormat: c.col.type === 'date' ? 'date' : null }));
    op.keyValues = keys;
    p.outputs.forEach((o, i) => {
      const col = cols.length + i;
      const formulas = keys.map((k, ki) => {
        const keysMap = {};
        cols.forEach((c, ci) => { keysMap['KEY' + (ci + 1)] = '$' + U.colLetter(ci) + (ki + 2); });
        keysMap.KEY = keysMap.KEY1;
        return F.expand(o.formula, { resolve, hostSheet: sheetName, row: null, allowRow: false, keys: keysMap });
      });
      op.outputs.push({ title: o.title, format: o.format, sheet: sheetName, col, rows: keys.map((_, ki) => ki + 1), formulas, sampleFormula: formulas[0], dsl: o.formula });
      op.blocks.push({ sheet: sheetName, row: 1, col, data: formulas.map(f => [f]) });
    });
    op.groupSheet = { name: sheetName, header };
  }

  /** 撤销待确认的计算 */
  function discardPending(wb) {
    if (!wb.pending) return;
    try { wb.hf.undo(); } catch (e) { replayAll(wb); }
    wb.pending = null;
  }

  /** 确认待定计算：更新数据模型（新增列可被后续需求引用） */
  function commitPending(wb) {
    const op = wb.pending;
    if (!op) return null;
    wb.pending = null;
    registerOp(wb, op);
    wb.ops.push(op);
    try { wb.hf.clearUndoStack(); } catch (e) { /* ignore */ }
    return op;
  }

  function registerOp(wb, op) {
    if (op.kind === 'column') {
      const src = wb.sheets.find(s => s.name === op.sheet);
      for (const o of op.outputs) {
        // 把结果写回数据模型，后续需求可以引用这一列
        for (let i = 0; i < o.rows.length; i++) {
          const r = o.rows[i];
          while (src.values[r].length <= o.col) src.values[r].push(null);
          const v = o.values[i];
          src.values[r][o.col] = isErrorValue(v) ? v.error : v;
          const z = { date: 'yyyy-mm-dd', percent: '0.00%' }[o.format];
          if (z && typeof v === 'number') { src.fmts[r] = src.fmts[r] || []; src.fmts[r][o.col] = z; }
          if (src.display[r]) src.display[r][o.col] = formatValue(v, o.format);
        }
        while (src.values[src.headerRow].length <= o.col) src.values[src.headerRow].push(null);
        src.values[src.headerRow][o.col] = o.title;
        if (src.display[src.headerRow]) src.display[src.headerRow][o.col] = o.title;
        src.lastCol = Math.max(src.lastCol, o.col);
        const col = { index: o.col, letter: U.colLetter(o.col), name: o.title, rawHeader: o.title, added: true, opId: op.id };
        computeColumnStats(src, col);
        src.columns.push(col);
      }
    } else if (op.kind === 'scalar') {
      let out = wb.outputs.find(x => x.kind === 'scalar');
      if (!out) { out = { kind: 'scalar', name: op.scalarSheet.name, items: [], nextRow: 1 }; wb.outputs.push(out); }
      for (const o of op.outputs) out.items.push({ opId: op.id, title: o.title, row: o.rows[0], formula: o.formulas[0], value: o.values[0], format: o.format, request: op.plan.request || '' });
      out.nextRow = op.scalarSheet.nextRow;
    } else {
      wb.outputs.push({ kind: 'group', name: op.groupSheet.name, opId: op.id, header: op.groupSheet.header, keyCols: op.keyCols, keyValues: op.keyValues, outputs: op.outputs });
    }
  }

  /** 删除一个已确认的计算：重建引擎并按顺序重放剩余计算 */
  function removeOp(wb, opId) {
    discardPending(wb);
    const keep = wb.ops.filter(o => o.id !== opId);
    return replay(wb, keep);
  }

  function replayAll(wb) { return replay(wb, wb.ops.slice()); }

  function replay(wb, ops) {
    // 还原原始数据
    for (const s of wb.sheets) {
      for (const row of s.values) if (row.length > s.origLastCol + 1) row.length = s.origLastCol + 1;
      for (const row of s.display) if (row && row.length > s.origLastCol + 1) row.length = s.origLastCol + 1;
      for (const row of s.fmts) if (row && row.length > s.origLastCol + 1) row.length = s.origLastCol + 1;
      s.lastCol = s.origLastCol;
      s.columns = s.columns.filter(c => !c.added);
    }
    wb.ops = [];
    buildEngine(wb);
    const dropped = [];
    for (const old of ops) {
      try {
        applyPending(wb, old.plan);
        commitPending(wb);
      } catch (e) {
        dropped.push({ op: old, reason: e.message });
      }
    }
    return dropped;
  }

  /** 修改表头行：清空已添加的计算 */
  function changeHeaderRow(wb, sheetIndex, headerRow) {
    discardPending(wb);
    const s = wb.sheets[sheetIndex];
    const ops = wb.ops.slice();
    // 先回到原始数据
    replay(wb, []);
    setHeaderRow(s, headerRow);
    const dropped = [];
    for (const old of ops) {
      try { applyPending(wb, old.plan); commitPending(wb); } catch (e) { dropped.push({ op: old, reason: e.message }); }
    }
    return dropped;
  }

  EFA.Workbook = {
    readFile, fromBuffer, setHeaderRow, detectHeader, displayValue, formatValue, isErrorValue,
    applyPending, commitPending, discardPending, removeOp, changeHeaderRow, resolver, isDateFmt,
    ENGINE_CONFIG,
  };
})(window.EFA = window.EFA || {});
