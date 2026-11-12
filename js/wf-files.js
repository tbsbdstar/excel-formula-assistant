/* 工作流配置器：读取样例文件的结构（表头、列类型、少量样本值）。只在本机浏览器内读取，不上传。 */
(function (EFA) {
  'use strict';
  const U = EFA.U;
  const F = {};

  F.SAMPLE_ROWS = 60;          // 每个工作表只读取前若干行，用来识别表头和列类型
  F.MAX_FILE_MB = 60;

  /** 与执行器一致的表头规范化：换行/连续空白 → 单个空格，去首尾空格 */
  F.normHeader = function (v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return dateText(v);
    if (typeof v === 'number' && Number.isInteger(v)) return String(v);
    return String(v).replace(/\n/g, ' ').replace(/\s+/g, ' ').trim();
  };

  function dateText(d) {
    const p = n => (n < 10 ? '0' : '') + n;
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  function isBlank(v) {
    return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  }
  F.isBlank = isBlank;

  /** 读取一个文件 → { fileName, size, sheets: [{ name, rows, fullRows, fullCols, headerRow }] } */
  F.read = async function (file) {
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (!['xlsx', 'xlsm'].includes(ext)) {
      const e = new Error(ext === 'xls'
        ? '「' + file.name + '」是旧版 .xls 文件。执行器只能处理 .xlsx，请先在 Excel/WPS 里“另存为”.xlsx 再载入。'
        : '「' + file.name + '」不是 Excel 文件（.xlsx）。');
      e.kind = 'user';
      throw e;
    }
    if (file.size > F.MAX_FILE_MB * 1024 * 1024) {
      const e = new Error('「' + file.name + '」超过 ' + F.MAX_FILE_MB + 'MB。配置时只需要结构，可以先删掉大部分数据行，另存一份小样例再载入。');
      e.kind = 'user';
      throw e;
    }
    const buf = await file.arrayBuffer();
    let wb;
    try {
      wb = XLSX.read(buf, { type: 'array', sheetRows: F.SAMPLE_ROWS, dense: true, cellDates: true, cellFormula: false, cellHTML: false, cellStyles: false });
    } catch (err) {
      const e = new Error('无法读取「' + file.name + '」：文件可能已损坏、加了密码，或不是真正的 Excel 文件。');
      e.kind = 'user';
      throw e;
    }
    const sheets = wb.SheetNames.map(name => readSheet(name, wb.Sheets[name]));
    return { fileName: file.name, size: file.size, sheets };
  };

  function readSheet(name, ws) {
    const rows = [];
    const data = ws['!data'] || [];
    let maxC = -1;
    for (let r = 0; r < data.length; r++) {
      const src = data[r];
      if (!src) { rows.push([]); continue; }
      const row = [];
      for (let c = 0; c < src.length; c++) {
        const cell = src[c];
        if (!cell || cell.t === 'z') { row.push(null); continue; }
        let v = cell.v;
        if (cell.t === 'e') v = cell.w || '#ERR';
        row.push(v === undefined ? null : v);
        if (!isBlank(v)) maxC = Math.max(maxC, c);
      }
      rows.push(row);
    }
    // 去掉末尾空行
    while (rows.length && rows[rows.length - 1].every(isBlank)) rows.pop();
    const full = ws['!fullref'] || ws['!ref'];
    let fullRows = rows.length, fullCols = maxC + 1;
    if (full) {
      try {
        const rg = XLSX.utils.decode_range(full);
        fullRows = Math.max(fullRows, rg.e.r + 1);
        fullCols = Math.max(fullCols, rg.e.c + 1);
      } catch (e) { /* ignore */ }
    }
    const sheet = { name, rows, fullRows, fullCols, sampleCols: maxC + 1 };
    sheet.headerRow = detectHeader(sheet);
    return sheet;
  }

  /** 自动识别表头行（从 1 开始）：前 15 行里第一行“多数是文字、且宽度与下方数据相当”的行 */
  function detectHeader(sheet) {
    const rows = sheet.rows;
    const width = r => (rows[r] || []).filter(v => !isBlank(v)).length;
    let firstNonEmpty = -1;
    for (let r = 0; r < Math.min(rows.length, 15); r++) {
      const row = rows[r] || [];
      const vals = row.filter(v => !isBlank(v));
      if (!vals.length) continue;
      if (firstNonEmpty < 0) firstNonEmpty = r;
      const texts = vals.filter(v => typeof v === 'string' && isNaN(Number(v))).length;
      let below = 0;
      for (let k = 1; k <= 3 && r + k < rows.length; k++) below = Math.max(below, width(r + k));
      if (texts >= Math.ceil(vals.length * 0.6) && vals.length >= Math.ceil(below * 0.6)) return r + 1;
    }
    return Math.max(0, firstNonEmpty) + 1;
  }
  F.detectHeader = detectHeader;

  /** 某个工作表在指定表头行下的列信息：[{ idx, letter, name, type, samples, sensitive }] */
  F.columns = function (sheet, headerRow) {
    if (!sheet) return [];
    const hr = (headerRow || 1) - 1;
    const head = sheet.rows[hr] || [];
    const width = Math.max(head.length, sheet.sampleCols || 0);
    const out = [];
    for (let c = 0; c < width; c++) {
      const name = F.normHeader(head[c]);
      const vals = [];
      for (let r = hr + 1; r < sheet.rows.length && vals.length < 40; r++) {
        const v = (sheet.rows[r] || [])[c];
        if (!isBlank(v)) vals.push(v);
      }
      if (!name && !vals.length) continue;
      out.push({ idx: c, letter: U.colLetter(c), name, type: typeOf(vals), samples: vals.slice(0, 3).map(show), sensitive: sensitiveOf(name, vals) });
    }
    return out;
  };

  /** 表头名称数组（下标=列号，从 0 开始；空表头为 ''） */
  F.headerList = function (sheet, headerRow) {
    if (!sheet) return null;
    const head = sheet.rows[(headerRow || 1) - 1] || [];
    const width = Math.max(head.length, sheet.sampleCols || 0);
    const out = [];
    for (let c = 0; c < width; c++) out.push(F.normHeader(head[c]));
    while (out.length && !out[out.length - 1]) out.pop();
    return out;
  };

  /** 表头行之下的数据行数（按读取到的文件尺寸估算） */
  F.dataRows = function (sheet, headerRow) {
    return Math.max(0, (sheet.fullRows || 0) - (headerRow || 1));
  };

  function show(v) {
    if (v instanceof Date) return dateText(v);
    if (typeof v === 'number') return U.formatNumber(v);
    const s = String(v);
    return s.length > 18 ? s.slice(0, 17) + '…' : s;
  }

  function typeOf(vals) {
    if (!vals.length) return 'empty';
    let num = 0, date = 0, textNum = 0;
    for (const v of vals) {
      if (v instanceof Date) date++;
      else if (typeof v === 'number') num++;
      else if (typeof v === 'string' && /^\s*-?\d+(\.\d+)?\s*$/.test(v)) textNum++;
    }
    const n = vals.length;
    if (date / n >= 0.8) return 'date';
    if (num / n >= 0.8) return 'number';
    if (textNum / n >= 0.8) return 'textnum';
    return 'text';
  }
  F.TYPE_LABEL = { number: '数字', date: '日期', text: '文本', textnum: '文本数字', empty: '空' };

  const ID_RE = /^\d{17}[\dXx]$/;
  const PHONE_RE = /^1[3-9]\d{9}$/;
  function sensitiveOf(name, vals) {
    if (/身份证|证件号/.test(name)) return '身份证号';
    if (/手机|电话|联系方式/.test(name)) return '手机号';
    if (/^(姓名|名字|联系人|收货人|收件人)$/.test(name)) return '姓名';
    const sample = vals.slice(0, 20).map(v => String(v).trim());
    if (sample.length >= 3) {
      if (sample.filter(s => ID_RE.test(s)).length / sample.length >= 0.6) return '身份证号';
      if (sample.filter(s => PHONE_RE.test(s)).length / sample.length >= 0.6) return '手机号';
    }
    return '';
  }

  /** 根据样例文件名推荐“文件名规则”：去掉日期、编号等每次都会变的部分 */
  F.suggestPattern = function (fileName) {
    const dot = fileName.lastIndexOf('.');
    let stem = dot > 0 ? fileName.slice(0, dot) : fileName;
    const ext = dot > 0 ? fileName.slice(dot).toLowerCase() : '.xlsx';
    stem = stem
      .replace(/[（(]\s*\d+\s*[)）]/g, '*')                                  // (1) 副本编号
      .replace(/\d+(?:[年月日号周.\-_/]\d*)*[年月日号周]?/g, '*')             // 日期、期数等数字
      .replace(/(?:^|[\s_\-—])(?:副本|最终版?|终版|new|copy)(?=$|[\s_\-—])/gi, '*')
      .replace(/[\s_\-—.]*\*[\s_\-—.*]*/g, '*')
      .replace(/^\*+|\*+$/g, '');
    if (!stem) return '*' + ext;
    return '*' + stem + '*' + ext;
  };

  /** 与执行器一致：文件名规则（* ? 通配，不区分大小写） */
  F.matchPattern = function (pattern, name) {
    if (!pattern) return true;
    const re = '^' + String(pattern).toLowerCase().replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$';
    try { return new RegExp(re).test(String(name).toLowerCase()); } catch (e) { return false; }
  };

  EFA.WFFiles = F;
})(window.EFA = window.EFA || {});
