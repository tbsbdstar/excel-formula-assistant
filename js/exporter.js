/* 导出 .xlsx：在原文件基础上写入新增列、结果表；可选“结果值”或“公式”两种模式 */
(function (EFA) {
  'use strict';
  const U = EFA.U, W = EFA.Workbook;

  const NUMFMT = {
    date: 'yyyy-mm-dd', percent: '0.00%', integer: '0', money: '#,##0.00', number: '0.00',
    month: 'yyyy"年"m"月"', year: 'yyyy"年"',
  };
  const HEADER_STYLE = {
    font: { bold: true, color: { argb: 'FF1F2937' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F3EC' } },
    alignment: { vertical: 'middle', horizontal: 'center' },
    border: { bottom: { style: 'thin', color: { argb: 'FF9CA3AF' } } },
  };

  function toCellValue(v) {
    if (v === null || v === undefined) return null;
    if (W.isErrorValue(v)) return { error: v.error };
    return v;
  }

  function clone(o) { return o ? JSON.parse(JSON.stringify(o)) : {}; }

  /** 原文件不是 xlsx（或无法读取）时，按数据重建工作簿 */
  function rebuild(ewb, wb) {
    for (const s of wb.sheets) {
      const ws = ewb.addWorksheet(s.name);
      for (let r = 0; r <= s.lastRow; r++) {
        const row = s.values[r] || [];
        for (let c = 0; c <= s.origLastCol; c++) {
          const v = row[c];
          const f = s.formulas[r + ',' + c];
          if (v === null || v === undefined || v === '') { if (!f) continue; }
          const cell = ws.getCell(r + 1, c + 1);
          if (f) cell.value = { formula: f, result: typeof v === 'string' && /^#/.test(v) ? { error: v } : v };
          else if (typeof v === 'string' && /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/.test(v)) cell.value = { error: v };
          else cell.value = v;
          const z = s.fmts[r] && s.fmts[r][c];
          if (z && z !== 'General') cell.numFmt = z;
        }
      }
      if (s.cols) s.cols.forEach((ci, i) => { if (ci && ci.wch) ws.getColumn(i + 1).width = ci.wch + 1; });
      for (const m of s.merges || []) {
        try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch (e) { /* ignore */ }
      }
      if (s.headerRow >= 0 && s.lastRow >= 0) {
        ws.getRow(s.headerRow + 1).font = { bold: true };
      }
    }
  }

  function setCell(cell, formula, value, mode, fmt) {
    if (mode === 'formula' && formula) {
      cell.value = { formula: formula.replace(/^=/, ''), result: toCellValue(value) };
    } else {
      cell.value = toCellValue(value);
    }
    cell.numFmt = NUMFMT[fmt] || 'General';
  }

  function styleHeader(row, n) {
    for (let c = 1; c <= n; c++) Object.assign(row.getCell(c), clone(HEADER_STYLE));
  }

  async function exportWorkbook(wb, mode) {
    const ewb = new ExcelJS.Workbook();
    const notes = [];
    let loaded = false;
    if (wb.kind === 'xlsx' && wb.buffer) {
      try {
        await ewb.xlsx.load(wb.buffer.slice(0));
        loaded = true;
      } catch (e) {
        console.warn('原文件无法按原样载入，改为重建', e);
        notes.push('原文件中部分格式（如图表、透视表）无法保留，已按数据重建');
      }
    }
    if (!loaded) {
      if (wb.kind !== 'xlsx' && !notes.length) notes.push('原文件为 ' + wb.kind.toUpperCase() + '，已转换为 .xlsx（原有字体颜色等样式可能丢失）');
      if (wb.skipped && wb.skipped.length) notes.push('超出处理上限、未载入的工作表没有包含在新文件里');
      const fresh = new ExcelJS.Workbook();
      rebuild(fresh, wb);
      return finish(fresh, wb, mode, notes);
    }
    return finish(ewb, wb, mode, notes);
  }

  async function finish(ewb, wb, mode, notes) {
    for (const op of wb.ops) {
      if (op.kind === 'column') writeColumnOp(ewb, op, mode);
    }
    // 计算结果表（单值）
    const sc = wb.outputs.find(o => o.kind === 'scalar');
    if (sc && sc.items.length) {
      const ws = ewb.getWorksheet(sc.name) || ewb.addWorksheet(sc.name);
      ws.getRow(1).values = ['计算项', '结果', '需求描述'];
      styleHeader(ws.getRow(1), 3);
      sc.items.forEach(it => {
        const r = it.row + 1;
        ws.getCell(r, 1).value = it.title;
        setCell(ws.getCell(r, 2), it.formula, it.value, mode, it.format);
        ws.getCell(r, 3).value = it.request || '';
      });
      ws.getColumn(1).width = 34; ws.getColumn(2).width = 18; ws.getColumn(3).width = 40;
      ws.views = [{ state: 'frozen', ySplit: 1 }];
    }
    // 分组汇总表
    for (const g of wb.outputs.filter(o => o.kind === 'group')) {
      const ws = ewb.getWorksheet(g.name) || ewb.addWorksheet(g.name);
      ws.getRow(1).values = g.header;
      styleHeader(ws.getRow(1), g.header.length);
      g.keyValues.forEach((keys, ki) => {
        keys.forEach((k, ci) => {
          const cell = ws.getCell(ki + 2, ci + 1);
          cell.value = k;
          const kc = g.keyCols[ci];
          if (kc.unit === 'month') cell.numFmt = NUMFMT.month;
          else if (kc.unit === 'year') cell.numFmt = NUMFMT.year;
          else if (kc.srcFormat === 'date') cell.numFmt = NUMFMT.date;
        });
      });
      g.outputs.forEach(o => {
        o.rows.forEach((r, i) => setCell(ws.getCell(r + 1, o.col + 1), o.formulas[i], o.values[i], mode, o.format));
      });
      g.header.forEach((h, i) => { ws.getColumn(i + 1).width = Math.max(12, Math.min(30, String(h).length * 2 + 4)); });
      ws.views = [{ state: 'frozen', ySplit: 1 }];
    }
    if (mode === 'formula') {
      ewb.calcProperties = ewb.calcProperties || {};
      ewb.calcProperties.fullCalcOnLoad = true;
    }
    const buf = await ewb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const base = wb.fileName.replace(/\.[^.]+$/, '');
    return { blob, fileName: base + (mode === 'formula' ? '_公式版.xlsx' : '_计算结果.xlsx'), notes };
  }

  function writeColumnOp(ewb, op, mode) {
    const ws = ewb.getWorksheet(op.sheet);
    if (!ws) return;
    for (const o of op.outputs) {
      const col = o.col + 1;
      const head = ws.getCell(o.headerRow + 1, col);
      const leftHead = ws.getCell(o.headerRow + 1, col - 1);
      head.style = clone(leftHead.style);
      if (!head.font || !head.font.bold) head.font = Object.assign({}, head.font || {}, { bold: true });
      head.value = o.title;
      o.rows.forEach((r, i) => {
        const cell = ws.getCell(r + 1, col);
        const left = ws.getCell(r + 1, col - 1);
        const st = clone(left.style);
        delete st.numFmt;
        cell.style = st;
        setCell(cell, o.formulas[i], o.values[i], mode, o.format);
      });
      const w = ws.getColumn(col - 1).width;
      ws.getColumn(col).width = Math.max(w || 10, Math.min(30, String(o.title).length * 2 + 4));
    }
  }

  EFA.Exporter = { exportWorkbook };
})(window.EFA = window.EFA || {});
