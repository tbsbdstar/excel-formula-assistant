/* 公式工具：扫描、列引用语法展开、函数白名单校验、预览引擎适配、语法高亮
 *
 * 列引用语法（模板与 AI 统一使用，展开后才是真正的 Excel 公式）：
 *   [列名]          该列整列数据区域，如 $D$2:$D$101（跨表时自动带表名）
 *   [工作表!列名]   其他工作表中该列的数据区域
 *   [@列名]         同一行的单元格，如 D2（仅“新增列”可用）
 *   [^列名]         从首行到当前行的累计区域，如 $D$2:D2（仅“新增列”可用）
 *   {KEY} {KEY2}    分组汇总表中当前行的分组值单元格（仅“分组汇总”可用）
 */
(function (EFA) {
  'use strict';
  const U = EFA.U;

  /** 把公式切成片段：str（"文本"）、sheetq（'表名'）、dsl（[...]）、key（{KEY}）、other */
  function segments(f) {
    const out = [];
    let i = 0, buf = '';
    const flush = () => { if (buf) { out.push({ t: 'other', v: buf }); buf = ''; } };
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"') {
        flush();
        let j = i + 1;
        while (j < f.length) {
          if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; }
          j++;
        }
        out.push({ t: 'str', v: f.slice(i, j + 1) });
        i = j + 1;
      } else if (ch === "'") {
        flush();
        let j = i + 1;
        while (j < f.length) {
          if (f[j] === "'") { if (f[j + 1] === "'") { j += 2; continue; } break; }
          j++;
        }
        out.push({ t: 'sheetq', v: f.slice(i, j + 1) });
        i = j + 1;
      } else if (ch === '[') {
        flush();
        const j = f.indexOf(']', i + 1);
        if (j < 0) { buf += f.slice(i); break; }
        out.push({ t: 'dsl', v: f.slice(i, j + 1) });
        i = j + 1;
      } else if (ch === '{' && /^\{KEY\d?\}/.test(f.slice(i, i + 6))) {
        flush();
        const j = f.indexOf('}', i);
        out.push({ t: 'key', v: f.slice(i, j + 1) });
        i = j + 1;
      } else {
        buf += ch;
        i++;
      }
    }
    flush();
    return out;
  }

  /** 解析一个 [..] 引用 */
  function parseDslRef(token) {
    let body = token.slice(1, -1).trim();
    let mode = 'range';
    if (body[0] === '@') { mode = 'row'; body = body.slice(1).trim(); }
    else if (body[0] === '^') { mode = 'cum'; body = body.slice(1).trim(); }
    let sheet = null, column = body;
    const bang = body.indexOf('!');
    if (bang > 0) { sheet = body.slice(0, bang).trim().replace(/^'|'$/g, ''); column = body.slice(bang + 1).trim(); }
    return { mode, sheet, column };
  }

  /** 列出公式中用到的列引用（用于校验和展示） */
  function listRefs(dsl) {
    return segments(dsl).filter(s => s.t === 'dsl').map(s => parseDslRef(s.v));
  }

  /**
   * 展开列引用语法为真实 Excel 公式
   * ctx = { resolve(sheetName|null, colName) -> {sheet, col}|null, hostSheet, row, keys, allowRow, allowKeys }
   */
  function expand(dsl, ctx) {
    let f = String(dsl).trim();
    if (!f.startsWith('=')) f = '=' + f;
    const segs = segments(f);
    let out = '';
    for (const s of segs) {
      if (s.t === 'dsl') {
        const ref = parseDslRef(s.v);
        let hit = ctx.resolve(ref.sheet, ref.column);
        // 列名本身可能含 “!”
        if (!hit && ref.sheet) hit = ctx.resolve(null, ref.sheet + '!' + ref.column);
        // [表!起始列:结束列] 多列区域（VLOOKUP 用）
        if (!hit && ref.mode === 'range' && ref.column.indexOf(':') > 0) {
          const [c1, c2] = ref.column.split(':').map(x => x.trim());
          const h1 = ctx.resolve(ref.sheet, c1), h2 = ctx.resolve(ref.sheet, c2);
          if (h1 && h2 && h1.sheet === h2.sheet) {
            const sh = h1.sheet;
            const lo = Math.min(h1.col.index, h2.col.index), hi = Math.max(h1.col.index, h2.col.index);
            const prefix = sh.name === ctx.hostSheet ? '' : U.quoteSheet(sh.name) + '!';
            out += prefix + '$' + U.colLetter(lo) + '$' + (sh.dataStart + 1) + ':$' + U.colLetter(hi) + '$' + (sh.dataEnd + 1);
            continue;
          }
        }
        if (!hit) throw new Error('找不到列「' + (ref.sheet ? ref.sheet + '!' : '') + ref.column + '」');
        out += rangeText(hit, ref.mode, ctx);
      } else if (s.t === 'key') {
        if (!ctx.keys) throw new Error('只有“分组汇总”才能使用 {KEY}');
        const name = s.v.slice(1, -1);
        const k = ctx.keys[name] || (name === 'KEY' ? ctx.keys.KEY1 : null);
        if (!k) throw new Error('无效的分组占位符 ' + s.v);
        out += k;
      } else {
        out += s.v;
      }
    }
    return out;
  }

  function rangeText(hit, mode, ctx) {
    const sh = hit.sheet, col = hit.col;
    const L = col.letter;
    const a = sh.dataStart + 1, b = sh.dataEnd + 1;
    const prefix = sh.name === ctx.hostSheet ? '' : U.quoteSheet(sh.name) + '!';
    if (mode === 'range') return prefix + '$' + L + '$' + a + ':$' + L + '$' + b;
    if (!ctx.allowRow) throw new Error('「[@' + col.name + ']」这类同行引用只能用于“新增列”');
    if (sh.name !== ctx.hostSheet) throw new Error('同行引用只能引用当前工作表的列：' + col.name);
    if (mode === 'row') return L + ctx.row;
    return '$' + L + '$' + a + ':' + L + ctx.row; // cum
  }

  /** 公式中用到的函数名（大写，去重） */
  function functionsUsed(f) {
    const names = new Set();
    for (const s of segments(f)) {
      if (s.t !== 'other') continue;
      const re = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
      let m;
      while ((m = re.exec(s.v))) names.add(m[1].toUpperCase());
    }
    return Array.from(names);
  }

  /** 校验函数是否全部属于 Excel 2016 且预览可算 */
  function checkFunctions(f) {
    const used = functionsUsed(f);
    const problems = [];
    const allowed = EFA.Functions.allowedSet();
    for (const n of used) {
      if (!EFA.Functions.isExcel2016(n)) {
        const alt = EFA.Functions.newerOnly(n);
        problems.push(alt ? n + ' 是 Excel 2019/365 才有的函数，Excel 2016 / WPS 旧版不可用（可改用 ' + alt + '）'
          : '不认识的函数 ' + n);
      } else if (!allowed.has(n)) {
        problems.push(n + ' 暂不支持在浏览器内预览计算');
      }
    }
    // 其他新版语法
    for (const s of segments(f)) {
      if (s.t !== 'other') continue;
      if (/@/.test(s.v)) problems.push('包含 @ 隐式交集语法（新版 Excel 才有）');
      if (/[A-Z]\d+#/.test(s.v)) problems.push('包含 # 溢出引用（新版 Excel 才有）');
    }
    return { used, problems };
  }

  /** 转成预览引擎可识别的写法（引擎不认裸 TRUE/FALSE） */
  function toEngine(f) {
    return segments(f).map(s => {
      if (s.t !== 'other') return s.v;
      return s.v.replace(/(^|[^A-Za-z0-9_.$])(TRUE|FALSE)(?![A-Za-z0-9_.(])/gi, (m, p, w) => p + w.toUpperCase() + '()');
    }).join('');
  }

  /** 语法高亮（HTML） */
  function highlight(f) {
    const esc = U.escapeHtml;
    let html = '';
    for (const s of segments(f)) {
      if (s.t === 'str') html += '<span class="fx-s">' + esc(s.v) + '</span>';
      else if (s.t === 'sheetq') html += '<span class="fx-r">' + esc(s.v) + '</span>';
      else if (s.t === 'dsl' || s.t === 'key') html += '<span class="fx-r">' + esc(s.v) + '</span>';
      else {
        html += esc(s.v).replace(/([A-Za-z_][A-Za-z0-9_.]*)(\s*\()|(!?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|(\b\d+(?:\.\d+)?\b)/g,
          (m, fn, paren, ref, num) => {
            if (fn) return '<span class="fx-f">' + fn + '</span>' + paren;
            if (ref) return '<span class="fx-r">' + ref + '</span>';
            return '<span class="fx-n">' + num + '</span>';
          });
      }
    }
    return html;
  }

  EFA.Formula = { segments, parseDslRef, listRefs, expand, functionsUsed, checkFunctions, toEngine, highlight };
})(window.EFA = window.EFA || {});
