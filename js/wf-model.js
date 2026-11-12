/* 工作流配置器：配置模型、逐步推算表头、参数校验、步骤摘要、导入导出。
 * 配置格式与执行器一致（excel-workflow/1），步骤对象就是最终导出的 JSON，界面直接修改它。 */
(function (EFA) {
  'use strict';
  const U = EFA.U, F = EFA.WFFiles, CAT = EFA.WF_CATALOG;
  const M = {};

  M.SCHEMA = 'excel-workflow/1';
  M.COMMON = ['op', 'title', 'file', 'sheet', 'header_row', 'enabled', 'note'];
  M.OPS = {};
  CAT.ops.forEach(d => { M.OPS[d.op] = d; (d.aliases || []).forEach(a => { M.OPS[a] = d; }); });
  M.GROUPS = CAT.groups;

  M.COND_OPS = [
    { v: '=', label: '等于', sym: '=' }, { v: '!=', label: '不等于', sym: '≠' },
    { v: '>', label: '大于', sym: '>' }, { v: '>=', label: '大于等于', sym: '≥' },
    { v: '<', label: '小于', sym: '<' }, { v: '<=', label: '小于等于', sym: '≤' },
    { v: 'contains', label: '包含', sym: '包含' }, { v: 'not_contains', label: '不包含', sym: '不包含' },
    { v: 'starts_with', label: '开头是', sym: '开头是' }, { v: 'ends_with', label: '结尾是', sym: '结尾是' },
    { v: 'empty', label: '为空', sym: '为空', noValue: true }, { v: 'not_empty', label: '不为空', sym: '不为空', noValue: true },
    { v: 'in', label: '是以下之一', sym: '是以下之一', list: true }, { v: 'not_in', label: '不是以下之一', sym: '不是以下之一', list: true },
    { v: 'regex', label: '符合正则', sym: '符合正则' },
    { v: 'is_number', label: '是数字', sym: '是数字', noValue: true }, { v: 'is_text', label: '是文本', sym: '是文本', noValue: true },
    { v: 'is_error', label: '是错误值', sym: '是错误值', noValue: true },
  ];
  const OP_ALIAS = { '==': '=', '<>': '!=', 等于: '=', 不等于: '!=', 大于: '>', 大于等于: '>=', 小于: '<', 小于等于: '<=',
    包含: 'contains', 不包含: 'not_contains', 开头是: 'starts_with', 结尾是: 'ends_with', 为空: 'empty', 不为空: 'not_empty', 属于: 'in', 不属于: 'not_in' };
  M.canonOp = op => OP_ALIAS[op] || op;
  M.condOp = op => M.COND_OPS.find(o => o.v === M.canonOp(op));

  const NAMED_COLORS = { 红: 'FFC7CE', 黄: 'FFEB9C', 绿: 'C6EFCE', 蓝: 'BDD7EE', 橙: 'FCE4D6', 灰: 'D9D9D9' };
  M.NAMED_COLORS = NAMED_COLORS;
  /** 与执行器一致的颜色解析，返回 6 位十六进制或 null */
  M.parseColor = function (v) {
    let s = String(v == null ? '' : v).trim().replace(/^#/, '').toUpperCase();
    s = NAMED_COLORS[s] || s;
    if (/^[0-9A-F]{6}$/.test(s)) return s;
    if (/^[0-9A-F]{8}$/.test(s)) return s.slice(2);
    return null;
  };

  M.newConfig = () => ({ schema: M.SCHEMA, name: '', description: '', files: [], options: { recalc: 'auto' }, steps: [] });

  /* ------------------------------------------------------------------ 状态 */
  M.state = { cfg: M.newConfig(), samples: {} };   // samples[文件id] = { fileName, size, sheets }

  M.mainId = function (cfg) {
    const f = (cfg || M.state.cfg).files.find(x => x.role === 'main');
    return f ? f.id : null;
  };
  M.fileSpec = id => M.state.cfg.files.find(f => f.id === id) || null;
  M.fileLabel = function (id) {
    const f = M.fileSpec(id);
    return f ? (f.label || f.id) : String(id);
  };
  M.newFileId = function (role) {
    const taken = new Set(M.state.cfg.files.map(f => f.id));
    if (role === 'main' && !taken.has('main')) return 'main';
    return U.uniqueName('ref', taken);
  };

  /** 解析工作表：sheet 可为名称、从 1 开始的序号或空（=文件设置的工作表，再空=第一个） */
  M.sampleSheet = function (fileId, sheet) {
    const smp = M.state.samples[fileId];
    if (!smp) return null;
    const spec = M.fileSpec(fileId) || {};
    const want = (sheet === undefined || sheet === null || sheet === '') ? spec.sheet : sheet;
    if (want === undefined || want === null || want === '') return smp.sheets[0] || null;
    if (typeof want === 'number') return smp.sheets[want - 1] || null;
    return smp.sheets.find(s => s.name === want) || smp.sheets.find(s => s.name.trim() === String(want).trim()) || null;
  };

  /* ------------------------------------------------------------------ 列定位（与执行器 SheetView.col / cols 一致） */
  const LETTERS = /^[A-Za-z]{1,3}$/;
  M.isLetters = s => LETTERS.test(String(s));

  M.resolveCol = function (headers, ref) {
    if (ref && typeof ref === 'object') {
      if ('letter' in ref) return M.isLetters(ref.letter) ? { idx: U.colIndex(ref.letter), how: 'letter' } : { error: '列字母「' + ref.letter + '」不正确' };
      if ('index' in ref) return Number(ref.index) >= 1 ? { idx: Number(ref.index) - 1, how: 'index' } : { error: '列序号应从 1 开始' };
      if ('name' in ref) ref = ref.name;
      else return { error: '列的写法不正确' };
    }
    if (typeof ref === 'number') return { idx: ref - 1, how: 'index' };
    const name = F.normHeader(ref);
    if (!name) return { error: '列名为空' };
    if (headers) {
      let i = headers.indexOf(name);
      if (i >= 0) return { idx: i, how: 'name' };
      const compact = name.replace(/ /g, '');
      i = headers.findIndex(h => h && h.replace(/ /g, '') === compact);
      if (i >= 0) return { idx: i, how: 'name' };
    }
    if (M.isLetters(name)) return { idx: U.colIndex(name), how: 'letter' };
    const m = /^([A-Za-z]{1,3})列$/.exec(name);
    if (m) return { idx: U.colIndex(m[1]), how: 'letter' };
    if (!headers) return { idx: -1, how: 'unknown' };
    return { error: '找不到列「' + name + '」' };
  };

  /** 多列：列表 / 'U:AB' / '表头1:表头2' / {from,to} */
  M.resolveCols = function (headers, spec) {
    const out = { idxs: [], errors: [] };
    const one = ref => {
      const r = M.resolveCol(headers, ref);
      if (r.error) out.errors.push(r.error); else if (r.idx >= 0) out.idxs.push(r.idx);
    };
    const range = (a, b) => {
      const ra = M.resolveCol(headers, a), rb = M.resolveCol(headers, b);
      if (ra.error) out.errors.push(ra.error);
      if (rb.error) out.errors.push(rb.error);
      if (ra.idx >= 0 && rb.idx >= 0) for (let i = Math.min(ra.idx, rb.idx); i <= Math.max(ra.idx, rb.idx); i++) out.idxs.push(i);
    };
    const rangeText = s => {
      const [a, b] = s.split(':', 2).map(x => x.trim());
      const isHdr = x => headers && headers.includes(F.normHeader(x));
      if (M.isLetters(a) && M.isLetters(b) && !isHdr(a) && !isHdr(b)) range({ letter: a }, { letter: b });
      else range(a, b);
    };
    if (spec === undefined || spec === null || spec === '') { out.errors.push('没有选择列'); return out; }
    if (Array.isArray(spec)) {
      if (!spec.length) out.errors.push('没有选择列');
      spec.forEach(s => (typeof s === 'string' && s.includes(':') ? rangeText(s) : one(s)));
    } else if (typeof spec === 'object' && 'from' in spec) range(spec.from, spec.to);
    else if (typeof spec === 'string' && spec.includes(':')) rangeText(spec);
    else one(spec);
    return out;
  };

  M.colText = function (spec) {
    if (spec === undefined || spec === null) return '';
    const t = x => (x && typeof x === 'object') ? (x.letter ? x.letter + '列' : x.index ? '第' + x.index + '列' : x.name || JSON.stringify(x)) : String(x);
    if (Array.isArray(spec)) return spec.map(t).join('、');
    if (typeof spec === 'object' && 'from' in spec) return t(spec.from) + ' 到 ' + t(spec.to);
    return t(spec);
  };

  /* ------------------------------------------------------------------ 逐步推算表头 */
  /** 返回 simulate 结果：view(i, fileId, sheet, headerRow) → 第 i 步执行前该表的表头（null=没有样例，无法得知） */
  M.simulate = function (cfg) {
    cfg = cfg || M.state.cfg;
    const mainId = M.mainId(cfg);
    const tables = new Map();        // 已被改动过的表：key → { headers, name }
    const extraSheets = [];          // 复制进主表的工作表名
    const snaps = [];

    function keyOf(fileId, sheetName, hr) { return fileId + '\u0001' + sheetName + '\u0001' + hr; }
    function locate(map, fileId, sheet, headerRow) {
      const spec = cfg.files.find(f => f.id === fileId) || {};
      const hr = Number(headerRow || spec.header_row || 1);
      let sheetName;
      const smpSheet = M.sampleSheet(fileId, sheet);
      if (smpSheet) sheetName = smpSheet.name;
      else if (sheet !== undefined && sheet !== null && sheet !== '') sheetName = String(sheet);
      else if (spec.sheet !== undefined && spec.sheet !== null && spec.sheet !== '') sheetName = String(spec.sheet);
      else sheetName = '';
      const key = keyOf(fileId, sheetName, hr);
      if (map.has(key)) return { key, sheetName, hr, headers: map.get(key).headers };
      const base = smpSheet ? F.headerList(smpSheet, hr) : null;
      return { key, sheetName, hr, headers: base };
    }

    for (let i = 0; i < cfg.steps.length; i++) {
      snaps.push(new Map(Array.from(tables, ([k, v]) => [k, { headers: v.headers ? v.headers.slice() : null }])));
      const s = cfg.steps[i];
      if (!s || s.enabled === false) continue;
      const def = M.OPS[s.op];
      if (!def || def.writes === false) continue;
      const fileId = s.file || mainId;
      if (fileId !== mainId) continue;
      const t = locate(tables, fileId, s.sheet, s.header_row);
      if (!t.headers && s.op !== 'copy_sheet') continue;
      const h = t.headers ? t.headers.slice() : null;
      const col = ref => { const r = M.resolveCol(h, ref); return r.error ? -1 : r.idx; };
      let changed = true;
      switch (s.op) {
        case 'delete_columns': {
          const r = M.resolveCols(h, s.columns);
          Array.from(new Set(r.idxs)).sort((a, b) => b - a).forEach(i2 => { if (i2 < h.length) h.splice(i2, 1); });
          break;
        }
        case 'insert_column': {
          let k = h.length;
          if (s.before !== undefined) k = col(s.before);
          else if (s.after !== undefined) { k = col(s.after); if (k >= 0) k += 1; }
          else if (s.at !== undefined) k = typeof s.at === 'number' ? s.at - 1 : (M.isLetters(s.at) ? U.colIndex(s.at) : -1);
          if (k < 0) { changed = false; break; }
          while (h.length < k) h.push('');
          h.splice(k, 0, F.normHeader(s.header));
          break;
        }
        case 'rename_header': {
          const k = col(s.column);
          if (k >= 0) { while (h.length <= k) h.push(''); h[k] = F.normHeader(s.to); }
          break;
        }
        case 'set_cells': {
          Object.keys(s.cells || {}).forEach(addr => {
            const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(addr.trim());
            if (m && Number(m[2]) === t.hr) {
              const k = U.colIndex(m[1]);
              while (h.length <= k) h.push('');
              const v = s.cells[addr];
              h[k] = (typeof v === 'string' && v.startsWith('=')) ? '' : F.normHeader(v);
            }
          });
          break;
        }
        case 'copy_matched':
          if (s.create_missing) {
            M.copyPairs(s).forEach(p => { if (M.resolveCol(h, p.to).error) h.push(F.normHeader(p.to)); });
          }
          break;
        case 'compare': {
          const ensure = name => { if (name && M.resolveCol(h, name).error) h.push(F.normHeader(name)); };
          ensure(s.status_column || '比对结果');
          ensure(s.diff_column === undefined ? '差异列' : s.diff_column);
          break;
        }
        case 'copy_sheet': {
          const src = s.from || {};
          const srcSheet = M.sampleSheet(src.file || mainId, src.sheet);
          const name = s.name || (srcSheet ? srcSheet.name : src.sheet);
          if (name) {
            if (!extraSheets.includes(name)) extraSheets.push(name);
            const srcSpec = cfg.files.find(f => f.id === src.file) || {};
            const hdr = srcSheet ? F.headerList(srcSheet, srcSpec.header_row || 1) : null;
            tables.set(keyOf(mainId, name, Number(cfg.files.find(f => f.id === mainId).header_row || 1)), { headers: hdr });
          }
          changed = false;
          break;
        }
        default:
          changed = false;
      }
      if (changed) tables.set(t.key, { headers: h });
    }
    snaps.push(new Map(tables));
    return {
      /** 第 i 步执行前的表头；i = steps.length 表示全部步骤之后 */
      headersAt(i, fileId, sheet, headerRow) {
        const snap = snaps[Math.min(i, snaps.length - 1)] || new Map();
        return locate(snap, fileId || mainId, sheet, headerRow).headers;
      },
      extraSheets,
    };
  };

  /** copy_matched 的列对：[{from, to}] */
  M.copyPairs = function (s) {
    if (Array.isArray(s.columns)) return s.columns.map(p => (p && typeof p === 'object') ? { from: p.from, to: p.to === undefined ? p.from : p.to } : { from: p, to: p });
    if (s.from !== undefined) {
      const fr = Array.isArray(s.from) ? s.from : [s.from];
      const to = s.to === undefined ? fr : (Array.isArray(s.to) ? s.to : [s.to]);
      return fr.map((f, i) => ({ from: f, to: to[i] === undefined ? f : to[i] }));
    }
    return [];
  };

  /** 第 i 步之前定义过（且启用）的匹配：[{id, ref, step}] */
  M.matchesBefore = function (i, cfg) {
    cfg = cfg || M.state.cfg;
    const out = [];
    for (let k = 0; k < i && k < cfg.steps.length; k++) {
      const s = cfg.steps[k];
      if (s && s.op === 'match' && s.enabled !== false && s.id) out.push({ id: s.id, ref: s.ref, step: s, index: k });
    }
    return out;
  };

  /** 第 i 步匹配到的参照表信息（用于参照表列名提示） */
  M.refOfMatch = function (i, matchId, sim) {
    const m = M.matchesBefore(i).filter(x => x.id === matchId).pop();
    if (!m) return null;
    const headers = sim.headersAt(i, m.ref, m.step.ref_sheet, m.step.ref_header_row);
    return { fileId: m.ref, headers, step: m.step };
  };

  /** 某一步作用的目标表 */
  M.stepTarget = function (s) {
    return { fileId: s.file || M.mainId(), sheet: s.sheet, headerRow: s.header_row };
  };

  /* ------------------------------------------------------------------ 值的解析 */
  M.parseScalar = function (text) {
    if (text === null || text === undefined) return null;
    const s = String(text);
    const t = s.trim();
    if (t === '') return null;
    if (/^-?(0|[1-9]\d*)(\.\d+)?$/.test(t)) return Number(t);
    return s;
  };
  M.scalarText = v => (v === null || v === undefined) ? '' : String(v);
  M.parseList = text => String(text || '').split(/[,，、;；\n]/).map(x => x.trim()).filter(Boolean).map(M.parseScalar);

  M.cellAddrOk = a => /^\$?[A-Za-z]{1,3}\$?\d+$/.test(String(a || '').trim());
  M.rangeOk = a => /^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$/.test(String(a || '').trim());

  /* ------------------------------------------------------------------ 条件 */
  M.condText = function (c) {
    if (!c || typeof c !== 'object') return '';
    if (Array.isArray(c.all)) return c.all.map(x => wrap(x)).join(' 且 ');
    if (Array.isArray(c.any)) return c.any.map(x => wrap(x)).join(' 或 ');
    if (c.not) return '不满足（' + M.condText(c.not) + '）';
    const side = c.ref_col !== undefined ? '参照表' : '';
    const col = c.ref_col !== undefined ? c.ref_col : c.col;
    const op = M.condOp(c.op);
    let rhs = '';
    if (c.col2 !== undefined) rhs = '「' + c.col2 + '」列';
    else if (c.ref_col2 !== undefined) rhs = '参照表「' + c.ref_col2 + '」列';
    else if (op && op.noValue) rhs = '';
    else if (Array.isArray(c.value)) rhs = c.value.join('、');
    else if (c.value !== undefined) rhs = typeof c.value === 'string' ? '“' + c.value + '”' : String(c.value);
    return side + M.colText(col) + ' ' + (op ? op.sym : c.op) + (rhs ? ' ' + rhs : '');
    function wrap(x) { return (x && (x.all || x.any)) ? '（' + M.condText(x) + '）' : M.condText(x); }
  };

  /** 检查条件，push 错误到 errs；headers=主表表头，refHeaders=参照表表头（ref_col 用） */
  function checkCond(c, errs, headers, refHeaders, allowRef, path) {
    if (c === undefined || c === null) return;
    if (typeof c !== 'object' || Array.isArray(c)) { errs.push('条件写法不正确'); return; }
    for (const k of ['all', 'any']) {
      if (k in c) {
        if (!Array.isArray(c[k]) || !c[k].length) errs.push('条件组是空的，请添加条件或删除这个组');
        else c[k].forEach(x => checkCond(x, errs, headers, refHeaders, allowRef));
        return;
      }
    }
    if ('not' in c) { checkCond(c.not, errs, headers, refHeaders, allowRef); return; }
    if (c.col === undefined && c.ref_col === undefined) { errs.push('条件缺少要判断的列'); return; }
    if (c.ref_col !== undefined && !allowRef) errs.push('这里的条件不能使用参照表的列');
    const op = M.condOp(c.op);
    if (!op) { errs.push('条件的判断方式「' + c.op + '」不支持'); return; }
    if (!op.noValue && c.value === undefined && c.col2 === undefined && c.ref_col2 === undefined) errs.push('条件「' + M.condText(c) + '」缺少比较的值');
    if (op.list && c.value !== undefined && !Array.isArray(c.value)) errs.push('“' + op.label + '”的值应是多个值的列表');
    if (op.v === 'regex' && typeof c.value === 'string') { try { new RegExp(c.value); } catch (e) { errs.push('正则表达式写法不正确：' + c.value); } }
    const chk = (hdr, ref, who) => {
      if (ref === undefined) return;
      const r = M.resolveCol(hdr, ref);
      if (r.error) errs.push(who + r.error);
    };
    chk(headers, c.col, '条件中');
    chk(headers, c.col2, '条件中');
    if (refHeaders !== undefined) { chk(refHeaders, c.ref_col, '条件中参照表'); chk(refHeaders, c.ref_col2, '条件中参照表'); }
  }

  /* ------------------------------------------------------------------ 校验 */
  M.allowedKeys = function (def) {
    const set = new Set(M.COMMON);
    def.fields.forEach(f => (f.keys || [f.key]).forEach(k => set.add(k)));
    (def.hidden || []).forEach(k => set.add(k));
    return set;
  };

  M.validate = function (cfg) {
    cfg = cfg || M.state.cfg;
    const res = { global: [], files: {}, steps: [], errors: 0, warns: 0 };
    const add = (arr, level, msg) => { arr.push({ level, msg }); if (level === 'error') res.errors++; else res.warns++; };
    const sim = M.simulate(cfg);
    res.sim = sim;
    const mainId = M.mainId(cfg);
    const ids = new Set();

    // 文件
    const mains = cfg.files.filter(f => f.role === 'main');
    if (!cfg.files.length) add(res.global, 'error', '还没有文件：请载入样例文件，或添加文件');
    else if (mains.length !== 1) add(res.global, 'error', '必须有且只有一个主表（现在有 ' + mains.length + ' 个）');
    cfg.files.forEach(f => {
      const arr = res.files[f.id] = [];
      if (!/^[A-Za-z0-9_\-一-鿿]+$/.test(String(f.id))) add(arr, 'error', '文件代号只能包含字母、数字、下划线或中文');
      if (ids.has(f.id)) add(arr, 'error', '文件代号「' + f.id + '」重复');
      ids.add(f.id);
      if (!['main', 'ref'].includes(f.role)) add(arr, 'error', '角色只能是主表或参照表');
      if (f.header_row !== undefined && !(Number.isInteger(f.header_row) && f.header_row >= 1)) add(arr, 'error', '表头行应是 ≥1 的整数');
      const smp = cfgSample(f.id);
      if (!f.match) add(arr, 'warn', '没有文件名规则：运行时会提示把文件拖进窗口');
      else if (smp && !F.matchPattern(f.match, smp.fileName)) add(arr, 'warn', '样例文件名「' + smp.fileName + '」不符合文件名规则，运行时可能找不到');
      if (f.match) {
        cfg.files.forEach(o => {
          const os = cfgSample(o.id);
          if (o !== f && os && F.matchPattern(f.match, os.fileName)) add(arr, 'warn', '规则也能匹配到「' + (o.label || o.id) + '」的文件（' + os.fileName + '），两个文件同时放在 input 里时会让你选择');
        });
      }
      if (smp && f.sheet !== undefined && f.sheet !== '' && !M.sampleSheet(f.id, f.sheet)) add(arr, 'error', '样例文件里没有工作表「' + f.sheet + '」');
    });

    // 步骤
    if (!cfg.steps.length) add(res.global, 'error', '还没有处理步骤');
    const matchIdsSeen = new Set();
    cfg.steps.forEach((s, i) => {
      const arr = [];
      res.steps.push(arr);
      const def = M.OPS[s && s.op];
      if (!def) { add(arr, 'error', '操作「' + (s && s.op) + '」不存在'); return; }
      const allowed = M.allowedKeys(def);
      Object.keys(s).forEach(k => { if (!allowed.has(k)) add(arr, 'error', '有执行器不认识的参数：' + k); });
      def.fields.forEach(fd => {
        if (!fd.required) return;
        const keys = fd.requiredKeys || fd.keys || [fd.key];
        const has = keys.some(k => k in s && (k === 'value' || (s[k] !== undefined && s[k] !== null && !(typeof s[k] === 'string' && s[k].trim() === '') && !(Array.isArray(s[k]) && !s[k].length))));
        if (!has) add(arr, 'error', '请填写「' + fd.label + '」');
      });
      if (s.enabled === false) return;
      const tgt = M.stepTarget(s);
      if (s.file !== undefined && !ids.has(s.file)) add(arr, 'error', '文件「' + s.file + '」不在文件列表里');
      if (def.writes !== false && tgt.fileId !== mainId && s.file !== undefined) add(arr, 'error', '只有主表可以被修改（参照表只读）');
      if (s.header_row !== undefined && !(Number.isInteger(s.header_row) && s.header_row >= 1)) add(arr, 'error', '表头行应是 ≥1 的整数');
      if (s.sheet !== undefined && s.sheet !== '' && cfgSample(tgt.fileId) && !M.sampleSheet(tgt.fileId, s.sheet) && !sim.extraSheets.includes(s.sheet)) {
        add(arr, 'error', '「' + M.fileLabel(tgt.fileId) + '」的样例里没有工作表「' + s.sheet + '」');
      }
      const headers = sim.headersAt(i, tgt.fileId, s.sheet, s.header_row);
      const errs = [];
      const warns = [];
      checkFields(s, def, i, headers, sim, errs, warns, matchIdsSeen, cfg);
      errs.forEach(m => add(arr, 'error', m));
      warns.forEach(m => add(arr, 'warn', m));
      if (s.op === 'match' && s.id) matchIdsSeen.add(s.id);
    });
    freshFormulaPass(cfg, res, add);
    const recalc = (cfg.options || {}).recalc || 'auto';
    if (!['auto', 'off', 'required'].includes(recalc)) add(res.global, 'error', '重算设置只能是 auto / off / required');
    if (!String(cfg.name || '').trim()) add(res.global, 'warn', '还没有给工作流起名字');
    return res;
  };

  function cfgSample(id) { return M.state.samples[id] || null; }

  /** 执行器不计算公式：本次运行中新写入公式的列，在后面的条件、排序、去重、匹配、校验里都还没有值（当作空值）。提前提醒。 */
  function freshFormulaPass(cfg, res, add) {
    const fresh = new Map();   // 列名 → 写入公式的步骤号
    const mainId = M.mainId(cfg);
    const names = v => (v === undefined || v === null ? [] : (Array.isArray(v) ? v : [v])).filter(x => typeof x === 'string' && !x.includes(':')).map(F.normHeader);
    const condCols = c => {
      if (!c || typeof c !== 'object') return [];
      if (Array.isArray(c.all)) return c.all.flatMap(condCols);
      if (Array.isArray(c.any)) return c.any.flatMap(condCols);
      if (c.not) return condCols(c.not);
      return names(c.col).concat(names(c.col2));
    };
    cfg.steps.forEach((s, i) => {
      if (!s || s.enabled === false || (s.file && s.file !== mainId)) return;
      const reads = [];
      const push = (cols, what) => cols.forEach(n => reads.push([n, what]));
      push(condCols(s.where), '条件');
      if (s.op === 'sort_rows') push(names((Array.isArray(s.by) ? s.by : [s.by]).map(b => (b && typeof b === 'object') ? b.column : b)), '排序');
      if (s.op === 'dedupe_rows') push(names(s.keys), '去重');
      if (s.op === 'match') push(names(s.main_keys), '匹配键');
      if (s.op === 'compare' && Array.isArray(s.compare_columns)) push(names(s.compare_columns.map(p => (p && typeof p === 'object') ? p.main : p)), '比对');
      if (s.from_col !== undefined) push(names(s.from_col), '取值');
      if (s.op === 'validate') (s.checks || []).forEach(ck => { push(names(ck.columns), '校验'); push(condCols(ck.where), '校验条件'); });
      const seen = new Set();
      reads.forEach(([n, what]) => {
        if (!fresh.has(n) || seen.has(n)) return;
        seen.add(n);
        add(res.steps[i], 'warn', what + '用到的「' + n + '」是第 ' + fresh.get(n) + ' 步新写入的公式。执行器不计算公式，执行时它还没有结果，会被当作空值。'
          + '建议改用参与计算的原始列来判断（例如不判断“差额<0”，而是比较“单价”和“标准单价”两列）。');
      });
      // 记录本步新写入公式的列
      const hasFormula = s.formula !== undefined || s.formula_from !== undefined;
      if (s.op === 'insert_column' && hasFormula) fresh.set(F.normHeader(s.header), i + 1);
      if ((s.op === 'fill' || s.op === 'write_if') && hasFormula) names(s.column).forEach(n => fresh.set(n, i + 1));
      if (s.op === 'fill_down') names(s.columns).forEach(n => fresh.set(n, i + 1));
      if (s.op === 'copy_matched' && s.mode === 'formula') M.copyPairs(s).forEach(p => names(p.to).forEach(n => fresh.set(n, i + 1)));
      if (s.op === 'rename_header') {
        const from = names(s.column)[0];
        if (from && fresh.has(from)) { fresh.set(F.normHeader(s.to), fresh.get(from)); fresh.delete(from); }
      }
      if (s.op === 'delete_columns') names(s.columns).forEach(n => fresh.delete(n));
    });
  }

  function checkFields(s, def, i, headers, sim, errs, warns, matchIdsSeen, cfg) {
    const colErr = (ref, who) => { if (ref === undefined || ref === null || ref === '') return; const r = M.resolveCol(headers, ref); if (r.error) errs.push((who || '') + r.error); };
    const colsErr = (spec, who) => { if (spec === undefined) return; const r = M.resolveCols(headers, spec); r.errors.forEach(e => errs.push((who || '') + e)); };
    const matchInfo = s.match !== undefined ? M.refOfMatch(i, s.match, sim) : null;
    def.fields.forEach(fd => {
      const v = s[fd.key];
      switch (fd.type) {
        case 'column': colErr(v, fd.label + '：'); break;
        case 'columns':
          if (v === undefined || (Array.isArray(v) && !v.length)) break;
          if (fd.source === 'ref') {
            const refH = sim.headersAt(i, s.ref, s.ref_sheet, s.ref_header_row);
            const r = M.resolveCols(refH, v); r.errors.forEach(e => errs.push(fd.label + '：' + e));
            const mk = Array.isArray(s.main_keys) ? s.main_keys.length : 1;
            const rk = Array.isArray(v) ? v.length : 1;
            if (rk !== mk) errs.push('参照表键列有 ' + rk + ' 列，主表键列有 ' + mk + ' 列，数量要一致');
          } else colsErr(v, fd.label + '：');
          if (s.op === 'match' && fd.key === 'main_keys' && s.ref_keys === undefined && s.ref) {
            const refH = sim.headersAt(i, s.ref, s.ref_sheet, s.ref_header_row);
            if (refH) (Array.isArray(v) ? v : [v]).forEach(k => { if (M.resolveCol(refH, k).error) errs.push('参照表里没有同名的键列「' + k + '」，请在“参照表对应的键列”里指定'); });
          }
          break;
        case 'condition': {
          if (v === undefined) break;
          let refH;
          if (fd.source === 'ref') { checkCond(v, errs, sim.headersAt(i, s.ref, s.ref_sheet, s.ref_header_row), undefined, false); break; }
          if (fd.allowRef && matchInfo) refH = matchInfo.headers;
          checkCond(v, errs, headers, fd.allowRef ? (refH === undefined ? null : refH) : undefined, !!fd.allowRef);
          break;
        }
        case 'content': {
          if (s.formula !== undefined) {
            const f = String(s.formula).trim();
            if (!f || f === '=') errs.push('请填写公式');
            const re = /\[@?([^\[\]]+)\]/g; let m;
            while ((m = re.exec(f))) colErr(m[1].trim(), '公式中：');
            const open = (f.match(/\(/g) || []).length, close = (f.match(/\)/g) || []).length;
            if (open !== close) errs.push('公式的括号不配对');
          }
          if (s.from_col !== undefined) colErr(s.from_col, '取值列：');
          if (s.formula_from !== undefined) {
            const ff = s.formula_from || {};
            if (!M.cellAddrOk(ff.cell)) errs.push('复制公式的来源单元格地址不正确（如 P2）');
            if (ff.file !== undefined && !cfg.files.some(f => f.id === ff.file)) errs.push('复制公式的来源文件「' + ff.file + '」不在文件列表里');
          }
          const n = ['value', 'formula', 'formula_from', 'from_col'].filter(k => s[k] !== undefined).length;
          if (n > 1) errs.push('写入内容只能选一种');
          break;
        }
        case 'position': {
          const n = ['before', 'after', 'at', 'end'].filter(k => s[k] !== undefined).length;
          if (n > 1) errs.push('插入位置只能选一种');
          if (s.before !== undefined) colErr(s.before, '插入位置：');
          if (s.after !== undefined) colErr(s.after, '插入位置：');
          if (s.at !== undefined && !(typeof s.at === 'number' ? s.at >= 1 : M.isLetters(s.at))) errs.push('插入位置的列字母不正确');
          break;
        }
        case 'rowMode': {
          const n = ['where', 'rows', 'blank'].filter(k => s[k] !== undefined).length;
          if (n > 1) errs.push('删除方式只能选一种');
          if (s.where !== undefined) checkCond(s.where, errs, headers);
          if (s.rows !== undefined && !/^\s*\d+\s*:\s*\d+\s*$/.test(String(s.rows))) errs.push('行号范围应写成“开始:结束”，如 5:20');
          break;
        }
        case 'rowsAt':
          if (s.before_row !== undefined && !(Number.isInteger(s.before_row) && s.before_row >= 1)) errs.push('插入位置的行号应是 ≥1 的整数');
          break;
        case 'rowRange':
          if (v !== undefined && !/^\s*\d+\s*:\s*\d+\s*$/.test(String(v))) errs.push(fd.label + '应写成“开始:结束”，如 5:20');
          break;
        case 'number':
          if (v !== undefined && (typeof v !== 'number' || !isFinite(v))) errs.push(fd.label + '应是数字');
          else if (v !== undefined && fd.min !== undefined && v < fd.min) errs.push(fd.label + '不能小于 ' + fd.min);
          break;
        case 'color':
          if (v !== undefined && v !== null && v !== '' && !M.parseColor(v)) errs.push(fd.label + '「' + v + '」不是合法的颜色');
          break;
        case 'numfmt':
          if (v !== undefined && !String(v).trim()) errs.push('请填写' + fd.label);
          break;
        case 'enum':
          if (v !== undefined && !fd.options.some(o => o.value === v)) {
            if (fd.allowColumn && typeof v === 'string') colErr(v, fd.label + '：');
            else errs.push(fd.label + '的取值「' + v + '」不正确');
          }
          break;
        case 'sortBy':
          if (v === undefined) break;
          (Array.isArray(v) ? v : [v]).forEach(b => {
            const o = (b && typeof b === 'object' && 'column' in b) ? b : { column: b };
            colErr(o.column, '排序列：');
            if (o.order !== undefined && !['asc', 'desc'].includes(o.order)) errs.push('排序方向只能是升序或降序');
          });
          break;
        case 'cells':
          if (v === undefined) break;
          if (typeof v !== 'object' || Array.isArray(v) || !Object.keys(v).length) { errs.push('请至少填写一个单元格'); break; }
          Object.keys(v).forEach(a => { if (!M.cellAddrOk(a)) errs.push('单元格地址「' + a + '」不正确（如 B2）'); });
          break;
        case 'matchId':
          if (v !== undefined && matchIdsSeen.has(v)) warns.push('匹配名称「' + v + '」与前面的匹配重名，后面的步骤会使用这一个');
          break;
        case 'fileRef':
          if (v !== undefined && !cfg.files.some(f => f.id === v)) errs.push('参照表「' + v + '」不在文件列表里');
          break;
        case 'matchRef':
          if (v !== undefined && !M.matchesBefore(i, cfg).some(m => m.id === v)) errs.push('在这一步之前没有名为「' + v + '」的“建立匹配”步骤');
          break;
        case 'copyColumns': {
          const pairs = M.copyPairs(s);
          if (!pairs.length) break;
          const refH = matchInfo ? matchInfo.headers : null;
          pairs.forEach(p => {
            if (p.from === undefined || p.from === '') { errs.push('复制列：有一行没有选择参照表的列'); return; }
            const r = M.resolveCol(refH, p.from); if (r.error) errs.push('参照表' + r.error);
            if (!s.create_missing) { const t = M.resolveCol(headers, p.to); if (t.error) errs.push('主表' + t.error + '（可勾选“主表没有目标列时自动加在最后”）'); }
          });
          break;
        }
        case 'comparePairs': {
          if (v === undefined) break;
          if (!Array.isArray(v) || !v.length) { errs.push('请至少选择一组要比较的列'); break; }
          const refH = matchInfo ? matchInfo.headers : null;
          v.forEach(p => {
            const a = (p && typeof p === 'object') ? p.main : p, b = (p && typeof p === 'object') ? (p.ref === undefined ? p.main : p.ref) : p;
            colErr(a, '主表：');
            const r = M.resolveCol(refH, b); if (r.error) errs.push('参照表' + r.error);
          });
          break;
        }
        case 'rangeSource': case 'sheetSource': {
          if (v === undefined) break;
          if (typeof v !== 'object') { errs.push(fd.label + '写法不正确'); break; }
          const fid = v.file || M.mainId(cfg);
          if (!cfg.files.some(f => f.id === fid)) errs.push('来源文件「' + v.file + '」不在文件列表里');
          else if (cfgSample(fid) && v.sheet !== undefined && v.sheet !== '' && !M.sampleSheet(fid, v.sheet)) errs.push('「' + M.fileLabel(fid) + '」的样例里没有工作表「' + v.sheet + '」');
          if (fd.type === 'rangeSource' && !M.rangeOk(v.range)) errs.push('来源区域应写成“左上:右下”，如 AF1:AI30');
          break;
        }
        case 'cellTarget': {
          if (v === undefined) break;
          const cell = typeof v === 'string' ? v : (v || {}).cell;
          if (!M.cellAddrOk(cell)) errs.push('粘贴位置应是单元格地址，如 AF1');
          break;
        }
        case 'contentParts':
          if (v !== undefined && v !== 'all') {
            const arr = Array.isArray(v) ? v : [v];
            if (!arr.length) errs.push('请至少选择一种要复制的内容');
            arr.forEach(x => { if (!['values', 'formulas', 'styles', 'widths', 'merges'].includes(x)) errs.push('复制内容「' + x + '」不认识'); });
          }
          break;
        case 'notFound':
          if (v !== undefined && v !== null) { if (!v.column) errs.push('未匹配标记：请选择写到哪一列'); else colErr(v.column, '未匹配标记：'); }
          break;
        case 'checks':
          if (v === undefined) break;
          if (!Array.isArray(v) || !v.length) { errs.push('请至少添加一项检查'); break; }
          v.forEach((ck, k) => {
            const no = '第 ' + (k + 1) + ' 项检查：';
            const ckFile = s.file || M.mainId(cfg);
            const ckHeaders = sim.headersAt(i, ckFile, s.sheet, s.header_row);
            switch (ck.type) {
              case 'required_columns': if (!(ck.columns || []).length) errs.push(no + '请选择列'); break;
              case 'not_empty': case 'unique':
                if (!(ck.columns || []).length) errs.push(no + '请选择列');
                (ck.columns || []).forEach(c => { const r = M.resolveCol(ckHeaders, c); if (r.error) errs.push(no + r.error); });
                break;
              case 'row_count':
                if (ck.min !== undefined && typeof ck.min !== 'number') errs.push(no + '最少行数应是数字');
                if (ck.max !== undefined && typeof ck.max !== 'number') errs.push(no + '最多行数应是数字');
                break;
              case 'headers_match':
                if (!cfg.files.some(f => f.id === (ck.file || 'ref'))) errs.push(no + '对照文件「' + (ck.file || 'ref') + '」不在文件列表里');
                if (ck.columns !== undefined && !/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(String(ck.columns))) errs.push(no + '列范围应写成列字母区间，如 A:AB');
                break;
              case 'condition': {
                if (!ck.where) errs.push(no + '请设置条件');
                const ce = []; checkCond(ck.where, ce, ckHeaders); ce.forEach(e => errs.push(no + e));
                if (ck.expect !== undefined && !['none', 'all', 'any'].includes(ck.expect)) errs.push(no + '期望结果不正确');
                break;
              }
              default: errs.push(no + '检查类型「' + ck.type + '」不支持');
            }
          });
          break;
        default:
      }
    });
    if (!headers && s.file === undefined && M.state.cfg.files.length) {
      const fid = M.stepTarget(s).fileId;
      if (fid && !cfgSample(fid)) warns.push('「' + M.fileLabel(fid) + '」没有载入样例文件，无法检查列名是否存在');
    }
  }

  /* ------------------------------------------------------------------ 步骤摘要 */
  M.summary = function (s) {
    const def = M.OPS[s.op];
    if (!def) return '未知操作 ' + s.op;
    const cond = s.where ? '（' + M.condText(s.where) + ' 的行）' : '';
    const content = () => {
      if (s.value !== undefined) return s.value === null ? '清空' : '“' + s.value + '”';
      if (s.formula !== undefined) return '公式 ' + s.formula;
      if (s.from_col !== undefined) return '「' + M.colText(s.from_col) + '」列的值';
      if (s.formula_from !== undefined) return '复制 ' + ((s.formula_from || {}).cell || '?') + ' 的公式';
      return '';
    };
    switch (s.op) {
      case 'delete_columns': return '删除：' + (M.colText(s.columns) || '（未选择列）');
      case 'insert_column': {
        let pos = '最后';
        if (s.before !== undefined) pos = '「' + M.colText(s.before) + '」之前';
        else if (s.after !== undefined) pos = '「' + M.colText(s.after) + '」之后';
        else if (s.at !== undefined) pos = s.at + ' 列';
        const c = content();
        return '在' + pos + '插入「' + (s.header || '?') + '」' + (c ? '，写入 ' + c : '') + cond;
      }
      case 'rename_header': return '「' + M.colText(s.column) + '」改为「' + (s.to || '?') + '」';
      case 'set_cells': return Object.keys(s.cells || {}).map(k => k + ' = ' + s.cells[k]).join('；') || '（未填写）';
      case 'number_format': return M.colText(s.columns) + ' 设为 ' + (s.format || '?') + cond;
      case 'set_width': return M.colText(s.columns) + ' 列宽 ' + (s.width === undefined ? '?' : s.width);
      case 'delete_rows':
        if (s.blank) return '删除整行为空的行';
        if (s.rows) return '删除第 ' + s.rows + ' 行';
        return '删除 ' + (s.where ? M.condText(s.where) : '?') + ' 的行';
      case 'keep_rows': return '只保留 ' + (s.where ? M.condText(s.where) : '?') + ' 的行';
      case 'insert_rows': return (s.at_end ? '在最后' : '在第 ' + (s.before_row || '?') + ' 行之前') + '插入 ' + (s.count || 1) + ' 行';
      case 'dedupe_rows': return '按 ' + (s.keys ? M.colText(s.keys) : '整行') + ' 去重，保留' + (s.keep === 'last' ? '最后' : '第一') + '次出现';
      case 'sort_rows': return '按 ' + (Array.isArray(s.by) ? s.by : [s.by]).filter(x => x !== undefined).map(b => (b && typeof b === 'object' && 'column' in b) ? M.colText(b.column) + (b.order === 'desc' ? '降序' : '升序') : M.colText(b) + '升序').join('、');
      case 'fill': case 'write_if': return '「' + M.colText(s.column) + '」写入 ' + (content() || '?') + cond + (s.only_empty ? '，只写空单元格' : '');
      case 'fill_down': return M.colText(s.columns) + ' 向下填充' + (s.from_row ? '（从第 ' + s.from_row + ' 行）' : '');
      case 'replace_text': return M.colText(s.columns) + (s.replace ? ' 中“' + (s.find || '') + '”替换为“' + s.replace + '”' : ' 中删除“' + (s.find || '') + '”') + cond;
      case 'format': case 'color': {
        const parts = [];
        if (s.fill) parts.push('底色 ' + s.fill);
        if (s.font_color) parts.push('字色 ' + s.font_color);
        if (s.bold === true) parts.push('加粗');
        if (s.bold === false) parts.push('取消加粗');
        if (s.number_format) parts.push('格式 ' + s.number_format);
        if (s.clear_fill) parts.push('清除底色');
        return M.colText(s.columns) + ' ' + (parts.join('、') || '（未设置格式）') + cond;
      }
      case 'match': return '按 ' + M.colText(s.main_keys) + ' 匹配「' + M.fileLabel(s.ref) + '」' + (s.ref_keys ? '（对应 ' + M.colText(s.ref_keys) + '）' : '') + '，命名为 ' + (s.id || '?');
      case 'copy_matched': {
        const pairs = M.copyPairs(s).map(p => p.from === p.to ? String(p.from) : p.from + '→' + p.to).join('、');
        return '用匹配 ' + (s.match || '?') + ' 复制' + (s.mode === 'formula' ? '公式' : '值') + '：' + (pairs || '?') + cond + (s.fill ? '，并着色' : '');
      }
      case 'compare': return '用匹配 ' + (s.match || '?') + ' 比对' + (Array.isArray(s.compare_columns) ? '：' + s.compare_columns.map(p => (p && typeof p === 'object') ? p.main : p).join('、') : '所有同名列') + (s.tolerance ? '，容差 ' + s.tolerance : '');
      case 'copy_range': {
        const f = s.from || {}, t = s.to || {};
        return '「' + M.fileLabel(f.file || M.mainId()) + '」' + (f.sheet ? f.sheet + '!' : '') + (f.range || '?') + ' → 主表 ' + ((typeof t === 'string' ? '' : t.sheet ? t.sheet + '!' : '')) + ((typeof t === 'string' ? t : t.cell) || '?');
      }
      case 'copy_sheet': {
        const f = s.from || {};
        return '把「' + M.fileLabel(f.file || '?') + '」的 ' + (f.sheet || '第一张表') + ' 复制进主表' + (s.name ? '，命名为 ' + s.name : '');
      }
      case 'validate': return (s.checks || []).map(ck => checkLabel(ck)).join('；') || '（未添加检查）';
      default: return def.label;
    }
  };

  const CHECK_TYPES = {
    required_columns: '必须有这些列', not_empty: '不能为空', unique: '不能重复', row_count: '数据行数范围',
    headers_match: '表头与另一个文件一致', condition: '自定义条件',
  };
  M.CHECK_TYPES = CHECK_TYPES;
  function checkLabel(ck) {
    const cols = M.colText(ck.columns);
    switch (ck.type) {
      case 'required_columns': return '必须有：' + cols;
      case 'not_empty': return cols + ' 不能为空';
      case 'unique': return cols + ' 不能重复';
      case 'row_count': return '行数 ' + (ck.min === undefined ? 0 : ck.min) + '～' + (ck.max === undefined ? '不限' : ck.max);
      case 'headers_match': return '表头与「' + M.fileLabel(ck.file || 'ref') + '」一致';
      case 'condition': return (ck.expect === 'all' ? '所有行都应 ' : ck.expect === 'any' ? '至少一行 ' : '不应有 ') + M.condText(ck.where) + (ck.expect === 'all' || ck.expect === 'any' ? '' : ' 的行');
      default: return ck.type;
    }
  }

  /* ------------------------------------------------------------------ 导入导出 */
  /** 导出：去掉空值、按固定顺序排列键 */
  M.exportConfig = function (cfg) {
    cfg = cfg || M.state.cfg;
    const out = { schema: M.SCHEMA, name: String(cfg.name || '').trim() || '未命名工作流' };
    if (String(cfg.description || '').trim()) out.description = cfg.description.trim();
    out.files = cfg.files.map(f => {
      const o = { id: f.id, role: f.role };
      if (f.label) o.label = f.label;
      if (f.match) o.match = f.match;
      if (f.sheet !== undefined && f.sheet !== '') o.sheet = f.sheet;
      if (f.header_row !== undefined && f.header_row !== 1) o.header_row = f.header_row;
      Object.keys(f).forEach(k => { if (!(k in o) && !['id', 'role', 'label', 'match', 'sheet', 'header_row'].includes(k)) o[k] = f[k]; });
      return o;
    });
    const fname = ((cfg.output || {}).file_name || '').trim();
    if (fname) out.output = Object.assign({}, cfg.output, { file_name: fname });
    out.options = Object.assign({ recalc: 'auto' }, cfg.options || {});
    out.steps = cfg.steps.map(s => {
      const o = { op: s.op };
      if (s.title) o.title = s.title;
      Object.keys(s).forEach(k => { if (k !== 'op' && k !== 'title') o[k] = s[k]; });
      return o;
    });
    Object.keys(cfg).forEach(k => { if (!(k in out) && !['schema', 'name', 'description', 'files', 'output', 'options', 'steps'].includes(k) && !k.startsWith('_')) out[k] = cfg[k]; });
    return out;
  };

  M.toJson = cfg => JSON.stringify(M.exportConfig(cfg), null, 2);

  /** 导入：返回规范化后的配置；格式明显不对时抛错 */
  M.importConfig = function (obj) {
    if (!obj || typeof obj !== 'object' || Array.isArray(obj)) throw new Error('文件内容不是工作流配置');
    const schema = obj.schema || M.SCHEMA;
    if (!String(schema).startsWith('excel-workflow/')) throw new Error('这不是 Excel 工作流配置文件（schema=' + schema + '）');
    if (schema !== M.SCHEMA) throw new Error('配置版本 ' + schema + ' 比本配置器新，请使用新版配置器');
    const cfg = JSON.parse(JSON.stringify(obj));
    cfg.schema = M.SCHEMA;
    cfg.name = cfg.name || '';
    cfg.description = cfg.description || '';
    cfg.files = Array.isArray(cfg.files) ? cfg.files.filter(f => f && typeof f === 'object') : [];
    cfg.files.forEach(f => { if (!f.role) f.role = 'ref'; });
    cfg.options = Object.assign({ recalc: 'auto' }, cfg.options || {});
    cfg.steps = Array.isArray(cfg.steps) ? cfg.steps.filter(s => s && typeof s === 'object') : [];
    delete cfg._path;
    return cfg;
  };

  EFA.WFModel = M;
})(window.EFA = window.EFA || {});
