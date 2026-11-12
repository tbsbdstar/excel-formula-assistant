/* 模板引擎：把中文需求解析成“计算计划”（不需要 AI）
 *
 * 流程：规范化 → 识别列名/取值/数字/日期等实体（替换成 ⟨C0⟩ ⟨V1⟩ ⟨N2⟩ 这类占位）
 *       → 抽取条件、汇总方式、分组依据 → 按优先级逐个模板尝试 → 输出计划
 */
(function (EFA) {
  'use strict';
  const U = EFA.U;

  /* ======================= 词表 ======================= */

  const OP_GROUPS = [
    ['>=', '大于等于|大于或等于|大于或者等于|不低于|不少于|不小于|至少|>='],
    ['<=', '小于等于|小于或等于|小于或者等于|不高于|不超过|不大于|不多于|至多|<='],
    ['<>', '不等于|不是|不为|不属于|<>|!='],
    ['>', '大于|超过|高于|多于|超出|>'],
    ['<', '小于|低于|少于|不足|不到|未达到|未满|<'],
    ['=', '等于|==|=|是|为|属于'],
  ];
  const OP_RE = OP_GROUPS.map(g => g[1]).join('|');
  function opOf(word) {
    for (const [op, alts] of OP_GROUPS) if (alts.split('|').includes(word)) return op;
    return '=';
  }
  const UNIT = '(?:万元|元|分|个|件|岁|天|次|人|台|箱|公斤|斤|kg|米|小时|单|笔)?';
  const OP_SYMBOL = { '=': '=', '<>': '≠', '>': '>', '>=': '≥', '<': '<', '<=': '≤' };

  const AGG_DEFS = [
    ['distinct', /(不重复|去重|唯一)的?(值的)?(个数|数量|数目|计数|有多少|多少|几)|有多少(种|类|个不同)|多少个不同/g],
    ['avg', /平均|均值/g],
    ['median', /中位数/g],
    ['max', /最大|最高|最多|峰值|最贵|最晚/g],
    ['min', /最小|最低|最少|最便宜|最早/g],
    ['sum', /求和|总和|合计|总计|汇总|总额|加总|累加|加起来|之和|的和|总金额|总量|总销量|总共|一共|总数|总值/g],
    ['count', /订单数|客户数|员工数|学生数|人次|计数|个数|多少个|多少条|多少人|多少笔|多少单|多少次|多少行|多少家|多少名|多少位|多少项|几个|几条|几人|几笔|几单|几次|几行|几家|几名|几位|人数|条数|笔数|单数|次数|行数|家数|记录数|数量/g],
    ['countWeak', /有多少|有几/g],
  ];
  const AGG_NAMES = { sum: '合计', avg: '平均值', max: '最大值', min: '最小值', count: '数量', median: '中位数', distinct: '不重复个数' };

  /* ======================= 规范化与实体识别 ======================= */

  function normalize(text) {
    return U.halfWidth(String(text || ''))
      .replace(/[⟨⟩\u0001]/g, '')
      .replace(/≥/g, '>=').replace(/≤/g, '<=').replace(/≠/g, '<>')
      .replace(/×/g, '*').replace(/÷/g, '/')
      .replace(/。/g, ',')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function aliasesOf(col) {
    const out = new Set([col.name]);
    if (col.rawHeader) out.add(col.rawHeader);
    const stripped = col.name.replace(/[(（][^)）]*[)）]/g, '').trim();
    if (stripped) out.add(stripped);
    return Array.from(out).filter(a => a.length >= 2);
  }

  function parse(raw, wb, sheet) {
    const text = normalize(raw);
    const taken = new Uint8Array(text.length);
    const spans = [];
    const free = (s, e) => { for (let i = s; i < e; i++) if (taken[i]) return false; return true; };
    const add = (s, e, ent) => { for (let i = s; i < e; i++) taken[i] = 1; spans.push({ s, e, ent }); };
    const others = wb.sheets.filter(s => s !== sheet);

    const findCol = (name, sheets) => {
      const n = String(name).trim();
      for (const sh of sheets) {
        const c = sh.columns.find(c => c.name === n || c.rawHeader === n) ||
          sh.columns.find(c => aliasesOf(c).includes(n));
        if (c) return { sheet: sh, col: c };
      }
      return null;
    };
    const findValue = (val) => {
      for (const c of sheet.columns) {
        if (c.distinct && c.distinct.has(val)) return c;
      }
      return null;
    };

    // 1. 引号里的内容
    const QUOTE = /[“"「『‘'【《]([^“”"「」『』‘’'【】《》]{1,40})[”"」』’'】》]/g;
    let m;
    while ((m = QUOTE.exec(text))) {
      const inner = m[1].trim();
      if (!inner) continue;
      const hit = findCol(inner, [sheet].concat(others));
      const sh = wb.sheets.find(s => s.name === inner);
      let ent;
      if (hit) ent = { t: 'C', col: hit.col, sheet: hit.sheet };
      else if (sh) ent = { t: 'S', sheet: sh };
      else {
        const vc = findValue(inner);
        ent = vc ? { t: 'V', col: vc, sheet, value: inner } : { t: 'L', text: inner };
      }
      add(m.index, m.index + m[0].length, ent);
    }

    // 2. 工作表名
    for (const sh of others.slice().sort((a, b) => b.name.length - a.name.length)) {
      if (sh.name.length < 2) continue;
      let i = -1;
      while ((i = text.indexOf(sh.name, i + 1)) >= 0) {
        let e = i + sh.name.length;
        if (text[e] === '表' && !/表$/.test(sh.name)) e++;
        if (free(i, e)) add(i, e, { t: 'S', sheet: sh });
      }
    }

    // 3. 列名（当前表优先，其次其他表；长名优先）
    const cands = [];
    for (const sh of [sheet].concat(others)) {
      for (const c of sh.columns) for (const a of aliasesOf(c)) cands.push({ a, sheet: sh, col: c, pri: sh === sheet ? 0 : 1 });
    }
    // 列名的后缀简称：“工资”→“基本工资”（仅当前表、且只对应一列时）
    const sufMap = new Map();
    for (const c of sheet.columns) {
      const n = c.name.replace(/[(（][^)）]*[)）]/g, '');
      for (let k = 1; k <= n.length - 2; k++) {
        const suf = n.slice(k);
        if (!/^[一-龥]{2,}$/.test(suf)) continue;
        sufMap.set(suf, (sufMap.get(suf) || []).concat([c]));
      }
    }
    for (const [suf, cs] of sufMap) {
      if (cs.length !== 1 || sheet.columns.some(c => c.name === suf)) continue;
      cands.push({ a: suf, sheet, col: cs[0], pri: 2 });
    }
    cands.sort((x, y) => (y.a.length - x.a.length) || (x.pri - y.pri));
    for (const cd of cands) {
      let i = -1;
      while ((i = text.indexOf(cd.a, i + 1)) >= 0) {
        if (free(i, i + cd.a.length)) add(i, i + cd.a.length, { t: 'C', col: cd.col, sheet: cd.sheet });
      }
    }
    // 列字母：C列、列C
    const LETTER = /(^|[^A-Za-z])([A-Z]{1,2})\s*列|列\s*([A-Z]{1,2})(?![A-Za-z])/g;
    while ((m = LETTER.exec(text))) {
      const L = m[2] || m[3];
      const col = sheet.columns.find(c => c.letter === L);
      const s = m.index + (m[1] ? m[1].length : 0);
      const e = m.index + m[0].length;
      if (col && free(s, e)) add(s, e, { t: 'C', col, sheet });
    }

    // 4. 数据中的取值（如“销售部”），仅当前表的文本列
    const vals = [];
    for (const c of sheet.columns) {
      if (!c.distinct || c.type === 'number' || c.type === 'date') continue;
      for (const v of c.distinct.keys()) {
        if (typeof v !== 'string') continue;
        const s = v.trim();
        if (!s || s.length > 30 || /^-?\d+(\.\d+)?$/.test(s)) continue;
        if (s.length === 1 && !(c.distinctCount <= 10 && /[一-龥]/.test(s))) continue;
        vals.push({ v: s, col: c });
      }
    }
    vals.sort((a, b) => b.v.length - a.v.length);
    for (const vv of vals) {
      let i = -1;
      while ((i = text.indexOf(vv.v, i + 1)) >= 0) {
        if (free(i, i + vv.v.length)) add(i, i + vv.v.length, { t: 'V', col: vv.col, sheet, value: vv.v });
      }
    }

    const scanRe = (re, fn) => {
      re.lastIndex = 0;
      let mm;
      while ((mm = re.exec(text))) {
        if (!free(mm.index, mm.index + mm[0].length)) continue;
        const ent = fn(mm);
        if (ent) add(mm.index, mm.index + mm[0].length, ent);
      }
    };

    // 5. 日期
    scanRe(/(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*[日号]?/g, mm => {
      const y = +mm[1], mo = +mm[2], d = +mm[3];
      return (mo >= 1 && mo <= 12 && d >= 1 && d <= 31 && y > 1900) ? { t: 'D', y, m: mo, d, gran: 'day' } : null;
    });
    scanRe(/(\d{4})\s*(?:年|-|\/)\s*(\d{1,2})\s*月?(?:份)?(?!\d)/g, mm => {
      const y = +mm[1], mo = +mm[2];
      return (mo >= 1 && mo <= 12 && y > 1900) ? { t: 'D', y, m: mo, d: 1, gran: 'month' } : null;
    });
    scanRe(/(\d{4})\s*年(?:度)?(?!\s*\d)/g, mm => (+mm[1] > 1900 ? { t: 'D', y: +mm[1], m: 1, d: 1, gran: 'year' } : null));
    scanRe(/(\d{1,2}|[一二三四五六七八九十]{1,3})\s*月(?:份)?(?![\d])/g, mm => {
      if (mm.index > 0 && /[个第]/.test(text[mm.index - 1])) return null;
      const n = /\d/.test(mm[1]) ? +mm[1] : U.parseCnNumber(mm[1]);
      return n >= 1 && n <= 12 ? { t: 'M', m: n } : null;
    });

    // 6. 数字
    scanRe(/(?:(?<=^|[\s(=<>,:于为是])-)?\d+(?:,\d{3})*(?:\.\d+)?\s*(?:万|千|[kKwW](?![A-Za-z]))?\s*%?/g, mm => {
      let s = mm[0].replace(/\s/g, '');
      let mult = 1;
      if (/%$/.test(s)) { mult /= 100; s = s.slice(0, -1); }
      if (/万$|[wW]$/.test(s)) { mult *= 10000; s = s.slice(0, -1); }
      else if (/千$|[kK]$/.test(s)) { mult *= 1000; s = s.slice(0, -1); }
      const v = Number(s.replace(/,/g, '')) * mult;
      return isFinite(v) ? { t: 'N', v: Math.round(v * 1e10) / 1e10, raw: mm[0].trim() } : null;
    });
    scanRe(/[零一二两三四五六七八九十百千万]+(?=\s*(?:个|位|名|分|天|岁|年|月|倍|次|元|件|小数|列|行|字|号|周|星期))/g, mm => {
      const v = U.parseCnNumber(mm[0]);
      return v === null ? null : { t: 'N', v, raw: mm[0] };
    });

    // 生成签名串
    spans.sort((a, b) => a.s - b.s);
    const ents = [];
    let sig = '', pos = 0;
    const sigPos = [];
    for (const sp of spans) {
      sig += text.slice(pos, sp.s);
      const idx = ents.length;
      sp.ent.idx = idx;
      sp.ent.pos = sig.length;
      ents.push(sp.ent);
      sig += '⟨' + sp.ent.t + idx + '⟩';
      pos = sp.e;
      sigPos.push(sp.ent.pos);
    }
    sig += text.slice(pos);

    const q = { raw, text, sig, ents, wb, sheet, others };
    q.has = re => re.test(q.sig);
    q.colEnts = ents.filter(e => e.t === 'C');
    q.activeCols = q.colEnts.filter(e => e.sheet === sheet).map(e => e.col);
    return q;
  }

  /* ======================= 引用与常量写法 ======================= */

  function mkRef(q) {
    const base = q.planSheet || q.sheet;
    const nameOf = (col, sh) => (sh && sh !== base ? sh.name + '!' : '') + col.name;
    return {
      R: (col, sh) => '[' + nameOf(col, sh || colSheet(q, col)) + ']',
      C: col => '[@' + col.name + ']',
      X: col => '[^' + col.name + ']',
    };
  }
  function colSheet(q, col) {
    for (const sh of q.wb.sheets) if (sh.columns.includes(col)) return sh;
    return q.sheet;
  }

  function dateExpr(serial) {
    const p = U.serialToParts(serial);
    return 'DATE(' + p.y + ',' + p.m + ',' + p.d + ')';
  }
  function dateText(serial) { return U.formatSerialDate(serial); }

  /** 条件值在公式中的写法 */
  function litExpr(c, v) {
    if (c.isDate && typeof v === 'number') return dateExpr(v);
    if (typeof v === 'number') return String(v);
    return U.excelString(v);
  }

  /* ======================= 条件抽取 ======================= */

  function isTextCol(col) { return col && (col.type === 'text' || col.type === 'textnum' || col.type === 'mixed' || col.type === 'bool'); }

  function typedValue(col, ent, rawText) {
    if (ent && ent.t === 'N') return isTextCol(col) ? String(ent.raw.replace(/\s/g, '')) : ent.v;
    if (ent && (ent.t === 'V' || ent.t === 'L')) {
      const s = ent.t === 'V' ? ent.value : ent.text;
      if (!isTextCol(col) && /^-?\d+(\.\d+)?$/.test(s)) return Number(s);
      return s;
    }
    if (rawText != null) {
      if (!isTextCol(col) && /^-?\d+(\.\d+)?$/.test(rawText)) return Number(rawText);
      return rawText;
    }
    return null;
  }

  /** 日期实体 → [起, 止) 区间 */
  function dateRange(d) {
    const start = U.partsToSerial(d.y, d.m, d.d);
    let end;
    if (d.gran === 'day') end = start + 1;
    else if (d.gran === 'month') end = d.m === 12 ? U.partsToSerial(d.y + 1, 1, 1) : U.partsToSerial(d.y, d.m + 1, 1);
    else end = U.partsToSerial(d.y + 1, 1, 1);
    return { start, end };
  }

  function dateCond(col, op, d) {
    const { start, end } = dateRange(d);
    if (d.gran === 'day' && op !== '=' && op !== '<>') return { kind: 'cmp', col, op, value: start, isDate: true };
    switch (op) {
      case '>': return { kind: 'cmp', col, op: '>=', value: end, isDate: true };
      case '>=': return { kind: 'cmp', col, op: '>=', value: start, isDate: true };
      case '<': return { kind: 'cmp', col, op: '<', value: start, isDate: true };
      case '<=': return { kind: 'cmp', col, op: '<', value: end, isDate: true };
      case '<>': return d.gran === 'day' ? { kind: 'cmp', col, op: '<>', value: start, isDate: true } : null;
      default:
        return d.gran === 'day' ? { kind: 'cmp', col, op: '=', value: start, isDate: true }
          : { kind: 'range', col, lo: start, loOp: '>=', hi: end, hiOp: '<', isDate: true };
    }
  }

  function dateColOf(q, near) {
    if (near && near.type === 'date') return near;
    const mentioned = q.activeCols.find(c => c.type === 'date');
    if (mentioned) return mentioned;
    const all = q.sheet.columns.filter(c => c.type === 'date');
    return all.length === 1 ? all[0] : (all[0] || null);
  }

  /**
   * 从签名串（或其中一段）抽取条件
   * opts.defaultCol：只写了数字没写列名时默认针对的列
   */
  function extractConditions(q, sig, opts) {
    opts = opts || {};
    const conds = [];
    const consumed = new Set();
    let work = sig;
    const E = i => q.ents[+i];
    const blank = (mm) => { work = work.slice(0, mm.index) + ' '.repeat(mm[0].length) + work.slice(mm.index + mm[0].length); };
    const run = (re, fn) => {
      re.lastIndex = 0;
      let mm;
      const found = [];
      while ((mm = re.exec(work))) { found.push({ index: mm.index, 0: mm[0], m: mm }); if (mm[0] === '') re.lastIndex++; }
      for (const f of found) {
        const c = fn(f.m);
        if (c) { c.pos = f.index; conds.push(c); blank(f); }
      }
    };
    const colOr = (ci) => {
      if (ci !== undefined && ci !== null && E(ci) && E(ci).t === 'C') { consumed.add(+ci); return E(ci).col; }
      return opts.defaultCol || null;
    };

    // 区间：在 1000 到 2000 之间
    run(new RegExp('(?:⟨C(\\d+)⟩\\s*(?:的)?\\s*(?:值)?\\s*)?(?:在|介于|位于|处于)?\\s*⟨N(\\d+)⟩\\s*' + UNIT + '\\s*(?:到|至|-|~|和|与)\\s*⟨N(\\d+)⟩\\s*' + UNIT + '\\s*(?:之间|以内|范围内|区间内|区间)?', 'g'), mm => {
      const col = colOr(mm[1]);
      if (!col || isTextCol(col)) return null;
      consumed.add(+mm[2]); consumed.add(+mm[3]);
      const a = E(mm[2]).v, b = E(mm[3]).v;
      return { kind: 'range', col, lo: Math.min(a, b), loOp: '>=', hi: Math.max(a, b), hiOp: '<=' };
    });

    // 后缀：60 分以上 / 1 万以下
    run(new RegExp('(?:⟨C(\\d+)⟩\\s*(?:的)?\\s*(?:在)?\\s*)?⟨N(\\d+)⟩\\s*' + UNIT + '\\s*(及以上|或以上|以上|及以下|或以下|以下|以内|之内)', 'g'), mm => {
      const col = colOr(mm[1]);
      if (!col || isTextCol(col)) return null;
      consumed.add(+mm[2]);
      const w = mm[3];
      const op = /以上/.test(w) ? '>=' : /及以下|或以下|以内|之内/.test(w) ? '<=' : '<';
      return { kind: 'cmp', col, op, value: E(mm[2]).v };
    });

    // 及格 / 不及格（成绩类常用说法，按 60 分）
    run(/(?:⟨C(\d+)⟩\s*(?:成绩)?\s*(?:的)?\s*)?(?:是)?(不及格|及格|不合格|合格)(?=的|人|学生|同学|有|数|,|;|$|\s)/g, mm => {
      const col = colOr(mm[1]);
      if (!col || col.type !== 'number') return null;
      return { kind: 'cmp', col, op: /不/.test(mm[2]) ? '<' : '>=', value: 60 };
    });

    // 为空 / 不为空
    run(/⟨C(\d+)⟩\s*(?:的)?\s*(?:值)?\s*(不为空|不是空的?|非空|不空|已填写?|有值|有数据|有内容|为空|是空的?|空白|未填写?|没填写?|没有填写?|没有值|没有内容)/g, mm => {
      const col = colOr(mm[1]);
      return col ? { kind: 'blank', col, empty: !/不为空|不是空|非空|不空|已填|有值|有数据|有内容/.test(mm[2]) } : null;
    });

    // 包含 / 开头 / 结尾
    run(/⟨C(\d+)⟩\s*(?:中|里|里面)?\s*(不包含|不含|包含|含有|含|带有|带)\s*(?:⟨([VLN])(\d+)⟩|([^\s⟨⟩,;的]{1,20}?))(?=的|且|并且|和|或|,|;|$|\s|⟨|字)/g, mm => {
      const col = colOr(mm[1]);
      if (!col) return null;
      let v;
      if (mm[3]) { consumed.add(+mm[4]); const e = E(mm[4]); v = e.t === 'N' ? e.raw : (e.value || e.text); }
      else v = mm[5];
      if (!v) return null;
      return { kind: 'text', col, op: /不/.test(mm[2]) ? 'notcontains' : 'contains', value: v };
    });
    run(/⟨C(\d+)⟩\s*(?:是)?以\s*(?:⟨([VLN])(\d+)⟩|([^\s⟨⟩,;]{1,20}?))\s*(开头|打头|开始|结尾|结束)/g, mm => {
      const col = colOr(mm[1]);
      if (!col) return null;
      let v;
      if (mm[2]) { consumed.add(+mm[3]); const e = E(mm[3]); v = e.t === 'N' ? e.raw : (e.value || e.text); } else v = mm[4];
      return { kind: 'text', col, op: /开头|打头|开始/.test(mm[5]) ? 'starts' : 'ends', value: v };
    });

    // 列 + 比较 + 值
    run(new RegExp('(?:⟨C(\\d+)⟩\\s*(?:的)?\\s*(?:值)?\\s*)?(' + OP_RE + ')\\s*⟨([NDVLM])(\\d+)⟩\\s*' + UNIT + '\\s*(?:以上|以下)?', 'g'), mm => {
      const op = opOf(mm[2]);
      const ve = E(mm[4]);
      let col = (mm[1] !== undefined) ? E(mm[1]).col : null;
      if (!col && ve.t === 'V') col = ve.col;
      if (!col && (ve.t === 'D' || ve.t === 'M')) col = dateColOf(q, null);
      if (!col) col = opts.defaultCol || null;
      if (!col) return null;
      // “是/为”后面跟数字且列是数值列时才当条件；避免“为500”这种赋值被误判
      if (mm[1] === undefined && /^(是|为)$/.test(mm[2]) && !opts.allowBareEq) return null;
      if (mm[1] !== undefined) consumed.add(+mm[1]);
      consumed.add(+mm[4]);
      if (ve.t === 'D') return dateCond(col, op, ve);
      if (ve.t === 'M') return op === '=' ? { kind: 'monthNum', col, month: ve.m } : null;
      const v = typedValue(col, ve);
      if (v === null) return null;
      return { kind: 'cmp', col, op, value: v };
    });

    // 日期 + 之后/之前
    run(/(?:⟨C(\d+)⟩\s*(?:在)?\s*)?⟨D(\d+)⟩\s*(?:\(含\))?\s*(之后|以后|后|起|开始|及以后|之前|以前|前)/g, mm => {
      const col = mm[1] !== undefined ? E(mm[1]).col : dateColOf(q, null);
      if (!col) return null;
      if (mm[1] !== undefined) consumed.add(+mm[1]);
      consumed.add(+mm[2]);
      const d = E(mm[2]);
      const after = /后|起|开始/.test(mm[3]);
      const { start, end } = dateRange(d);
      if (after) return { kind: 'cmp', col, op: (/起|开始|及以后/.test(mm[3]) ? '>=' : (d.gran === 'day' ? '>' : '>=')), value: (/起|开始|及以后/.test(mm[3]) ? start : (d.gran === 'day' ? start : end)), isDate: true };
      return { kind: 'cmp', col, op: '<', value: start, isDate: true };
    });

    // 自由文本取值：部门是研发中心
    run(/⟨C(\d+)⟩\s*(?:的)?\s*(不等于|不是|不为|等于|是|为)\s*(?!多少|几|谁|哪|什么|啥|否|空)([^\s⟨⟩,;的且和或]{1,20}?)(?=的|且|并且|而且|同时|和|或|,|;|\s|$|⟨|时|则|就)/g, mm => {
      const col = E(mm[1]).col;
      if (!isTextCol(col)) return null;
      consumed.add(+mm[1]);
      return { kind: 'cmp', col, op: opOf(mm[2]), value: mm[3] };
    });

    // 单独出现的取值：“销售部的销售额”
    for (const e of q.ents) {
      if (e.t !== 'V' || consumed.has(e.idx)) continue;
      if (work.indexOf('⟨V' + e.idx + '⟩') < 0) continue;
      consumed.add(e.idx);
      conds.push({ kind: 'cmp', col: e.col, op: '=', value: e.value, pos: work.indexOf('⟨V' + e.idx + '⟩') });
    }
    // 单独出现的月份：“3月份的销售额”
    for (const e of q.ents) {
      if (e.t !== 'M' || consumed.has(e.idx)) continue;
      if (work.indexOf('⟨M' + e.idx + '⟩') < 0) continue;
      if (opts.skipLooseDates) continue;
      const col = dateColOf(q, null);
      if (!col) continue;
      consumed.add(e.idx);
      conds.push({ kind: 'monthNum', col, month: e.m, pos: work.indexOf('⟨M' + e.idx + '⟩') });
    }
    // 单独出现的日期：“2024年3月的销售额”
    for (const e of q.ents) {
      if (e.t !== 'D' || consumed.has(e.idx)) continue;
      if (work.indexOf('⟨D' + e.idx + '⟩') < 0) continue;
      if (opts.skipLooseDates) continue;
      const col = dateColOf(q, null);
      if (!col) continue;
      consumed.add(e.idx);
      const c = dateCond(col, '=', e);
      if (c) { c.pos = work.indexOf('⟨D' + e.idx + '⟩'); conds.push(c); }
    }

    conds.sort((a, b) => a.pos - b.pos);
    // 同一列的多个“等于”合并为“或”
    const merged = [];
    const orMode = /或|或者|或是/.test(sig.replace(/(大于|小于)或(者)?等于|或以上|或以下/g, ''));
    for (const c of conds) {
      if (c.kind === 'cmp' && c.op === '=') {
        const prev = merged.find(p => (p.kind === 'cmp' && p.op === '=' && p.col === c.col) || (p.kind === 'or' && p.col === c.col));
        if (prev) {
          if (prev.kind === 'or') prev.conds.push(c);
          else { const idx = merged.indexOf(prev); merged[idx] = { kind: 'or', col: c.col, conds: [prev, c], pos: prev.pos }; }
          continue;
        }
      }
      merged.push(c);
    }
    return { conds: merged, consumed, orMode, rest: work };
  }

  function describeCond(c) {
    const n = '【' + c.col.name + '】';
    const val = v => (c.isDate && typeof v === 'number') ? dateText(v) : (typeof v === 'string' ? '“' + v + '”' : U.formatNumber(v));
    switch (c.kind) {
      case 'cmp': return n + ' ' + OP_SYMBOL[c.op] + ' ' + val(c.value);
      case 'range': return n + ' 在 ' + val(c.lo) + (c.hiOp === '<' && c.isDate ? ' 至 ' + dateText(c.hi - 1) : ' ～ ' + val(c.hi)) + ' 之间';
      case 'text': return n + ({ contains: ' 包含 ', notcontains: ' 不包含 ', starts: ' 以 ', ends: ' 以 ' })[c.op] + '“' + c.value + '”' + (c.op === 'starts' ? ' 开头' : c.op === 'ends' ? ' 结尾' : '');
      case 'blank': return n + (c.empty ? ' 为空' : ' 不为空');
      case 'monthNum': return n + ' 是 ' + c.month + ' 月';
      case 'or': return n + ' 是 ' + c.conds.map(x => val(x.value)).join(' 或 ');
      default: return n;
    }
  }

  /* ---------- 条件 → SUMIFS 条件对 / 布尔表达式 ---------- */

  function critPairs(c, R) {
    const r = R(c.col);
    const escText = v => (/^[<>=]/.test(v) ? '"=' + v.replace(/"/g, '""') + '"' : U.excelString(v.replace(/([*?~])/g, '~$1')));
    switch (c.kind) {
      case 'cmp': {
        if (c.op === '=') return [[r, typeof c.value === 'string' ? escText(c.value) : litExpr(c, c.value)]];
        if (c.isDate) return [[r, '"' + c.op + '"&' + dateExpr(c.value)]];
        return [[r, U.excelString(c.op + c.value)]];
      }
      case 'range':
        if (c.isDate) return [[r, '"' + c.loOp + '"&' + dateExpr(c.lo)], [r, '"' + c.hiOp + '"&' + dateExpr(c.hi)]];
        return [[r, U.excelString(c.loOp + c.lo)], [r, U.excelString(c.hiOp + c.hi)]];
      case 'text': {
        const v = c.value.replace(/([*?~])/g, '~$1');
        const pat = { contains: '*' + v + '*', notcontains: '<>*' + v + '*', starts: v + '*', ends: '*' + v }[c.op];
        return [[r, U.excelString(pat)]];
      }
      case 'blank': return [[r, c.empty ? '""' : '"<>"']];
      default: return null; // monthNum / or 不能直接写成条件对
    }
  }

  function boolAtoms(c, R) {
    const r = R(c.col);
    switch (c.kind) {
      case 'cmp': return [r + (c.op === '<>' ? '<>' : c.op) + litExpr(c, c.value)];
      case 'range': return [r + c.loOp + litExpr(c, c.lo), r + c.hiOp + litExpr(c, c.hi)];
      case 'text': {
        const s = U.excelString(c.value);
        if (c.op === 'contains') return ['ISNUMBER(SEARCH(' + s + ',' + r + '))'];
        if (c.op === 'notcontains') return ['ISERROR(SEARCH(' + s + ',' + r + '))'];
        if (c.op === 'starts') return ['LEFT(' + r + ',' + c.value.length + ')=' + s];
        return ['RIGHT(' + r + ',' + c.value.length + ')=' + s];
      }
      case 'blank': return [r + (c.empty ? '=""' : '<>""')];
      case 'monthNum': return ['MONTH(' + r + ')=' + c.month];
      case 'or': return [null, c.conds.map(x => boolAtoms(x, R))];
      default: return [];
    }
  }

  /** 行内判断：AND(...)/OR(...) */
  function rowBool(conds, R, orMode) {
    const parts = [];
    for (const c of conds) {
      const a = boolAtoms(c, R);
      if (a[0] === null) parts.push('OR(' + a[1].map(x => x.length > 1 ? 'AND(' + x.join(',') + ')' : x[0]).join(',') + ')');
      else if (a.length > 1 && !orMode) parts.push(...a);
      else if (a.length > 1) parts.push('AND(' + a.join(',') + ')');
      else parts.push(a[0]);
    }
    if (parts.length === 1) return parts[0];
    return (orMode ? 'OR(' : 'AND(') + parts.join(',') + ')';
  }

  /** 数组判断（放在 SUMPRODUCT / AGGREGATE 里）：(A=1)*(B>2) */
  function arrBool(conds, R) {
    const parts = [];
    for (const c of conds) {
      const a = boolAtoms(c, R);
      if (a[0] === null) parts.push('((' + a[1].map(x => x.map(y => '(' + y + ')').join('*')).join(')+(') + ')>0)');
      else a.forEach(x => parts.push('(' + x + ')'));
    }
    return parts;
  }

  /* ======================= 汇总公式生成 ======================= */

  /**
   * type: sum|count|avg|max|min|median|distinct
   * target: 目标列（count 可为空）
   * conds: 条件
   * key: 分组键 {pairs: [[range, crit]], atoms: [bool]}
   */
  function aggFormula(q, type, target, conds, key, R) {
    key = key || { pairs: [], atoms: [] };
    const orConds = conds.filter(c => c.kind === 'or');
    const plain = conds.filter(c => c.kind !== 'or');
    let pairs = key.pairs.slice();
    let pairable = true;
    for (const c of plain) {
      const p = critPairs(c, R);
      if (!p) { pairable = false; break; }
      pairs = pairs.concat(p);
    }
    if (orConds.length > 1) pairable = false;
    const T = target ? R(target) : null;
    const flat = ps => ps.map(p => p[0] + ',' + p[1]).join(',');

    if (type === 'distinct') {
      if (!target) return null;
      if (!conds.length && !key.pairs.length) return '=SUMPRODUCT((' + T + '<>"")/COUNTIF(' + T + ',' + T + '&""))';
      if (!conds.length && key.atoms.length === 1 && key.pairs.length === 1) {
        const g = key.pairs[0][0];
        return '=SUMPRODUCT((' + key.atoms[0] + ')/COUNTIFS(' + g + ',' + g + '&"",' + T + ',' + T + '&""))';
      }
      return null;
    }

    if (['sum', 'count', 'avg'].includes(type) && pairable) {
      const alts = orConds.length ? orConds[0].conds.map(c => pairs.concat(critPairs(c, R))) : [pairs];
      const one = (ps, t) => {
        if (t === 'sum') return !ps.length ? 'SUM(' + T + ')' : ps.length === 1 ? 'SUMIF(' + ps[0][0] + ',' + ps[0][1] + ',' + T + ')' : 'SUMIFS(' + T + ',' + flat(ps) + ')';
        if (t === 'count') {
          if (!ps.length) return 'COUNTA(' + (T || R(repColumn(q))) + ')';
          return ps.length === 1 ? 'COUNTIF(' + ps[0][0] + ',' + ps[0][1] + ')' : 'COUNTIFS(' + flat(ps) + ')';
        }
        if (!ps.length) return 'AVERAGE(' + T + ')';
        return ps.length === 1 ? 'AVERAGEIF(' + ps[0][0] + ',' + ps[0][1] + ',' + T + ')' : 'AVERAGEIFS(' + T + ',' + flat(ps) + ')';
      };
      if (type === 'sum' && !T) return null;
      if (type === 'avg' && !T) return null;
      if (alts.length === 1) return '=' + one(alts[0], type);
      if (type === 'avg') return '=(' + alts.map(a => one(a, 'sum')).join('+') + ')/(' + alts.map(a => one(a.concat([[T, '"<>"']]), 'count')).join('+') + ')';
      return '=' + alts.map(a => one(a, type)).join('+');
    }

    // 数组写法
    const atoms = key.atoms.map(a => '(' + a + ')').concat(arrBool(conds, R));
    if (type === 'sum') return T ? '=SUMPRODUCT(' + atoms.concat([T]).join('*') + ')' : null;
    if (type === 'count') return atoms.length ? '=SUMPRODUCT(' + atoms.join('*') + ')' : null;
    if (type === 'avg') return T ? '=SUMPRODUCT(' + atoms.concat([T]).join('*') + ')/SUMPRODUCT(' + atoms.join('*') + ')' : null;
    if (!T) return null;
    const fnum = { max: 14, min: 15, median: 17 }[type];
    const k = type === 'median' ? 2 : 1;
    if (!atoms.length) return '=' + ({ max: 'MAX', min: 'MIN', median: 'MEDIAN' }[type]) + '(' + T + ')';
    return '=AGGREGATE(' + fnum + ',6,' + T + '/' + atoms.join('/') + ',' + k + ')';
  }

  /** 统计行数时用的代表列：姓名 > 编号 > 非空最多 */
  function repColumn(q) {
    const cols = (q.planSheet || q.sheet).columns.filter(c => !c.added);
    return cols.find(c => c.isName) || cols.find(c => c.isId) || cols.slice().sort((a, b) => b.nonEmpty - a.nonEmpty)[0];
  }

  /* ======================= 识别汇总方式、分组 ======================= */

  function findAggs(q) {
    const s = q.sig;
    const found = [];
    for (const [type, re] of AGG_DEFS) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(s))) {
        if (found.some(f => m.index < f.end && m.index + m[0].length > f.pos)) continue;
        found.push({ type, pos: m.index, end: m.index + m[0].length, word: m[0] });
      }
    }
    found.sort((a, b) => a.pos - b.pos);
    // “有多少”只在没有其他汇总词时才表示计数
    const strong = found.filter(f => f.type !== 'countWeak');
    if (strong.length) return strong;
    return found.map(f => Object.assign(f, { type: 'count' })).slice(0, 1);
  }

  function findGroups(q) {
    const groups = [];
    const E = i => q.ents[+i];
    const s = q.sig;
    let m;
    const add = (col, unit, pos) => { if (col && !groups.some(g => g.col === col && g.unit === unit)) groups.push({ col, unit, pos }); };
    const re1 = /(?:按|按照|根据|依据|依照|以|分)\s*(?:每个|各个|不同的?)?\s*⟨C(\d+)⟩(?:\s*(?:和|与|、|及|,)\s*⟨C(\d+)⟩)?(?!\s*(?:开头|结尾|打头|查找|查询|匹配|找|排名|排序|内|中|里|计算|算|提取|求出|得出|判断|生成|获取|取出|转换|推算|标记|从))/g;
    while ((m = re1.exec(s))) {
      add(E(m[1]).col, 'value', m.index);
      if (m[2] !== undefined) add(E(m[2]).col, 'value', m.index);
    }
    const re2 = /(?:每个|每一个|每位|每名|每家|每种|每类|每一|各个|各|每|不同的?)\s*⟨C(\d+)⟩/g;
    while ((m = re2.exec(s))) add(E(m[1]).col, 'value', m.index);
    const re3 = /⟨C(\d+)⟩\s*(?:分别|各自)/g;
    while ((m = re3.exec(s))) add(E(m[1]).col, 'value', m.index);
    // 按月 / 按年
    const dcol = dateColOf(q, null);
    if (dcol) {
      if (/(?:按|按照|每|各|分)\s*(?:个)?\s*(?:月|月份)|月度|逐月|每个月|各月份?|月份汇总|分月/.test(s)) add(dcol, 'month', s.search(/月/));
      else if (/(?:按|按照|每|各|分)\s*(?:个)?\s*(?:年|年份|年度)(?!龄)|逐年|每一年|各年份?|分年/.test(s)) add(dcol, 'year', s.search(/年/));
    }
    // “各班”“每个部”这类简称：取列名开头相同的分类列
    if (!groups.length) {
      const re4 = /(?:每个|各个|各|每|按|分)\s*([一-龥]{1,2})/g;
      while ((m = re4.exec(s))) {
        const w = m[1];
        const hits = q.sheet.columns.filter(c => c.isCategory && c.name.startsWith(w) && c.name !== w);
        if (hits.length === 1) { add(hits[0], 'value', m.index); break; }
      }
    }
    // “每个人 / 每位员工”等指代姓名列
    if (/(?:每个|每|各|按|每位)\s*(?:人|员工|销售员|业务员|学生|客户|同学)(?![一-龥]*⟨)/.test(s) && !groups.length) {
      const nc = q.sheet.columns.find(c => c.isName);
      if (nc) add(nc, 'value', 0);
    }
    // 数值列只有取值很少（如年级 1/2/3）时才适合当分组依据
    return groups.filter(g => g.unit !== 'value' || g.col.type !== 'number' || g.col.distinctCount <= 50);
  }

  /* ======================= 计划构造 ======================= */

  function plan(kind, outputs, extra) {
    return Object.assign({ kind, outputs, source: 'template' }, extra || {});
  }

  function titleHint(q) {
    const s = q.sig;
    let m = s.match(/(?:列名|列名称|标题|表头)\s*(?:叫做?|为|是|:|=)\s*(?:⟨L(\d+)⟩|([^\s,;⟨⟩.]{1,15}))/) ||
      s.match(/(?:新增|添加|增加|生成|插入|加)\s*(?:一个?)?\s*(?:新的?)?\s*列\s*(?:叫做?|名为|命名为|:|，|,)\s*(?:⟨L(\d+)⟩|([^\s,;⟨⟩.]{1,15}))/) ||
      s.match(/命名为\s*(?:⟨L(\d+)⟩|([^\s,;⟨⟩.]{1,15}))/);
    if (!m) return null;
    return m[1] !== undefined ? q.ents[+m[1]].text : m[2];
  }

  const numCols = q => q.colEnts.filter(e => e.sheet === q.sheet && (e.col.type === 'number' || e.col.type === 'textnum')).map(e => e.col);

  /* ======================= 各个模板 ======================= */

  // ---------- 身份证 ----------
  function tIdCard(q) {
    if (!/身份证/.test(q.text) && !q.activeCols.some(c => c.isIdCard)) return null;
    const id = q.activeCols.find(c => c.isIdCard) || q.sheet.columns.find(c => c.isIdCard);
    if (!id) return null;
    const x = '[@' + id.name + ']';
    const birth = 'DATE(MID(' + x + ',7,4),MID(' + x + ',11,2),MID(' + x + ',13,2))';
    const outs = [];
    const expl = [];
    if (/性别|男女|男还是女|男性|女性/.test(q.text)) {
      outs.push({ title: '性别', formula: '=IF(MOD(MID(' + x + ',17,1),2)=1,"男","女")', format: 'general' });
      expl.push('身份证第 17 位是性别码：奇数为男、偶数为女，用 MID 取出这一位，MOD 判断奇偶');
    }
    if (/年龄|岁|多大|周岁/.test(q.text)) {
      outs.push({ title: '年龄', formula: '=DATEDIF(' + birth + ',TODAY(),"Y")', format: 'integer' });
      expl.push('身份证第 7–14 位是出生年月日，用 MID 分别取出后由 DATE 组成日期，再用 DATEDIF 计算到今天的整年数');
    }
    if (/出生年份|出生的年份|哪年出生/.test(q.text)) {
      outs.push({ title: '出生年份', formula: '=VALUE(MID(' + x + ',7,4))', format: 'integer' });
      expl.push('身份证第 7–10 位是出生年份');
    } else if (/出生|生日/.test(q.text)) {
      outs.push({ title: '出生日期', formula: '=' + birth, format: 'date' });
      expl.push('身份证第 7–14 位是出生年月日，用 MID 取出年、月、日，再用 DATE 组成真正的日期');
    }
    if (/脱敏|隐藏|打码|遮挡|星号|\*/.test(q.text)) {
      outs.push({ title: id.name + '(脱敏)', formula: '=REPLACE(' + x + ',7,8,"********")', format: 'general' });
      expl.push('用 REPLACE 把第 7 位起的 8 位（出生日期）替换成星号');
    }
    if (!outs.length) return null;
    return plan('column', outs, {
      interpretation: '根据【' + id.name + '】提取' + outs.map(o => '“' + o.title + '”').join('、'),
      explanation: expl.join('；') + '。适用于 18 位身份证号。',
    });
  }

  // ---------- 跨表查找 ----------
  function tLookup(q) {
    if (!/查找|查询|查出|查|匹配|对应|引用|带出|调取|找出|取出|获取|获得|拿到|取过来|vlookup|填入|填上|带过来|关联/i.test(q.text)) return null;
    const sheetEnt = q.ents.find(e => e.t === 'S');
    let T = sheetEnt ? sheetEnt.sheet : null;
    if (!T) {
      const other = q.colEnts.find(e => e.sheet !== q.sheet);
      if (other) T = other.sheet;
    }
    if (!T || T === q.sheet) return null;
    const A = q.sheet;
    const byName = (sh, name) => sh.columns.find(c => c.name === name || aliasesOf(c).includes(name));
    const mentioned = q.colEnts.map(e => e.col.name);
    // 关键列：两表都有、且被提到的列
    let keyName = mentioned.find(n => byName(A, n) && byName(T, n));
    if (!keyName) {
      const common = A.columns.filter(c => !c.added && byName(T, c.name));
      const pick = common.find(c => c.isId) || common.find(c => c.isName) || common[0];
      if (pick) keyName = pick.name;
    }
    if (!keyName) return null;
    const keyA = byName(A, keyName), keyT = byName(T, keyName);
    let rets = mentioned.filter(n => n !== keyName && byName(T, n) && !q.colEnts.some(e => e.col.name === n && e.sheet === A && !byName(T, n)));
    rets = Array.from(new Set(rets)).map(n => byName(T, n)).filter(c => c && c !== keyT);
    if (!rets.length) return null;
    const outs = [];
    const notes = [];
    let keyExpr = '[@' + keyA.name + ']';
    const aNum = keyA.type === 'number', tNum = keyT.type === 'number';
    if (aNum && (keyT.type === 'textnum' || keyT.type === 'text')) { keyExpr = keyExpr + '&""'; notes.push('两张表里【' + keyName + '】一个是数字、一个是文本，公式里用 &"" 把它转成文本再匹配'); }
    if (!aNum && keyA.type === 'textnum' && tNum) { keyExpr = 'VALUE(' + keyExpr + ')'; notes.push('两张表里【' + keyName + '】一个是文本、一个是数字，公式里用 VALUE 转成数字再匹配'); }
    const notFound = /空白|留空|为空|显示空/.test(q.text) ? '""' : '"未找到"';
    for (const r of rets) {
      let f;
      if (r.index > keyT.index) {
        const n = r.index - keyT.index + 1;
        f = '=IFERROR(VLOOKUP(' + keyExpr + ',[' + T.name + '!' + keyT.name + ':' + r.name + '],' + n + ',FALSE),' + notFound + ')';
      } else {
        f = '=IFERROR(INDEX([' + T.name + '!' + r.name + '],MATCH(' + keyExpr + ',[' + T.name + '!' + keyT.name + '],0)),' + notFound + ')';
      }
      const exists = A.columns.some(c => c.name === r.name);
      outs.push({ title: exists ? r.name + '(' + T.name + ')' : r.name, formula: f, format: r.type === 'date' ? 'date' : 'general' });
    }
    const usedV = outs.some(o => /VLOOKUP/.test(o.formula));
    const usedI = outs.some(o => /INDEX/.test(o.formula));
    let expl = '用本行的【' + keyName + '】去「' + T.name + '」里找到同一个' + keyName + '，再取出对应的' + rets.map(r => '【' + r.name + '】').join('、') + '。';
    if (usedV) expl += 'VLOOKUP 的第 3 个参数是要取第几列（从查找列算起），FALSE 表示必须精确匹配。';
    if (usedI) expl += '要取的列在查找列左边时 VLOOKUP 做不到，所以用 MATCH 先找到行号、再用 INDEX 取值。';
    expl += '外层 IFERROR 让找不到时显示' + (notFound === '""' ? '空白' : '“未找到”') + '而不是 #N/A。';
    if (notes.length) expl += notes.join('；') + '。';
    return plan('column', outs, {
      interpretation: '按【' + keyName + '】从「' + T.name + '」查找' + rets.map(r => '【' + r.name + '】').join('、') + '，填到当前表新列',
      explanation: expl,
    });
  }

  // ---------- 四则运算 ----------
  const OPND = '(?:[(]\\s*)*⟨[CN]\\d+⟩(?:\\s*[)])*';
  const OPR = '(?:\\+|-|\\*|/|乘以|乘上|乘|除以|加上|加|减去|减掉|减)';
  const EXPR_RE = new RegExp(OPND + '(?:\\s*' + OPR + '\\s*' + OPND + ')+', 'g');

  function tExpr(q) {
    if (/如果|若|否则|标记|则显示|则为/.test(q.sig)) return null;
    let best = null;
    EXPR_RE.lastIndex = 0;
    let m;
    while ((m = EXPR_RE.exec(q.sig))) {
      if (!/⟨C\d+⟩/.test(m[0])) continue;
      if (!best || m[0].length > best[0].length) best = m;
    }
    let exprSig = best ? best[0] : null;
    let titleFromEq = null;
    // A 和 B 相乘 / 之差 / 之比
    if (!exprSig) {
      const m2 = q.sig.match(/⟨C(\d+)⟩\s*(?:和|与|跟|、)\s*⟨C(\d+)⟩\s*(?:的)?\s*(相乘|乘积|之积|相除|之比|比值|之差|的差|相减|差值|差额)/);
      if (!m2) return null;
      const op = /乘|积/.test(m2[3]) ? '*' : /除|比/.test(m2[3]) ? '/' : '-';
      exprSig = '⟨C' + m2[1] + '⟩' + op + '⟨C' + m2[2] + '⟩';
    }
    const E = i => q.ents[+i];
    const cols = [];
    let bad = false;
    let f = exprSig.replace(/⟨([CN])(\d+)⟩/g, (mm, t, i) => {
      const e = E(i);
      if (t === 'N') return String(e.v);
      if (e.sheet !== q.sheet) { bad = true; return ''; }
      cols.push(e.col);
      return '[@' + e.col.name + ']';
    });
    if (bad || !cols.length) return null;
    f = f.replace(/乘以|乘上|乘/g, '*').replace(/除以/g, '/').replace(/加上|加/g, '+').replace(/减去|减掉|减/g, '-').replace(/\s+/g, '');
    let depth = 0;
    for (const ch of f) { if (ch === '(') depth++; if (ch === ')') depth--; if (depth < 0) return null; }
    if (depth !== 0) return null;
    // 标题：“利润=收入-成本”
    const eq = q.sig.match(/(?:^|[\s,;:])(?:计算|求|算出|新增|添加)?\s*([一-龥A-Za-z0-9_()（）]{1,15})\s*=\s*(?:[(]\s*)*⟨/);
    if (eq) titleFromEq = eq[1].replace(/^(计算|求|算出|新增一列|新增|添加一列|添加|一列)/, '');
    const eqCol = q.sig.match(/⟨C(\d+)⟩\s*=\s*(?:[(]\s*)*⟨/);
    if (!titleFromEq && eqCol && exprSig.indexOf('⟨C' + eqCol[1] + '⟩') < 0) titleFromEq = q.ents[+eqCol[1]].col.name;
    const pretty = exprSig.replace(/⟨([CN])(\d+)⟩/g, (mm, t, i) => (t === 'N' ? E(i).raw : E(i).col.name))
      .replace(/\*/g, '×').replace(/\//g, '÷').replace(/\s+/g, '');
    let title = titleHint(q) || titleFromEq || pretty.replace(/乘以|乘上|乘/g, '×').replace(/除以/g, '÷').replace(/加上|加/g, '+').replace(/减去|减掉|减/g, '-');
    if (title.length > 20) title = '计算结果';
    let format = 'general';
    if (/率|比例|占比|百分比|百分数/.test(title + q.text)) format = 'percent';
    const dateCols = cols.filter(c => c.type === 'date');
    if (dateCols.length === 2 && /-/.test(f)) format = 'integer';
    else if (dateCols.length === 1 && /[+-]/.test(f) && cols.length === 1) format = 'date';
    // 保留小数
    const rd = q.sig.match(/保留\s*⟨N(\d+)⟩\s*位(?:小数)?|四舍五入/);
    if (rd) {
      const n = rd[1] !== undefined ? E(rd[1]).v : 0;
      f = 'ROUND(' + f + ',' + n + ')';
    }
    return plan('column', [{ title, formula: '=' + f, format }], {
      interpretation: '每一行计算：' + pretty + (rd ? '（四舍五入）' : ''),
      explanation: '直接用单元格做加减乘除，' + cols.map(c => '【' + c.name + '】').join('、') + '取本行的值。' +
        (format === 'percent' ? '结果设置为百分比格式。' : '') + (rd ? 'ROUND 用来四舍五入保留小数位。' : '') +
        (dateCols.length === 2 ? '日期相减得到相差的天数。' : ''),
    });
  }

  // ---------- 排名 ----------
  function tRank(q) {
    if (!/排名|名次|排第几|排位|排行|第几名/.test(q.sig)) return null;
    if (/前\s*⟨N\d+⟩\s*(名|位)/.test(q.sig) && /谁|哪些|哪几|是/.test(q.sig)) return null;
    // 组内排名：“在部门内…排名”“每个班级…排名”“分部门排名”（“按销售额排名”的“按”指排名依据，不是分组）
    const groupM = q.sig.match(/(?:在|每个|各个|各|分|同一)\s*⟨C(\d+)⟩\s*(?:内|中|里|内部|之内)?/) ||
      q.sig.match(/⟨C(\d+)⟩\s*(?:内|内部|之内)/);
    let gcol = groupM ? q.ents[+groupM[1]].col : null;
    if (gcol && gcol.type === 'number') gcol = null;
    if (!gcol) {
      const m2 = q.sig.match(/(部门|班级|班|组|小组|区域|门店|地区)内/);
      if (m2) gcol = q.sheet.columns.find(c => c.name.indexOf(m2[1]) >= 0) || null;
    }
    const targets = numCols(q).filter(c => c !== gcol && c.type === 'number');
    const target = targets[0];
    if (!target) return null;
    const asc = /升序|从低到高|从小到大|由低到高|由小到大|越小越|越少越|越低越|最小的?排第一|最少的?排第一|最低的?排第一/.test(q.text);
    let f, title, expl;
    if (gcol && gcol !== target) {
      f = '=COUNTIFS([' + gcol.name + '],[@' + gcol.name + '],[' + target.name + '],"' + (asc ? '<' : '>') + '"&[@' + target.name + '])+1';
      title = gcol.name + '内排名';
      expl = '用 COUNTIFS 统计“同一' + gcol.name + '里，' + target.name + (asc ? '比自己小' : '比自己大') + '的人数”，再加 1 就是名次；并列时名次相同。';
    } else {
      f = '=RANK([@' + target.name + '],[' + target.name + '],' + (asc ? 1 : 0) + ')';
      title = '排名';
      expl = 'RANK 返回本行【' + target.name + '】在整列中的名次；第 3 个参数 ' + (asc ? '1 表示从小到大' : '0 表示从大到小') + '排，并列时名次相同。';
    }
    return plan('column', [{ title: titleHint(q) || title, formula: f, format: 'integer' }], {
      interpretation: '按【' + target.name + '】' + (asc ? '从低到高' : '从高到低') + '排名' + (gcol && gcol !== target ? '（在每个【' + gcol.name + '】内部分别排）' : ''),
      explanation: expl,
    });
  }

  // ---------- 条件分档 / 判断 ----------
  const ASSIGN = '(?:则显示为|则显示|则为|则标记为|则返回|就显示|就标记为|就是|就为|显示为|显示成|显示|标记为|标记成|标记|标为|标注为|标注|评为|评定为|算作|算为|算|记为|记作|返回|输出|填写|填|写|定为|视为|划为|归为|设为|则|为|是|:|(?<![<>!=])=(?!=)|->)';
  const ELSE_RE = /^\s*(?:否则|不然|其余|其他情况|其它情况|其他的?|其它的?|剩下的?|反之|以外的?|别的)\s*(?:的)?\s*/;
  const BAD_LABEL = /多少|几个|几人|求和|合计|总和|平均|最大|最小|最高|最低|统计|计数|人数|个数|数量|有几|排名/;

  function parseLabel(q, s) {
    s = s.trim().replace(/^(?:则|就|那么)\s*/, '');
    let m = s.match(/^⟨([LVN])(\d+)⟩/);
    if (m) {
      const e = q.ents[+m[2]];
      if (e.t === 'N') return { value: e.v, text: e.raw };
      return { value: e.t === 'L' ? e.text : e.value, text: e.t === 'L' ? e.text : e.value };
    }
    m = s.match(/^([^\s⟨⟩,;.、]{1,12})/);
    if (!m) return null;
    let t = m[1].replace(/(的|了|吧|呢)$/, '');
    if (!t || BAD_LABEL.test(t)) return null;
    if (/^(空|空白|留空|不显示|空着)$/.test(t)) return { value: '', text: '空白' };
    if (/^-?\d+(\.\d+)?$/.test(t)) return { value: Number(t), text: t };
    return { value: t, text: t };
  }

  function tTiers(q, aggs) {
    const s = q.sig;
    const hasWhether = /是否|是不是|有没有|有无/.test(s);
    if (!/如果|若|假如|要是|则|否则|不然|其余|其他|其它|标记|标注|标为|显示|返回|评为|评定|判定|判断|算作|记为|等级|评级|档|分级|达标|合格|及格|以上|以下|之间|是否|是不是/.test(s)) return null;
    // 有统计词、却没有明确的“如果/标记/显示”等说法时，更可能是统计需求
    if (aggs.length && !/如果|若|假如|要是|则|否则|不然|标记|标注|标为|显示|返回|评为|评定|判定|判断|等级|评级|分级|是否|是不是/.test(s)) return null;
    const numMentioned = numCols(q);
    const subject = numMentioned[0] || q.activeCols.find(c => c.type === 'date') || q.activeCols[0] || null;
    const clauses = s.split(/[,;]|(?=否则|不然)/).map(x => x.trim()).filter(Boolean);
    const rules = [];
    let elseLabel = null;
    let pendingConds = null;
    const condOpts = { defaultCol: subject, allowBareEq: false, skipLooseDates: false };

    for (let cl of clauses) {
      cl = cl.replace(/^(?:如果|若|假如|要是|当|如)\s*/, '');
      if (ELSE_RE.test(cl)) {
        const rest = cl.replace(ELSE_RE, '').replace(new RegExp('^' + ASSIGN + '\\s*'), '');
        const lb = parseLabel(q, rest);
        if (lb) elseLabel = lb;
        continue;
      }
      // 找最后一个赋值词
      const re = new RegExp(ASSIGN, 'g');
      let m, last = null;
      while ((m = re.exec(cl))) { if (m[0]) last = m; else re.lastIndex++; }
      let condPart = cl, label = null;
      if (last) {
        const lb = parseLabel(q, cl.slice(last.index + last[0].length));
        if (lb && cl.slice(last.index + last[0].length).trim()) {
          // “是/为”后面若是取值实体，更可能是条件而不是结果
          const after = cl.slice(last.index + last[0].length).trim();
          const isCondVal = /^(是|为|=)$/.test(last[0]) && /^⟨V\d+⟩/.test(after) && !/的\s*$/.test(cl.slice(0, last.index));
          if (!isCondVal) { condPart = cl.slice(0, last.index); label = lb; }
        }
      }
      if (!label) {
        // 无赋值词：“90分以上优秀”
        const m2 = cl.match(/^(.*(?:以上|以下|以内|之间|的|⟨N\d+⟩\s*(?:分|元|岁|个|件|天|次)?))\s*((?:⟨[LV]\d+⟩)|[^\s⟨⟩,;.、]{1,12})$/);
        if (m2) { const lb = parseLabel(q, m2[2]); if (lb) { condPart = m2[1]; label = lb; } }
      }
      const cs = condPart.trim() ? extractConditions(q, condPart, condOpts) : { conds: [], orMode: false };
      if (label) {
        let conds = cs.conds.length ? cs.conds : (pendingConds ? pendingConds.conds : []);
        const orMode = cs.conds.length ? cs.orMode : (pendingConds ? pendingConds.orMode : false);
        if (!conds.length) continue;
        rules.push({ conds, orMode, label });
        pendingConds = null;
      } else if (cs.conds.length) {
        pendingConds = cs;
      }
    }

    // “判断销售额是否大于 10000”
    if (!rules.length && hasWhether) {
      const part = s.slice(Math.max(0, s.search(/是否|是不是|有没有|有无/) - 40));
      const cleaned = part.replace(/是否|是不是|有没有|有无/g, ' ');
      const cs = extractConditions(q, cleaned, Object.assign({}, condOpts, { allowBareEq: true }));
      if (cs.conds.length) {
        const yes = /达标/.test(s) ? '达标' : '是', no = /达标/.test(s) ? '未达标' : '否';
        rules.push({ conds: cs.conds, orMode: cs.orMode, label: { value: yes, text: yes } });
        elseLabel = elseLabel || { value: no, text: no };
      }
    }
    if (!rules.length) return null;
    if (/多少|几个|几人|人数|求和|合计|平均|总和/.test(s) && !rules.some(r => r.label.text.length > 0)) return null;

    const R = mkRef(q);
    const cell = col => R.C(col);
    for (const r of rules) for (const c of r.conds) {
      const cs = c.kind === 'or' ? c.conds : [c];
      if (cs.some(x => colSheet(q, x.col) !== q.sheet)) return null;
    }
    const lit = lb => (typeof lb.value === 'number' ? String(lb.value) : U.excelString(lb.value));
    const elseExpr = elseLabel ? lit(elseLabel) : '""';

    // 同一数值列的分档：按阈值从高到低嵌套
    const single = rules.every(r => r.conds.length === 1 && ['cmp', 'range'].includes(r.conds[0].kind) && !r.conds[0].isDate);
    const sameCol = single && rules.every(r => r.conds[0].col === rules[0].conds[0].col);
    let f;
    let tiersText = [];
    if (sameCol && rules.length > 1 && rules.every(r => r.conds[0].kind === 'range' || ['>', '>='].includes(r.conds[0].op) || ['<', '<='].includes(r.conds[0].op))) {
      const col = rules[0].conds[0].col;
      const lowers = [], bottoms = [];
      for (const r of rules) {
        const c = r.conds[0];
        if (c.kind === 'range') lowers.push({ op: '>=', v: c.lo, label: r.label });
        else if (c.op === '>' || c.op === '>=') lowers.push({ op: c.op, v: c.value, label: r.label });
        else bottoms.push({ op: c.op, v: c.value, label: r.label });
      }
      lowers.sort((a, b) => b.v - a.v);
      let tail = bottoms.length ? lit(bottoms[0].label) : elseExpr;
      if (bottoms.length > 1 || (bottoms.length && elseLabel)) tail = null;
      if (tail !== null) {
        f = tail;
        for (let i = lowers.length - 1; i >= 0; i--) f = 'IF(' + cell(col) + lowers[i].op + lowers[i].v + ',' + lit(lowers[i].label) + ',' + f + ')';
        tiersText = lowers.map(l => OP_SYMBOL[l.op] + U.formatNumber(l.v) + ' →“' + l.label.text + '”');
        if (bottoms.length) tiersText.push('其余 →“' + bottoms[0].label.text + '”');
        else tiersText.push('其余 →' + (elseLabel ? '“' + elseLabel.text + '”' : '空白'));
      }
    }
    if (!f) {
      f = elseExpr;
      for (let i = rules.length - 1; i >= 0; i--) {
        const r = rules[i];
        f = 'IF(' + rowBool(r.conds, cell, r.orMode) + ',' + lit(r.label) + ',' + f + ')';
      }
      tiersText = rules.map(r => r.conds.map(describeCond).join(r.orMode ? ' 或 ' : ' 且 ') + ' →“' + r.label.text + '”');
      tiersText.push('其余 →' + (elseLabel ? '“' + elseLabel.text + '”' : '空白'));
    }
    const labels = rules.map(r => String(r.label.text)).concat(elseLabel ? [String(elseLabel.text)] : []);
    let title = titleHint(q);
    if (!title) {
      const wm = q.text.match(/(?:是否|是不是|有没有)\s*([^,;.\s]{1,10})/);
      if (labels.some(l => /^(优秀|良好|中等|及格|不及格|合格|不合格|A|B|C|D|甲|乙|丙|丁|高|中|低)$/.test(l))) title = labels.some(l => /及格/.test(l)) && labels.length <= 2 ? '是否及格' : '等级';
      else if (wm) title = '是否' + wm[1].replace(/(的|了|吗)$/, '');
      else if (rules.length === 1 && !elseLabel && typeof rules[0].label.value === 'string' && rules[0].label.value.length <= 6) title = '是否' + rules[0].label.value;
      else if (labels.every(l => /^(是|否|达标|未达标|Y|N)$/.test(l))) title = '是否达标';
      else title = '分类';
    }
    const ifCount = (f.match(/IF\(/g) || []).length;
    return plan('column', [{ title, formula: '=' + f, format: rules.every(r => typeof r.label.value === 'number') ? 'general' : 'general' }], {
      interpretation: '逐行判断：' + tiersText.join('；'),
      explanation: 'IF(条件, 满足时的结果, 不满足时的结果)。' + (ifCount > 1 ? '多个档位时把 IF 一层层嵌套：先判断最高一档，不满足再往下判断。' : '') +
        (/AND\(/.test(f) ? 'AND 表示几个条件要同时满足。' : '') + (/OR\(/.test(f) ? 'OR 表示满足其中一个即可。' : '') +
        (/SEARCH/.test(f) ? 'ISNUMBER(SEARCH(…)) 用来判断是否包含某段文字。' : ''),
    });
  }

  // ---------- 重复 ----------
  function tDuplicate(q) {
    if (!/重复|出现(了)?几次|出现次数|第几次出现|出现的次数/.test(q.sig)) return null;
    const col = q.activeCols[0] || q.sheet.columns.find(c => c.isName) || q.sheet.columns.find(c => c.isId);
    if (!col) return null;
    if (/去重|不重复/.test(q.sig) && /多少|几|个数|数量|计数/.test(q.sig)) return null; // 交给“不重复计数”
    if (/第几次|出现次序|次序/.test(q.sig)) {
      return plan('column', [{ title: '第几次出现', formula: '=COUNTIF([^' + col.name + '],[@' + col.name + '])', format: 'integer' }], {
        interpretation: '标出每个【' + col.name + '】是第几次出现',
        explanation: 'COUNTIF 的范围从第一行“滚动”到当前行（起点锁定用 $，终点不锁），数一数到本行为止这个值出现了几次。',
      });
    }
    if (/出现(了)?几次|出现次数|出现的次数|重复(了)?几次|重复次数/.test(q.sig)) {
      return plan('column', [{ title: '出现次数', formula: '=COUNTIF([' + col.name + '],[@' + col.name + '])', format: 'integer' }], {
        interpretation: '统计每个【' + col.name + '】在整列中出现的次数',
        explanation: 'COUNTIF(整列, 本行的值) 统计这个值在整列出现了几次。',
      });
    }
    if (/多少|几个|几条|统计|计数/.test(q.sig) && !/标记|标出|找出|判断|是否/.test(q.sig)) {
      return plan('scalar', [{ title: col.name + '重复的行数', formula: '=SUMPRODUCT(--(COUNTIF([' + col.name + '],[' + col.name + '])>1))', format: 'integer' }], {
        interpretation: '统计【' + col.name + '】中出现不止一次的行数',
        explanation: 'COUNTIF 对每一行算出它出现的次数，>1 表示重复，SUMPRODUCT 把满足条件的行数加起来。',
      });
    }
    const later = /第二次|再次|后面|之后|以后/.test(q.sig);
    const f = later ? '=IF(COUNTIF([^' + col.name + '],[@' + col.name + '])>1,"重复","")' : '=IF(COUNTIF([' + col.name + '],[@' + col.name + '])>1,"重复","")';
    return plan('column', [{ title: '是否重复', formula: f, format: 'general' }], {
      interpretation: later ? '【' + col.name + '】第二次及以后出现的行标记为“重复”' : '【' + col.name + '】出现不止一次的行都标记为“重复”',
      explanation: later ? 'COUNTIF 只数到当前行为止，次数 >1 说明前面已经出现过，首次出现的那行保持空白。'
        : 'COUNTIF 统计本行的值在整列中出现了几次，超过 1 次即为重复。',
    });
  }

  // ---------- 日期派生 ----------
  function tDate(q, aggs) {
    const s = q.sig;
    const dates = q.activeCols.filter(c => c.type === 'date');
    const outs = [];
    const expl = [];
    let interp = '';
    const pickDate = (pred) => dates.find(pred) || q.sheet.columns.find(c => c.type === 'date' && pred(c)) || null;
    // “按年/按月统计…”是分组汇总，不是提取年份
    if (aggs.length || findGroups(q).length) {
      if (!/年龄|工龄|司龄/.test(s) || findGroups(q).length) return null;
    }

    if (/年龄|周岁|多大|多少岁|几岁/.test(s)) {
      const b = dates.find(c => c.isBirth) || dates[0] || pickDate(c => c.isBirth);
      if (!b) return null;
      outs.push({ title: '年龄', formula: '=DATEDIF([@' + b.name + '],TODAY(),"Y")', format: 'integer' });
      expl.push('DATEDIF(开始日期, 结束日期, "Y") 计算相差的整年数，TODAY() 是今天');
      interp = '根据【' + b.name + '】计算到今天的周岁';
    } else if (/工龄|司龄|入职年限|工作年限|在职年限|年资|入职多久|工作了?几年/.test(s)) {
      const h = dates.find(c => c.isHire) || dates[0] || pickDate(c => c.isHire);
      if (!h) return null;
      const unit = /月/.test(s.replace(/年月/g, '')) && /几个月|月数|多少个月/.test(s) ? 'M' : 'Y';
      outs.push({ title: unit === 'M' ? '在职月数' : '工龄', formula: '=DATEDIF([@' + h.name + '],TODAY(),"' + unit + '")', format: 'integer' });
      expl.push('DATEDIF 计算从入职日期到今天（TODAY）相差的整' + (unit === 'M' ? '月' : '年') + '数');
      interp = '根据【' + h.name + '】计算到今天的' + (unit === 'M' ? '月数' : '年数');
    }
    if (outs.length) return plan('column', outs, { interpretation: interp, explanation: expl.join('；') + '。' });

    const d = dates[0];
    if (!d) return null;
    const x = '[@' + d.name + ']';
    // 两个日期相差
    if (dates.length >= 2 && /相差|间隔|之间|差几|差多少|多少天|天数|时长|用时|耗时|周期|月数|年数|工作日/.test(s)) {
      const a = dates[0], b = dates[1];
      if (/工作日/.test(s)) {
        outs.push({ title: '工作日天数', formula: '=NETWORKDAYS([@' + a.name + '],[@' + b.name + '])', format: 'integer' });
        expl.push('NETWORKDAYS 计算两个日期之间的工作日（周一到周五）天数，含首尾两天');
      } else if (/月数|几个月|多少个月/.test(s)) {
        outs.push({ title: '相差月数', formula: '=DATEDIF([@' + a.name + '],[@' + b.name + '],"M")', format: 'integer' });
        expl.push('DATEDIF(…,"M") 计算相差的整月数，要求前一个日期不晚于后一个');
      } else if (/年数|几年/.test(s)) {
        outs.push({ title: '相差年数', formula: '=DATEDIF([@' + a.name + '],[@' + b.name + '],"Y")', format: 'integer' });
        expl.push('DATEDIF(…,"Y") 计算相差的整年数');
      } else {
        outs.push({ title: '相差天数', formula: '=[@' + b.name + ']-[@' + a.name + ']', format: 'integer' });
        expl.push('日期在 Excel 里本质是数字，两个日期直接相减就是相差的天数');
      }
      return plan('column', outs, { interpretation: '计算【' + a.name + '】到【' + b.name + '】之间的' + outs[0].title.replace('相差', ''), explanation: expl.join('；') + '。' });
    }
    // 加减天数/月份
    const add = s.match(/(?:加|加上|往后|后推|推迟|延后|顺延|\+|过)\s*⟨N(\d+)⟩\s*(个)?\s*(工作日|天|日|月|年|周)/) || s.match(/⟨N(\d+)⟩\s*(个)?\s*(工作日|天|日|月|年|周)\s*(?:后|之后|以后)/);
    const sub = s.match(/(?:减|减去|往前|提前|前推)\s*⟨N(\d+)⟩\s*(个)?\s*(工作日|天|日|月|年|周)/) || s.match(/⟨N(\d+)⟩\s*(个)?\s*(工作日|天|日|月|年|周)\s*(?:前|之前|以前)/);
    const mm = add || sub;
    if (mm) {
      const n = q.ents[+mm[1]].v * (sub && !add ? -1 : 1);
      const unit = mm[3];
      let f, desc;
      if (unit === '工作日') { f = '=WORKDAY(' + x + ',' + n + ')'; desc = 'WORKDAY 计算若干个工作日之后的日期（跳过周末）'; }
      else if (unit === '月') { f = '=EDATE(' + x + ',' + n + ')'; desc = 'EDATE 计算若干个月之后的日期，会自动处理大小月'; }
      else if (unit === '年') { f = '=EDATE(' + x + ',' + (n * 12) + ')'; desc = 'EDATE 按月数推算，' + Math.abs(n) + ' 年 = ' + Math.abs(n * 12) + ' 个月'; }
      else if (unit === '周') { f = '=' + x + (n >= 0 ? '+' : '-') + Math.abs(n * 7); desc = '日期加减数字即加减天数，一周 7 天'; }
      else { f = '=' + x + (n >= 0 ? '+' : '-') + Math.abs(n); desc = '日期在 Excel 里是数字，加减数字就是加减天数'; }
      return plan('column', [{ title: titleHint(q) || (/到期/.test(s) ? '到期日' : '计算日期'), formula: f, format: 'date' }], {
        interpretation: '【' + d.name + '】' + (n >= 0 ? '往后' : '往前') + Math.abs(n) + ' ' + (unit === '日' ? '天' : unit),
        explanation: desc + '。结果设置为日期格式。',
      });
    }
    const parts = [];
    if (/年月|月份\(含年\)|哪年哪月/.test(s)) parts.push({ title: '年月', formula: '=TEXT(' + x + ',"yyyy-mm")', format: 'general', e: 'TEXT 按“年-月”格式输出' });
    else {
      if (/年份|哪一?年|提取年|取年|的年(?!龄)/.test(s)) parts.push({ title: '年份', formula: '=YEAR(' + x + ')', format: 'integer', e: 'YEAR 取年份' });
      if (/月份|哪个?月|提取月|取月|的月(?!末|底|初)/.test(s)) parts.push({ title: '月份', formula: '=MONTH(' + x + ')', format: 'integer', e: 'MONTH 取月份' });
    }
    if (/几号|取日|提取日|日期中的日|哪一?天(?!数)/.test(s)) parts.push({ title: '日', formula: '=DAY(' + x + ')', format: 'integer', e: 'DAY 取日' });
    if (/季度/.test(s)) parts.push({ title: '季度', formula: '="第"&ROUNDUP(MONTH(' + x + ')/3,0)&"季度"', format: 'general', e: '月份除以 3 再向上取整（ROUNDUP）就是第几季度' });
    if (/星期|周几|礼拜/.test(s)) parts.push({ title: '星期', formula: '="星期"&MID("日一二三四五六",WEEKDAY(' + x + '),1)', format: 'general', e: 'WEEKDAY 返回 1（周日）到 7（周六），再用 MID 从“日一二三四五六”中取出对应的字' });
    if (/周末|工作日还是/.test(s)) parts.push({ title: '是否周末', formula: '=IF(WEEKDAY(' + x + ',2)>5,"周末","工作日")', format: 'general', e: 'WEEKDAY 第 2 参数为 2 时周一=1…周日=7，大于 5 即周末' });
    if (/月末|月底|最后一天/.test(s)) parts.push({ title: '当月最后一天', formula: '=EOMONTH(' + x + ',0)', format: 'date', e: 'EOMONTH(日期,0) 返回当月最后一天' });
    if (/月初|第一天|1号/.test(s)) parts.push({ title: '当月第一天', formula: '=EOMONTH(' + x + ',-1)+1', format: 'date', e: '上个月最后一天再加 1 天就是本月第一天' });
    if (/距今|到今天|至今|过去了?多少天|已经过了?|已过/.test(s)) parts.push({ title: '距今天数', formula: '=TODAY()-' + x, format: 'integer', e: 'TODAY() 减去日期得到相差天数，每天打开文件会自动更新' });
    else if (/还剩|剩余|还有多少天|还有几天|倒计时|距离/.test(s)) parts.push({ title: '剩余天数', formula: '=' + x + '-TODAY()', format: 'integer', e: '日期减去今天（TODAY）得到剩余天数，负数表示已过期' });
    if (!parts.length) return null;
    return plan('column', parts.map(p => ({ title: p.title, formula: p.formula, format: p.format })), {
      interpretation: '从【' + d.name + '】提取' + parts.map(p => '“' + p.title + '”').join('、'),
      explanation: parts.map(p => p.e).join('；') + '。',
    });
  }

  // ---------- 文本处理 ----------
  const SEP_WORDS = { 空格: ' ', 横杠: '-', 横线: '-', 短横线: '-', 中划线: '-', 减号: '-', 斜杠: '/', 逗号: ',', 顿号: '、', 下划线: '_', 竖线: '|', 冒号: ':', 点: '.' };
  function tText(q) {
    const s = q.sig;
    const cols = q.activeCols;
    if (!cols.length) return null;
    const x = c => '[@' + c.name + ']';
    const E = i => q.ents[+i];

    if (/合并|拼接|连接|连起来|组合|串起来|合在一起|合成|拼起来|拼在一起/.test(s) && cols.length >= 2) {
      let sep = '';
      const sm = s.match(/(?:用|以|中间加|中间用|中间插入|分隔符(?:为|是|用)?|并用)\s*(?:⟨L(\d+)⟩|(空格|横杠|横线|短横线|中划线|减号|斜杠|逗号|顿号|下划线|竖线|冒号|点)|([-_/|,、:.~ ]))/);
      if (sm) sep = sm[1] !== undefined ? E(sm[1]).text : sm[2] ? SEP_WORDS[sm[2]] : sm[3];
      const parts = cols.map(x);
      const f = '=' + parts.join(sep ? '&' + U.excelString(sep) + '&' : '&');
      return plan('column', [{ title: titleHint(q) || cols.map(c => c.name).join(sep || '') , formula: f, format: 'general' }], {
        interpretation: '把' + cols.map(c => '【' + c.name + '】').join('、') + '连接成一个文本' + (sep ? '，中间用“' + sep + '”分隔' : ''),
        explanation: '& 是文本连接符，可以把多个单元格和文字首尾相连' + (sep ? '，用双引号括起来的“' + sep + '”就是分隔符' : '') + '。',
      });
    }
    const c = cols[0];
    // 手机号 / 号码脱敏
    if (/隐藏|脱敏|打码|星号|遮挡|中间四位/.test(s) && (c.isPhone || /手机|电话/.test(c.name))) {
      return plan('column', [{ title: c.name + '(脱敏)', formula: '=REPLACE(' + x(c) + ',4,4,"****")', format: 'general' }], {
        interpretation: '把【' + c.name + '】中间 4 位替换成星号',
        explanation: 'REPLACE(文本, 开始位置, 替换几位, 新文本)：从第 4 位开始的 4 位换成 ****。',
      });
    }
    let m;
    if ((m = s.match(/第\s*⟨N(\d+)⟩\s*(?:个|位)?(?:字符|字|位)?\s*(?:到|至|-)\s*(?:第)?\s*⟨N(\d+)⟩\s*(?:个|位)/))) {
      const a = E(m[1]).v, b = E(m[2]).v;
      return midPlan(c, a, b - a + 1);
    }
    if ((m = s.match(/(?:从)?第\s*⟨N(\d+)⟩\s*(?:个|位)?(?:字符|字|位)?\s*(?:开始|起)\s*(?:的|取|提取|截取|往后)?\s*⟨N(\d+)⟩\s*(?:个|位)/))) {
      return midPlan(c, E(m[1]).v, E(m[2]).v);
    }
    function midPlan(col, start, len) {
      return plan('column', [{ title: titleHint(q) || col.name + '(第' + start + '-' + (start + len - 1) + '位)', formula: '=MID(' + x(col) + ',' + start + ',' + len + ')', format: 'general' }], {
        interpretation: '从【' + col.name + '】第 ' + start + ' 位开始取 ' + len + ' 位',
        explanation: 'MID(文本, 开始位置, 取几位) 从中间截取字符。',
      });
    }
    if ((m = s.match(/(?:前|左边?|左侧|开头的?)\s*⟨N(\d+)⟩\s*(?:个|位)?\s*(?:字符|字|位|个字|个字符|数字|位数字)?/)) && /字|位|字符|截取|提取|取/.test(s)) {
      const n = E(m[1]).v;
      return plan('column', [{ title: titleHint(q) || c.name + '(前' + n + '位)', formula: '=LEFT(' + x(c) + ',' + n + ')', format: 'general' }], {
        interpretation: '取【' + c.name + '】左边 ' + n + ' 个字符', explanation: 'LEFT(文本, 个数) 从左边开始截取。',
      });
    }
    if ((m = s.match(/(?:后|右边?|右侧|最后|末尾|结尾的?)\s*⟨N(\d+)⟩\s*(?:个|位)?\s*(?:字符|字|位|个字|个字符|数字|位数字)?/)) && /字|位|字符|截取|提取|取/.test(s)) {
      const n = E(m[1]).v;
      return plan('column', [{ title: titleHint(q) || c.name + '(后' + n + '位)', formula: '=RIGHT(' + x(c) + ',' + n + ')', format: 'general' }], {
        interpretation: '取【' + c.name + '】右边 ' + n + ' 个字符', explanation: 'RIGHT(文本, 个数) 从右边开始截取。',
      });
    }
    if ((m = s.match(/⟨L(\d+)⟩\s*(?:之前|前面|以前|左边|前边)/))) {
      const sep = U.excelString(E(m[1]).text);
      return plan('column', [{ title: titleHint(q) || c.name + '(前段)', formula: '=IFERROR(LEFT(' + x(c) + ',FIND(' + sep + ',' + x(c) + ')-1),' + x(c) + ')', format: 'general' }], {
        interpretation: '取【' + c.name + '】中“' + E(m[1]).text + '”前面的部分',
        explanation: 'FIND 找到分隔符的位置，LEFT 取它前面的字符；没有分隔符时 IFERROR 返回原文。',
      });
    }
    if ((m = s.match(/⟨L(\d+)⟩\s*(?:之后|后面|以后|右边|后边)/))) {
      const sepTxt = E(m[1]).text, sep = U.excelString(sepTxt);
      return plan('column', [{ title: titleHint(q) || c.name + '(后段)', formula: '=IFERROR(MID(' + x(c) + ',FIND(' + sep + ',' + x(c) + ')+' + sepTxt.length + ',LEN(' + x(c) + ')),"")', format: 'general' }], {
        interpretation: '取【' + c.name + '】中“' + sepTxt + '”后面的部分',
        explanation: 'FIND 找到分隔符的位置，MID 从它后面开始一直取到末尾。',
      });
    }
    if (/去掉?空格|去除空格|删除空格|清除空格|去空格|删掉空格|多余空格/.test(s)) {
      const all = /所有|全部|中间/.test(s);
      return plan('column', [{ title: c.name + '(去空格)', formula: all ? '=SUBSTITUTE(' + x(c) + '," ","")' : '=TRIM(' + x(c) + ')', format: 'general' }], {
        interpretation: all ? '删除【' + c.name + '】里所有空格' : '去掉【' + c.name + '】首尾和多余的空格',
        explanation: all ? 'SUBSTITUTE 把空格全部替换成空文本。' : 'TRIM 去掉首尾空格，并把中间连续空格变成一个。',
      });
    }
    if ((m = s.match(/(?:把|将)?\s*(?:⟨([LVN])(\d+)⟩|([^\s⟨⟩,]{1,10}?))\s*(?:替换为|替换成|换成|改为|改成|变成)\s*(?:⟨([LVN])(\d+)⟩|([^\s⟨⟩,]{0,10}))/)) && /替换|换成|改为|改成|变成/.test(s)) {
      const tok = (t, i, raw) => (t ? (E(i).t === 'N' ? E(i).raw : (E(i).text || E(i).value)) : raw);
      const from = tok(m[1], m[2], m[3]), to = tok(m[4], m[5], m[6]) || '';
      if (from && !/^[一-龥]*列$/.test(from)) {
        return plan('column', [{ title: c.name + '(替换后)', formula: '=SUBSTITUTE(' + x(c) + ',' + U.excelString(from) + ',' + U.excelString(to) + ')', format: 'general' }], {
          interpretation: '把【' + c.name + '】中的“' + from + '”替换为“' + to + '”',
          explanation: 'SUBSTITUTE(文本, 旧文本, 新文本) 把所有出现的旧文本换成新文本。',
        });
      }
    }
    if ((m = s.match(/(?:删除|删掉|去掉|去除|移除)\s*⟨L(\d+)⟩/))) {
      const from = E(m[1]).text;
      return plan('column', [{ title: c.name + '(删除后)', formula: '=SUBSTITUTE(' + x(c) + ',' + U.excelString(from) + ',"")', format: 'general' }], {
        interpretation: '删除【' + c.name + '】中的“' + from + '”', explanation: 'SUBSTITUTE 把指定文字替换成空文本，相当于删除。',
      });
    }
    if (/长度|字数|字符数|几个字|多少个字|多少位|几位数/.test(s)) {
      return plan('column', [{ title: c.name + '长度', formula: '=LEN(' + x(c) + ')', format: 'integer' }], {
        interpretation: '计算【' + c.name + '】的字符个数', explanation: 'LEN 返回文本的字符数，汉字、字母、数字都算 1 个。',
      });
    }
    if (/首字母大写/.test(s)) return plan('column', [{ title: c.name + '(首字母大写)', formula: '=PROPER(' + x(c) + ')', format: 'general' }], { interpretation: '【' + c.name + '】每个单词首字母大写', explanation: 'PROPER 把每个英文单词的首字母变大写。' });
    if (/大写/.test(s)) return plan('column', [{ title: c.name + '(大写)', formula: '=UPPER(' + x(c) + ')', format: 'general' }], { interpretation: '【' + c.name + '】转成大写字母', explanation: 'UPPER 把英文字母全部转成大写。' });
    if (/小写/.test(s)) return plan('column', [{ title: c.name + '(小写)', formula: '=LOWER(' + x(c) + ')', format: 'general' }], { interpretation: '【' + c.name + '】转成小写字母', explanation: 'LOWER 把英文字母全部转成小写。' });
    if ((m = s.match(/(?:补零|补0|补足|补齐|凑够|不足)\s*(?:到|成)?\s*⟨N(\d+)⟩\s*位/)) || (m = s.match(/⟨N(\d+)⟩\s*位\s*(?:不足)?(?:前面)?(?:补零|补0)/))) {
      const n = E(m[1]).v;
      return plan('column', [{ title: c.name + '(' + n + '位)', formula: '=TEXT(' + x(c) + ',"' + '0'.repeat(n) + '")', format: 'general' }], {
        interpretation: '把【' + c.name + '】补足 ' + n + ' 位，不足的前面补 0', explanation: 'TEXT 按格式“' + '0'.repeat(n) + '”输出，位数不够时左边自动补 0。',
      });
    }
    if (/转(?:为|成|换为|换成)?数字|转数值|文本转数字|变成数字|转换成数字|转成数值/.test(s)) {
      return plan('column', [{ title: c.name + '(数值)', formula: '=VALUE(' + x(c) + ')', format: 'general' }], {
        interpretation: '把【' + c.name + '】从文本转换成数字', explanation: 'VALUE 把“看起来像数字的文本”转换成真正的数字，转换后才能正常求和。',
      });
    }
    return null;
  }

  // ---------- 四舍五入 / 取整 ----------
  function tRound(q) {
    const s = q.sig;
    if (!/保留\s*⟨N\d+⟩\s*位|四舍五入|取整|进位|舍去小数|去掉小数|保留整数|整数部分|抹零/.test(s)) return null;
    const c = numCols(q)[0];
    if (!c) return null;
    const m = s.match(/保留\s*⟨N(\d+)⟩\s*位/) || s.match(/(?:到|精确到)\s*⟨N(\d+)⟩\s*位/);
    const n = m ? q.ents[+m[1]].v : 0;
    const x = '[@' + c.name + ']';
    let f, title, expl;
    if (/向上|进位|进一/.test(s)) { f = '=ROUNDUP(' + x + ',' + n + ')'; title = c.name + '(向上取整)'; expl = 'ROUNDUP 不管尾数大小都往上进位'; }
    else if (/向下|舍去|去掉小数|整数部分|抹零|截断/.test(s)) { f = '=ROUNDDOWN(' + x + ',' + n + ')'; title = c.name + '(向下取整)'; expl = 'ROUNDDOWN 直接舍去多余的小数位'; }
    else { f = '=ROUND(' + x + ',' + n + ')'; title = c.name + (n ? '(保留' + n + '位)' : '(取整)'); expl = 'ROUND 四舍五入'; }
    if (n) expl += '，第 2 个参数 ' + n + ' 表示保留 ' + n + ' 位小数'; else expl += '，第 2 个参数 0 表示取到整数';
    return plan('column', [{ title, formula: f, format: 'general' }], { interpretation: '对【' + c.name + '】' + title.replace(c.name, '').replace(/[()]/g, ''), explanation: expl + '。注意：结果是真正改变了数值，不只是显示格式。' });
  }

  // ---------- 累计 ----------
  function tCumulative(q) {
    const s = q.sig;
    if (!/累计|累加|累积|滚动求和|滚动合计|逐行累|截至当前|截止到当前|到当前行/.test(s)) return null;
    if (/是多少|一共|总共|合计是|多少钱/.test(s)) return null;
    const c = numCols(q).find(x => x.type === 'number');
    if (!c) return null;
    const groups = findGroups(q).filter(g => g.unit === 'value' && g.col !== c);
    if (groups.length) {
      const g = groups[0].col;
      return plan('column', [{ title: g.name + '内累计' + c.name, formula: '=SUMIF([^' + g.name + '],[@' + g.name + '],[^' + c.name + '])', format: 'general' }], {
        interpretation: '按【' + g.name + '】分别逐行累计【' + c.name + '】',
        explanation: 'SUMIF 的两个区域都从第一行“滚动”到当前行，只把与本行同一' + g.name + '的数加起来，就是该' + g.name + '截至本行的累计。',
      });
    }
    return plan('column', [{ title: '累计' + c.name, formula: '=SUM([^' + c.name + '])', format: 'general' }], {
      interpretation: '逐行累计【' + c.name + '】',
      explanation: 'SUM 的区域起点用 $ 锁定在第一行、终点是当前行，公式往下填充时区域逐行变长，就得到累计值。',
    });
  }

  // ---------- 占比 ----------
  function tRatio(q) {
    const s = q.sig;
    if (!/占比|比例|百分比|比重|份额|占总|占全部|占合计|占所在|占所属|占本/.test(s)) return null;
    const nums = numCols(q).filter(c => c.type === 'number');
    if (!nums.length) return null;
    const groups = findGroups(q);
    const catGroup = groups.find(g => g.unit !== 'value' || (g.col.distinctCount < (q.sheet.dataEnd - q.sheet.dataStart + 1) * 0.7));
    if (catGroup && !/所在|所属|本|同/.test(s)) return null; // “各部门占比”交给分组汇总
    let m = s.match(/⟨C(\d+)⟩\s*(?:的)?\s*占\s*⟨C(\d+)⟩/);
    if (m && q.ents[+m[1]].col.type === 'number' && q.ents[+m[2]].col.type === 'number') {
      const a = q.ents[+m[1]].col, b = q.ents[+m[2]].col;
      return plan('column', [{ title: titleHint(q) || a.name + '占' + b.name + '比例', formula: '=[@' + a.name + ']/[@' + b.name + ']', format: 'percent' }], {
        interpretation: '每行计算【' + a.name + '】÷【' + b.name + '】', explanation: '两个数直接相除，结果设置为百分比格式。',
      });
    }
    const c = nums[0];
    m = s.match(/(?:所在|所属|本|同一?|该)\s*⟨C(\d+)⟩/);
    if (m) {
      const g = q.ents[+m[1]].col;
      return plan('column', [{ title: '占' + g.name + '比例', formula: '=[@' + c.name + ']/SUMIF([' + g.name + '],[@' + g.name + '],[' + c.name + '])', format: 'percent' }], {
        interpretation: '每行【' + c.name + '】占所在【' + g.name + '】合计的比例',
        explanation: 'SUMIF 先算出本行所在' + g.name + '的' + c.name + '合计，再用本行的数去除。结果为百分比格式。',
      });
    }
    return plan('column', [{ title: c.name + '占比', formula: '=[@' + c.name + ']/SUM([' + c.name + '])', format: 'percent' }], {
      interpretation: '每行【' + c.name + '】占整列合计的比例',
      explanation: '本行的数 ÷ SUM(整列)，整列区域用 $ 锁定，往下填充时分母不变。结果为百分比格式。',
    });
  }

  // ---------- 行内多列汇总：语文+数学+英语 ----------
  const SCORE_RE = /语文|数学|英语|物理|化学|生物|历史|地理|政治|道德|科学|体育|音乐|美术|信息|成绩|分数|得分/;
  function tRowMulti(q, aggs) {
    if (/分别|各自|每列|每一列|各列/.test(q.sig)) return null;
    let nums = numCols(q).filter(c => c.type === 'number');
    const scoreWord = q.text.match(/总分|平均分|最高分|最低分/);
    if (nums.length === 0 && scoreWord) nums = q.sheet.columns.filter(c => c.type === 'number' && !c.isId && SCORE_RE.test(c.name));
    if (nums.length < 2) return null;
    const groups = findGroups(q).filter(g => g.col.distinctCount < (q.sheet.dataEnd - q.sheet.dataStart + 1) * 0.9 || g.unit !== 'value');
    if (groups.length && !/每行|每一行|每人|每个人|每位|每名|每条|逐行|每个学生|每个员工|每个同学/.test(q.text)) return null;
    const score = !!scoreWord || nums.every(c => SCORE_RE.test(c.name));
    if (!score && !/每行|每一行|每人|每个人|每位|每名|每条|逐行|每个学生|每个员工|每个同学|相加|加起来|之和/.test(q.text)) return null;
    let a = aggs.find(x => ['sum', 'avg', 'max', 'min'].includes(x.type));
    if (!a && scoreWord) a = { type: { 总分: 'sum', 平均分: 'avg', 最高分: 'max', 最低分: 'min' }[scoreWord[0]] };
    if (!a && /相加|加起来|之和/.test(q.text)) a = { type: 'sum' };
    if (!a) return null;
    const fn = { sum: 'SUM', avg: 'AVERAGE', max: 'MAX', min: 'MIN' }[a.type];
    const title = titleHint(q) || ({ sum: score ? '总分' : '合计', avg: score ? '平均分' : '平均值', max: score ? '最高分' : '最大值', min: score ? '最低分' : '最小值' })[a.type];
    const f = '=' + fn + '(' + nums.map(c => '[@' + c.name + ']').join(',') + ')';
    return plan('column', [{ title, formula: f, format: a.type === 'avg' ? 'general' : 'general' }], {
      interpretation: '每一行对' + nums.map(c => '【' + c.name + '】').join('、') + '求' + AGG_NAMES[a.type],
      explanation: fn + ' 可以同时接收多个单元格，逗号分隔，结果是本行这几列的' + AGG_NAMES[a.type] + '。',
    });
  }

  // ---------- 行内显示所在分组的汇总值 ----------
  function tRowGroupAgg(q, aggs) {
    const m = q.sig.match(/(?<!不)(?:所在|所属|同一|相同|同|本|对应|该)\s*(?:的)?\s*⟨C(\d+)⟩/);
    if (!m || !aggs.length) return null;
    const g = q.ents[+m[1]].col;
    const a = aggs[0];
    const target = numCols(q).find(c => c !== g && c.type === 'number');
    let f, title;
    if (a.type === 'count') { f = '=COUNTIF([' + g.name + '],[@' + g.name + '])'; title = g.name + '人数'; }
    else if (!target) return null;
    else if (a.type === 'sum') { f = '=SUMIF([' + g.name + '],[@' + g.name + '],[' + target.name + '])'; title = g.name + target.name + '合计'; }
    else if (a.type === 'avg') { f = '=AVERAGEIF([' + g.name + '],[@' + g.name + '],[' + target.name + '])'; title = g.name + '平均' + target.name; }
    else if (a.type === 'max') { f = '=AGGREGATE(14,6,[' + target.name + ']/([' + g.name + ']=[@' + g.name + ']),1)'; title = g.name + '最高' + target.name; }
    else if (a.type === 'min') { f = '=AGGREGATE(15,6,[' + target.name + ']/([' + g.name + ']=[@' + g.name + ']),1)'; title = g.name + '最低' + target.name; }
    else return null;
    if (/差|差距|相差|高出|低于/.test(q.sig) && target && a.type !== 'count') {
      f = '=[@' + target.name + ']-' + f.slice(1);
      title = '与' + title + '的差';
    }
    return plan('column', [{ title: titleHint(q) || title, formula: f, format: 'general' }], {
      interpretation: '每行显示本行所在【' + g.name + '】的' + (target && a.type !== 'count' ? '【' + target.name + '】' : '') + AGG_NAMES[a.type] + (/与/.test(title) ? '，并计算本行与它的差' : ''),
      explanation: '条件用“本行的' + g.name + '”（[@' + g.name + ']），这样每一行算出来的都是自己所在' + g.name + '的统计值。' +
        (/AGGREGATE/.test(f) ? 'AGGREGATE(14/15,6,…) 是 Excel 2016 里做“带条件最大/最小值”的写法，外面的 INDEX(…,0) 让它按数组计算，不需要按 Ctrl+Shift+Enter。' : ''),
    });
  }

  // ---------- 最高的是谁 / 第 N 大 ----------
  function tExtreme(q, conds) {
    const s = q.sig;
    let kth = null, dir = null;
    let m = s.match(/第\s*⟨N(\d+)⟩\s*(大|高|多|名)/);
    if (m) { kth = q.ents[+m[1]].v; dir = 'max'; }
    m = s.match(/第\s*⟨N(\d+)⟩\s*(小|低|少)/);
    if (m) { kth = q.ents[+m[1]].v; dir = 'min'; }
    if (!dir) {
      if (/最高|最大|最多|最好|第一名|冠军|最贵/.test(s)) dir = 'max';
      else if (/最低|最小|最少|最差|最后一名|倒数第一|最便宜/.test(s)) dir = 'min';
    }
    if (!dir) return null;
    const who = /谁|哪个|哪位|哪一|哪家|哪天|哪一天|什么|名字|姓名|是哪|对应的|哪款|哪种/.test(s);
    const target = numCols(q).find(c => c.type === 'number' || c.type === 'date');
    if (!target) return null;
    if (conds.length) return null;
    const R = mkRef(q);
    const T = R.R(target);
    const val = kth ? (dir === 'max' ? 'LARGE(' : 'SMALL(') + T + ',' + kth + ')' : (dir === 'max' ? 'MAX(' : 'MIN(') + T + ')';
    const word = kth ? '第' + kth + (dir === 'max' ? '大' : '小') : (dir === 'max' ? '最高' : '最低');
    if (!who) {
      if (!kth) return null; // 普通的最大最小交给汇总模板
      return plan('scalar', [{ title: target.name + word + '值', formula: '=' + val, format: target.type === 'date' ? 'date' : 'general' }], {
        interpretation: '【' + target.name + '】中' + word + '的值', explanation: (dir === 'max' ? 'LARGE' : 'SMALL') + '(区域, k) 返回第 k ' + (dir === 'max' ? '大' : '小') + '的值。',
      });
    }
    const idCol = q.activeCols.find(c => c !== target && c.type !== 'number') || q.sheet.columns.find(c => c.isName) || q.sheet.columns.find(c => c.type === 'text');
    if (!idCol) return null;
    const f = '=INDEX(' + R.R(idCol) + ',MATCH(' + val + ',' + T + ',0))';
    return plan('scalar', [{ title: target.name + word + '的' + idCol.name, formula: f, format: idCol.type === 'date' ? 'date' : 'general' }], {
      interpretation: '找出【' + target.name + '】' + word + '的那一行，返回它的【' + idCol.name + '】',
      explanation: (kth ? (dir === 'max' ? 'LARGE' : 'SMALL') + ' 求出第 ' + kth + ' 名的数值' : (dir === 'max' ? 'MAX' : 'MIN') + ' 求出' + word + '值') +
        '，MATCH 找到它在第几行，INDEX 再取出同一行的' + idCol.name + '。若有并列，只返回第一个。',
    });
  }

  // ---------- 单值查找：张三的工资 ----------
  function tValueLookup(q, conds, aggs) {
    if (aggs.length) return null;
    if (conds.length !== 1 || conds[0].kind !== 'cmp' || conds[0].op !== '=') return null;
    const c = conds[0];
    if (typeof c.value !== 'string') return null;
    const count = c.col.distinct ? c.col.distinct.get(c.value) : null;
    if (count !== 1) return null;
    const target = q.activeCols.find(x => x !== c.col);
    if (!target) return null;
    const R = mkRef(q);
    let f;
    if (target.index > c.col.index) f = '=VLOOKUP(' + U.excelString(c.value) + ',[' + c.col.name + ':' + target.name + '],' + (target.index - c.col.index + 1) + ',FALSE)';
    else f = '=INDEX(' + R.R(target) + ',MATCH(' + U.excelString(c.value) + ',' + R.R(c.col) + ',0))';
    return plan('scalar', [{ title: c.value + '的' + target.name, formula: f, format: target.type === 'date' ? 'date' : 'general' }], {
      interpretation: '在【' + c.col.name + '】中找到“' + c.value + '”，返回同一行的【' + target.name + '】',
      explanation: /VLOOKUP/.test(f) ? 'VLOOKUP(要找的值, 查找区域, 返回第几列, FALSE)，FALSE 表示精确匹配。' : 'MATCH 找到它在第几行，INDEX 取出同一行另一列的值。',
    });
  }

  // ---------- 汇总（单值 / 分组） ----------
  function aggTitle(type, target, word) {
    if (type === 'count') {
      const w = (word || '').match(/订单数|客户数|员工数|学生数|人次|人数|条数|笔数|单数|次数|行数|家数|记录数/);
      return w ? w[0] : '数量';
    }
    return (target ? target.name : '') + AGG_NAMES[type];
  }

  function pairMetrics(q, aggs, targets, conds) {
    const metrics = [];
    if (!aggs.length) return metrics;
    if (aggs.length === 1) {
      const a = aggs[0];
      if (a.type === 'count' || !targets.length) metrics.push({ type: a.type, target: a.type === 'count' ? null : targets[0] || null, word: a.word });
      else targets.forEach(t => metrics.push({ type: a.type, target: t, word: a.word }));
      return metrics;
    }
    for (const a of aggs) {
      if (a.type === 'count') { metrics.push({ type: 'count', target: null, word: a.word }); continue; }
      let best = null, bestD = Infinity;
      for (const t of targets) {
        const ent = q.colEnts.find(e => e.col === t);
        const d = ent ? Math.abs(ent.pos - a.pos) : 999;
        if (d < bestD) { bestD = d; best = t; }
      }
      metrics.push({ type: a.type, target: best || targets[0] || null, word: a.word });
    }
    return metrics;
  }

  function tAggregate(q, condRes, aggs) {
    const conds = condRes.conds;
    const groups = findGroups(q);
    const groupCols = new Set(groups.map(g => g.col));
    const consumedCols = new Set();
    q.ents.forEach(e => { if (e.t === 'C' && condRes.consumed.has(e.idx)) consumedCols.add(e.col); });
    // “华东区域的销售额”里的“区域”只是在说明条件，不是统计对象
    conds.forEach(c => consumedCols.add(c.col));
    // 目标列：被提到、不是分组列、不是只在条件里出现的列
    let targets = [];
    for (const e of q.colEnts) {
      if (groupCols.has(e.col)) continue;
      if (consumedCols.has(e.col)) continue;
      if (!targets.includes(e.col)) targets.push(e.col);
    }
    // 跨表：所有提到的列都在同一张别的表里
    const sheetsUsed = new Set(q.colEnts.map(e => e.sheet));
    if (sheetsUsed.size === 1 && !sheetsUsed.has(q.sheet)) q.planSheet = Array.from(sheetsUsed)[0];
    else if (q.colEnts.some(e => e.sheet !== q.sheet)) return null;

    const numericTargets = targets.filter(c => c.type === 'number' || c.type === 'date');
    let useAggs = aggs.slice();
    const ratio = /占比|比例|百分比|比重|份额/.test(q.sig);
    // 没写汇总方式：有分组时默认求和；“销售部的销售额（是多少）”默认求和
    if (!useAggs.length) {
      if (groups.length && numericTargets.length) useAggs = [{ type: 'sum', pos: 0, word: '' }];
      else if (ratio && groups.length && numericTargets.length) useAggs = [{ type: 'sum', pos: 0, word: '' }];
      else if (conds.length && numericTargets.length && /多少|是多少|总|共/.test(q.sig)) useAggs = [{ type: 'sum', pos: 0, word: '' }];
      else if (conds.length && numericTargets.length === 1 && targets.length === 1) useAggs = [{ type: 'sum', pos: 0, word: '' }];
      else return null;
    }
    // “总数”：有数值目标时是求和，否则是计数
    useAggs = useAggs.map(a => (a.type === 'sum' && /总数|一共|总共/.test(a.word) && !numericTargets.length) ? Object.assign({}, a, { type: 'count' }) : a);
    // 数值类汇总只针对数值列；若目标列都不是数值且只有一个条件列是数值，则用它
    let aggTargets = numericTargets;
    if (!aggTargets.length && useAggs.some(a => !['count', 'distinct'].includes(a.type))) {
      const condNum = conds.map(c => c.col).find(c => c.type === 'number');
      if (condNum) aggTargets = [condNum];
    }
    if (useAggs.some(a => a.type === 'distinct')) aggTargets = targets.length ? targets : aggTargets;
    const metrics = [];
    for (const m of pairMetrics(q, useAggs, aggTargets, conds)) {
      if (!metrics.some(x => x.type === m.type && x.target === m.target)) metrics.push(m);
    }
    if (!metrics.length) return null;
    if (metrics.some(m => !['count'].includes(m.type) && !m.target)) return null;

    const R = mkRef(q).R;
    const condText = conds.map(describeCond).join(condRes.orMode ? ' 或 ' : ' 且 ');
    if (condRes.orMode && conds.length > 1 && !conds.every(c => c.kind === 'or' || c.col === conds[0].col)) return null;

    if (groups.length) {
      const gs = groups.slice(0, 2);
      const key = { pairs: [], atoms: [] };
      gs.forEach((g, i) => {
        const K = '{KEY' + (i + 1) + '}';
        const r = R(g.col);
        if (g.unit === 'value') { key.pairs.push([r, K]); key.atoms.push(r + '=' + K); }
        else {
          const months = g.unit === 'month' ? 1 : 12;
          key.pairs.push([r, '">="&' + K], [r, '"<"&EDATE(' + K + ',' + months + ')']);
          key.atoms.push(r + '>=' + K, r + '<EDATE(' + K + ',' + months + ')');
        }
      });
      const outs = [];
      for (const m of metrics) {
        const f = aggFormula(q, m.type, m.target, conds, key, R);
        if (!f) return null;
        outs.push({ title: aggTitle(m.type, m.target, m.word), formula: f, format: m.type === 'count' || m.type === 'distinct' ? 'integer' : (m.target && m.target.type === 'date' ? 'date' : 'general') });
      }
      if (ratio) {
        const t = metrics.find(m => m.target && m.target.type === 'number');
        if (t) {
          const sumF = aggFormula(q, 'sum', t.target, conds, key, R);
          const totF = aggFormula(q, 'sum', t.target, conds, null, R);
          outs.push({ title: t.target.name + '占比', formula: '=' + sumF.slice(1) + '/' + totF.slice(1), format: 'percent' });
        }
      }
      const gText = gs.map(g => '【' + g.col.name + '】' + (g.unit === 'month' ? '（按月）' : g.unit === 'year' ? '（按年）' : '')).join('、');
      const used = outs.map(o => o.formula).join(' ');
      let expl = '新建一张汇总表，' + (gs.length > 1 ? '前几列' : '第一列') + '列出每个不同的' + gs.map(g => g.unit === 'month' ? '月份' : g.unit === 'year' ? '年份' : g.col.name).join('、') +
        '，旁边的公式用本行的分组值（如 $A2）作为条件，往下每一行各算各的组。';
      if (/SUMIFS?\(/.test(used)) expl += 'SUMIF/SUMIFS 只把满足条件的行加起来；';
      if (/COUNTIFS?\(/.test(used)) expl += 'COUNTIF/COUNTIFS 统计满足条件的行数；';
      if (/AVERAGEIFS?\(/.test(used)) expl += 'AVERAGEIF/AVERAGEIFS 只对满足条件的行求平均；';
      if (/AGGREGATE/.test(used)) expl += 'AGGREGATE(14/15,6,INDEX(区域/(条件),0),1) 是 Excel 2016 求“带条件最大/最小值”的写法，不满足条件的行会变成错误值并被忽略；';
      if (/EDATE/.test(used)) expl += '按月/年分组时，条件写成“日期 ≥ 该月第一天 且 < 下个月第一天”（EDATE 推算下个月）；';
      return plan('group', outs, {
        groupBy: gs.map(g => ({ column: g.col.name, unit: g.unit })),
        sheet: (q.planSheet || q.sheet).name,
        interpretation: '按' + gText + '分组，统计' + metrics.map(m => (m.target ? '【' + m.target.name + '】的' : '') + AGG_NAMES[m.type]).join('、') + (ratio ? '及占比' : '') + (conds.length ? '；条件：' + condText : ''),
        explanation: expl.replace(/；$/, '。'),
      });
    }

    const outs = [];
    for (const m of metrics) {
      const f = aggFormula(q, m.type, m.target, conds, null, R);
      if (!f) return null;
      let title = aggTitle(m.type, m.target, m.word);
      if (m.type === 'count' && !m.target && !conds.length) title = '总行数';
      if (conds.length) title += '（' + conds.map(c => describeCond(c).replace(/[【】“”]/g, '')).join(condRes.orMode ? '或' : '，') + '）';
      outs.push({ title, formula: f, format: m.type === 'count' || m.type === 'distinct' ? 'integer' : (m.target && m.target.type === 'date' ? 'date' : 'general') });
    }
    const used = outs.map(o => o.formula).join(' ');
    let expl = '';
    if (/SUMIFS\(/.test(used)) expl += 'SUMIFS(求和区域, 条件区域1, 条件1, 条件区域2, 条件2…) 只把同时满足所有条件的行加起来。';
    else if (/SUMIF\(/.test(used)) expl += 'SUMIF(条件区域, 条件, 求和区域) 只把满足条件的行加起来。';
    if (/COUNTIFS\(/.test(used)) expl += 'COUNTIFS 统计同时满足多个条件的行数。';
    else if (/COUNTIF\(/.test(used)) expl += 'COUNTIF(区域, 条件) 统计满足条件的单元格个数。';
    if (/AVERAGEIFS\(/.test(used)) expl += 'AVERAGEIFS 对同时满足条件的行求平均。';
    else if (/AVERAGEIF\(/.test(used)) expl += 'AVERAGEIF(条件区域, 条件, 求平均区域) 对满足条件的行求平均。';
    if (/AGGREGATE/.test(used)) expl += 'Excel 2016 没有 MAXIFS/MINIFS，用 AGGREGATE(14或15, 6, INDEX(区域/(条件),0), 1) 代替：不满足条件的行除以 0 变成错误值，第 2 参数 6 表示忽略错误值，INDEX(…,0) 让它按数组计算。';
    if (/SUMPRODUCT\(\(/.test(used) && !/COUNTIF\(/.test(used)) expl += 'SUMPRODUCT 把每个条件算成 1/0 再相乘求和，可以表达 SUMIFS 写不了的条件（如按月份）。';
    if (/COUNTIF\([^)]*&""\)/.test(used)) expl += '不重复计数：COUNTIF 算出每个值出现的次数，取倒数后相加，每个不同的值正好贡献 1。';
    if (/COUNTA\(/.test(used)) expl += 'COUNTA 统计非空单元格个数，用来数数据有多少行。';
    if (!expl) expl = '直接对整列使用 ' + Array.from(new Set(used.match(/[A-Z]+(?=\()/g) || [])).join('、') + ' 函数。';
    if (/"\*/.test(used)) expl += '条件里的 * 是通配符，代表任意多个字符。';
    if (/&DATE\(/.test(used)) expl += '日期条件写成 ">="&DATE(年,月,日) 的形式。';
    return plan('scalar', outs, {
      sheet: (q.planSheet || q.sheet).name,
      interpretation: metrics.map(m => (m.target ? '对【' + m.target.name + '】求' : '统计') + AGG_NAMES[m.type]).join('，') + (conds.length ? '；条件：' + condText : ''),
      explanation: expl,
    });
  }

  /* ======================= 入口 ======================= */

  function match(request, wb, sheet) {
    const q = parse(request, wb, sheet);
    const aggs = findAggs(q);
    const steps = [
      () => tIdCard(q),
      () => tLookup(q),
      () => tRank(q),
      () => tDuplicate(q),
      () => tTiers(q, aggs),
      () => tExpr(q),
      () => tDate(q, aggs),
      () => tCumulative(q),
      () => tRatio(q),
      () => tRound(q),
      () => tText(q),
      () => tRowGroupAgg(q, aggs),
      () => tRowMulti(q, aggs),
    ];
    let result = null;
    for (const s of steps) {
      try { result = s(); } catch (e) { console.warn('模板出错', e); result = null; }
      if (result) break;
    }
    if (!result) {
      const condRes = extractConditions(q, q.sig, { defaultCol: numCols(q)[0] || null });
      result = tExtreme(q, condRes.conds) || tValueLookup(q, condRes.conds, aggs) || tAggregate(q, condRes, aggs);
    }
    if (result) {
      result.sheet = result.sheet || (q.planSheet || q.sheet).name;
      result.request = request;
    }
    return { plan: result, parsed: q };
  }

  /** 根据当前表的列生成示例需求 */
  function examples(sheet, wb) {
    const cols = sheet.columns.filter(c => !c.added);
    const num = cols.filter(c => c.type === 'number' && !c.isId);
    const CAT_RE = /部门|区域|地区|类别|分类|类型|产品|品类|班级|门店|渠道|城市|省份|性别|状态|级别/;
    const cats = cols.filter(c => c.isCategory && !c.isName && !c.isId && c.distinctCount >= 2);
    const cat = cats.find(c => CAT_RE.test(c.name)) || cats.sort((a, b) => a.distinctCount - b.distinctCount)[0];
    const date = cols.find(c => c.type === 'date' && !c.isBirth && !c.isHire);
    const people = cols.some(c => /^姓名$|^名字$/.test(c.name));
    const id = cols.find(c => c.isIdCard);
    const out = [];
    const n0 = num.find(c => /额|金额|收入|工资|薪|成绩|分数|利润|成本|费用/.test(c.name)) || num[num.length - 1];
    if (cat && n0) out.push('按' + cat.name + '汇总' + n0.name + '的合计');
    if (n0 && n0.numStats) {
      const t = n0.numStats.median;
      const nice = t >= 100 ? Math.round(t / Math.pow(10, Math.floor(Math.log10(t)) - 1)) * Math.pow(10, Math.floor(Math.log10(t)) - 1) : Math.round(t);
      out.push(n0.name + '大于' + nice + '的有多少' + (people ? '人' : '条'));
    }
    if (id) out.push('根据' + id.name + '计算年龄');
    if (num.length >= 2 && /单价|价格/.test(num.map(c => c.name).join()) && /数量|件数/.test(num.map(c => c.name).join())) {
      const p = num.find(c => /单价|价格/.test(c.name)), qn = num.find(c => /数量|件数/.test(c.name));
      out.push(p.name + '乘以' + qn.name);
    }
    if (n0) out.push('按' + n0.name + '从高到低排名');
    if (date && n0) out.push('按月统计' + n0.name + '合计');
    if (num.length >= 3 && num.every(c => /语文|数学|英语|物理|化学|成绩|分/.test(c.name))) out.unshift('计算' + num.slice(0, 3).map(c => c.name).join('、') + '的总分');
    if (n0 && num.length < 3) out.push(n0.name + '大于等于' + (n0.numStats ? Math.round(n0.numStats.median) : 60) + '标记为达标，否则不达标');
    // 查找示例：对方表里该编号列必须唯一（一对一）
    const other = wb.sheets.find(s => s !== sheet && s.columns.some(c => c.distinctCount === c.nonEmpty && cols.some(x => x.name === c.name && x.isId)));
    if (other) {
      const key = cols.find(x => x.isId && other.columns.some(c => c.name === x.name));
      const rc = other.columns.filter(c => !cols.some(x => x.name === c.name) && c.type !== 'empty' && !c.isIdCard && !c.isPhone);
      const ret = rc.find(c => CAT_RE.test(c.name)) || rc[0];
      if (key && ret) out.push('根据' + key.name + '从' + other.name + '查找' + ret.name);
    }
    if (cat && !n0) out.push('统计每个' + cat.name + '有多少行');
    return Array.from(new Set(out)).slice(0, 6);
  }

  EFA.Templates = { match, parse, examples, extractConditions, describeCond, normalize };
})(window.EFA = window.EFA || {});
