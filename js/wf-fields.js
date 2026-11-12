/* 工作流配置器：步骤参数编辑控件。每种字段类型一个渲染函数，直接读写步骤对象。 */
(function (EFA) {
  'use strict';
  const M = EFA.WFModel, F = EFA.WFFiles;
  const W = {};
  let uid = 0;

  /* ------------------------------------------------------------------ DOM 小工具 */
  function h(tag, props, ...kids) {
    const e = document.createElement(tag);
    if (props) {
      for (const k of Object.keys(props)) {
        const v = props[k];
        if (v === undefined || v === null || v === false) continue;
        if (k === 'class') e.className = v;
        else if (k === 'text') e.textContent = v;
        else if (k === 'html') e.innerHTML = v;
        else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
        else if (k === 'value') e.value = v;
        else if (k === 'checked') e.checked = !!v;
        else e.setAttribute(k, v === true ? '' : v);
      }
    }
    for (const c of kids.flat()) if (c !== null && c !== undefined && c !== false) e.append(c instanceof Node ? c : document.createTextNode(String(c)));
    return e;
  }
  W.h = h;

  function select(options, value, onChange, cls) {
    const s = h('select', { class: cls || 'in', onchange: () => onChange(options[s.selectedIndex].value) });
    options.forEach(o => s.append(h('option', { text: o.label })));
    const i = options.findIndex(o => o.value === value);
    s.selectedIndex = i >= 0 ? i : 0;
    return s;
  }
  W.select = select;

  /** 列名输入（带候选列表）。source: 'main' | 'ref' | 'refFile' | 'ck' */
  function colInput(ctx, value, onChange, opt) {
    opt = opt || {};
    const listId = ctx.datalist(opt.source || 'main');
    const inp = h('input', { class: 'in col-in', list: listId, value: value === undefined || value === null ? '' : M.colText(value), placeholder: opt.placeholder || '列名或列字母', spellcheck: 'false' });
    const mark = () => {
      const hdr = ctx.headersOf(opt.source || 'main');
      const v = inp.value.trim();
      const r = v ? M.resolveCol(hdr, v) : {};
      inp.classList.toggle('bad', !!r.error);
      inp.classList.toggle('by-letter', r.how === 'letter' && !!hdr);
      inp.title = r.error ? r.error + (hdr ? '。现有列：' + hdr.filter(Boolean).slice(0, 30).join('、') : '') : (r.how === 'letter' ? '按列字母定位（表头里没有这个名字）' : '');
    };
    inp.addEventListener('input', () => { onChange(inp.value.trim() === '' ? undefined : inp.value.trim()); mark(); });
    mark();
    ctx.onRefresh(mark);
    return inp;
  }
  W.colInput = colInput;

  /** 多列选择：标签 + 输入框，可输入区间（如 U:AB、开始列:结束列） */
  function colsInput(ctx, value, onChange, opt) {
    opt = opt || {};
    const box = h('div', { class: 'chips-in' });
    let arr = value === undefined ? [] : (Array.isArray(value) ? value.slice() : [value]);
    const listId = ctx.datalist(opt.source || 'main');
    const tags = h('span', { style: 'display:contents' });
    const inp = h('input', { class: 'tag-in', list: listId, spellcheck: 'false', 'aria-label': opt.placeholder || '添加列' });
    const draw = () => {
      tags.textContent = '';
      const hdr = ctx.headersOf(opt.source || 'main');
      arr.forEach((c, i) => {
        const txt = M.colText(c);
        const r = M.resolveCols(hdr, typeof c === 'string' && c.includes(':') ? c : [c]);
        tags.append(h('span', { class: 'tag' + (r.errors.length ? ' bad' : ''), title: r.errors.join('；') },
          txt, h('button', { type: 'button', class: 'x', title: '移除', 'aria-label': '移除 ' + txt, onclick: () => { arr.splice(i, 1); commit(); } }, '×')));
      });
      inp.placeholder = arr.length ? '继续添加…' : (opt.placeholder || '选择或输入列名，回车添加');
    };
    const add = () => {
      const v = inp.value.trim();
      if (!v) return;
      v.split(/[,，、]/).map(x => x.trim()).filter(Boolean).forEach(x => { if (!arr.some(a => M.colText(a) === x)) arr.push(x); });
      inp.value = '';
      commit(true);
    };
    inp.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); add(); }
      else if (e.key === 'Backspace' && !inp.value && arr.length) { arr.pop(); commit(true); }
    });
    inp.addEventListener('change', add);
    const commit = (focus) => {
      onChange(arr.length ? arr.slice() : undefined);
      draw();
      if (focus) inp.focus();
    };
    box.append(tags, inp);
    draw();
    ctx.onRefresh(draw);
    box.addEventListener('click', e => { if (e.target === box) inp.focus(); });
    return box;
  }
  W.colsInput = colsInput;

  function textInput(value, onChange, opt) {
    opt = opt || {};
    const inp = h('input', { class: 'in' + (opt.mono ? ' mono' : ''), value: value === undefined || value === null ? '' : String(value), placeholder: opt.placeholder || '', list: opt.list, spellcheck: 'false' });
    inp.addEventListener('input', () => onChange(inp.value));
    return inp;
  }

  function numInput(value, onChange, opt) {
    opt = opt || {};
    const inp = h('input', { class: 'in num-in', type: 'number', value: value === undefined || value === null ? '' : String(value), placeholder: opt.placeholder || '', min: opt.min, step: 'any' });
    inp.addEventListener('input', () => {
      const t = inp.value.trim();
      onChange(t === '' ? undefined : Number(t));
    });
    return inp;
  }

  function checkbox(checked, label, onChange) {
    const c = h('input', { type: 'checkbox', checked });
    c.addEventListener('change', () => onChange(c.checked));
    return h('label', { class: 'ck' }, c, h('span', null, label));
  }

  function row(...kids) { return h('div', { class: 'row' }, ...kids); }

  /* ------------------------------------------------------------------ 颜色 */
  const SWATCHES = ['红', '黄', '绿', '蓝', '橙', '灰'];
  function colorInput(fd, value, onChange) {
    const wrap = h('div', { class: 'color-in' });
    const draw = () => {
      wrap.textContent = '';
      const cur = value === undefined ? (fd.default || '') : value;
      const curHex = cur === null ? null : M.parseColor(cur);
      const noneLabel = fd.nullable ? '不标色' : '不设置';
      wrap.append(h('button', { type: 'button', class: 'sw none' + (cur === null || (!cur && !fd.default) ? ' on' : ''), title: noneLabel, 'aria-label': noneLabel,
        onclick: () => set(fd.nullable ? null : undefined) }, '无'));
      SWATCHES.forEach(n => {
        const hex = M.NAMED_COLORS[n];
        wrap.append(h('button', { type: 'button', class: 'sw' + (curHex === hex ? ' on' : ''), style: 'background:#' + hex, title: n + '（' + hex + '）', 'aria-label': n,
          onclick: () => set(hex) }));
      });
      const pick = h('input', { type: 'color', class: 'sw-pick', value: '#' + (curHex || 'FFFFFF').toLowerCase(), title: '自定义颜色' });
      pick.addEventListener('input', () => set(pick.value.replace('#', '').toUpperCase(), true));
      const txt = h('input', { class: 'in mono hex-in', value: cur === null ? '' : String(cur || ''), placeholder: '如 FCE4D6', maxlength: 9, spellcheck: 'false' });
      txt.addEventListener('input', () => {
        const t = txt.value.trim();
        value = t === '' ? (fd.nullable ? null : undefined) : t;
        txt.classList.toggle('bad', !!t && !M.parseColor(t));
        onChange(value === (fd.default || undefined) ? undefined : value);
      });
      wrap.append(pick, txt);
      if (fd.default && value === undefined) wrap.append(h('span', { class: 'hint' }, '默认'));
    };
    const set = (v, keepFocus) => {
      value = v;
      onChange(v === fd.default ? undefined : v);
      if (!keepFocus) draw();
    };
    draw();
    return wrap;
  }

  /* ------------------------------------------------------------------ 条件构建器 */
  /** 条件对象 ↔ 编辑树 */
  function toTree(c) {
    if (!c) return { type: 'group', mode: 'all', children: [] };
    if (Array.isArray(c.all)) return { type: 'group', mode: 'all', children: c.all.map(toNode) };
    if (Array.isArray(c.any)) return { type: 'group', mode: 'any', children: c.any.map(toNode) };
    if (c.not) {
      const inner = c.not;
      if (Array.isArray(inner.any)) return { type: 'group', mode: 'none', children: inner.any.map(toNode) };
      return { type: 'group', mode: 'none', children: [toNode(inner)] };
    }
    return { type: 'group', mode: 'all', children: [toNode(c)] };
  }
  function toNode(c) {
    if (c && (c.all || c.any || c.not)) return toTree(c);
    const leaf = { type: 'leaf', extra: {} };
    Object.keys(c || {}).forEach(k => { if (!['col', 'ref_col', 'op', 'value', 'col2', 'ref_col2'].includes(k)) leaf.extra[k] = c[k]; });
    leaf.side = c.ref_col !== undefined ? 'ref' : 'main';
    leaf.col = c.ref_col !== undefined ? c.ref_col : c.col;
    leaf.op = M.canonOp(c.op || '=');
    if (c.col2 !== undefined) { leaf.rhs = 'col'; leaf.col2 = c.col2; }
    else if (c.ref_col2 !== undefined) { leaf.rhs = 'refcol'; leaf.col2 = c.ref_col2; }
    else { leaf.rhs = 'value'; leaf.value = c.value; }
    return leaf;
  }
  function fromTree(t, top) {
    if (t.type === 'leaf') {
      const o = {};
      o[t.side === 'ref' ? 'ref_col' : 'col'] = t.col === undefined ? '' : t.col;
      o.op = t.op;
      const op = M.condOp(t.op);
      if (!(op && op.noValue)) {
        if (t.rhs === 'col') o.col2 = t.col2 === undefined ? '' : t.col2;
        else if (t.rhs === 'refcol') o.ref_col2 = t.col2 === undefined ? '' : t.col2;
        else if (t.value !== undefined) o.value = t.value;
      }
      return Object.assign(o, t.extra || {});
    }
    const kids = t.children.map(k => fromTree(k, false)).filter(Boolean);
    if (!kids.length) return top ? undefined : null;
    if (t.mode === 'all' && kids.length === 1 && top) return kids[0];
    if (t.mode === 'none') return { not: kids.length === 1 ? kids[0] : { any: kids } };
    return t.mode === 'any' ? { any: kids } : { all: kids };
  }

  function condEditor(ctx, value, onChange, opt) {
    opt = opt || {};
    const box = h('div', { class: 'cond' });
    let tree = toTree(value);
    const commit = () => onChange(fromTree(tree, true));
    const redraw = () => { box.textContent = ''; box.append(groupView(tree, true)); };

    function groupView(g, top) {
      const el = h('div', { class: 'cgroup' + (top ? ' top' : '') });
      if (g.children.length > 1 || !top || g.mode !== 'all') {
        el.append(h('div', { class: 'cg-head' },
          select([{ value: 'all', label: '以下条件全部满足' }, { value: 'any', label: '以下条件满足任意一个' }, { value: 'none', label: '以下条件都不满足' }], g.mode, v => { g.mode = v; commit(); }, 'in in-sm'),
          !top ? h('button', { type: 'button', class: 'icon-btn', title: '删除这个条件组', 'aria-label': '删除条件组', onclick: () => { removeNode(tree, g); commit(); redraw(); } }, '×') : null));
      }
      g.children.forEach(k => el.append(k.type === 'leaf' ? leafView(k, g) : groupView(k, false)));
      el.append(h('div', { class: 'cg-actions' },
        h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { g.children.push({ type: 'leaf', side: 'main', op: '=', rhs: 'value', extra: {} }); commit(); redraw(); } }, g.children.length ? '＋ 再加一个条件' : '＋ 添加条件'),
        g.children.length ? h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { g.children.push({ type: 'group', mode: 'any', children: [{ type: 'leaf', side: 'main', op: '=', rhs: 'value', extra: {} }] }); commit(); redraw(); } }, '＋ 条件组（或）') : null));
      return el;
    }

    function leafView(l, parent) {
      const el = h('div', { class: 'cleaf' });
      if (opt.allowRef) {
        el.append(select([{ value: 'main', label: '主表' }, { value: 'ref', label: '参照表' }], l.side, v => { l.side = v; commit(); redraw(); }, 'in in-sm side'));
      }
      el.append(colInput(ctx, l.col, v => { l.col = v; commit(); }, { source: l.side === 'ref' ? 'ref' : (opt.source || 'main'), placeholder: '列' }));
      el.append(select(M.COND_OPS.map(o => ({ value: o.v, label: o.label })), l.op, v => {
        const was = M.condOp(l.op), now = M.condOp(v);
        l.op = v;
        if (now.list && !Array.isArray(l.value) && l.value !== undefined) l.value = [l.value];
        if (!now.list && Array.isArray(l.value)) l.value = l.value[0];
        commit();
        if (!!was.noValue !== !!now.noValue || !!was.list !== !!now.list) redraw();
      }, 'in in-sm op'));
      const op = M.condOp(l.op);
      if (!(op && op.noValue)) {
        const rhsOpts = [{ value: 'value', label: '值' }, { value: 'col', label: '另一列' }];
        if (opt.allowRef) rhsOpts.push({ value: 'refcol', label: '参照表的列' });
        el.append(select(rhsOpts, l.rhs, v => { l.rhs = v; if (v === 'value') delete l.col2; else delete l.value; commit(); redraw(); }, 'in in-sm rhs'));
        if (l.rhs === 'value') {
          const txt = Array.isArray(l.value) ? l.value.join('、') : M.scalarText(l.value);
          const inp = textInput(txt, t => {
            l.value = op && op.list ? M.parseList(t) : (t === '' ? undefined : M.parseScalar(t));
            commit();
          }, { placeholder: op && op.list ? '多个值用顿号或逗号分隔' : '比较的值' });
          inp.classList.add('val');
          el.append(inp);
        } else {
          el.append(colInput(ctx, l.col2, v => { l.col2 = v; commit(); }, { source: l.rhs === 'refcol' ? 'ref' : (opt.source || 'main'), placeholder: '比较的列' }));
        }
      }
      el.append(h('button', { type: 'button', class: 'icon-btn', title: '删除这个条件', 'aria-label': '删除条件', onclick: () => { parent.children.splice(parent.children.indexOf(l), 1); commit(); redraw(); } }, '×'));
      return el;
    }

    function removeNode(root, node) {
      const walk = g => {
        const i = g.children.indexOf(node);
        if (i >= 0) { g.children.splice(i, 1); return true; }
        return g.children.some(k => k.type === 'group' && walk(k));
      };
      walk(root);
    }

    redraw();
    return box;
  }
  W.condEditor = condEditor;

  /* ------------------------------------------------------------------ 各字段类型 */
  const R = {};

  R.text = (fd, ctx) => textInput(ctx.step[fd.key], v => ctx.set(fd.key, v === '' && !fd.keepEmpty ? undefined : (v === '' ? undefined : v)), { placeholder: fd.placeholder });
  R.number = (fd, ctx) => numInput(ctx.step[fd.key], v => ctx.set(fd.key, v), { placeholder: fd.placeholder || (fd.default !== undefined ? '默认 ' + fd.default : ''), min: fd.min });
  R.bool = (fd, ctx) => checkbox(fd.key in ctx.step ? !!ctx.step[fd.key] : !!fd.default, fd.label, v => ctx.set(fd.key, v === !!fd.default ? undefined : v));
  R.enum = (fd, ctx) => {
    const cur = fd.key in ctx.step ? ctx.step[fd.key] : fd.default;
    let opts = fd.options;
    // 导入的配置里可能写了列名（如 style_from: "金额" 表示沿用该列格式），保留它
    if (fd.allowColumn && cur !== undefined && !opts.some(o => o.value === cur)) opts = opts.concat([{ value: cur, label: '沿用「' + cur + '」列的格式' }]);
    return select(opts, cur, v => ctx.set(fd.key, v === fd.default ? undefined : v));
  };
  R.triBool = (fd, ctx) => select(fd.options, fd.key in ctx.step ? ctx.step[fd.key] : '', v => ctx.set(fd.key, v === '' ? undefined : v));
  R.color = (fd, ctx) => colorInput(fd, ctx.step[fd.key], v => ctx.set(fd.key, v));
  R.numfmt = (fd, ctx) => {
    const id = 'nf' + (++uid);
    const dl = h('datalist', { id },
      ['#,##0.00', '#,##0', '0.00', '0%', '0.00%', 'yyyy-mm-dd', 'yyyy-mm-dd hh:mm', 'yyyy年m月d日', '@'].map(v => h('option', { value: v })));
    return h('div', null, textInput(ctx.step[fd.key], v => ctx.set(fd.key, v.trim() === '' ? undefined : v), { placeholder: '如 #,##0.00、0%、yyyy-mm-dd', list: id, mono: true }), dl);
  };
  R.column = (fd, ctx) => colInput(ctx, ctx.step[fd.key], v => ctx.set(fd.key, v), { source: fd.source });
  R.columns = (fd, ctx) => colsInput(ctx, ctx.step[fd.key], v => ctx.set(fd.key, v), { source: fd.source === 'ref' ? 'refFile' : fd.source });
  R.sheet = (fd, ctx) => {
    const id = ctx.sheetList(fd.source === 'ref' ? ctx.step.ref : null);
    return textInput(ctx.step[fd.key], v => ctx.set(fd.key, v.trim() === '' ? undefined : v.trim()), { placeholder: fd.placeholder, list: id });
  };
  R.condition = (fd, ctx) => condEditor(ctx, ctx.step[fd.key], v => ctx.set(fd.key, v), { allowRef: fd.allowRef, source: fd.source === 'ref' ? 'refFile' : 'main' });
  R.rowRange = (fd, ctx) => textInput(ctx.step[fd.key], v => ctx.set(fd.key, v.trim() === '' ? undefined : v.trim()), { placeholder: fd.placeholder, mono: true });

  R.matchId = (fd, ctx) => textInput(ctx.step[fd.key], v => ctx.set(fd.key, v.trim() === '' ? undefined : v.trim()), { placeholder: '如 m1' });

  R.fileRef = (fd, ctx) => {
    const files = M.state.cfg.files;
    if (!files.length) return h('div', { class: 'hint' }, '请先在左侧添加文件');
    const opts = [{ value: undefined, label: '— 选择 —' }].concat(files.map(f => ({ value: f.id, label: (f.label || f.id) + (f.role === 'main' ? '（主表）' : '') })));
    return select(opts, ctx.step[fd.key], v => { ctx.set(fd.key, v); ctx.rerender(); });
  };

  R.matchRef = (fd, ctx) => {
    const ms = M.matchesBefore(ctx.index);
    if (!ms.length) return h('div', { class: 'notice warn' }, '这一步之前还没有“建立匹配”步骤。请先添加一个“建立匹配”，再把它拖到这一步上面。');
    const opts = [{ value: undefined, label: '— 选择 —' }].concat(ms.map(m => ({ value: m.id, label: m.id + '：' + M.summary(m.step) })));
    if (ctx.step[fd.key] === undefined && ms.length === 1) { ctx.step[fd.key] = ms[0].id; setTimeout(ctx.changed); }
    return select(opts, ctx.step[fd.key], v => { ctx.set(fd.key, v); ctx.rerender(); });
  };

  R.content = (fd, ctx) => {
    const s = ctx.step;
    const modes = [];
    if (fd.allowNone) modes.push({ value: 'none', label: '留空' });
    modes.push({ value: 'value', label: '固定值' }, { value: 'formula', label: '公式' }, { value: 'from_col', label: '同一行另一列的值' }, { value: 'formula_from', label: '复制某个单元格的公式' });
    const cur = ['value', 'formula', 'from_col', 'formula_from'].find(k => k in s) || (fd.allowNone ? 'none' : 'value');
    const box = h('div', { class: 'stack' });
    box.append(select(modes, cur, v => {
      ['value', 'formula', 'from_col', 'formula_from'].forEach(k => delete s[k]);
      if (v === 'value') s.value = '';
      if (v === 'formula') s.formula = '=';
      if (v === 'from_col') s.from_col = '';
      if (v === 'formula_from') s.formula_from = { cell: '' };
      ctx.changed(); ctx.rerender();
    }));
    if (cur === 'value') {
      if (!('value' in s)) s.value = '';
      box.append(textInput(M.scalarText(s.value), t => { s.value = M.parseScalar(t); ctx.changed(); }, { placeholder: '要写入的文字或数字；留空=清空单元格' }));
    } else if (cur === 'formula') {
      box.append(textInput(s.formula, t => { s.formula = t; ctx.changed(); }, { placeholder: '=[@单价]*[@数量]', mono: true }),
        h('div', { class: 'hint' }, '[@列名] = 同一行的单元格，[列名] = 整列数据区域，{row} = 当前行号。例：=[@单价]*[@数量]、=SUMIF([客户],[@客户],[金额])'));
    } else if (cur === 'from_col') {
      box.append(colInput(ctx, s.from_col, v => { s.from_col = v === undefined ? '' : v; ctx.changed(); }));
    } else if (cur === 'formula_from') {
      const ff = s.formula_from = (s.formula_from && typeof s.formula_from === 'object') ? s.formula_from : { cell: '' };
      const files = M.state.cfg.files;
      const fileOpts = [{ value: undefined, label: '主表' }].concat(files.filter(f => f.role !== 'main').map(f => ({ value: f.id, label: f.label || f.id })));
      const sheetId = ctx.sheetList(ff.file || null);
      box.append(row(
        select(fileOpts, ff.file, v => { if (v === undefined) delete ff.file; else ff.file = v; ctx.changed(); ctx.rerender(); }),
        textInput(ff.sheet, v => { if (v.trim()) ff.sheet = v.trim(); else delete ff.sheet; ctx.changed(); }, { placeholder: '工作表（不填=默认）', list: sheetId }),
        textInput(ff.cell, v => { ff.cell = v.trim().toUpperCase(); ctx.changed(); }, { placeholder: '单元格，如 P2', mono: true })),
        h('div', { class: 'hint' }, '像 Excel 复制粘贴一样平移公式里的相对引用。来源公式引用了主表没有的工作表时会报错，可先用“复制整个工作表”把它复制进来。'));
    }
    return box;
  };

  R.position = (fd, ctx) => {
    const s = ctx.step;
    const cur = ['before', 'after', 'at', 'end'].find(k => k in s) || 'end';
    if (!['before', 'after', 'at', 'end'].some(k => k in s)) s.end = true;
    const box = h('div', { class: 'row' });
    box.append(select([{ value: 'end', label: '放在最后' }, { value: 'before', label: '在某列之前' }, { value: 'after', label: '在某列之后' }, { value: 'at', label: '指定列字母' }], cur, v => {
      ['before', 'after', 'at', 'end'].forEach(k => delete s[k]);
      if (v === 'end') s.end = true; else s[v] = '';
      ctx.changed(); ctx.rerender();
    }));
    if (cur === 'before' || cur === 'after') box.append(colInput(ctx, s[cur], v => { s[cur] = v === undefined ? '' : v; ctx.changed(); }));
    if (cur === 'at') box.append(textInput(s.at, v => { s.at = v.trim().toUpperCase(); ctx.changed(); }, { placeholder: '如 C', mono: true }));
    return box;
  };

  R.rowMode = (fd, ctx) => {
    const s = ctx.step;
    const cur = ['where', 'rows', 'blank'].find(k => k in s) || 'where';
    const box = h('div', { class: 'stack' });
    box.append(select([{ value: 'where', label: '符合条件的行' }, { value: 'rows', label: '指定行号范围' }, { value: 'blank', label: '整行为空的行' }], cur, v => {
      ['where', 'rows', 'blank'].forEach(k => delete s[k]);
      if (v === 'blank') s.blank = true;
      if (v === 'rows') s.rows = '';
      ctx.changed(); ctx.rerender();
    }));
    if (cur === 'where') box.append(condEditor(ctx, s.where, v => ctx.set('where', v)));
    if (cur === 'rows') box.append(textInput(s.rows, v => { s.rows = v.trim(); ctx.changed(); }, { placeholder: '如 5:20（第 5 行到第 20 行）', mono: true }));
    return box;
  };

  R.rowsAt = (fd, ctx) => {
    const s = ctx.step;
    if (!('before_row' in s) && !('at_end' in s)) s.at_end = true;
    const cur = 'before_row' in s ? 'before_row' : 'at_end';
    const box = h('div', { class: 'row' });
    box.append(select([{ value: 'at_end', label: '表格最后' }, { value: 'before_row', label: '在第几行之前' }], cur, v => {
      delete s.before_row; delete s.at_end;
      if (v === 'at_end') s.at_end = true; else s.before_row = 2;
      ctx.changed(); ctx.rerender();
    }));
    if (cur === 'before_row') box.append(numInput(s.before_row, v => { if (v === undefined) s.before_row = ''; else s.before_row = v; ctx.changed(); }, { min: 1 }));
    return box;
  };

  R.sortBy = (fd, ctx) => {
    const s = ctx.step;
    let list = s.by === undefined ? [] : (Array.isArray(s.by) ? s.by : [s.by]).map(b => (b && typeof b === 'object' && 'column' in b) ? Object.assign({}, b) : { column: b });
    if (!list.length) list = [{ column: '' }];
    const box = h('div', { class: 'stack' });
    const commit = () => { s.by = list.map(b => { const o = { column: b.column }; if (b.order === 'desc') o.order = 'desc'; return o; }); ctx.changed(); };
    const draw = () => {
      box.textContent = '';
      list.forEach((b, i) => box.append(row(
        h('span', { class: 'lbl-sm' }, i ? '然后按' : '先按'),
        colInput(ctx, b.column, v => { b.column = v === undefined ? '' : v; commit(); }),
        select([{ value: 'asc', label: '升序（小→大）' }, { value: 'desc', label: '降序（大→小）' }], b.order || 'asc', v => { b.order = v; commit(); }),
        list.length > 1 ? h('button', { type: 'button', class: 'icon-btn', 'aria-label': '删除', onclick: () => { list.splice(i, 1); commit(); draw(); } }, '×') : null)));
      box.append(h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { list.push({ column: '' }); commit(); draw(); } }, '＋ 再加一个排序列'));
    };
    if (s.by === undefined) commit();
    draw();
    return box;
  };

  R.cells = (fd, ctx) => {
    const s = ctx.step;
    let list = Object.keys(s.cells || {}).map(k => ({ addr: k, val: s.cells[k] }));
    if (!list.length) list = [{ addr: '', val: '' }];
    const box = h('div', { class: 'stack' });
    const commit = () => {
      const o = {};
      list.forEach(x => { if (x.addr.trim()) o[x.addr.trim().toUpperCase()] = x.val; });
      s.cells = o;
      ctx.changed();
    };
    const draw = () => {
      box.textContent = '';
      list.forEach((x, i) => box.append(row(
        textInput(x.addr, v => { x.addr = v; commit(); }, { placeholder: '单元格，如 AC1', mono: true }),
        textInput(M.scalarText(x.val), v => { x.val = v.trim().startsWith('=') ? v.trim() : M.parseScalar(v); commit(); }, { placeholder: '内容：文字、数字或以 = 开头的公式' }),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': '删除', onclick: () => { list.splice(i, 1); commit(); draw(); } }, '×'))));
      box.append(h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { list.push({ addr: '', val: '' }); draw(); } }, '＋ 再加一个单元格'));
    };
    draw();
    return box;
  };

  R.copyColumns = (fd, ctx) => {
    const s = ctx.step;
    if (s.from !== undefined && s.columns === undefined) {
      return h('div', { class: 'notice info' }, '这一步用 from / to 写法指定列（' + M.copyPairs(s).map(p => p.from + '→' + p.to).join('、') + '），可在下方“其他参数”里修改。');
    }
    let list = M.copyPairs(s).map(p => ({ from: p.from, to: p.to === p.from ? '' : p.to }));
    if (!list.length) list = [{ from: '', to: '' }];
    const box = h('div', { class: 'stack' });
    const commit = () => {
      s.columns = list.filter(p => p.from !== '' && p.from !== undefined).map(p => (p.to && p.to !== p.from) ? { from: p.from, to: p.to } : p.from);
      ctx.changed();
    };
    const draw = () => {
      box.textContent = '';
      box.append(h('div', { class: 'row head-row' }, h('span', null, '参照表的列'), h('span', null, '写到主表的列（不填=同名）')));
      list.forEach((p, i) => box.append(row(
        colInput(ctx, p.from, v => { p.from = v === undefined ? '' : v; commit(); }, { source: 'ref', placeholder: '参照表列' }),
        h('span', { class: 'arrow' }, '→'),
        colInput(ctx, p.to || undefined, v => { p.to = v === undefined ? '' : v; commit(); }, { placeholder: '同名' }),
        h('button', { type: 'button', class: 'icon-btn', 'aria-label': '删除', onclick: () => { list.splice(i, 1); commit(); draw(); } }, '×'))));
      box.append(h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { list.push({ from: '', to: '' }); draw(); } }, '＋ 再加一列'));
    };
    draw();
    return box;
  };

  R.comparePairs = (fd, ctx) => {
    const s = ctx.step;
    const custom = Array.isArray(s.compare_columns);
    const box = h('div', { class: 'stack' });
    box.append(select([{ value: false, label: '所有同名列（键列除外）' }, { value: true, label: '只比较指定的列' }], custom, v => {
      if (v) s.compare_columns = []; else delete s.compare_columns;
      ctx.changed(); ctx.rerender();
    }));
    if (custom) {
      const list = s.compare_columns.map(p => (p && typeof p === 'object') ? { main: p.main, ref: p.ref === p.main ? '' : (p.ref || '') } : { main: p, ref: '' });
      if (!list.length) list.push({ main: '', ref: '' });
      const commit = () => {
        s.compare_columns = list.filter(p => p.main).map(p => (p.ref && p.ref !== p.main) ? { main: p.main, ref: p.ref } : p.main);
        ctx.changed();
      };
      const inner = h('div', { class: 'stack' });
      const draw = () => {
        inner.textContent = '';
        list.forEach((p, i) => inner.append(row(
          colInput(ctx, p.main || undefined, v => { p.main = v || ''; commit(); }, { placeholder: '主表列' }),
          h('span', { class: 'arrow' }, '对比'),
          colInput(ctx, p.ref || undefined, v => { p.ref = v || ''; commit(); }, { source: 'ref', placeholder: '参照表列（不填=同名）' }),
          h('button', { type: 'button', class: 'icon-btn', 'aria-label': '删除', onclick: () => { list.splice(i, 1); commit(); draw(); } }, '×'))));
        inner.append(h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { list.push({ main: '', ref: '' }); draw(); } }, '＋ 再加一组'));
      };
      draw();
      box.append(inner);
    }
    return box;
  };

  R.optText = (fd, ctx) => {
    const s = ctx.step;
    const on = s[fd.key] !== null && s[fd.key] !== '';
    const box = h('div', { class: 'row' });
    box.append(checkbox(on, fd.label, v => {
      if (v) delete s[fd.key]; else s[fd.key] = null;
      ctx.changed(); ctx.rerender();
    }));
    if (on) box.append(textInput(s[fd.key], v => { if (!v.trim() || v.trim() === fd.defaultText) delete s[fd.key]; else s[fd.key] = v.trim(); ctx.changed(); }, { placeholder: fd.defaultText }));
    return box;
  };

  R.labels = (fd, ctx) => {
    const s = ctx.step;
    const lab = Object.assign({}, s.labels || {});
    const commit = () => {
      Object.keys(lab).forEach(k => { if (!lab[k]) delete lab[k]; });
      if (Object.keys(lab).length) s.labels = Object.assign({}, lab); else delete s.labels;
      ctx.changed();
    };
    const one = (k, def) => h('label', { class: 'mini' }, h('span', null, def + ' →'), textInput(lab[k], v => { lab[k] = v.trim(); commit(); }, { placeholder: def }));
    return h('div', { class: 'row wrap' }, one('same', '一致'), one('diff', '有差异'), one('main_only', '仅主表有'));
  };

  function fileSelect(value, onChange, opt) {
    opt = opt || {};
    const files = M.state.cfg.files.filter(f => !opt.refOnly || f.role !== 'main');
    const opts = files.map(f => ({ value: f.id, label: (f.label || f.id) + (f.role === 'main' ? '（主表）' : '') }));
    if (opt.allowDefault) opts.unshift({ value: undefined, label: opt.allowDefault });
    return select(opts, value, onChange);
  }

  R.rangeSource = (fd, ctx) => {
    const s = ctx.step;
    const v = s.from = (s.from && typeof s.from === 'object') ? s.from : {};
    if (v.file === undefined) { const r = M.state.cfg.files.find(f => f.role !== 'main'); if (r) v.file = r.id; }
    return h('div', { class: 'row wrap' },
      fileSelect(v.file, x => { if (x === undefined) delete v.file; else v.file = x; ctx.changed(); ctx.rerender(); }, { allowDefault: '主表' }),
      textInput(v.sheet, x => { if (x.trim()) v.sheet = x.trim(); else delete v.sheet; ctx.changed(); }, { placeholder: '工作表（不填=默认）', list: ctx.sheetList(v.file || null) }),
      textInput(v.range, x => { v.range = x.trim().toUpperCase(); ctx.changed(); }, { placeholder: '区域，如 AF1:AI30', mono: true }));
  };
  R.cellTarget = (fd, ctx) => {
    const s = ctx.step;
    const v = s.to = (typeof s.to === 'string') ? { cell: s.to } : ((s.to && typeof s.to === 'object') ? s.to : {});
    return h('div', { class: 'row wrap' },
      textInput(v.sheet, x => { if (x.trim()) v.sheet = x.trim(); else delete v.sheet; ctx.changed(); }, { placeholder: '主表工作表（不填=默认）', list: ctx.sheetList(null) }),
      textInput(v.cell, x => { v.cell = x.trim().toUpperCase(); ctx.changed(); }, { placeholder: '左上角单元格，如 AF1', mono: true }));
  };
  R.sheetSource = (fd, ctx) => {
    const s = ctx.step;
    const v = s.from = (s.from && typeof s.from === 'object') ? s.from : {};
    if (v.file === undefined) { const r = M.state.cfg.files.find(f => f.role !== 'main'); if (r) v.file = r.id; }
    return h('div', { class: 'row wrap' },
      fileSelect(v.file, x => { v.file = x; ctx.changed(); ctx.rerender(); }, { refOnly: true }),
      textInput(v.sheet, x => { if (x.trim()) v.sheet = x.trim(); else delete v.sheet; ctx.changed(); }, { placeholder: '工作表（不填=第一个）', list: ctx.sheetList(v.file || null) }));
  };

  const PARTS = [{ v: 'values', label: '值' }, { v: 'formulas', label: '公式' }, { v: 'styles', label: '格式（底色、字体、边框、数字格式）' }, { v: 'widths', label: '列宽' }, { v: 'merges', label: '合并单元格' }];
  R.contentParts = (fd, ctx) => {
    const s = ctx.step;
    const cur = new Set(s.content === undefined || s.content === 'all' ? PARTS.map(p => p.v) : (Array.isArray(s.content) ? s.content : [s.content]));
    return h('div', { class: 'row wrap' }, PARTS.map(p => checkbox(cur.has(p.v), p.label, on => {
      if (on) cur.add(p.v); else cur.delete(p.v);
      if (cur.size === PARTS.length) delete s.content; else s.content = PARTS.map(x => x.v).filter(x => cur.has(x));
      ctx.changed();
    })));
  };

  R.notFound = (fd, ctx) => {
    const s = ctx.step;
    const on = !!s.not_found;
    const box = h('div', { class: 'row wrap' });
    box.append(checkbox(on, '在未匹配到的行写入标记', v => {
      if (v) s.not_found = { column: '', value: '未找到' }; else delete s.not_found;
      ctx.changed(); ctx.rerender();
    }));
    if (on) {
      box.append(colInput(ctx, s.not_found.column || undefined, v => { s.not_found.column = v || ''; ctx.changed(); }, { placeholder: '写到哪一列' }),
        textInput(s.not_found.value, v => { s.not_found.value = v; ctx.changed(); }, { placeholder: '未找到' }));
    }
    return box;
  };

  R.checks = (fd, ctx) => {
    const s = ctx.step;
    if (!Array.isArray(s.checks)) s.checks = [];
    const box = h('div', { class: 'stack' });
    const types = Object.keys(M.CHECK_TYPES).map(k => ({ value: k, label: M.CHECK_TYPES[k] }));
    const draw = () => {
      box.textContent = '';
      s.checks.forEach((ck, i) => {
        const card = h('div', { class: 'ck-card' });
        card.append(h('div', { class: 'row' },
          h('span', { class: 'lbl-sm' }, (i + 1) + '.'),
          select(types, ck.type, v => { s.checks[i] = { type: v }; if (v === 'condition') s.checks[i].expect = 'none'; ctx.changed(); draw(); }),
          h('button', { type: 'button', class: 'icon-btn', 'aria-label': '删除这项检查', onclick: () => { s.checks.splice(i, 1); ctx.changed(); draw(); } }, '×')));
        const body = h('div', { class: 'ck-body' });
        const t = ck.type;
        const set = (k, v) => { if (v === undefined) delete ck[k]; else ck[k] = v; ctx.changed(); };
        if (['required_columns', 'not_empty', 'unique'].includes(t)) {
          body.append(colsInput(ctx, ck.columns, v => set('columns', v), { source: 'ck' }));
          if (t === 'unique') body.append(h('div', { class: 'hint' }, '选多列时，按多列组合判断是否重复'));
        } else if (t === 'row_count') {
          body.append(row(h('span', { class: 'lbl-sm' }, '至少'), numInput(ck.min, v => set('min', v), { min: 0, placeholder: '0' }), h('span', { class: 'lbl-sm' }, '行，最多'), numInput(ck.max, v => set('max', v), { min: 0, placeholder: '不限' }), h('span', { class: 'lbl-sm' }, '行')));
        } else if (t === 'headers_match') {
          body.append(row(h('span', { class: 'lbl-sm' }, '对照文件'), fileSelect(ck.file === undefined ? 'ref' : ck.file, v => set('file', v === 'ref' ? undefined : v)),
            textInput(ck.columns, v => set('columns', v.trim() ? v.trim().toUpperCase() : undefined), { placeholder: '列范围，如 A:AB（不填=全部列）', mono: true })));
        } else if (t === 'condition') {
          body.append(row(h('span', { class: 'lbl-sm' }, '期望'), select([{ value: 'none', label: '不应有符合条件的行' }, { value: 'all', label: '所有行都应符合' }, { value: 'any', label: '至少有一行符合' }], ck.expect || 'none', v => set('expect', v))));
          body.append(condEditor(ctx, ck.where, v => set('where', v), { source: 'ck' }));
        }
        body.append(h('div', { class: 'row wrap' },
          checkbox(ck.on_fail === 'stop', '不通过就中止（不生成结果文件）', v => set('on_fail', v ? 'stop' : undefined)),
          textInput(ck.message, v => set('message', v.trim() || undefined), { placeholder: '报告里显示的说明（可选）' })));
        card.append(body);
        box.append(card);
      });
      box.append(h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => { s.checks.push({ type: 'not_empty' }); ctx.changed(); draw(); } }, '＋ 添加检查项'));
    };
    draw();
    return box;
  };

  W.render = function (fd, ctx) {
    const fn = R[fd.type];
    if (!fn) return h('div', { class: 'hint' }, '（不支持的字段类型 ' + fd.type + '）');
    return fn(fd, ctx);
  };
  W.TYPES = Object.keys(R);

  EFA.WFFields = W;
})(window.EFA = window.EFA || {});
