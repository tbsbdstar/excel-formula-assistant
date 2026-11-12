/* 工作流配置器：页面交互 */
(function (EFA) {
  'use strict';
  const M = EFA.WFModel, F = EFA.WFFiles, W = EFA.WFFields, Z = EFA.WFZip, U = EFA.U;
  const h = W.h;
  const S = M.state;
  const $ = id => document.getElementById(id);
  const DRAFT_KEY = 'efa.wf.draft';

  let openIdx = null;          // 当前展开的步骤
  let val = null;              // 最近一次校验结果
  let simCache = null;
  let editorHooks = [];        // 展开的编辑器里需要随数据刷新的控件
  let toolkit = null;          // { name, size, buf }
  let pendingSampleFor = null; // “载入样例”针对的文件 id

  /* ------------------------------------------------------------------ 工具 */
  function toast(msg, ms) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => t.classList.remove('show'), ms || 3200);
  }
  function sim() { if (!simCache) simCache = M.simulate(); return simCache; }
  function invalidate() { simCache = null; val = M.validate(); simCache = val.sim; }
  const saveDraft = U.debounce(() => {
    try { localStorage.setItem(DRAFT_KEY, JSON.stringify(S.cfg)); } catch (e) { /* 浏览器禁用存储时忽略 */ }
  }, 400);
  function safeName(s) { return String(s || '工作流').replace(/[\\/:*?"<>|]/g, '_').trim() || '工作流'; }
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  function fmtSize(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + 'MB' : Math.max(1, Math.round(n / 1024)) + 'KB'; }
  function stepsReferencing(fileId) {
    return S.cfg.steps.filter(s => s.file === fileId || s.ref === fileId || (s.from && s.from.file === fileId) || (s.formula_from && s.formula_from.file === fileId) || (Array.isArray(s.checks) && s.checks.some(c => c.file === fileId)));
  }

  /* ------------------------------------------------------------------ 整体刷新 */
  function renderAll() {
    invalidate();
    renderFiles();
    renderSteps();
    renderExport();
  }

  /** 数据变化后的轻量刷新：不重建正在编辑的控件，只更新标题、状态和提示 */
  function changed() {
    invalidate();
    updateStepHeads();
    updateEditorMsgs();
    editorHooks.forEach(fn => { try { fn(); } catch (e) { /* ignore */ } });
    renderFileMsgs();
    renderCheckBox();
    saveDraft();
  }

  /* ------------------------------------------------------------------ 文件 */
  async function addSampleFiles(fileList, targetId) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    let added = 0;
    const created = [], skipped = [];
    for (const file of files) {
      let smp;
      try {
        smp = await F.read(file);
      } catch (e) {
        toast(e.message || String(e), 6000);
        continue;
      }
      let spec = targetId ? M.fileSpec(targetId) : null;
      if (!spec) spec = S.cfg.files.find(f => !S.samples[f.id] && f.match && F.matchPattern(f.match, file.name)) || null;
      // 属于已经有样例的“多个文件”：同类文件只需要一个样例，不再新建文件
      if (!spec && S.cfg.files.some(f => f.multiple && S.samples[f.id] && f.match && F.matchPattern(f.match, file.name))) {
        skipped.push(file.name);
        continue;
      }
      if (!spec) {
        const role = S.cfg.files.some(f => f.role === 'main') ? 'ref' : 'main';
        const pattern = F.suggestPattern(file.name);
        spec = { id: M.newFileId(role), role, label: pattern.replace(/\*/g, '').replace(/\.[^.]+$/, '') || file.name.replace(/\.[^.]+$/, ''), match: pattern };
        const hr = smp.sheets[0] ? smp.sheets[0].headerRow : 1;
        if (hr !== 1) spec.header_row = hr;
        S.cfg.files.push(spec);
        created.push(spec.id);
      }
      S.samples[spec.id] = smp;
      targetId = null;
      added++;
    }
    // 一次拖入的几个文件表头完全相同：可能是要合并的同类文件（各门店、各月），给出一键设置的提示
    S.mergeHint = null;
    if (created.length >= 2) {
      // 合并按表头名称对应列：只看有哪些列、不看顺序；大部分列相同（个别文件多一两列）也算同类
      const cols = created.map(id => { const f = M.fileSpec(id); return new Set((F.headerList(M.sampleSheet(id), f.header_row || 1) || []).filter(Boolean)); });
      const union = new Set(cols.flatMap(s => [...s]));
      const common = [...union].filter(c => cols.every(s => s.has(c)));
      if (common.length >= 3 && common.length / union.size >= 0.7) {
        S.mergeHint = { ids: created.slice(), partial: [...union].filter(c => !common.includes(c)) };
      }
    }
    if (added) {
      renderAll();
      saveDraft();
    }
    if (added || skipped.length) {
      toast((added ? '已载入 ' + added + ' 个样例文件' : '') + (skipped.length ? (added ? '；' : '') + skipped.join('、') + ' 属于“多个文件”，同类文件有一个样例就够了，未重复载入' : ''), 5000);
    }
  }

  /** 把几个同类文件设为：第一个作主表，其余用一个“多个文件”参照表表示，并加一步“追加行” */
  function applyMergeHint() {
    const ids = (S.mergeHint || {}).ids || [];
    if (ids.length < 2) return;
    const names = ids.map(id => S.samples[id] ? S.samples[id].fileName : '').filter(Boolean);
    const pattern = F.commonPattern(names);
    const mainSpec = M.fileSpec(ids[0]);
    S.cfg.files.forEach(f => { if (f.role === 'main' && f.id !== ids[0]) f.role = 'ref'; });
    mainSpec.role = 'main';
    delete mainSpec.multiple;
    mainSpec.match = pattern;
    const multi = M.fileSpec(ids[1]);
    multi.role = 'ref';
    multi.multiple = true;
    multi.match = pattern;
    const core = pattern.replace(/\*/g, '').replace(/\.[^.]+$/, '');
    multi.label = (core.length >= 2 ? core : '同类文件') + '（其余文件）';
    ids.slice(2).forEach(id => { S.cfg.files = S.cfg.files.filter(f => f.id !== id); delete S.samples[id]; });
    if (!S.cfg.steps.some(s => s.op === 'append_rows')) S.cfg.steps.unshift({ op: 'append_rows', from: { file: multi.id }, source_column: '来源文件' });
    S.mergeHint = null;
    openIdx = null;
    renderAll(); saveDraft();
    toast('已设置：符合「' + pattern + '」的文件里，第一个作主表，其余全部追加进来（第 1 步）。规则可以在左侧修改', 6000);
  }

  function renderFiles() {
    const list = $('fileList');
    list.textContent = '';
    $('fileCount').textContent = S.cfg.files.length ? '（' + S.cfg.files.length + '）' : '';
    const mainId = M.mainId();
    const hint = S.mergeHint && S.mergeHint.ids.every(id => M.fileSpec(id)) ? S.mergeHint : null;
    if (hint) {
      list.append(h('div', { class: 'merge-hint' },
        h('div', null, h('b', null, '这 ' + hint.ids.length + ' 个文件的列' + (hint.partial && hint.partial.length ? '基本相同' : '相同（顺序可以不同）') + '。'),
          '要把它们合并成一张表吗？（例如各门店、各月份的同一种表）' +
          (hint.partial && hint.partial.length ? '只在部分文件里有的列：' + hint.partial.join('、') + '（合并时可选择是否保留）。' : '')),
        h('div', { class: 'row' },
          h('button', { type: 'button', class: 'btn btn-primary btn-sm', onclick: applyMergeHint }, '设为合并：第一个作主表，其余追加'),
          h('button', { type: 'button', class: 'btn btn-text btn-sm', onclick: () => { S.mergeHint = null; renderFiles(); } }, '不用，它们用途不同'))));
    }
    S.cfg.files.forEach(spec => list.append(fileCard(spec, mainId)));
    renderFileMsgs();
  }

  function fileCard(spec, mainId) {
    const smp = S.samples[spec.id];
    const card = h('div', { class: 'file-card' + (spec.role === 'main' ? ' main' : ''), 'data-id': spec.id });
    const labelIn = h('input', { class: 'in label-in', value: spec.label || '', placeholder: '给这个文件起个名字，如：本周订单', 'aria-label': '文件名称' });
    labelIn.addEventListener('input', () => { spec.label = labelIn.value.trim() || undefined; changed(); });
    card.append(h('div', { class: 'fc-head' },
      h('span', { class: 'role ' + spec.role }, spec.role === 'main' ? '主表' : (spec.multiple ? '参照表 · 多个' : '参照表')),
      labelIn,
      h('button', { type: 'button', class: 'icon-btn', title: '移除这个文件', 'aria-label': '移除文件', onclick: () => removeFile(spec.id) }, '×')));
    const body = h('div', { class: 'fc-body' });
    // 角色
    body.append(h('div', { class: 'fc-line' }, h('span', null, '角色'),
      h('div', null, h('span', { class: 'seg-mini' },
        h('button', { type: 'button', class: spec.role === 'main' ? 'on' : '', onclick: () => setRole(spec.id, 'main') }, '主表（输出结果）'),
        h('button', { type: 'button', class: spec.role === 'ref' ? 'on' : '', onclick: () => setRole(spec.id, 'ref') }, '参照表（只读）')))));
    // 样例
    if (smp) {
      const sh = M.sampleSheet(spec.id, spec.sheet);
      const hr = spec.header_row || 1;
      body.append(h('div', { class: 'fc-sample' }, '样例：', h('b', null, smp.fileName),
        h('span', null, '· ' + smp.sheets.length + ' 个工作表' + (sh ? ' · 约 ' + F.dataRows(sh, hr).toLocaleString() + ' 行 × ' + sh.fullCols + ' 列' : '')),
        h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => pickSample(spec.id) }, '更换')));
    } else {
      body.append(h('div', { class: 'fc-sample none' }, '未载入样例文件：无法从列表选择列名、检查列是否存在。',
        h('button', { type: 'button', class: 'btn-text btn-xs', onclick: () => pickSample(spec.id) }, '载入样例')));
    }
    // 文件名规则
    const patIn = h('input', { class: 'in mono', value: spec.match || '', placeholder: '如 *本周订单*.xlsx', spellcheck: 'false' });
    const patMsg = h('div', { class: 'hint' });
    const markPat = () => {
      patMsg.className = 'hint';
      if (!spec.match) { patMsg.textContent = '不填：运行时会提示把文件拖进窗口'; return; }
      if (smp) {
        const ok = F.matchPattern(spec.match, smp.fileName);
        patMsg.className = ok ? 'pat-ok' : 'pat-bad';
        patMsg.textContent = ok ? '✓ 样例文件名符合规则' : '✗ 样例文件名「' + smp.fileName + '」不符合规则';
      } else patMsg.textContent = '* 代表任意文字，例如 *订单*.xlsx';
    };
    patIn.addEventListener('input', () => { spec.match = patIn.value.trim() || undefined; markPat(); changed(); });
    markPat();
    body.append(h('div', { class: 'fc-line' }, h('span', null, '文件名规则'), h('div', { class: 'stack', style: 'gap:3px' }, patIn, patMsg)));
    if (spec.role === 'ref') {
      const mc = h('input', { type: 'checkbox', checked: !!spec.multiple });
      mc.addEventListener('change', () => {
        if (mc.checked) {
          spec.multiple = true;
          // 规则通常要放宽，才能匹配到同一类的所有文件
          if (spec.match && smp && spec.match === F.suggestPattern(smp.fileName)) toast('提示：确认文件名规则能匹配到所有这类文件，例如 *门店*.xlsx', 5000);
        } else delete spec.multiple;
        renderAll(); saveDraft();
      });
      body.append(h('div', { class: 'fc-line' }, h('span', null, '多个文件'),
        h('label', { class: 'ck' }, mc, h('span', null, '把所有符合规则的文件都读进来（用于“追加行”合并）'))));
    }
    // 工作表与表头行
    let sheetCtl;
    if (smp && smp.sheets.length) {
      const opts = smp.sheets.map((s, i) => ({ value: i === 0 ? undefined : s.name, label: s.name + (i === 0 ? '（第一个）' : '') }));
      const cur = M.sampleSheet(spec.id, spec.sheet);
      sheetCtl = W.select(opts, cur && smp.sheets.indexOf(cur) > 0 ? cur.name : undefined, v => {
        if (v === undefined) delete spec.sheet; else spec.sheet = v;
        const sh = M.sampleSheet(spec.id, spec.sheet);
        if (sh) { if (sh.headerRow !== 1) spec.header_row = sh.headerRow; else delete spec.header_row; }
        renderAll(); saveDraft();
      });
    } else {
      sheetCtl = h('input', { class: 'in', value: spec.sheet === undefined ? '' : String(spec.sheet), placeholder: '不填=第一个工作表' });
      sheetCtl.addEventListener('input', () => { const v = sheetCtl.value.trim(); if (v) spec.sheet = /^\d+$/.test(v) ? Number(v) : v; else delete spec.sheet; changed(); });
    }
    const hrIn = h('input', { class: 'in num-in', type: 'number', min: 1, value: spec.header_row || 1, 'aria-label': '表头行' });
    hrIn.addEventListener('change', () => {
      const n = parseInt(hrIn.value, 10);
      if (n >= 1) { if (n === 1) delete spec.header_row; else spec.header_row = n; }
      renderAll(); saveDraft();
    });
    const sh = M.sampleSheet(spec.id, spec.sheet);
    body.append(h('div', { class: 'fc-line' }, h('span', null, '工作表'), sheetCtl));
    body.append(h('div', { class: 'fc-line' }, h('span', null, '表头行'), h('div', { class: 'row' }, hrIn,
      sh && sh.headerRow !== (spec.header_row || 1) ? h('span', { class: 'hint' }, '自动识别为第 ' + sh.headerRow + ' 行') : h('span', { class: 'hint' }, '表头下一行起为数据'))));
    // 列、敏感列、一致性
    if (sh) {
      const cols = F.columns(sh, spec.header_row || 1);
      const sens = cols.filter(c => c.sensitive);
      if (sens.length) body.append(h('div', { class: 'sens' }, '含个人信息列：' + sens.map(c => c.name + '（' + c.sensitive + '）').join('、') + '。文件只在本机读取，不会上传，也不会写进配置。'));
      if (spec.role === 'ref' && mainId && S.samples[mainId]) {
        const mh = F.headerList(M.sampleSheet(mainId), (M.fileSpec(mainId).header_row || 1)) || [];
        const rh = new Set(F.headerList(sh, spec.header_row || 1) || []);
        const same = mh.filter(x => x && rh.has(x));
        const missing = mh.filter(x => x && !rh.has(x));
        body.append(h('div', { class: 'consist' }, '与主表同名的列 ' + same.length + ' 个' +
          (missing.length ? '；主表有、这里没有：' + missing.slice(0, 8).join('、') + (missing.length > 8 ? ' 等 ' + missing.length + ' 个' : '') : '；主表的列这里都有')));
      }
      const tbl = h('table', { class: 'cols-table' });
      cols.forEach(c => tbl.append(h('tr', null, h('td', { class: 'lt' }, c.letter), h('td', { class: 'nm' }, c.name || '（无表头）'),
        h('td', null, h('span', { class: 'type t-' + c.type }, F.TYPE_LABEL[c.type])), h('td', { class: 'sm', title: c.sensitive ? '' : c.samples.join('、') }, c.sensitive ? '（不显示）' : c.samples.join('、')))));
      body.append(h('details', { class: 'cols-detail' }, h('summary', null, '查看列（' + cols.length + '）'), h('div', { class: 'cols-wrap' }, tbl)));
    }
    body.append(h('ul', { class: 'msgs file-msgs' }));
    card.append(body);
    return card;
  }

  function renderFileMsgs() {
    document.querySelectorAll('.file-card').forEach(card => {
      const ul = card.querySelector('.file-msgs');
      if (!ul || !val) return;
      ul.textContent = '';
      (val.files[card.getAttribute('data-id')] || []).filter(m => m.level === 'error' || !/规则也能匹配到|不符合文件名规则|没有文件名规则/.test(m.msg)).forEach(m => ul.append(h('li', { class: m.level }, m.msg)));
      (val.files[card.getAttribute('data-id')] || []).filter(m => /规则也能匹配到/.test(m.msg)).forEach(m => ul.append(h('li', { class: m.level }, m.msg)));
    });
  }

  function setRole(id, role) {
    S.cfg.files.forEach(f => {
      if (f.id === id) { f.role = role; if (role === 'main') delete f.multiple; }
      else if (role === 'main' && f.role === 'main') f.role = 'ref';
    });
    renderAll(); saveDraft();
  }

  function removeFile(id) {
    const refs = stepsReferencing(id);
    const label = M.fileLabel(id);
    if (!confirm('移除文件「' + label + '」？' + (refs.length ? '\n有 ' + refs.length + ' 个步骤用到了它，需要随后修改。' : ''))) return;
    S.cfg.files = S.cfg.files.filter(f => f.id !== id);
    delete S.samples[id];
    renderAll(); saveDraft();
  }

  function pickSample(id) {
    pendingSampleFor = id;
    $('fileInput').click();
  }

  /* ------------------------------------------------------------------ 步骤列表 */
  function renderSteps() {
    const list = $('stepList');
    list.textContent = '';
    editorHooks = [];
    $('stepCount').textContent = S.cfg.steps.length ? '（' + S.cfg.steps.length + '）' : '';
    $('stepsEmpty').hidden = S.cfg.steps.length > 0;
    S.cfg.steps.forEach((s, i) => list.append(stepCard(s, i)));
    if (openIdx !== null && openIdx >= S.cfg.steps.length) openIdx = null;
    if (openIdx !== null) renderEditor(openIdx);
  }

  function statusOf(i) {
    const msgs = (val && val.steps[i]) || [];
    if (msgs.some(m => m.level === 'error')) return 'error';
    if (msgs.some(m => m.level === 'warn')) return 'warn';
    return 'ok';
  }

  function stepCard(s, i) {
    const def = M.OPS[s.op];
    const li = h('li', { class: 'step-card' + (openIdx === i ? ' open' : '') + (s.enabled === false ? ' off' : ''), 'data-i': i });
    const sw = h('label', { class: 'switch', title: s.enabled === false ? '已停用（不执行）' : '启用中' },
      h('input', { type: 'checkbox', checked: s.enabled !== false, 'aria-label': '启用这一步' }), h('span'));
    sw.querySelector('input').addEventListener('change', e => {
      if (e.target.checked) delete s.enabled; else s.enabled = false;
      li.classList.toggle('off', s.enabled === false);
      sw.title = s.enabled === false ? '已停用（不执行）' : '启用中';
      changed();
    });
    const handle = h('button', { type: 'button', class: 'sc-handle', title: '拖动调整顺序', 'aria-label': '拖动调整顺序' }, '⋮⋮');
    handle.addEventListener('mousedown', () => { li.draggable = true; });
    handle.addEventListener('mouseup', () => { li.draggable = false; });
    const main = h('button', { type: 'button', class: 'sc-main', 'aria-expanded': openIdx === i ? 'true' : 'false', onclick: () => toggleStep(i) },
      h('div', { class: 'sc-title' }, h('span', { class: 'sc-t' }), h('span', { class: 'sc-op' }, def ? def.label : s.op)),
      h('div', { class: 'sc-sum' }));
    const st = statusOf(i);
    li.append(h('div', { class: 'sc-head' },
      handle,
      h('span', { class: 'sc-no' }, String(i + 1)),
      h('span', { class: 'sc-status ' + st, title: st === 'ok' ? '参数完整' : st === 'warn' ? '有需要留意的地方' : '有错误' }, st === 'ok' ? '✓' : st === 'warn' ? '!' : '✕'),
      main,
      h('div', { class: 'sc-tools' },
        sw,
        h('button', { type: 'button', class: 'icon-btn', title: '上移', 'aria-label': '上移', disabled: i === 0, onclick: () => moveStep(i, i - 1) }, '↑'),
        h('button', { type: 'button', class: 'icon-btn', title: '下移', 'aria-label': '下移', disabled: i === S.cfg.steps.length - 1, onclick: () => moveStep(i, i + 1) }, '↓'),
        h('button', { type: 'button', class: 'icon-btn', title: '复制这一步', 'aria-label': '复制这一步', onclick: () => duplicateStep(i) }, '⧉'),
        h('button', { type: 'button', class: 'icon-btn del', title: '删除这一步', 'aria-label': '删除这一步', onclick: () => deleteStep(i) }, '🗑'))));
    fillHead(li, s, i);
    // 拖动排序
    li.addEventListener('dragstart', e => { li.classList.add('dragging'); e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', String(i)); });
    li.addEventListener('dragend', () => { li.classList.remove('dragging'); li.draggable = false; clearDrop(); });
    li.addEventListener('dragover', e => {
      if (!document.querySelector('.step-card.dragging')) return;
      e.preventDefault();
      const r = li.getBoundingClientRect();
      clearDrop();
      li.classList.add(e.clientY < r.top + r.height / 2 ? 'drop-before' : 'drop-after');
    });
    li.addEventListener('drop', e => {
      e.preventDefault();
      const from = Number(e.dataTransfer.getData('text/plain'));
      const r = li.getBoundingClientRect();
      let to = i + (e.clientY < r.top + r.height / 2 ? 0 : 1);
      if (from < to) to--;
      clearDrop();
      if (!isNaN(from) && from !== to) moveStep(from, to);
    });
    return li;
  }
  function clearDrop() { document.querySelectorAll('.drop-before,.drop-after').forEach(x => x.classList.remove('drop-before', 'drop-after')); }

  function fillHead(li, s, i) {
    const def = M.OPS[s.op];
    li.querySelector('.sc-t').textContent = s.title || (def ? def.label : s.op);
    li.querySelector('.sc-op').hidden = !s.title;
    li.querySelector('.sc-sum').textContent = M.summary(s);
    li.querySelector('.sc-sum').title = M.summary(s);
    const st = statusOf(i);
    const badge = li.querySelector('.sc-status');
    badge.className = 'sc-status ' + st;
    badge.textContent = st === 'ok' ? '✓' : st === 'warn' ? '!' : '✕';
    badge.title = ((val && val.steps[i]) || []).map(m => m.msg).join('\n') || '参数完整';
  }

  function updateStepHeads() {
    document.querySelectorAll('#stepList > .step-card').forEach(li => {
      const i = Number(li.getAttribute('data-i'));
      if (S.cfg.steps[i]) fillHead(li, S.cfg.steps[i], i);
    });
  }

  function toggleStep(i) {
    openIdx = openIdx === i ? null : i;
    renderSteps();
    if (openIdx !== null) {
      const li = document.querySelector('.step-card[data-i="' + i + '"]');
      if (li) li.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  }

  function moveStep(from, to) {
    if (to < 0 || to >= S.cfg.steps.length) return;
    const [s] = S.cfg.steps.splice(from, 1);
    S.cfg.steps.splice(to, 0, s);
    if (openIdx === from) openIdx = to;
    else if (openIdx !== null) {
      if (from < openIdx && to >= openIdx) openIdx--;
      else if (from > openIdx && to <= openIdx) openIdx++;
    }
    invalidate(); renderSteps(); renderCheckBox(); saveDraft();
  }

  function duplicateStep(i) {
    const copy = JSON.parse(JSON.stringify(S.cfg.steps[i]));
    if (copy.op === 'match') copy.id = U.uniqueName(copy.id || 'm', new Set(S.cfg.steps.filter(s => s.op === 'match').map(s => s.id)));
    S.cfg.steps.splice(i + 1, 0, copy);
    openIdx = i + 1;
    invalidate(); renderSteps(); renderCheckBox(); saveDraft();
  }

  function deleteStep(i) {
    const s = S.cfg.steps[i];
    const def = M.OPS[s.op];
    if (!confirm('删除第 ' + (i + 1) + ' 步「' + (s.title || (def ? def.label : s.op)) + '」？')) return;
    S.cfg.steps.splice(i, 1);
    if (openIdx === i) openIdx = null; else if (openIdx !== null && openIdx > i) openIdx--;
    invalidate(); renderSteps(); renderCheckBox(); saveDraft();
  }

  /* ------------------------------------------------------------------ 添加步骤 */
  function openOpDialog() {
    const box = $('opGroups');
    box.textContent = '';
    M.GROUPS.forEach(g => {
      const ops = EFA.WF_CATALOG.ops.filter(d => d.group === g.id);
      box.append(h('div', { class: 'op-group' }, h('h3', null, g.label),
        h('div', { class: 'op-grid' }, ops.map(d => h('button', { type: 'button', class: 'op-btn', onclick: () => { $('opDialog').close(); addStep(d.op); } },
          h('b', null, d.label), h('span', null, d.desc))))));
    });
    $('opDialog').showModal();
  }

  function addStep(op) {
    const s = { op };
    const at = openIdx === null ? S.cfg.steps.length : openIdx + 1;
    if (op === 'match') {
      const taken = new Set(S.cfg.steps.filter(x => x.op === 'match').map(x => x.id));
      let n = 1;
      while (taken.has('m' + n)) n++;
      s.id = 'm' + n;
      const ref = S.cfg.files.find(f => f.role === 'ref');
      if (ref) s.ref = ref.id;
      s.main_keys = [];
    }
    if (op === 'copy_matched' || op === 'compare') {
      const ms = M.matchesBefore(at);
      if (ms.length) s.match = ms[ms.length - 1].id;
    }
    if (op === 'validate') s.checks = [{ type: 'not_empty' }];
    if (op === 'group_summary') { s.group_by = []; s.metrics = [{ func: 'sum', column: '' }]; }
    if (op === 'insert_column') s.end = true;
    S.cfg.steps.splice(at, 0, s);
    openIdx = at;
    invalidate(); renderSteps(); renderCheckBox(); saveDraft();
    const li = document.querySelector('.step-card[data-i="' + at + '"]');
    if (li) {
      li.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      const first = li.querySelector('.sc-body input, .sc-body select');
      if (first) setTimeout(() => first.focus(), 50);
    }
  }

  /* ------------------------------------------------------------------ 步骤编辑器 */
  function renderEditor(i) {
    const li = document.querySelector('.step-card[data-i="' + i + '"]');
    if (!li) return;
    const s = S.cfg.steps[i];
    const def = M.OPS[s.op];
    editorHooks = [];
    const old = li.querySelector('.sc-body');
    if (old) old.remove();
    const body = h('div', { class: 'sc-body' });
    const lists = h('div', { hidden: true });
    const listIds = {};
    const ctx = {
      step: s, index: i,
      set(k, v) { if (v === undefined) delete s[k]; else s[k] = v; changed(); },
      changed() { changed(); },
      rerender() { invalidate(); renderEditor(i); updateStepHeads(); renderCheckBox(); saveDraft(); },
      onRefresh(fn) { editorHooks.push(fn); },
      headersOf(source) {
        const sm = sim();
        if (source === 'ref') {
          if (s.op === 'match') return sm.headersAt(i, s.ref, s.ref_sheet, s.ref_header_row);
          const info = s.match !== undefined ? M.refOfMatch(i, s.match, sm) : null;
          return info ? info.headers : null;
        }
        if (source === 'refFile') return sm.headersAt(i, s.ref, s.ref_sheet, s.ref_header_row);
        if (source === 'from') return M.fromHeaders(i, s, sm);
        if (source === 'ck') return sm.headersAt(i, s.file || M.mainId(), s.sheet, s.header_row);
        const t = M.stepTarget(s);
        return sm.headersAt(i, t.fileId, t.sheet, t.headerRow);
      },
      datalist(source) {
        const key = 'c:' + source;
        if (listIds[key]) return listIds[key];
        const id = 'dl' + Math.random().toString(36).slice(2, 9);
        listIds[key] = id;
        const dl = h('datalist', { id });
        const fill = () => {
          dl.textContent = '';
          const hdr = ctx.headersOf(source) || [];
          hdr.forEach((name, k) => { if (name) dl.append(h('option', { value: name, label: U.colLetter(k) + ' 列' })); });
        };
        fill();
        editorHooks.push(fill);
        lists.append(dl);
        return id;
      },
      sheetList(fileId) {
        const fid = fileId || M.mainId();
        const key = 's:' + fid;
        if (listIds[key]) return listIds[key];
        const id = 'dl' + Math.random().toString(36).slice(2, 9);
        listIds[key] = id;
        const names = (S.samples[fid] ? S.samples[fid].sheets.map(x => x.name) : []).concat(fid === M.mainId() ? sim().extraSheets : []);
        lists.append(h('datalist', { id }, names.map(n => h('option', { value: n }))));
        return id;
      },
    };
    if (!def) {
      body.append(h('div', { class: 'notice err' }, '执行器没有操作「' + s.op + '」。可以删除这一步，或在下方“其他参数”里查看内容。'));
    } else {
      body.append(h('p', { class: 'sc-desc' }, def.desc));
      body.append(h('ul', { class: 'msgs step-msgs' }));
      def.fields.forEach(fd => {
        const ctl = W.render(fd, ctx);
        const boolish = fd.type === 'bool' || fd.type === 'optText' || fd.type === 'notFound';
        body.append(h('div', { class: 'fld' + (boolish ? ' boolish' : '') },
          h('div', { class: 'fl' }, boolish ? '' : fd.label, fd.required && !boolish ? h('span', { class: 'req', title: '必填' }, '*') : null),
          h('div', { class: 'fc' }, ctl, fd.help ? h('div', { class: 'hint' }, fd.help) : null)));
      });
    }
    body.append(advancedSection(s, i, def, ctx));
    body.append(h('div', { class: 'sc-foot' }, h('span', { class: 'hint' }, '修改会自动保存在本浏览器里'),
      h('button', { type: 'button', class: 'btn btn-ghost btn-sm', onclick: () => toggleStep(i) }, '完成')));
    body.append(lists);
    li.append(body);
    updateEditorMsgs();
  }

  function advancedSection(s, i, def, ctx) {
    const det = h('details', { class: 'sc-adv' });
    const known = def ? M.allowedKeys(def) : new Set(M.COMMON);
    const extraKeys = Object.keys(s).filter(k => !known.has(k) || (def && (def.hidden || []).includes(k) && !(def.op === 'copy_matched' && k === 'from' && s.columns !== undefined)));
    if (s.title || s.note || s.sheet !== undefined || s.header_row !== undefined || s.file !== undefined || extraKeys.length) det.open = true;
    det.append(h('summary', null, '更多设置（步骤名称、备注、工作表、表头行）'));
    const txt = (key, label, ph) => {
      const inp = h('input', { class: 'in', value: s[key] === undefined ? '' : String(s[key]), placeholder: ph || '' });
      inp.addEventListener('input', () => { const v = inp.value.trim(); if (v) s[key] = v; else delete s[key]; changed(); });
      return h('div', { class: 'fld' }, h('div', { class: 'fl' }, label), h('div', { class: 'fc' }, inp));
    };
    det.append(txt('title', '步骤名称', '不填=自动生成说明；会显示在校验报告里'));
    det.append(txt('note', '备注', '给自己看的说明'));
    if (def && def.op === 'validate') {
      const files = S.cfg.files;
      det.append(h('div', { class: 'fld' }, h('div', { class: 'fl' }, '检查哪个文件'), h('div', { class: 'fc' },
        W.select([{ value: undefined, label: '主表（处理后的结果）' }].concat(files.filter(f => f.role !== 'main').map(f => ({ value: f.id, label: f.label || f.id }))), s.file, v => ctx.set('file', v)))));
    }
    const sheetIn = h('input', { class: 'in', value: s.sheet === undefined ? '' : String(s.sheet), placeholder: '不填=文件设置的工作表', list: ctx.sheetList(s.file || null) });
    sheetIn.addEventListener('input', () => { const v = sheetIn.value.trim(); if (v) s.sheet = v; else delete s.sheet; changed(); });
    det.append(h('div', { class: 'fld' }, h('div', { class: 'fl' }, '工作表'), h('div', { class: 'fc' }, sheetIn, h('div', { class: 'hint' }, '这一步要处理同一文件里的其他工作表时填写'))));
    const hrIn = h('input', { class: 'in num-in', type: 'number', min: 1, value: s.header_row === undefined ? '' : s.header_row, placeholder: '默认' });
    hrIn.addEventListener('input', () => { const n = parseInt(hrIn.value, 10); if (n >= 1) s.header_row = n; else delete s.header_row; changed(); });
    det.append(h('div', { class: 'fld' }, h('div', { class: 'fl' }, '表头行'), h('div', { class: 'fc' }, hrIn)));
    // 其他参数（界面没有提供编辑的参数，原样保留）
    const extra = {};
    extraKeys.forEach(k => { extra[k] = s[k]; });
    const ta = h('textarea', { class: 'extra-json', spellcheck: 'false', placeholder: '{ }' });
    ta.value = extraKeys.length ? JSON.stringify(extra, null, 2) : '';
    const taMsg = h('div', { class: 'hint' }, '高级：界面上没有的参数可以在这里用 JSON 填写（例如 from / to 写法）。');
    ta.addEventListener('input', () => {
      let obj;
      try { obj = ta.value.trim() ? JSON.parse(ta.value) : {}; } catch (e) { ta.classList.add('bad'); taMsg.textContent = 'JSON 写法不正确：' + e.message; return; }
      if (!obj || typeof obj !== 'object' || Array.isArray(obj)) { ta.classList.add('bad'); taMsg.textContent = '应是 { } 形式的对象'; return; }
      ta.classList.remove('bad');
      taMsg.textContent = '高级：界面上没有的参数可以在这里用 JSON 填写。';
      Object.keys(extra).forEach(k => { if (!(k in obj)) delete s[k]; });
      Object.keys(obj).forEach(k => { if (k !== 'op') s[k] = obj[k]; });
      Object.keys(extra).forEach(k => delete extra[k]);
      Object.assign(extra, obj);
      changed();
    });
    det.append(h('div', { class: 'fld' }, h('div', { class: 'fl' }, '其他参数'), h('div', { class: 'fc' }, ta, taMsg)));
    return det;
  }

  function updateEditorMsgs() {
    if (openIdx === null) return;
    const li = document.querySelector('.step-card[data-i="' + openIdx + '"]');
    const ul = li && li.querySelector('.step-msgs');
    if (!ul) return;
    ul.textContent = '';
    ((val && val.steps[openIdx]) || []).forEach(m => ul.append(h('li', { class: m.level }, (m.level === 'error' ? '✕ ' : '! ') + m.msg)));
  }

  /* ------------------------------------------------------------------ 导出 */
  function renderExport() {
    $('cfgName').value = S.cfg.name || '';
    $('cfgDesc').value = S.cfg.description || '';
    $('cfgOutput').value = (S.cfg.output || {}).file_name || '';
    $('cfgRecalc').value = (S.cfg.options || {}).recalc || 'auto';
    renderCheckBox();
    renderToolkitLine();
  }

  function renderCheckBox() {
    const box = $('checkBox');
    box.textContent = '';
    if (!val) return;
    const items = [];
    val.global.forEach(m => items.push({ m, go: null }));
    Object.keys(val.files).forEach(id => val.files[id].filter(m => m.level === 'error').forEach(m => items.push({ m: { level: m.level, msg: '文件「' + M.fileLabel(id) + '」：' + m.msg }, go: null })));
    val.steps.forEach((arr, i) => arr.filter(m => m.level === 'error').forEach(m => items.push({ m: { level: 'error', msg: '第 ' + (i + 1) + ' 步：' + m.msg }, go: i })));
    const errs = items.filter(x => x.m.level === 'error');
    const warnSteps = val.steps.filter(a => a.some(m => m.level === 'warn')).length;
    if (!errs.length) {
      box.className = 'check-box ' + (items.length || warnSteps ? 'warn' : 'ok');
      box.append(h('div', null, '✓ 检查通过，可以保存' + (warnSteps ? '（' + warnSteps + ' 个步骤有提示，建议看一下）' : '')));
      const w = items.filter(x => x.m.level === 'warn');
      if (w.length) box.append(h('ul', null, w.map(x => h('li', null, x.m.msg))));
    } else {
      box.className = 'check-box error';
      box.append(h('div', null, '还有 ' + errs.length + ' 处需要修改：'));
      box.append(h('ul', null, errs.slice(0, 8).map(x => h('li', null, x.go === null ? x.m.msg : h('a', { onclick: () => { openIdx = x.go; renderSteps(); const li = document.querySelector('.step-card[data-i="' + x.go + '"]'); if (li) li.scrollIntoView({ block: 'center', behavior: 'smooth' }); } }, x.m.msg))),
        errs.length > 8 ? h('li', null, '…') : null));
    }
  }

  function readyToExport() {
    invalidate();
    renderCheckBox();
    if (val.errors) {
      toast('还有 ' + val.errors + ' 处需要修改，改好后才能保存（执行器会拒绝有错误的配置）', 5000);
      $('checkBox').scrollIntoView({ block: 'center', behavior: 'smooth' });
      return false;
    }
    return true;
  }

  async function saveJson() {
    if (!readyToExport()) return;
    const text = M.toJson();
    const name = safeName(S.cfg.name) + '.json';
    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: '工作流配置', accept: { 'application/json': ['.json'] } }] });
        const w = await handle.createWritable();
        await w.write(text);
        await w.close();
        toast('已保存「' + handle.name + '」。把它放进工具包的 configs 文件夹就能用了', 5000);
        return;
      } catch (e) {
        if (e && e.name === 'AbortError') return;
      }
    }
    download(new Blob([text], { type: 'application/json' }), name);
    toast('已下载「' + name + '」。把它放进工具包的 configs 文件夹就能用了', 5000);
  }

  async function buildPackage() {
    if (!readyToExport()) return;
    if (!toolkit) {
      toast('请选择“Excel 工作流工具”的 zip 压缩包（只需要选一次）', 4000);
      $('toolkitInput').click();
      buildPackage.pending = true;
      return;
    }
    try {
      const name = safeName(S.cfg.name);
      const blob = Z.build(toolkit.buf, [{ path: 'configs/' + name + '.json', text: M.toJson() }]);
      download(blob, 'Excel工作流工具_' + name + '.zip');
      toast('已生成「Excel工作流工具_' + name + '.zip」：解压后把表格放进 input，双击 run.bat 即可', 6000);
    } catch (e) {
      toast(e.message || String(e), 6000);
    }
  }

  async function onToolkitChosen(file) {
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      Z.inspect(buf);
      toolkit = { name: file.name, size: file.size, buf };
      Z.saveToolkit({ name: file.name, size: file.size, buf, savedAt: Date.now() });
      renderToolkitLine();
      if (buildPackage.pending) { buildPackage.pending = false; buildPackage(); }
    } catch (e) {
      buildPackage.pending = false;
      toast(e.message || String(e), 6000);
    }
  }

  function renderToolkitLine() {
    const line = $('toolkitLine');
    line.textContent = '';
    if (toolkit) {
      line.append('使用工具包：' + toolkit.name + '（' + fmtSize(toolkit.size) + '）',
        h('button', { type: 'button', onclick: () => $('toolkitInput').click() }, '更换'),
        h('button', { type: 'button', onclick: () => { toolkit = null; Z.forgetToolkit(); renderToolkitLine(); } }, '忘记'));
    } else {
      line.append('需要“Excel工作流工具_v*.zip”（第一次生成时选择，之后会记住）');
    }
  }

  function showJson() {
    $('jsonView').textContent = M.toJson();
    $('jsonDialog').showModal();
  }

  /* ------------------------------------------------------------------ AI 生成步骤 */
  let aiResult = null, aiAbort = null;

  function openAi() {
    const L = EFA.LLM;
    $('aiNoKey').hidden = L.isConfigured();
    const noFiles = !S.cfg.files.length;
    const missing = S.cfg.files.filter(f => !S.samples[f.id]);
    $('aiNoSample').hidden = !noFiles && !missing.length;
    $('aiNoSample').textContent = noFiles ? '还没有文件：请先在左侧载入样例文件（至少主表），AI 才知道表格有哪些列。'
      : '「' + missing.map(f => f.label || f.id).join('」「') + '」还没有载入样例：AI 不知道它有哪些列，生成的步骤可能用错列名。建议先在左侧载入样例。';
    $('aiMode').value = 'append';
    $('aiMode').disabled = !S.cfg.steps.length;
    $('aiDialog').showModal();
    setTimeout(() => $('aiRequest').focus(), 50);
  }

  async function runAi() {
    const req = $('aiRequest').value.trim();
    if (!req) { $('aiRequest').focus(); return; }
    if (!EFA.LLM.isConfigured()) { EFA.AISettings.open(() => { $('aiNoKey').hidden = EFA.LLM.isConfigured(); }); return; }
    if (!S.cfg.files.length) { toast('请先载入样例文件（至少主表）'); return; }
    if (aiAbort) aiAbort.abort();
    aiAbort = new AbortController();
    aiResult = null;
    $('aiResult').textContent = '';
    $('aiApply').hidden = true;
    $('aiGo').disabled = true;
    const status = $('aiStatus');
    status.className = 'ai-status busy';
    status.textContent = '正在生成…';
    try {
      const res = await EFA.WFAI.generate(req, { mode: $('aiMode').value, signal: aiAbort.signal, onStatus: t => { status.textContent = t; } });
      aiResult = res;
      status.className = 'ai-status';
      status.textContent = '';
      renderAiResult();
    } catch (e) {
      status.className = 'ai-status err';
      status.textContent = (e.kind === 'need_clarification' ? '需要补充说明：' : e.kind === 'unsupported' ? '做不到：' : '') + (e.message || String(e));
      if (e.kind === 'nokey') EFA.AISettings.open();
    } finally {
      $('aiGo').disabled = false;
      aiAbort = null;
    }
  }

  function renderAiResult() {
    const box = $('aiResult');
    box.textContent = '';
    const r = aiResult;
    if (!r) return;
    const nErr = r.issues.filter(a => a.some(m => m.level === 'error')).length;
    box.append(h('div', { class: 'ai-head' }, '生成了 ' + r.steps.length + ' 个步骤' + (nErr ? '，其中 ' + nErr + ' 个还有问题（可以先加入再修改）' : '，检查通过'),
      h('span', { class: 'hint' }, ' · ' + r.model + ' · 调用 ' + r.usage.calls + ' 次，约 ' + (r.usage.inTok + r.usage.outTok || r.usage.estimate).toLocaleString() + ' tokens')));
    const list = h('ol', { class: 'ai-steps' });
    r.steps.forEach((s, i) => {
      const def = M.OPS[s.op];
      const msgs = r.issues[i] || [];
      const st = msgs.some(m => m.level === 'error') ? 'error' : msgs.length ? 'warn' : 'ok';
      const cb = h('input', { type: 'checkbox', checked: true, 'data-i': i, 'aria-label': '加入这一步' });
      list.append(h('li', { class: 'ai-step' },
        h('label', { class: 'ai-pick' }, cb,
          h('span', { class: 'sc-status ' + st }, st === 'ok' ? '✓' : st === 'warn' ? '!' : '✕'),
          h('span', { class: 'ai-txt' }, h('b', null, s.title || (def ? def.label : s.op)), h('span', { class: 'sc-op' }, def ? def.label : '未知操作 ' + s.op),
            h('span', { class: 'ai-sum' }, def ? M.summary(s) : ''))),
        msgs.length ? h('ul', { class: 'msgs' }, msgs.map(m => h('li', { class: m.level }, m.msg))) : null));
    });
    box.append(list);
    if (r.explanation) box.append(h('div', { class: 'ai-expl' }, h('b', null, '说明：'), r.explanation));
    if (r.assumptions.length) box.append(h('div', { class: 'ai-expl' }, h('b', null, 'AI 的假设（请核对）：'), h('ul', null, r.assumptions.map(a => h('li', null, a)))));
    $('aiApply').hidden = false;
  }

  function applyAi() {
    if (!aiResult) return;
    const picked = [...document.querySelectorAll('#aiResult input[type=checkbox]')].filter(c => c.checked).map(c => Number(c.getAttribute('data-i')));
    if (!picked.length) { toast('没有勾选步骤'); return; }
    const steps = picked.map(i => JSON.parse(JSON.stringify(aiResult.steps[i])));
    if ($('aiMode').value === 'replace' && S.cfg.steps.length) {
      if (!confirm('用生成的 ' + steps.length + ' 个步骤替换现有的 ' + S.cfg.steps.length + ' 个步骤？')) return;
      S.cfg.steps = steps;
    } else {
      S.cfg.steps = S.cfg.steps.concat(steps);
    }
    openIdx = null;
    $('aiDialog').close();
    renderAll(); saveDraft();
    const bad = M.validate().steps.filter(a => a.some(m => m.level === 'error')).length;
    toast('已加入 ' + steps.length + ' 个步骤' + (bad ? '，有 ' + bad + ' 个步骤标红，请展开修改' : '，请检查每一步的参数'), 5000);
    const first = document.querySelector('#stepList > .step-card:last-child');
    if (first) first.scrollIntoView({ block: 'center', behavior: 'smooth' });
  }

  /* ------------------------------------------------------------------ 导入 / 新建 / 示例 */
  function loadConfig(obj, source) {
    let cfg;
    try { cfg = M.importConfig(obj); } catch (e) { toast(e.message, 6000); return false; }
    if ((S.cfg.steps.length || S.cfg.files.length) && !confirm('用「' + (cfg.name || source) + '」替换当前正在编辑的工作流？')) return false;
    const keep = {};
    cfg.files.forEach(f => { if (S.samples[f.id]) keep[f.id] = S.samples[f.id]; });
    S.cfg = cfg;
    S.samples = keep;
    openIdx = null;
    renderAll();
    saveDraft();
    return true;
  }

  async function openConfigFile(file) {
    if (!file) return;
    let text = await file.text();
    text = text.replace(/^﻿/, '');
    let obj;
    try { obj = JSON.parse(text); } catch (e) { toast('「' + file.name + '」不是有效的配置文件（JSON 格式不正确）', 6000); return; }
    if (loadConfig(obj, file.name)) toast('已打开「' + file.name + '」' + (S.cfg.files.some(f => !S.samples[f.id]) ? '。可以给每个文件“载入样例”，方便选择列名和检查' : ''), 5000);
  }

  function renderExamples() {
    const box = $('exampleList');
    box.textContent = '';
    (EFA.WF_EXAMPLES || []).forEach(ex => box.append(h('button', { type: 'button', onclick: () => {
      $('exampleMenu').open = false;
      if (loadConfig(JSON.parse(JSON.stringify(ex)), ex.name)) toast('已载入示例「' + ex.name + '」，可以在此基础上修改');
    } }, h('b', null, ex.name.replace(/^示例：/, '')), h('span', null, ex.description || ''))));
  }

  function newConfig() {
    if ((S.cfg.steps.length || S.cfg.files.length) && !confirm('清空当前工作流，重新开始？（建议先保存配置文件）')) return;
    S.cfg = M.newConfig();
    S.samples = {};
    openIdx = null;
    renderAll();
    saveDraft();
  }

  /* ------------------------------------------------------------------ 初始化 */
  function bind() {
    const drop = $('fileDrop'), input = $('fileInput');
    drop.addEventListener('click', () => { pendingSampleFor = null; input.click(); });
    drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pendingSampleFor = null; input.click(); } });
    input.addEventListener('change', () => { const t = pendingSampleFor; pendingSampleFor = null; addSampleFiles(input.files, t); input.value = ''; });
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('drag'); }));
    drop.addEventListener('drop', e => addSampleFiles(e.dataTransfer.files));
    // 拖到页面其他地方：.json 当作配置打开，.xlsx 当作样例
    window.addEventListener('dragover', e => e.preventDefault());
    window.addEventListener('drop', e => {
      if (e.defaultPrevented || !e.dataTransfer || !e.dataTransfer.files.length) return;
      e.preventDefault();
      const files = Array.from(e.dataTransfer.files);
      const json = files.find(f => /\.json$/i.test(f.name));
      if (json) openConfigFile(json); else addSampleFiles(files);
    });

    $('btnAddStep').addEventListener('click', openOpDialog);
    $('btnAi').addEventListener('click', openAi);
    $('btnAiFirst').addEventListener('click', openAi);
    $('aiGo').addEventListener('click', runAi);
    $('aiApply').addEventListener('click', applyAi);
    $('aiSettings').addEventListener('click', () => EFA.AISettings.open(() => { $('aiNoKey').hidden = EFA.LLM.isConfigured(); }));
    $('aiSetKey').addEventListener('click', () => EFA.AISettings.open(() => { $('aiNoKey').hidden = EFA.LLM.isConfigured(); }));
    $('aiDialog').addEventListener('close', () => { if (aiAbort) aiAbort.abort(); });
    $('aiRequest').addEventListener('keydown', e => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); runAi(); } });
    $('btnAddFirst').addEventListener('click', openOpDialog);
    $('btnOpen').addEventListener('click', () => $('openInput').click());
    $('openInput').addEventListener('change', e => { openConfigFile(e.target.files[0]); e.target.value = ''; });
    $('btnNew').addEventListener('click', newConfig);
    $('btnSave').addEventListener('click', saveJson);
    $('btnPackage').addEventListener('click', buildPackage);
    $('btnJson').addEventListener('click', showJson);
    $('toolkitInput').addEventListener('change', e => { onToolkitChosen(e.target.files[0]); e.target.value = ''; });
    $('btnCopyJson').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText($('jsonView').textContent); toast('已复制'); } catch (e) { toast('复制失败，请手动选择文字复制'); }
    });
    $('cfgName').addEventListener('input', e => { S.cfg.name = e.target.value; changed(); });
    $('cfgDesc').addEventListener('input', e => { S.cfg.description = e.target.value; saveDraft(); });
    $('cfgOutput').addEventListener('input', e => {
      const v = e.target.value.trim();
      if (v) S.cfg.output = Object.assign({}, S.cfg.output, { file_name: v }); else if (S.cfg.output) delete S.cfg.output.file_name;
      saveDraft();
    });
    $('cfgRecalc').addEventListener('change', e => { S.cfg.options = Object.assign({}, S.cfg.options, { recalc: e.target.value }); changed(); });
    document.addEventListener('click', e => { const m = $('exampleMenu'); if (m.open && !m.contains(e.target)) m.open = false; });
  }

  async function init() {
    bind();
    renderExamples();
    let restored = false;
    try {
      const raw = localStorage.getItem(DRAFT_KEY);
      if (raw) {
        const cfg = M.importConfig(JSON.parse(raw));
        if (cfg.steps.length || cfg.files.length) { S.cfg = cfg; restored = true; }
      }
    } catch (e) { /* 草稿损坏或存储不可用：从空白开始 */ }
    renderAll();
    if (restored) toast('已恢复上次编辑的工作流' + (S.cfg.files.length ? '（样例文件需要重新载入）' : ''), 4500);
    const rec = await Z.loadToolkit();
    if (rec && rec.buf) { toolkit = rec; renderToolkitLine(); }
  }

  EFA.WFApp = { state: S, renderAll, addStep, loadConfig };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})(window.EFA = window.EFA || {});
