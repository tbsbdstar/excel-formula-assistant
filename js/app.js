/* 页面交互 */
(function (EFA) {
  'use strict';
  const U = EFA.U, W = EFA.Workbook, F = EFA.Formula, T = EFA.Templates, L = EFA.LLM;
  const $ = id => document.getElementById(id);
  const esc = U.escapeHtml;

  const PREVIEW_ROWS = 100;
  const state = { wb: null, view: { type: 'sheet' }, mode: 'value', busy: false, abort: null };

  const TYPE_LABEL = { number: '数字', date: '日期', text: '文本', textnum: '文本数字', bool: '逻辑', mixed: '混合', empty: '空' };
  const KIND_LABEL = { column: '新增列', scalar: '单个结果', group: '分组汇总表' };

  /* ======================= 通用 ======================= */

  let toastTimer = null;
  function toast(msg, ms) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), ms || 2600);
  }

  const activeSheet = () => state.wb.sheets[state.wb.activeIndex];

  /* ======================= 载入文件 ======================= */

  async function loadFile(file) {
    if (!file) return;
    if (state.wb && state.wb.ops.length && !confirm('更换文件会清空已添加的计算，确定吗？')) return;
    if (file.size > 60 * 1024 * 1024) { toast('文件超过 60MB，浏览器处理会很慢，建议先删减数据'); return; }
    setLoadingData('正在读取「' + file.name + '」…');
    await U.sleep(30);
    try {
      const wb = await W.readFile(file);
      onLoaded(wb);
    } catch (e) {
      console.error(e);
      resetDataView();
      toast(e.message || '文件读取失败', 4000);
    }
  }

  async function loadSample() {
    setLoadingData('正在生成示例数据…');
    await U.sleep(30);
    try {
      const buf = EFA.Samples.buildSampleWorkbook();
      onLoaded(W.fromBuffer(buf, '示例数据.xlsx', 'xlsx'));
    } catch (e) {
      console.error(e);
      resetDataView();
      toast('示例数据生成失败：' + e.message, 4000);
    }
  }

  function setLoadingData(msg) {
    const dz = $('dropzone');
    dz.querySelector('.drop-inner').innerHTML = '<div class="state" style="justify-content:center"><div class="spinner"></div><div><h3>' + esc(msg) + '</h3><p>数据量大时需要几秒钟</p></div></div>';
  }

  const DROP_HTML = document.getElementById('dropzone').querySelector('.drop-inner').innerHTML;
  function resetDataView() {
    $('dropzone').querySelector('.drop-inner').innerHTML = DROP_HTML;
    bindDropButtons();
  }

  function onLoaded(wb) {
    state.wb = wb;
    state.view = { type: 'sheet' };
    $('dropzone').hidden = true;
    resetDataView();
    $('dataView').hidden = false;
    $('reqInput').disabled = false;
    $('btnGenerate').disabled = false;
    $('reqInput').placeholder = '用一句话描述需求，例如：' + (T.examples(activeSheet(), wb)[0] || '计算销售额的合计');
    $('resultArea').innerHTML = '';
    renderAll();
    if (wb.warnings.length) toast(wb.warnings.join('；'), 4000);
    $('reqInput').focus();
  }

  /* ======================= 渲染：数据区 ======================= */

  function renderAll() {
    renderFileBar();
    renderTabs();
    renderView();
    renderChips();
    renderApplied();
  }

  function renderFileBar() {
    const wb = state.wb;
    $('fileName').textContent = wb.fileName;
    const rows = wb.sheets.reduce((n, s) => n + Math.max(0, s.dataEnd - s.dataStart + 1), 0);
    $('fileMeta').textContent = wb.sheets.length + ' 个工作表 · 共 ' + rows.toLocaleString() + ' 行数据' + (wb.kind !== 'xlsx' ? ' · 下载时将转换为 .xlsx' : '');
    document.querySelector('.file-icon').textContent = wb.kind;
  }

  function renderTabs() {
    const wb = state.wb;
    const tabs = $('sheetTabs');
    let html = '';
    wb.sheets.forEach((s, i) => {
      const active = state.view.type === 'sheet' && i === wb.activeIndex;
      const added = s.columns.filter(c => c.added).length;
      html += '<button class="sheet-tab' + (active ? ' active' : '') + '" role="tab" aria-selected="' + active + '" data-sheet="' + i + '">' + esc(s.name) +
        (added ? '<span class="tag">+' + added + '列</span>' : '') + '</button>';
    });
    wb.outputs.forEach(o => {
      const active = state.view.type === 'output' && state.view.name === o.name;
      html += '<button class="sheet-tab' + (active ? ' active' : '') + '" role="tab" aria-selected="' + active + '" data-output="' + esc(o.name) + '">' + esc(o.name) + '<span class="tag">新表</span></button>';
    });
    tabs.innerHTML = html;
  }

  function renderView() {
    if (state.view.type === 'output') {
      const o = state.wb.outputs.find(x => x.name === state.view.name);
      if (o) return renderOutput(o);
      state.view = { type: 'sheet' };
    }
    renderSheetMeta();
    renderSheetTable();
  }

  function renderSheetMeta() {
    const s = activeSheet();
    const meta = $('sheetMeta');
    if (s.lastRow < 0) { meta.innerHTML = '<span>这个工作表是空的</span>'; return; }
    let opts = '';
    for (let r = 0; r <= Math.min(s.lastRow, 19); r++) opts += '<option value="' + r + '"' + (r === s.headerRow ? ' selected' : '') + '>第 ' + (r + 1) + ' 行</option>';
    const n = Math.max(0, s.dataEnd - s.dataStart + 1);
    let html = '<label>表头在 <select id="headerSel" aria-label="表头所在行">' + opts + '</select></label>' +
      '<span>数据 ' + n.toLocaleString() + ' 行（第 ' + (s.dataStart + 1) + '–' + (s.dataEnd + 1) + ' 行）</span>' +
      '<span>' + s.columns.filter(c => !c.added).length + ' 列</span>';
    s.warnings.forEach(w => { html += '<span class="warn">' + esc(w) + '</span>'; });
    const tn = s.columns.filter(c => c.type === 'textnum' && !c.added);
    if (tn.length) html += '<span class="warn" title="' + esc(tn.map(c => c.warning).join('\n')) + '">' + esc(tn.map(c => c.name).join('、')) + ' 是“文本形式的数字”</span>';
    meta.innerHTML = html;
  }

  function cellHtml(text, raw, extraCls) {
    let cls = extraCls || '';
    if (typeof raw === 'number') cls += ' num';
    if (raw && typeof raw === 'object' && 'error' in raw) cls += ' err';
    if (typeof raw === 'string' && /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?)$/.test(raw)) cls += ' err';
    return '<td class="' + cls.trim() + '" title="' + esc(text) + '">' + esc(text) + '</td>';
  }

  function renderSheetTable() {
    const s = activeSheet();
    const wrap = $('tableWrap');
    if (s.lastRow < 0) { wrap.innerHTML = '<div class="more-rows">没有数据</div>'; return; }
    const pend = state.wb.pending && state.wb.pending.kind === 'column' && state.wb.pending.sheet === s.name ? state.wb.pending : null;
    const cols = s.columns;
    let head = '<thead><tr><th class="rn"></th>';
    cols.forEach(c => {
      head += '<th class="' + (c.added ? 'new' : '') + '"><div class="col-head"><span class="col-letter">' + c.letter + '</span><span class="col-name">' + esc(c.name) + '</span>' +
        '<span class="type t-' + c.type + '"' + (c.warning ? ' title="' + esc(c.warning) + '"' : '') + '>' + TYPE_LABEL[c.type] + '</span></div></th>';
    });
    if (pend) pend.outputs.forEach(o => {
      head += '<th class="new"><div class="col-head"><span class="col-letter">' + U.colLetter(o.col) + '</span><span class="col-name">' + esc(o.title) + '</span><span class="type">预览</span></div></th>';
    });
    head += '</tr></thead>';
    let body = '<tbody>';
    const end = Math.min(s.dataEnd, s.dataStart + PREVIEW_ROWS - 1);
    for (let r = s.dataStart; r <= end; r++) {
      body += '<tr><td class="rn">' + (r + 1) + '</td>';
      cols.forEach(c => {
        const raw = s.values[r] ? s.values[r][c.index] : null;
        body += cellHtml(W.displayValue(s, r, c.index), raw, c.added ? 'new' : '');
      });
      if (pend) pend.outputs.forEach(o => {
        const v = o.values[r - s.dataStart];
        body += cellHtml(W.formatValue(v, o.format), v, 'new');
      });
      body += '</tr>';
    }
    body += '</tbody>';
    const total = s.dataEnd - s.dataStart + 1;
    let foot = '';
    if (total > PREVIEW_ROWS) foot += '<div class="more-rows">仅预览前 ' + PREVIEW_ROWS + ' 行，计算和下载包含全部 ' + total.toLocaleString() + ' 行。</div>';
    if (s.headerRow > 0) foot = '<div class="more-rows">表头上方还有 ' + s.headerRow + ' 行（如标题），不参与计算。</div>' + foot;
    wrap.innerHTML = '<table class="grid">' + head + body + '</table>' + foot;
  }

  function renderOutput(o) {
    $('sheetMeta').innerHTML = '<span>本工具新建的工作表，下载后会出现在文件里</span>';
    const wrap = $('tableWrap');
    let html = '<table class="grid"><thead><tr><th class="rn"></th>';
    if (o.kind === 'scalar') {
      html += '<th>计算项</th><th>结果</th><th>需求描述</th></tr></thead><tbody>';
      o.items.forEach(it => {
        html += '<tr><td class="rn">' + (it.row + 1) + '</td>' + cellHtml(it.title) + cellHtml(W.formatValue(it.value, it.format), it.value) + cellHtml(it.request) + '</tr>';
      });
    } else {
      o.header.forEach(h => { html += '<th>' + esc(h) + '</th>'; });
      html += '</tr></thead><tbody>';
      o.keyValues.slice(0, 500).forEach((keys, ki) => {
        html += '<tr><td class="rn">' + (ki + 2) + '</td>';
        keys.forEach((k, ci) => html += cellHtml(keyText(k, o.keyCols[ci]), k));
        o.outputs.forEach(out => { const v = out.values[ki]; html += cellHtml(W.formatValue(v, out.format), v); });
        html += '</tr>';
      });
    }
    html += '</tbody></table>';
    wrap.innerHTML = html;
  }

  function keyText(k, kc) {
    if (typeof k !== 'number') return String(k);
    if (kc.unit === 'month') return U.formatSerialMonth(k);
    if (kc.unit === 'year') return U.serialToParts(k).y + '年';
    if (kc.srcFormat === 'date') return U.formatSerialDate(k);
    return U.formatNumber(k);
  }

  function renderChips() {
    const ex = T.examples(activeSheet(), state.wb);
    $('exampleChips').innerHTML = ex.map(x => '<button class="chip" type="button">' + esc(x) + '</button>').join('');
  }

  /* ======================= 生成 ======================= */

  async function generate(forceAI) {
    if (state.busy || !state.wb) return;
    const req = $('reqInput').value.trim();
    if (!req) { $('reqInput').focus(); toast('先写下你想算什么'); return; }
    if (state.view.type !== 'sheet') { state.view = { type: 'sheet' }; renderTabs(); }
    const wb = state.wb;
    const sheet = activeSheet();
    W.discardPending(wb);
    renderView();
    setBusy(true);
    showLoading('正在理解需求…', false);
    await U.sleep(20);

    let tplError = null;
    if (!forceAI) {
      let plan = null;
      try { plan = T.match(req, wb, sheet).plan; } catch (e) { console.warn(e); }
      if (plan) {
        try {
          W.applyPending(wb, plan);
          setBusy(false);
          showResult(wb.pending);
          return;
        } catch (e) {
          console.warn('模板结果无法使用', e);
          tplError = e;
        }
      }
    }

    if (!L.isConfigured()) {
      setBusy(false);
      showUnmatched(req, tplError, forceAI);
      return;
    }

    state.abort = new AbortController();
    showLoading('AI 正在生成公式…', true);
    try {
      await L.generate(req, wb, sheet, {
        signal: state.abort.signal,
        onStatus: msg => updateLoading(msg),
        validate: plan => { W.applyPending(wb, plan); return null; },
      });
      showResult(wb.pending);
    } catch (e) {
      console.warn(e);
      W.discardPending(wb);
      renderView();
      showError(e);
    } finally {
      state.abort = null;
      setBusy(false);
    }
  }

  function setBusy(b) {
    state.busy = b;
    $('btnGenerate').disabled = b || !state.wb;
    $('btnGenerate').textContent = b ? '生成中…' : '生成公式';
  }

  function showLoading(msg, cancellable) {
    $('resultArea').innerHTML = '<div class="card"><div class="state"><div class="spinner"></div><div><h3 id="loadMsg">' + esc(msg) + '</h3>' +
      (cancellable ? '<p>一般需要 5～30 秒</p><button class="btn btn-ghost btn-sm" type="button" data-act="cancel">取消</button>' : '') + '</div></div></div>';
  }
  function updateLoading(msg) { const el = $('loadMsg'); if (el) el.textContent = msg; }

  function showUnmatched(req, tplError, forceAI) {
    const s = activeSheet();
    const ex = T.examples(s, state.wb).slice(0, 4);
    const reason = tplError ? '<p>内置模板尝试生成，但公式无法使用：' + esc(tplError.message) + '</p>' : '<p>内置模板没能理解这个需求。</p>';
    $('resultArea').innerHTML = '<div class="card"><div class="state"><div style="font-size:22px">🤔</div><div>' +
      '<h3>' + (forceAI ? '需要先设置 AI' : '暂时没看懂这个需求') + '</h3>' + (forceAI ? '' : reason) +
      '<p>可以试试：</p><ul>' +
      '<li>直接写出列名，例如「' + esc((s.columns[0] || {}).name || '销售额') + '」</li>' +
      '<li>把条件和要算的东西分开说，例如“部门是销售部的销售额合计”</li>' +
      (ex.length ? '<li>参考示例：' + ex.map(x => '<button class="btn-text" data-fill="' + esc(x) + '">' + esc(x) + '</button>').join('、') + '</li>' : '') +
      '</ul>' +
      '<p style="margin-top:10px">或者接入 AI（DeepSeek / 通义千问），处理更复杂的需求。</p>' +
      '<button class="btn btn-primary btn-sm" type="button" data-act="settings">设置 API Key</button>' +
      '</div></div></div>';
  }

  function showError(e) {
    const isKind = e.kind === 'unsupported' || e.kind === 'need_clarification';
    $('resultArea').innerHTML = '<div class="card"><div class="state"><div style="font-size:22px">' + (isKind ? '💬' : '⚠️') + '</div><div>' +
      '<h3>' + (e.kind === 'unsupported' ? '这个需求做不了' : e.kind === 'need_clarification' ? '需要说得更具体一点' : '生成失败') + '</h3>' +
      '<p>' + esc(e.message || String(e)) + '</p>' +
      (/API Key|设置/.test(e.message || '') ? '<button class="btn btn-ghost btn-sm" type="button" data-act="settings">打开 AI 设置</button>' : '') +
      '</div></div></div>';
  }

  /** 结果卡片 */
  function showResult(op) {
    if (!op) return;
    renderView();
    const p = op.plan;
    const src = p.source === 'ai' ? '<span class="badge badge-ai">AI 生成 · ' + esc(L.modelName(p.model || '')) + '</span>' : '<span class="badge badge-tpl">内置模板</span>';
    let body = '';
    if (p.interpretation) body += '<p class="interp"><b>我的理解：</b>' + esc(p.interpretation) + '</p>';

    const funcs = new Set();
    op.outputs.forEach((o, i) => {
      (o.functions || []).forEach(f => funcs.add(f));
      let label;
      if (op.kind === 'column') label = '新列「' + o.title + '」· 第 ' + (o.rows[0] + 1) + ' 行的公式' + (o.rows.length > 1 ? '（已向下填充到第 ' + (o.rows[o.rows.length - 1] + 1) + ' 行）' : '');
      else if (op.kind === 'group') label = '「' + o.title + '」· 汇总表第 2 行的公式（每组一行）';
      else label = '「' + o.title + '」';
      body += '<div class="fx-box"><div class="fx-label"><span>' + esc(label) + '</span><button class="copy" type="button" data-copy="' + i + '">复制公式</button></div>' +
        '<code class="fx-code">' + F.highlight(o.sampleFormula) + '</code></div>';
    });
    if (funcs.size) {
      body += '<div class="fns">用到的函数：' + Array.from(funcs).map(f => '<span class="fn" title="' + esc(EFA.Functions.describe(f) || f) + '">' + esc(f) + '</span>').join('') + '</div>';
    }
    if (p.explanation) body += '<p class="explain">' + esc(p.explanation) + '</p>';
    body += previewHtml(op);
    body += noticesHtml(op);

    const where = op.kind === 'column' ? '在「' + op.sheet + '」右侧新增 ' + op.outputs.length + ' 列'
      : op.kind === 'scalar' ? '写入「' + op.scalarSheet.name + '」表' : '新建「' + op.groupSheet.name + '」表';
    const aiBtn = p.source !== 'ai' ? '<button class="btn btn-ghost btn-sm" type="button" data-act="ai" title="结果不对时，让 AI 按原需求重新生成">不对？用 AI 重新生成</button>' : '';
    $('resultArea').innerHTML = '<div class="card">' +
      '<div class="card-head">' + src + '<span class="kind">' + KIND_LABEL[op.kind] + ' · ' + esc(where) + '</span></div>' +
      '<div class="card-body">' + body + '</div>' +
      '<div class="card-actions"><button class="btn btn-primary" type="button" data-act="commit">添加到表格</button>' +
      '<button class="btn btn-ghost" type="button" data-act="discard">放弃</button><span class="spacer"></span>' + aiBtn + '</div></div>';
  }

  function previewHtml(op) {
    if (op.kind === 'scalar') {
      return '<div class="pv-title">计算结果</div><div class="scalar-list">' + op.outputs.map(o =>
        '<div class="scalar"><div class="s-title">' + esc(o.title) + '</div><div class="s-val">' + esc(W.formatValue(o.values[0], o.format)) + '</div></div>').join('') + '</div>';
    }
    if (op.kind === 'group') {
      const keyCols = op.keyCols;
      let h = '<table class="grid"><thead><tr>' + keyCols.map(k => '<th>' + esc(k.name) + '</th>').join('') + op.outputs.map(o => '<th class="new">' + esc(o.title) + '</th>').join('') + '</tr></thead><tbody>';
      op.keyValues.slice(0, 12).forEach((keys, ki) => {
        h += '<tr>' + keys.map((k, ci) => cellHtml(keyText(k, keyCols[ci]), k)).join('') + op.outputs.map(o => cellHtml(W.formatValue(o.values[ki], o.format), o.values[ki], 'new')).join('') + '</tr>';
      });
      h += '</tbody></table>';
      const more = op.keyValues.length > 12 ? '<div class="more-rows">共 ' + op.keyValues.length + ' 组，仅显示前 12 组</div>' : '';
      return '<div class="pv-title">结果预览（共 ' + op.keyValues.length + ' 组）</div><div class="pv-wrap">' + h + '</div>' + more;
    }
    // 新增列：显示公式用到的列 + 新列
    const sheet = state.wb.sheets.find(s => s.name === op.sheet);
    const refCols = [];
    op.outputs.forEach(o => F.listRefs(o.dsl).forEach(r => {
      if (r.sheet && r.sheet !== sheet.name) return;
      const c = sheet.columns.find(x => x.name === r.column);
      if (c && !refCols.includes(c)) refCols.push(c);
    }));
    const showCols = refCols.slice(0, 3);
    if (!showCols.length) { const nc = sheet.columns.find(c => c.isName) || sheet.columns[0]; if (nc) showCols.push(nc); }
    let h = '<table class="grid"><thead><tr><th class="rn"></th>' + showCols.map(c => '<th>' + esc(c.name) + '</th>').join('') +
      op.outputs.map(o => '<th class="new">' + esc(o.title) + '</th>').join('') + '</tr></thead><tbody>';
    const n = Math.min(8, op.outputs[0].rows.length);
    for (let i = 0; i < n; i++) {
      const r = op.outputs[0].rows[i];
      h += '<tr><td class="rn">' + (r + 1) + '</td>' + showCols.map(c => cellHtml(W.displayValue(sheet, r, c.index), sheet.values[r][c.index])).join('') +
        op.outputs.map(o => cellHtml(W.formatValue(o.values[i], o.format), o.values[i], 'new')).join('') + '</tr>';
    }
    h += '</tbody></table>';
    return '<div class="pv-title">结果预览（前 ' + n + ' 行，左侧数据表中也能看到）</div><div class="pv-wrap">' + h + '</div>';
  }

  function noticesHtml(op) {
    let h = '';
    op.outputs.forEach(o => {
      if (o.errors && o.errors.total) {
        h += '<div class="notice warn">「' + esc(o.title) + '」有 ' + o.errors.text + '。这通常是数据本身的问题（如空值、找不到匹配），公式本身没有错。</div>';
      }
      const nf = o.values.filter(v => v === '未找到').length;
      if (nf) h += '<div class="notice info">「' + esc(o.title) + '」有 ' + nf + ' 行没有找到匹配，显示为“未找到”。</div>';
    });
    const sheet = state.wb.sheets.find(s => s.name === op.sheet);
    const used = new Set();
    op.outputs.forEach(o => F.listRefs(o.dsl).forEach(r => used.add(r.column)));
    sheet.columns.filter(c => c.type === 'textnum' && used.has(c.name)).forEach(c => {
      h += '<div class="notice warn">' + esc(c.warning) + '。如果结果为 0，可以先用“把' + esc(c.name) + '转成数字”生成一列数值再计算。</div>';
    });
    if (op.outputs.some(o => /TODAY\(\)/.test(o.sampleFormula))) h += '<div class="notice info">公式里用了 TODAY()，在 Excel 中每天打开都会按当天日期重新计算；“结果值”模式下保存的是今天算出的值。</div>';
    return h;
  }

  /* ======================= 确认 / 放弃 / 删除 ======================= */

  function commit() {
    const wb = state.wb;
    const op = W.commitPending(wb);
    if (!op) return;
    const titles = op.outputs.map(o => '「' + o.title + '」').join('、');
    toast(op.kind === 'column' ? '已新增列 ' + titles : op.kind === 'scalar' ? '已写入“计算结果”表：' + titles : '已新建汇总表「' + op.groupSheet.name + '」');
    $('resultArea').innerHTML = '<div class="notice info">✓ 已添加。可以继续描述下一个需求（新增的列也能被引用），或在下方下载结果。</div>';
    $('reqInput').value = '';
    if (op.kind !== 'column') state.view = { type: 'output', name: op.kind === 'scalar' ? op.scalarSheet.name : op.groupSheet.name };
    renderAll();
    $('reqInput').focus();
  }

  function discard() {
    W.discardPending(state.wb);
    $('resultArea').innerHTML = '';
    renderView();
  }

  function removeOp(id) {
    const dropped = W.removeOp(state.wb, id);
    if (state.view.type === 'output' && !state.wb.outputs.some(o => o.name === state.view.name)) state.view = { type: 'sheet' };
    renderAll();
    if (dropped.length) toast('另有 ' + dropped.length + ' 个计算依赖被删除的列，也一并移除了', 4000);
    else toast('已移除');
  }

  function renderApplied() {
    const wb = state.wb;
    const ops = wb.ops;
    $('appliedStep').hidden = !ops.length;
    $('exportStep').hidden = false;
    $('appliedCount').textContent = ops.length ? '(' + ops.length + ')' : '';
    $('appliedList').innerHTML = ops.map(op => {
      const where = op.kind === 'column' ? op.sheet : op.kind === 'scalar' ? op.scalarSheet.name : op.groupSheet.name;
      return '<li><span class="k">' + KIND_LABEL[op.kind] + '</span><div class="t"><div>' + esc(op.outputs.map(o => o.title).join('、')) + '</div>' +
        '<div class="sub">' + esc(where) + (op.plan.request ? ' · ' + esc(op.plan.request) : '') + '</div></div>' +
        '<button class="rm" type="button" title="移除" aria-label="移除" data-rm="' + op.id + '">×</button></li>';
    }).join('');
    $('btnDownload').disabled = !ops.length;
    renderModeDesc();
  }

  function renderModeDesc() {
    $('modeDesc').textContent = state.mode === 'value'
      ? '保存计算好的数值：打开就能用，不会因为数据变化而改变。'
      : '保存公式（同时附带当前结果）：可在 Excel / WPS 中继续编辑，原数据修改后会自动重算。';
    $('exportNote').textContent = !state.wb || !state.wb.ops.length ? '添加至少一个计算后即可下载。' : '原表的内容和格式保持不变，只新增列和新表。';
  }

  async function download() {
    const wb = state.wb;
    if (!wb || !wb.ops.length) return;
    const btn = $('btnDownload');
    btn.disabled = true;
    btn.textContent = '正在生成文件…';
    try {
      const { blob, fileName, notes } = await EFA.Exporter.exportWorkbook(wb, state.mode);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      toast('已下载 ' + fileName + (notes.length ? '（' + notes.join('；') + '）' : ''), notes.length ? 5000 : 2600);
      if (wb.pending) toast('提示：还有一个结果没点“添加到表格”，它不在下载的文件里', 4000);
    } catch (e) {
      console.error(e);
      toast('生成文件失败：' + (e.message || e), 5000);
    } finally {
      btn.disabled = false;
      btn.textContent = '下载 .xlsx';
    }
  }

  /* ======================= AI 设置 ======================= */

  function renderAiState() {
    const on = L.isConfigured();
    $('aiDot').classList.toggle('on', on);
    $('aiLabel').textContent = on ? 'AI 已启用 · ' + L.providerName() : 'AI 设置';
  }

  const CUSTOM = '__custom__';
  function fillProvider(pid) {
    const s = L.getSettings(pid);
    const p = L.PROVIDERS[s.provider];
    $('setProvider').innerHTML = L.PROVIDER_LIST.map(x => '<option value="' + x.id + '"' + (x.id === s.provider ? ' selected' : '') + '>' + esc(x.name) + '</option>').join('');
    $('keyLabel').textContent = p.keyLabel;
    $('keyWhere').textContent = '在' + p.keyWhere + ' 创建。只保存在当前浏览器标签页，关闭页面即清除。';
    $('setKey').value = s.apiKey;
    const preset = p.models.some(m => m.id === s.model);
    $('setModel').innerHTML = p.models.map(m => '<option value="' + m.id + '"' + (m.id === s.model ? ' selected' : '') + '>' + esc(m.name) + '</option>').join('') +
      '<option value="' + CUSTOM + '"' + (preset ? '' : ' selected') + '>自定义模型名…</option>';
    $('setCustomModel').value = preset ? '' : s.model;
    $('customModelField').hidden = preset;
    $('wsField').hidden = !p.workspaceBase;
    $('setWorkspace').value = s.workspace;
    $('setBase').value = s.baseURL;
    $('setBase').placeholder = p.base;
    $('setSamples').checked = s.sendSamples;
  }

  function openSettings() {
    fillProvider();
    $('settingsDlg').showModal();
    setTimeout(() => $('setKey').focus(), 50);
  }

  function saveSettings() {
    const key = $('setKey').value.trim();
    const pid = $('setProvider').value;
    let model = $('setModel').value;
    if (model === CUSTOM) model = $('setCustomModel').value.trim() || L.PROVIDERS[pid].models[0].id;
    if (key && !/^sk-/.test(key)) toast('API Key 一般以 sk- 开头，请检查是否复制完整');
    L.saveSettings({ provider: pid, apiKey: key, model, baseURL: $('setBase').value, workspace: $('setWorkspace').value, sendSamples: $('setSamples').checked });
    renderAiState();
    toast(key ? 'AI 已启用：' + L.PROVIDERS[pid].name + ' · ' + model + '（Key 仅保存在本标签页）' : '已保存设置');
  }

  /* ======================= 事件 ======================= */

  function bindDropButtons() {
    const pick = $('btnPick'), sample = $('btnSample');
    if (pick) pick.onclick = () => $('fileInput').click();
    if (sample) sample.onclick = loadSample;
  }

  function init() {
    bindDropButtons();
    $('fileInput').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; loadFile(f); });
    $('btnReplace').addEventListener('click', () => $('fileInput').click());

    const dz = $('dropzone');
    dz.addEventListener('keydown', e => { if (e.key === 'Enter' && e.target === dz) $('fileInput').click(); });
    let dragDepth = 0;
    window.addEventListener('dragenter', e => { if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) { dragDepth++; dz.classList.add('drag'); } });
    window.addEventListener('dragleave', () => { dragDepth = Math.max(0, dragDepth - 1); if (!dragDepth) dz.classList.remove('drag'); });
    window.addEventListener('dragover', e => e.preventDefault());
    window.addEventListener('drop', e => {
      e.preventDefault();
      dragDepth = 0;
      dz.classList.remove('drag');
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) loadFile(f);
    });

    $('sheetTabs').addEventListener('click', e => {
      const b = e.target.closest('.sheet-tab');
      if (!b) return;
      if (b.dataset.sheet !== undefined) {
        const idx = +b.dataset.sheet;
        if (state.view.type === 'sheet' && idx === state.wb.activeIndex) return;
        W.discardPending(state.wb);
        if ($('resultArea').querySelector('[data-act="commit"]')) $('resultArea').innerHTML = '';
        state.wb.activeIndex = idx;
        state.view = { type: 'sheet' };
        renderAll();
      } else {
        state.view = { type: 'output', name: b.dataset.output };
        renderTabs();
        renderView();
      }
    });

    $('sheetMeta').addEventListener('change', e => {
      if (e.target.id !== 'headerSel') return;
      const wb = state.wb;
      if (wb.ops.length && !confirm('修改表头行后，会按新的表头重新计算已添加的内容，引用不到的列会被移除。继续吗？')) { renderSheetMeta(); return; }
      const dropped = W.changeHeaderRow(wb, wb.activeIndex, +e.target.value);
      $('resultArea').innerHTML = '';
      renderAll();
      toast(dropped.length ? '表头已更新，' + dropped.length + ' 个计算因找不到列被移除' : '表头已更新');
    });

    $('btnGenerate').addEventListener('click', () => generate(false));
    $('reqInput').addEventListener('keydown', e => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); generate(false); }
    });
    $('exampleChips').addEventListener('click', e => {
      const c = e.target.closest('.chip');
      if (!c) return;
      $('reqInput').value = c.textContent;
      generate(false);
    });

    $('resultArea').addEventListener('click', e => {
      const t = e.target.closest('[data-act],[data-copy],[data-fill]');
      if (!t) return;
      if (t.dataset.fill) { $('reqInput').value = t.dataset.fill; generate(false); return; }
      if (t.dataset.copy !== undefined) {
        const o = state.wb.pending && state.wb.pending.outputs[+t.dataset.copy];
        if (!o) return;
        copyText(o.sampleFormula).then(() => toast('公式已复制（这是第一行的公式，粘贴到 Excel 后可向下填充）'), () => toast('复制失败，请手动选择公式复制'));
        return;
      }
      const act = t.dataset.act;
      if (act === 'commit') commit();
      else if (act === 'discard') discard();
      else if (act === 'settings') openSettings();
      else if (act === 'cancel') { if (state.abort) state.abort.abort(); }
      else if (act === 'ai') {
        if (!L.isConfigured()) { openSettings(); return; }
        generate(true);
      }
    });

    $('appliedList').addEventListener('click', e => {
      const b = e.target.closest('[data-rm]');
      if (b) removeOp(b.dataset.rm);
    });

    document.querySelectorAll('.seg-btn').forEach(b => b.addEventListener('click', () => {
      state.mode = b.dataset.mode;
      document.querySelectorAll('.seg-btn').forEach(x => { x.classList.toggle('active', x === b); x.setAttribute('aria-checked', x === b); });
      renderModeDesc();
    }));
    $('btnDownload').addEventListener('click', download);

    $('btnSettings').addEventListener('click', openSettings);
    $('settingsForm').addEventListener('submit', () => saveSettings());
    $('btnCancelSet').addEventListener('click', () => $('settingsDlg').close());
    $('btnClearKey').addEventListener('click', () => { $('setKey').value = ''; L.clearKey($('setProvider').value); renderAiState(); toast('已清除 API Key'); });
    $('setProvider').addEventListener('change', e => fillProvider(e.target.value));
    $('setModel').addEventListener('change', e => { $('customModelField').hidden = e.target.value !== CUSTOM; if (e.target.value === CUSTOM) $('setCustomModel').focus(); });

    renderAiState();
    renderModeDesc();
  }

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) return navigator.clipboard.writeText(text);
    return new Promise((resolve, reject) => {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      try { document.execCommand('copy') ? resolve() : reject(); } catch (e) { reject(e); } finally { ta.remove(); }
    });
  }

  // 依赖检查（离线打开时库文件缺失）
  if (!window.XLSX || !window.ExcelJS || !window.HyperFormula) {
    document.body.innerHTML = '<p style="padding:24px;font:15px sans-serif">组件加载失败：请确认 vendor 文件夹与 index.html 放在一起。</p>';
    return;
  }
  init();
  EFA.App = { state, generate, loadSample };
})(window.EFA = window.EFA || {});
