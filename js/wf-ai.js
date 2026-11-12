/* 工作流配置器：用一句话生成步骤（AI 只生成配置，不生成代码）。
 * 提示词里的操作说明由操作目录（wf-catalog.js）自动生成，与执行器保持一致；
 * AI 返回的步骤用与手动编辑相同的规则校验，有错误时把错误反馈给 AI 修正一次，最后由用户确认才加入。 */
(function (EFA) {
  'use strict';
  const M = EFA.WFModel, F = EFA.WFFiles, L = EFA.LLM, CAT = EFA.WF_CATALOG;
  const A = {};

  A.LIMITS = { maxInputTokens: 12000, maxCalls: 3, samplesPerColumn: 2 };

  /* ------------------------------------------------------------------ 操作说明（给 AI 看） */
  const TYPE_DOC = {
    text: '文字', number: '数字', bool: 'true/false', triBool: 'true/false',
    column: '列名', columns: '列名数组，如 ["金额","数量"]；也可写列字母区间 "U:AB"',
    condition: '条件，见“条件写法”', numfmt: 'Excel 数字格式，如 "#,##0.00"、"0%"、"yyyy-mm-dd"',
    color: '颜色：6 位十六进制如 "FCE4D6"，或 红/黄/绿/蓝/橙/灰',
    rowRange: '行号区间字符串，如 "5:20"', matchId: '给这次匹配起的名字，如 "m1"', matchRef: '前面某个 match 步骤的 id',
    fileRef: '参照表的文件 id', sheet: '工作表名',
    sortBy: '[{"column":"金额","order":"desc"},{"column":"日期"}]（order：asc 升序 / desc 降序）',
    cells: '{"AC1":"备注","B2":"=SUM(B3:B9)"}',
    copyColumns: '["区域"] 或 [{"from":"参照表列","to":"主表列"}]',
    comparePairs: '[{"main":"主表列","ref":"参照表列"}]，不写=比较所有同名列',
    optText: '文字；写 null 表示不生成',
    labels: '{"same":"一致","diff":"有差异","main_only":"仅主表有"}',
    rangeSource: '{"file":"文件id","sheet":"工作表","range":"AF1:AI30"}',
    cellTarget: '{"cell":"AF1","sheet":"主表工作表（可省略）"}',
    sheetSource: '{"file":"文件id","sheet":"工作表（可省略）"}',
    contentParts: '["values","formulas","styles","widths","merges"] 的子集，不写=全部',
    notFound: '{"column":"写标记的主表列","value":"未找到"}',
    metrics: '[{"column":"金额","func":"sum"},{"func":"count"}]（func：sum 合计 / count 记录数 / avg 平均 / max 最大 / min 最小 / count_distinct 不重复个数，可加 "title"）',
    checks: '[{"type":"not_empty","columns":["单号"],"on_fail":"stop"}]（type：required_columns / not_empty / unique / row_count{min,max} / headers_match{file,columns:"A:Z"} / condition{where,expect:none|all|any}；可加 "message"）',
  };
  const COMPOSITE_DOC = {
    content: '写入内容，四选一：value 固定值 / formula 公式模板（[@列名]=同一行的单元格，[列名]=整列数据区域，如 "=[@单价]*[@数量]"）/ from_col 同一行另一列的值 / formula_from {"file","sheet","cell"} 复制某单元格公式',
    position: '插入位置，四选一：before 列名 / after 列名 / at 列字母 / end: true',
    rowMode: '删除哪些行，三选一：where 条件 / rows "5:20" / blank: true（整行为空）',
    rowsAt: '插入位置：before_row 行号 或 at_end: true',
  };

  A.catalogText = function () {
    return CAT.ops.map(d => {
      const params = d.fields.map(f => {
        if (COMPOSITE_DOC[f.type]) return '  - ' + (f.keys || [f.key]).join('/') + (f.required ? '（必填）' : '') + '：' + COMPOSITE_DOC[f.type];
        let t = TYPE_DOC[f.type] || f.type;
        if (f.type === 'enum') t = f.options.map(o => JSON.stringify(o.value) + ' ' + o.label).join(' / ') + '（默认 ' + JSON.stringify(f.default) + '）';
        if (f.type === 'bool' && f.default !== undefined) t += '，默认 ' + f.default;
        if (f.source === 'ref') t += '（参照表的列）';
        if (f.source === 'from') t += '（来源文件的列）';
        if (f.allowRef) t += '（可用 ref_col 引用匹配到的参照表列）';
        return '  - ' + f.key + (f.required ? '（必填）' : '') + '：' + f.label + '｜' + t;
      });
      return '### ' + d.op + '：' + d.label + (d.writes === false ? '（只读）' : '') + '\n' + d.desc + '\n' + params.join('\n');
    }).join('\n\n');
  };

  A.systemPrompt = function () {
    const ops = M.COND_OPS.map(o => o.v).join(' ');
    return [
      '你是“Excel 工作流配置器”的助手。用户会描述要对 Excel 表格反复做的处理，你从下面的内置操作中选择并组合成步骤，输出 JSON 配置。',
      '本地执行器会按这些步骤处理用户每次放进来的新文件。你只能输出配置，绝对不要输出任何程序代码。',
      '',
      '## 规则',
      '1. 只能使用下面列出的操作和参数名，不要编造操作或参数。每个步骤是一个对象，必须有 "op"，可加简短的中文 "title"。',
      '2. 列名必须是数据里真实存在的表头，或前面步骤新建、改名后的名字；不要猜测不存在的列。',
      '3. 文件用 files 里的 id 指代；只有主表（role=main）可以被修改，参照表只读。“多个文件”的参照表只能用于 append_rows。',
      '4. 步骤按执行顺序排列：建立匹配（match）必须在使用它的 copy_matched / compare 之前；删除或改名列之后，后面的步骤要用新的列情况。',
      '5. 公式只能用 Excel 2016 已有的函数（不能用 IFS、SWITCH、XLOOKUP、TEXTJOIN、CONCAT、MAXIFS、MINIFS、FILTER、UNIQUE 等）。',
      '6. 执行器只能实时算出“同一行单元格之间的四则运算”这类简单公式（可用 ROUND/ABS/MAX/MIN/INT）。其他本次新写入的公式在执行时没有结果，条件、排序、去重、匹配、分组里不要用这些列，改用参与计算的原始列。',
      '7. 需求里没提到的处理不要加；用户描述不够明确、会导致结果完全不同时，status 用 "need_clarification"，在 message 里只问一个最关键的问题。做不到的需求 status 用 "unsupported" 并说明原因。',
      '',
      '## 条件写法',
      '{"col":"列名","op":"判断方式","value":值} ；两列比较 {"col":"实发","op":"=","col2":"应发"} ；组合 {"all":[条件…]} / {"any":[条件…]} / {"not":条件}',
      '判断方式 op：' + ops + '（empty/not_empty/is_number/is_text/is_error 不需要 value；in/not_in 的 value 是数组）',
      '',
      '## 可用操作',
      A.catalogText(),
      '',
      '## 输出格式（只输出这个 JSON 对象）',
      '{"status":"ok","steps":[{"op":"…","title":"…",…}],"explanation":"用中文分条说明每一步做什么","assumptions":["你做的假设，没有就空数组"],"message":""}',
    ].join('\n');
  };

  /* ------------------------------------------------------------------ 数据描述（不含文件内容） */
  A.context = function (cfg, opt) {
    opt = opt || {};
    const maxCols = opt.maxCols || 200;
    const withSamples = !!opt.withSamples;
    const out = ['## 文件'];
    cfg.files.forEach(f => {
      const smp = M.state.samples[f.id];
      const role = f.role === 'main' ? '主表，处理并输出结果' : '参照表，只读' + (f.multiple ? '，多个文件（一类文件可能有很多个）' : '');
      out.push('- id=' + f.id + '（' + role + '）名称：' + (f.label || f.id));
      if (!smp) { out.push('  （没有样例文件，不知道表头）'); return; }
      const sh = M.sampleSheet(f.id, f.sheet);
      out.push('  工作表：' + smp.sheets.map(s => s.name).join('、') + (sh ? '；使用「' + sh.name + '」，表头在第 ' + (f.header_row || 1) + ' 行，约 ' + F.dataRows(sh, f.header_row || 1) + ' 行数据' : ''));
      if (!sh) return;
      const cols = F.columns(sh, f.header_row || 1);
      const shown = cols.slice(0, maxCols);
      out.push('  列：' + shown.map(c => c.letter + ' ' + (c.name || '（无表头）') + '（' + F.TYPE_LABEL[c.type] +
        (withSamples && !c.sensitive && c.samples.length ? '，例：' + c.samples.slice(0, A.LIMITS.samplesPerColumn).join(' / ') : '') + '）').join('；'));
      if (cols.length > shown.length) out.push('  （另有 ' + (cols.length - shown.length) + ' 列未列出）');
      const fcols = F.formulaColumns(sh, f.header_row || 1);
      if (fcols.length) out.push('  其中是公式的列：' + fcols.join('、'));
    });
    if (opt.existing && cfg.steps.length) {
      out.push('', '## 已有步骤（新步骤会接在它们后面执行）');
      cfg.steps.forEach((s, i) => out.push((i + 1) + '. ' + JSON.stringify(s)));
    }
    return out.join('\n');
  };

  function contextWithinBudget(cfg, opt) {
    const st = L.getSettings();
    const sys = A.systemPrompt();
    const budget = A.LIMITS.maxInputTokens - L.estimateTokens(sys) - L.estimateTokens(opt.request) - 300;
    const levels = [
      { withSamples: st.sendSamples, maxCols: 200, existing: opt.existing },
      { withSamples: false, maxCols: 200, existing: opt.existing },
      { withSamples: false, maxCols: 60, existing: opt.existing },
      { withSamples: false, maxCols: 30, existing: false },
    ];
    for (let i = 0; i < levels.length; i++) {
      const ctx = A.context(cfg, levels[i]);
      if (L.estimateTokens(ctx) <= budget || i === levels.length - 1) {
        if (L.estimateTokens(ctx) > budget) {
          const e = new Error('文件的列太多，发给 AI 的内容超出单次上限。可以先删掉样例里用不到的列再试，或手动添加步骤。');
          e.kind = 'limit';
          throw e;
        }
        return { ctx, sys, level: i };
      }
    }
  }

  /* ------------------------------------------------------------------ 校验 AI 给出的步骤 */
  function normalize(out) {
    if (!out || typeof out !== 'object') throw new Error('AI 返回格式不对');
    const status = ['ok', 'need_clarification', 'unsupported'].includes(out.status) ? out.status : 'ok';
    const steps = Array.isArray(out.steps) ? out.steps.filter(s => s && typeof s === 'object' && typeof s.op === 'string') : [];
    if (status === 'ok' && !steps.length) throw new Error('AI 没有给出步骤');
    return {
      status, steps: JSON.parse(JSON.stringify(steps)),
      message: typeof out.message === 'string' ? out.message : '',
      explanation: typeof out.explanation === 'string' ? out.explanation : (Array.isArray(out.explanation) ? out.explanation.join('\n') : ''),
      assumptions: Array.isArray(out.assumptions) ? out.assumptions.filter(x => typeof x === 'string' && x.trim()) : [],
    };
  }

  /** 把新步骤放进临时配置里，用与界面相同的规则校验，返回每个新步骤的问题 */
  A.check = function (cfg, newSteps, mode) {
    const base = mode === 'replace' ? [] : cfg.steps;
    const temp = Object.assign({}, cfg, { steps: base.concat(newSteps) });
    const saved = M.state.cfg;
    M.state.cfg = temp;
    let res;
    try { res = M.validate(temp); } finally { M.state.cfg = saved; }
    const off = base.length;
    return newSteps.map((s, k) => res.steps[off + k] || []);
  };

  /**
   * 生成步骤。opts: { mode: 'append'|'replace', signal, onStatus }
   * 返回 { steps, issues, explanation, assumptions, usage, model }
   */
  A.generate = async function (request, opts) {
    opts = opts || {};
    const cfg = M.state.cfg;
    const st = L.getSettings();
    if (!st.apiKey) { const e = new Error('还没有设置 API Key'); e.kind = 'nokey'; throw e; }
    const p = L.PROVIDERS[st.provider];
    L.checkDaily();
    const mode = opts.mode === 'replace' ? 'replace' : 'append';
    const { ctx, sys, level } = contextWithinBudget(cfg, { request, existing: mode === 'append' });
    const messages = [
      { role: 'system', content: sys },
      { role: 'user', content: ctx + '\n\n## 需求\n' + request + '\n\n请按要求只输出 JSON。' },
    ];
    const est = L.estimateTokens(sys) + L.estimateTokens(messages[1].content);
    const usage = { calls: 0, inTok: 0, outTok: 0, estimate: est, trimmed: level > 0 };
    let jsonMode = true, blank = 0, last = null, lastErr = null;
    while (usage.calls < A.LIMITS.maxCalls) {
      opts.onStatus && opts.onStatus((last ? p.name + ' 正在修正步骤…' : p.name + ' 正在生成步骤…') + '（本次发送约 ' + est.toLocaleString() + ' tokens）');
      let res;
      try {
        usage.calls++;
        res = await L.chat(st, messages, opts, jsonMode);
        if (res.usage) { usage.inTok += res.usage.prompt_tokens || 0; usage.outTok += res.usage.completion_tokens || 0; }
      } catch (e) {
        if (jsonMode && (e.status === 400 || e.status === 422) && /response_format|json/i.test(e.raw || '')) { jsonMode = false; continue; }
        throw e;
      }
      if (res.finish === 'length') throw new Error('AI 输出被截断：要做的处理太多了，可以分几次生成');
      if (!res.content.trim()) {
        if (blank++ < 1) continue;
        throw new Error('AI 返回了空内容，请稍后再试');
      }
      let out;
      try { out = normalize(L.extractJson(res.content)); } catch (e) {
        lastErr = e;
        messages.push({ role: 'assistant', content: res.content });
        messages.push({ role: 'user', content: '你的输出无法使用：' + e.message + '。请只输出符合格式要求的 JSON 对象。' });
        continue;
      }
      if (out.status !== 'ok') {
        const e = new Error(out.message || (out.status === 'unsupported' ? '这个需求用内置操作做不到' : '需求不够明确，请补充说明'));
        e.kind = out.status;
        throw e;
      }
      const issues = A.check(cfg, out.steps, mode);
      last = Object.assign(out, { issues, usage, model: p.name + ' · ' + res.model });
      const errs = [];
      issues.forEach((arr, k) => arr.filter(m => m.level === 'error').forEach(m => errs.push('第 ' + (k + 1) + ' 步（' + out.steps[k].op + '）：' + m.msg)));
      if (!errs.length || usage.calls >= A.LIMITS.maxCalls) return last;
      messages.push({ role: 'assistant', content: res.content });
      messages.push({ role: 'user', content: '这些步骤有问题：\n' + errs.join('\n') + '\n请修正后重新输出完整的 JSON（所有步骤），仍需满足全部规则。' });
    }
    if (last) return last;
    throw lastErr || new Error('AI 没有给出可用的步骤');
  };

  EFA.WFAI = A;
})(window.EFA = window.EFA || {});
