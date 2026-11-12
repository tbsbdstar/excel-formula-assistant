/* AI 兜底：模板识别不了的需求交给国内大模型（DeepSeek / 通义千问），用户自备 API Key
 * - 两家都提供 OpenAI 兼容接口，且允许浏览器直接调用，所以不需要后端
 * - API Key 只存 sessionStorage，关闭标签页即清除
 * - 发送内容：列名、类型、行数、（可选）少量示例值、用户需求；不上传文件本身；
 *   身份证号、手机号、姓名列永远不发送示例值
 */
(function (EFA) {
  'use strict';
  const U = EFA.U;

  const PROVIDERS = {
    deepseek: {
      id: 'deepseek',
      name: 'DeepSeek',
      keyLabel: 'DeepSeek API Key',
      keyWhere: 'DeepSeek 开放平台（platform.deepseek.com）→ API keys',
      base: 'https://api.deepseek.com',
      models: [
        { id: 'deepseek-flash', name: 'deepseek-flash（默认，速度快、价格低）', thinking: false },
        { id: 'deepseek-v4-pro', name: 'deepseek-v4-pro（更强，较慢）', thinking: true },
      ],
      // DeepSeek 默认开启思考；快速模型关闭思考，强模型保持思考
      extra: m => (m ? { thinking: { type: m.thinking ? 'enabled' : 'disabled' } } : {}),
    },
    qwen: {
      id: 'qwen',
      name: '通义千问',
      keyLabel: '阿里云百炼 API Key',
      keyWhere: '阿里云百炼控制台（bailian.console.aliyun.com）→ API Key',
      base: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
      workspaceBase: ws => 'https://' + ws + '.cn-beijing.maas.aliyuncs.com/compatible-mode/v1',
      models: [
        { id: 'qwen-plus', name: 'qwen-plus（默认，均衡）', thinking: false },
        { id: 'qwen-max', name: 'qwen-max（更强）', thinking: false },
        { id: 'qwen3.8-max', name: 'qwen3.8-max（最新旗舰，较慢）', thinking: true },
      ],
      extra: () => ({}),
    },
  };
  const PROVIDER_LIST = [PROVIDERS.deepseek, PROVIDERS.qwen];
  const KEY_PREFIX = 'efa.key.';
  const PREF_STORE = 'efa.aiPrefs';

  /* ---------- 设置 ---------- */
  function safeGet(store, k) { try { return store.getItem(k); } catch (e) { return null; } }
  function safeSet(store, k, v) { try { if (v == null || v === '') store.removeItem(k); else store.setItem(k, v); } catch (e) { /* ignore */ } }

  function loadPrefs() {
    try { return JSON.parse(safeGet(localStorage, PREF_STORE) || '{}') || {}; } catch (e) { return {}; }
  }

  /** 当前设置（provider 可指定，用于设置弹窗里切换服务商时读取对应的值） */
  function getSettings(providerId) {
    const prefs = loadPrefs();
    const pid = PROVIDERS[providerId] ? providerId : (PROVIDERS[prefs.provider] ? prefs.provider : 'deepseek');
    const per = (prefs.per && prefs.per[pid]) || {};
    const p = PROVIDERS[pid];
    const daily = Number(prefs.dailyCalls);
    return {
      provider: pid,
      apiKey: safeGet(sessionStorage, KEY_PREFIX + pid) || '',
      model: per.model || p.models[0].id,
      baseURL: per.baseURL || '',
      workspace: per.workspace || '',
      sendSamples: prefs.sendSamples !== false,
      dailyCalls: Number.isFinite(daily) && daily >= 0 && prefs.dailyCalls !== undefined && prefs.dailyCalls !== '' ? Math.floor(daily) : EFA.LIMITS.ai.defaultDailyCalls,
    };
  }

  function saveSettings(s) {
    const prefs = loadPrefs();
    prefs.provider = s.provider;
    prefs.sendSamples = !!s.sendSamples;
    if (s.dailyCalls !== undefined) prefs.dailyCalls = s.dailyCalls;
    prefs.per = prefs.per || {};
    prefs.per[s.provider] = { model: (s.model || '').trim(), baseURL: (s.baseURL || '').trim(), workspace: (s.workspace || '').trim() };
    safeSet(localStorage, PREF_STORE, JSON.stringify(prefs));
    safeSet(sessionStorage, KEY_PREFIX + s.provider, (s.apiKey || '').trim());
  }

  function clearKey(providerId) { safeSet(sessionStorage, KEY_PREFIX + providerId, ''); }

  /* ---------- 用量统计与每日上限（存在本浏览器） ---------- */
  const USAGE_STORE = 'efa.aiUsage';
  function today() {
    const d = new Date();
    return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
  }
  function getUsage() {
    let u = null;
    try { u = JSON.parse(safeGet(localStorage, USAGE_STORE) || 'null'); } catch (e) { u = null; }
    if (!u || u.date !== today()) u = { date: today(), calls: 0, inTok: 0, outTok: 0 };
    u.limit = getSettings().dailyCalls;
    return u;
  }
  function recordUsage(usage) {
    const u = getUsage();
    u.calls += 1;
    if (usage) { u.inTok += usage.prompt_tokens || 0; u.outTok += usage.completion_tokens || 0; }
    delete u.limit;
    safeSet(localStorage, USAGE_STORE, JSON.stringify(u));
  }
  function checkDaily() {
    const u = getUsage();
    if (u.limit > 0 && u.calls >= u.limit) {
      const e = new Error('今天的 AI 调用次数已达上限（' + u.limit + ' 次）。可以在“AI 设置”里调高上限，或明天再试；内置模板不受影响。');
      e.kind = 'limit';
      throw e;
    }
  }

  /** 粗略估算 token：中文约 0.7 个/字，其他约 1 个/3.5 字符，再留 10% 余量 */
  function estimateTokens(text) {
    const s = String(text || '');
    const cjk = (s.match(/[　-鿿＀-￯]/g) || []).length;
    return Math.ceil((cjk * 0.7 + (s.length - cjk) / 3.5) * 1.1);
  }
  const isConfigured = () => !!getSettings().apiKey;
  const providerName = () => PROVIDERS[getSettings().provider].name;
  const modelName = id => id || '';

  function endpoint(st) {
    const p = PROVIDERS[st.provider];
    let base = st.baseURL || (st.workspace && p.workspaceBase ? p.workspaceBase(st.workspace) : p.base);
    base = base.replace(/\/+$/, '');
    if (/\/chat\/completions$/.test(base)) return base;
    return base + '/chat/completions';
  }

  /* ---------- 提示词 ---------- */
  function allowedList() {
    return Array.from(EFA.Functions.allowedSet()).sort().join(', ');
  }

  const EXAMPLE = {
    status: 'ok', message: '', kind: 'column', sheet: '销售明细', groupBy: [],
    outputs: [{ title: '提成', formula: '=ROUND([@销售额]*0.05,2)', format: 'general' }],
    interpretation: '每一行按销售额的 5% 计算提成，保留两位小数',
    explanation: '[@销售额] 取本行的销售额，乘以 0.05 就是 5%；ROUND(…,2) 四舍五入保留两位小数。',
  };

  function systemPrompt() {
    return [
      '你是 Excel 公式专家，帮不熟悉函数的办公人员把一句中文需求变成正确的 Excel 公式。只输出一个 JSON 对象。',
      '',
      '## 结果类型（kind）',
      '- column：在当前工作表右侧新增一列或几列，每一行一个结果（如按条件打标签、计算年龄、跨表查找、行内运算、排名）。',
      '- scalar：得到一个或几个单独的数（如总和、满足条件的行数、最大值、某人的某项值）。会写到新的“计算结果”表里。',
      '- group：按某列（或日期的月/年）分组，每组一行统计（如各部门销售额合计）。会新建一张汇总表，第一列是分组值。',
      '',
      '## 引用写法（必须使用，不要直接写单元格地址如 D2、D:D）',
      '- [列名]：该列的整列数据区域。',
      '- [工作表名!列名]：其他工作表中某列的数据区域。',
      '- [工作表名!起始列名:结束列名]：其他工作表中从起始列到结束列的多列区域（用于 VLOOKUP 的查找区域）。当前表也可以写 [起始列名:结束列名]。',
      '- [@列名]：同一行的单元格（只能用于 column）。',
      '- [^列名]：从第一行到当前行的累计区域（只能用于 column，用于累计、第几次出现）。',
      '- {KEY}：分组汇总表中当前行的分组值（只能用于 group；第二个分组列用 {KEY2}）。按月分组时 {KEY} 是该月 1 日的日期，条件写成 ">="&{KEY} 和 "<"&EDATE({KEY},1)；按年分组时用 EDATE({KEY},12)。',
      '列名必须与提供的列名完全一致。',
      '',
      '## 兼容性硬性要求（Excel 2016 / WPS）',
      '- 只能使用这些函数：' + allowedList() + '。',
      '- 禁止：IFS、SWITCH、MAXIFS、MINIFS、CONCAT、TEXTJOIN、XLOOKUP、XMATCH、FILTER、UNIQUE、SORT、SORTBY、SEQUENCE、LET、LAMBDA 等新版函数；禁止动态数组、@、# 溢出引用。',
      '- 禁止需要 Ctrl+Shift+Enter 的数组公式。需要逐行判断再汇总时，只能放在 SUMPRODUCT 里，或用 AGGREGATE(14或15或16或17,6,区域/(条件),k)。带条件的最大/最小值用 AGGREGATE(14,6,[值]/([条件列]="x"),1) / AGGREGATE(15,…)。',
      '- 不要用科学计数法（如 1E+10）。日期常量用 DATE(年,月,日)。文本放在英文双引号里。精确匹配的 VLOOKUP 第 4 参数写 FALSE。',
      '- 优先用最常见、最好懂的写法（SUMIF/SUMIFS、COUNTIF/COUNTIFS、AVERAGEIF(S)、IF、VLOOKUP、INDEX+MATCH、DATEDIF 等）。',
      '- 查找类结果建议外面包一层 IFERROR(…,"未找到")。',
      '',
      '## JSON 输出格式（字段都必须有）',
      '- status："ok" 表示已给出公式；"unsupported" 表示 Excel 2016 公式无法实现或超出范围（如需要 VBA、透视表、跨文件）；"need_clarification" 表示需求有歧义无法判断。非 ok 时 message 用一两句中文说明原因或需要用户补充什么，outputs 为空数组。',
      '- message：字符串，ok 时为空字符串。',
      '- kind："column" / "scalar" / "group"。',
      '- sheet：公式基于的工作表名（通常是当前工作表）。',
      '- groupBy：仅 group 需要，形如 [{"column":"部门","unit":"value"}]，unit 为 value（按取值）、month（按月）、year（按年）；其他类型给空数组。',
      '- outputs：数组，每项 {"title":简短中文标题, "formula":以 = 开头并使用上面引用写法的公式, "format":"general"/"integer"/"number"/"percent"/"date"/"money" 之一}。',
      '- interpretation：一句话复述你理解的需求（含条件）。',
      '- explanation：2～4 句通俗中文，讲清公式每部分在做什么，面向不懂函数的人。',
      '',
      '示例（仅示意格式）：',
      JSON.stringify(EXAMPLE),
      '',
      '只输出 JSON 对象本身，不要输出任何其他文字，不要用 Markdown 代码块。',
    ].join('\n');
  }

  function describeColumn(sheet, col, withSamples) {
    const typeName = { number: '数字', date: '日期', text: '文本', textnum: '文本形式的数字', bool: '逻辑值', mixed: '混合', empty: '空' }[col.type] || col.type;
    let s = '  - 「' + col.name + '」（' + col.letter + ' 列，' + typeName + (col.added ? '，本工具新增的列' : '') + '）';
    if (col.isIdCard || col.isPhone || col.isName) {
      // 个人信息列：只说明类型，不发送任何样例
      s += '，' + (col.isIdCard ? '身份证号' : col.isPhone ? '手机号' : '人名') + '（样例已隐藏）';
    } else if (withSamples) {
      if (col.isCategory && col.distinct && col.distinctCount <= 30) {
        s += '，共 ' + col.distinctCount + ' 种取值：' + Array.from(col.distinct.keys()).slice(0, 30).join('、');
      } else if (col.samples && col.samples.length) {
        s += '，示例：' + col.samples.slice(0, 3).join(' | ');
      }
      if (col.numStats && col.type === 'number') s += '（范围 ' + U.formatNumber(col.numStats.min) + ' ~ ' + U.formatNumber(col.numStats.max) + '）';
    }
    return s;
  }

  /** 需求里是否提到了这一列 / 这张表 */
  const mentions = (request, name) => !!request && String(name).length >= 2 && request.indexOf(name) >= 0;

  /** 选出要描述的列：需求里提到的列优先，其余按顺序补足到上限 */
  function pickColumns(sheet, request) {
    const max = EFA.LIMITS.ai.maxColsPerSheet;
    if (sheet.columns.length <= max) return { cols: sheet.columns, hidden: 0 };
    const hit = sheet.columns.filter(c => mentions(request, c.name));
    const rest = sheet.columns.filter(c => !hit.includes(c));
    const cols = hit.concat(rest.slice(0, Math.max(0, max - hit.length)));
    cols.sort((a, b) => a.index - b.index);
    return { cols, hidden: sheet.columns.length - cols.length };
  }

  /**
   * 表结构描述。level 越高越精简：
   * 0 全量；1 其他表不带样例；2 只保留需求里提到的其他表；3 当前表也不带样例
   */
  function buildContext(wb, sheet, withSamples, request, level) {
    level = level || 0;
    const A = EFA.LIMITS.ai;
    const lines = [];
    lines.push('当前工作表：「' + sheet.name + '」，表头在第 ' + (sheet.headerRow + 1) + ' 行，数据从第 ' + (sheet.dataStart + 1) + ' 行到第 ' + (sheet.dataEnd + 1) + ' 行（共 ' + (sheet.dataEnd - sheet.dataStart + 1) + ' 行）。列：');
    const own = pickColumns(sheet, request);
    own.cols.forEach(c => lines.push(describeColumn(sheet, c, withSamples && level < 3)));
    if (own.hidden) lines.push('  （另有 ' + own.hidden + ' 列未列出）');
    let others = wb.sheets.filter(s => s !== sheet && s.columns.length);
    const mentioned = s => mentions(request, s.name) || s.columns.some(c => mentions(request, c.name) && !sheet.columns.some(x => x.name === c.name));
    others.sort((a, b) => (mentioned(b) ? 1 : 0) - (mentioned(a) ? 1 : 0));
    if (level >= 2) others = others.filter(mentioned);
    const omitted = Math.max(0, others.length - A.maxOtherSheets);
    others = others.slice(0, A.maxOtherSheets);
    if (others.length) {
      lines.push('');
      lines.push('同一文件中的其他工作表：');
      others.forEach(s => {
        const pc = pickColumns(s, request);
        lines.push('- 「' + s.name + '」（' + (s.dataEnd - s.dataStart + 1) + ' 行数据）列：');
        pc.cols.forEach(c => lines.push(describeColumn(s, c, withSamples && level < 1)));
        if (pc.hidden) lines.push('  （另有 ' + pc.hidden + ' 列未列出）');
      });
    }
    const totalOthers = wb.sheets.filter(s => s !== sheet && s.columns.length).length;
    if (totalOthers > others.length) lines.push('（文件中另有 ' + (totalOthers - others.length) + ' 个工作表未列出' + (omitted ? '' : '，与需求无关') + '）');
    lines.push('');
    lines.push('今天的日期：' + U.formatSerialDate(U.todaySerial()));
    return lines.join('\n');
  }

  /** 在 token 预算内生成表结构描述；精简到最后仍超出则报错 */
  function contextWithinBudget(wb, sheet, withSamples, request) {
    const A = EFA.LIMITS.ai;
    const budget = A.maxInputTokens - estimateTokens(systemPrompt()) - estimateTokens(request) - 80;
    let ctx = '';
    for (let level = 0; level <= 3; level++) {
      ctx = buildContext(wb, sheet, withSamples, request, level);
      if (estimateTokens(ctx) <= budget) return { ctx, level };
    }
    const e = new Error('这个表格的列太多，超出了单次发给 AI 的上限（约 ' + A.maxInputTokens + ' tokens）。可以删掉无关的列，或在需求里写明要用到的列名后重试。');
    e.kind = 'limit';
    throw e;
  }

  /** 估算一次生成要发送的 token 数（供页面显示） */
  function estimateRequest(wb, sheet, request) {
    const st = getSettings();
    const { ctx } = contextWithinBudget(wb, sheet, st.sendSamples, request);
    return estimateTokens(systemPrompt()) + estimateTokens(ctx) + estimateTokens(request);
  }

  /* ---------- 解析与校验 AI 返回 ---------- */

  function extractJson(text) {
    let t = String(text || '').trim();
    t = t.replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '').trim();
    const a = t.indexOf('{'), b = t.lastIndexOf('}');
    if (a < 0 || b <= a) throw new Error('AI 返回的内容不是 JSON');
    return JSON.parse(t.slice(a, b + 1));
  }

  const FORMATS = ['general', 'integer', 'number', 'percent', 'date', 'money'];
  function checkShape(out) {
    if (!out || typeof out !== 'object') throw new Error('AI 返回格式不对');
    const status = ['ok', 'unsupported', 'need_clarification'].includes(out.status) ? out.status : 'ok';
    const kind = ['column', 'scalar', 'group'].includes(out.kind) ? out.kind : null;
    const outputs = Array.isArray(out.outputs) ? out.outputs.filter(o => o && typeof o.formula === 'string' && o.formula.trim()) : [];
    if (status === 'ok' && (!kind || !outputs.length)) throw new Error('AI 没有给出完整的公式');
    return {
      status, kind,
      message: typeof out.message === 'string' ? out.message : '',
      sheet: typeof out.sheet === 'string' ? out.sheet : '',
      groupBy: Array.isArray(out.groupBy) ? out.groupBy.filter(g => g && g.column).map(g => ({ column: String(g.column), unit: ['value', 'month', 'year'].includes(g.unit) ? g.unit : 'value' })) : [],
      outputs: outputs.map(o => ({ title: String(o.title || '计算结果'), formula: o.formula.trim(), format: FORMATS.includes(o.format) ? o.format : 'general' })),
      interpretation: typeof out.interpretation === 'string' ? out.interpretation : '',
      explanation: typeof out.explanation === 'string' ? out.explanation : '',
    };
  }

  /* ---------- 调用 ---------- */

  function httpError(status, body, st) {
    const msg = body && body.error ? (body.error.message || body.error.code || '') : '';
    const qwen = st.provider === 'qwen';
    switch (status) {
      case 400: case 422: return 'AI 请求参数有误：' + (msg || status);
      case 401: return 'API Key 无效' + (qwen ? '（百炼的 Key 要与地区/工作空间对应）' : '') + '，请在“AI 设置”里检查';
      case 402: return '账户余额不足，请到 ' + PROVIDERS[st.provider].name + ' 平台充值';
      case 403: return '没有权限调用该模型' + (qwen ? '，请确认已在百炼控制台开通该模型' : '') + (msg ? '：' + msg : '');
      case 404: return '找不到该模型或接口地址不正确' + (msg ? '：' + msg : '');
      case 429: return '请求太频繁或额度已用完，请稍后再试';
      case 500: case 502: case 503: case 504: return 'AI 服务繁忙，请稍后再试';
      default: return 'AI 服务出错（' + status + '）' + (msg ? '：' + msg : '');
    }
  }

  async function chat(st, messages, opts, useJsonMode) {
    const p = PROVIDERS[st.provider];
    const preset = p.models.find(m => m.id === st.model) || null;
    const A = EFA.LIMITS.ai;
    // 自定义模型不知道是否带思考过程，按思考型给足输出上限，避免被截断
    const thinking = preset ? preset.thinking : true;
    const body = Object.assign({
      model: st.model,
      messages,
      stream: false,
      max_tokens: thinking ? A.maxOutputTokensThinking : A.maxOutputTokens,
    }, useJsonMode ? { response_format: { type: 'json_object' } } : {}, p.extra(preset));
    checkDaily();

    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(new DOMException('timeout', 'TimeoutError')), 180000);
    const onAbort = () => ctl.abort();
    if (opts.signal) opts.signal.addEventListener('abort', onAbort);
    let resp;
    try {
      resp = await fetch(endpoint(st), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + st.apiKey },
        body: JSON.stringify(body),
        signal: ctl.signal,
      });
    } catch (e) {
      if (opts.signal && opts.signal.aborted) throw new Error('已取消');
      if (ctl.signal.aborted) throw new Error('AI 响应超时，请稍后再试或换用更快的模型');
      throw new Error('无法连接 ' + p.name + '，请检查网络或“接口地址”设置');
    } finally {
      clearTimeout(timer);
      if (opts.signal) opts.signal.removeEventListener('abort', onAbort);
    }
    let data = null;
    try { data = await resp.json(); } catch (e) { data = null; }
    // 请求已到达服务商就计入次数（失败的请求一般不计费，但也计数，防止反复重试）
    recordUsage(data && data.usage);
    if (!resp.ok) {
      const e = new Error(httpError(resp.status, data, st));
      e.status = resp.status;
      e.raw = data && data.error ? (data.error.message || '') : '';
      throw e;
    }
    const choice = data && data.choices && data.choices[0];
    if (!choice || !choice.message) throw new Error('AI 没有返回结果');
    return { content: choice.message.content || '', finish: choice.finish_reason, model: data.model || st.model, usage: data.usage || null };
  }

  /**
   * 生成计划。validate(plan) 抛出的错误会作为反馈让 AI 修正一次。
   * opts: { signal, onStatus, validate }
   */
  async function generate(request, wb, sheet, opts) {
    opts = opts || {};
    const st = getSettings();
    if (!st.apiKey) throw new Error('还没有设置 API Key');
    const p = PROVIDERS[st.provider];
    checkDaily();
    const { ctx, level } = contextWithinBudget(wb, sheet, st.sendSamples, request);
    const messages = [
      { role: 'system', content: systemPrompt() },
      { role: 'user', content: '## 数据\n' + ctx + '\n\n## 需求\n' + request + '\n\n请按要求输出 JSON。' },
    ];
    const est = estimateTokens(messages[0].content) + estimateTokens(messages[1].content);
    const spent = { calls: 0, inTok: 0, outTok: 0, estimate: est, trimmed: level > 0 };
    let jsonMode = true;
    let lastErr = null;
    let blankRetry = 0;
    for (let attempt = 0; attempt < 2; attempt++) {
      if (spent.calls >= EFA.LIMITS.ai.maxCallsPerRequest) break;
      opts.onStatus && opts.onStatus((attempt ? p.name + ' 正在修正公式…' : p.name + ' 正在生成公式…') + '（本次发送约 ' + est.toLocaleString() + ' tokens）');
      let res;
      try {
        spent.calls++;
        res = await chat(st, messages, opts, jsonMode);
        if (res.usage) { spent.inTok += res.usage.prompt_tokens || 0; spent.outTok += res.usage.completion_tokens || 0; }
      } catch (e) {
        // 个别模型不支持 JSON 模式：去掉该参数重试一次
        if (jsonMode && (e.status === 400 || e.status === 422) && /response_format|json/i.test(e.raw || '')) {
          jsonMode = false;
          attempt--;
          continue;
        }
        throw e;
      }
      if (res.finish === 'length') throw new Error('AI 输出被截断，请把需求说得更简短些');
      if (!res.content.trim()) {
        if (blankRetry++ < 1) { attempt--; continue; }
        throw new Error('AI 返回了空内容，请稍后再试');
      }
      let out;
      try { out = checkShape(extractJson(res.content)); } catch (e) {
        lastErr = e;
        messages.push({ role: 'assistant', content: res.content });
        messages.push({ role: 'user', content: '你的输出无法解析为要求的 JSON：' + e.message + '。请只输出符合格式要求的 JSON 对象。' });
        continue;
      }
      if (out.status !== 'ok') {
        const e = new Error(out.message || (out.status === 'unsupported' ? '这个需求无法用 Excel 2016 公式实现' : '需求不够明确'));
        e.kind = out.status;
        throw e;
      }
      const plan = {
        kind: out.kind,
        sheet: wb.sheets.some(s => s.name === out.sheet) ? out.sheet : sheet.name,
        groupBy: out.groupBy,
        outputs: out.outputs,
        interpretation: out.interpretation,
        explanation: out.explanation,
        source: 'ai',
        model: p.name + ' · ' + res.model,
        request,
        usage: spent,
      };
      if (!opts.validate) return plan;
      try {
        const extra = opts.validate(plan);
        return extra || plan;
      } catch (e) {
        lastErr = e;
        if (!e.validation && !/找不到列|无效|不支持|语法|Excel 2019|不认识|没有生成|分组/.test(e.message || '')) throw e;
        messages.push({ role: 'assistant', content: res.content });
        messages.push({ role: 'user', content: '上面的公式无法使用：' + e.message + '。请修正后重新输出 JSON，仍需满足全部要求。' });
      }
    }
    throw lastErr || new Error('AI 生成的公式无法使用');
  }

  EFA.LLM = {
    PROVIDERS, PROVIDER_LIST, getSettings, saveSettings, clearKey, isConfigured, providerName, modelName, generate,
    buildContext, contextWithinBudget, estimateTokens, estimateRequest, getUsage, endpoint, systemPrompt,
  };
})(window.EFA = window.EFA || {});
