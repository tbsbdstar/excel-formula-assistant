/* 为预览引擎补充 Excel 2016 常用但引擎原生缺失的函数：RANK / RANK.EQ / AVERAGEIFS / AGGREGATE
 * 只影响浏览器内的预览计算；导出到文件里的仍是标准 Excel 公式，由 Excel / WPS 自己计算。
 */
(function (EFA) {
  'use strict';
  if (!window.HyperFormula) return;
  const HF = window.HyperFormula;
  const T = HF.FunctionArgumentType;
  const { CellError, ErrorType } = HF;

  const isEmpty = v => v === null || v === undefined || typeof v === 'symbol' || v === '';
  const raw = v => {
    if (typeof v === 'symbol') return null;
    if (v && typeof v === 'object' && !(v instanceof CellError) && 'val' in v) return v.val;
    return v;
  };
  const isErr = v => v instanceof CellError;

  function rangeValues(r) {
    if (r && typeof r.valuesFromTopLeftCorner === 'function') return r.valuesFromTopLeftCorner();
    return [r];
  }

  function wildcardRegex(s) {
    let re = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '~' && i + 1 < s.length && /[*?~]/.test(s[i + 1])) { re += '\\' + s[++i]; continue; }
      if (ch === '*') re += '[\\s\\S]*';
      else if (ch === '?') re += '[\\s\\S]';
      else re += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&');
    }
    return new RegExp('^' + re + '$', 'i');
  }

  /** Excel 风格的条件匹配（COUNTIF 一族的规则） */
  function makeCriterion(c) {
    c = raw(c);
    if (typeof c === 'number') {
      return v => { const r = raw(v); return (typeof r === 'number' && r === c) || (typeof r === 'string' && r.trim() !== '' && Number(r) === c); };
    }
    if (typeof c === 'boolean') return v => raw(v) === c;
    let s = c == null ? '' : String(c);
    let op = '=';
    const m = s.match(/^(>=|<=|<>|>|<|=)/);
    if (m) { op = m[1]; s = s.slice(m[1].length); }
    const num = s.trim() !== '' && !isNaN(Number(s)) ? Number(s) : null;
    if (op === '=' || op === '<>') {
      let test;
      if (s === '') test = v => isEmpty(raw(v));
      else if (num !== null) test = v => { const r = raw(v); return (typeof r === 'number' && r === num) || (typeof r === 'string' && r.trim() !== '' && Number(r) === num); };
      else {
        const re = wildcardRegex(s);
        test = v => { const r = raw(v); return !isEmpty(r) && !isErr(r) && re.test(String(r)); };
      }
      return op === '=' ? test : (v => !isErr(v) && !test(v));
    }
    const cmp = (a, b) => (op === '>' ? a > b : op === '<' ? a < b : op === '>=' ? a >= b : a <= b);
    if (num !== null) return v => { const r = raw(v); return typeof r === 'number' && cmp(r, num); };
    const sl = s.toLowerCase();
    return v => { const r = raw(v); return typeof r === 'string' && cmp(r.toLowerCase().localeCompare(sl, 'zh'), 0); };
  }

  function quantile(sorted, p) {
    if (!sorted.length) return null;
    const h = (sorted.length - 1) * p;
    const lo = Math.floor(h), hi = Math.ceil(h);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (h - lo);
  }

  class EfaPlugin extends HF.FunctionPlugin {
    rank(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('RANK'), (num, range, order) => {
        const vals = rangeValues(range).map(raw).filter(v => typeof v === 'number');
        if (!vals.includes(num)) return new CellError(ErrorType.NA);
        const n = order ? vals.filter(v => v < num).length : vals.filter(v => v > num).length;
        return n + 1;
      });
    }

    rankEq(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('RANK.EQ'), (num, range, order) => {
        const vals = rangeValues(range).map(raw).filter(v => typeof v === 'number');
        if (!vals.includes(num)) return new CellError(ErrorType.NA);
        const n = order ? vals.filter(v => v < num).length : vals.filter(v => v > num).length;
        return n + 1;
      });
    }

    averageifs(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('AVERAGEIFS'), (avgRange, ...rest) => {
        const avg = rangeValues(avgRange);
        const pairs = [];
        for (let i = 0; i + 1 < rest.length; i += 2) pairs.push([rangeValues(rest[i]), makeCriterion(rest[i + 1])]);
        for (const [r] of pairs) if (r.length !== avg.length) return new CellError(ErrorType.VALUE);
        let sum = 0, cnt = 0;
        for (let i = 0; i < avg.length; i++) {
          if (!pairs.every(([r, t]) => t(r[i]))) continue;
          const v = avg[i];
          if (isErr(v)) return v;
          const r = raw(v);
          if (typeof r === 'number') { sum += r; cnt++; }
        }
        return cnt ? sum / cnt : new CellError(ErrorType.DIV_BY_ZERO);
      });
    }

    aggregate(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('AGGREGATE'), (fn, opt, arr, k) => {
        const ignoreErr = [2, 3, 6, 7].includes(opt);
        const nums = [];
        for (const v of rangeValues(arr)) {
          if (isErr(v)) { if (ignoreErr) continue; return v; }
          const r = raw(v);
          if (typeof r === 'number') nums.push(r);
        }
        nums.sort((a, b) => a - b);
        switch (fn) {
          case 1: return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : new CellError(ErrorType.DIV_BY_ZERO);
          case 2: return nums.length;
          case 4: return nums.length ? nums[nums.length - 1] : 0;
          case 5: return nums.length ? nums[0] : 0;
          case 9: return nums.reduce((a, b) => a + b, 0);
          case 12: return nums.length ? quantile(nums, 0.5) : new CellError(ErrorType.NUM);
          case 14: { const i = (k || 1); return i >= 1 && i <= nums.length ? nums[nums.length - i] : new CellError(ErrorType.NUM); }
          case 15: { const i = (k || 1); return i >= 1 && i <= nums.length ? nums[i - 1] : new CellError(ErrorType.NUM); }
          case 16: case 18: return nums.length && k >= 0 && k <= 1 ? quantile(nums, k) : new CellError(ErrorType.NUM);
          case 17: case 19: return nums.length && k >= 0 && k <= 4 ? quantile(nums, k / 4) : new CellError(ErrorType.NUM);
          default: return new CellError(ErrorType.VALUE);
        }
      });
    }
  }

  const rankParams = [
    { argumentType: T.NUMBER },
    { argumentType: T.RANGE },
    { argumentType: T.NUMBER, defaultValue: 0 },
  ];
  EfaPlugin.implementedFunctions = {
    RANK: { method: 'rank', parameters: rankParams },
    'RANK.EQ': { method: 'rankEq', parameters: rankParams },
    AVERAGEIFS: {
      method: 'averageifs',
      parameters: [{ argumentType: T.RANGE }, { argumentType: T.RANGE }, { argumentType: T.NOERROR }],
      repeatLastArgs: 2,
    },
    AGGREGATE: {
      method: 'aggregate',
      parameters: [
        { argumentType: T.NUMBER },
        { argumentType: T.NUMBER },
        { argumentType: T.RANGE },
        { argumentType: T.NUMBER, optionalArg: true },
      ],
    },
  };

  const names = { RANK: 'RANK', 'RANK.EQ': 'RANK.EQ', AVERAGEIFS: 'AVERAGEIFS', AGGREGATE: 'AGGREGATE' };
  try {
    HF.registerFunctionPlugin(EfaPlugin, { enGB: names, enUS: names });
  } catch (e) {
    console.warn('补充函数注册失败', e);
  }

  EFA.criterion = makeCriterion;
})(window.EFA = window.EFA || {});
