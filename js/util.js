/* 通用工具函数 */
(function (EFA) {
  'use strict';
  const U = {};

  U.colLetter = function (idx0) {
    let n = idx0 + 1, s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  U.colIndex = function (letters) {
    let n = 0;
    for (const ch of String(letters).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  };

  U.escapeHtml = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  };

  /** 工作表名在公式里一律加单引号，最稳妥 */
  U.quoteSheet = function (name) {
    return "'" + String(name).replace(/'/g, "''") + "'";
  };

  /** 公式中的文本常量 */
  U.excelString = function (s) {
    return '"' + String(s).replace(/"/g, '""') + '"';
  };

  /* ---------- 日期（1900 日期系统，序列号 ≥ 61 与 Excel 一致） ---------- */
  U.serialToParts = function (serial) {
    const ms = Math.round((serial - 25569) * 86400000);
    const d = new Date(ms);
    return {
      y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
      hh: d.getUTCHours(), mi: d.getUTCMinutes(), ss: d.getUTCSeconds(),
    };
  };
  U.partsToSerial = function (y, m, d) {
    return Date.UTC(y, m - 1, d) / 86400000 + 25569;
  };
  U.todaySerial = function () {
    const n = new Date();
    return U.partsToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate());
  };
  const pad = n => (n < 10 ? '0' : '') + n;
  U.formatSerialDate = function (serial, withTime) {
    if (typeof serial !== 'number' || !isFinite(serial)) return String(serial);
    const p = U.serialToParts(serial);
    let s = p.y + '-' + pad(p.m) + '-' + pad(p.d);
    if (withTime && (p.hh || p.mi || p.ss)) s += ' ' + pad(p.hh) + ':' + pad(p.mi) + (p.ss ? ':' + pad(p.ss) : '');
    return s;
  };
  U.formatSerialMonth = function (serial) {
    const p = U.serialToParts(serial);
    return p.y + '年' + p.m + '月';
  };

  /* ---------- 文本规范化 ---------- */
  U.halfWidth = function (str) {
    return String(str)
      .replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
      .replace(/　/g, ' ');
  };

  const CN_DIGIT = { 零: 0, 〇: 0, 一: 1, 二: 2, 两: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9 };
  const CN_UNIT = { 十: 10, 百: 100, 千: 1000, 万: 10000 };
  /** 简单中文数字转换：三、十二、二十五、一百、两千五百 */
  U.parseCnNumber = function (s) {
    if (!s) return null;
    let total = 0, section = 0, num = 0, seen = false;
    for (const ch of s) {
      if (ch in CN_DIGIT) { num = CN_DIGIT[ch]; seen = true; }
      else if (ch in CN_UNIT) {
        const u = CN_UNIT[ch];
        if (u === 10000) { section = (section + num) * u; total += section; section = 0; }
        else section += (num || (u === 10 ? 1 : 0)) * u;
        num = 0; seen = true;
      } else return null;
    }
    return seen ? total + section + num : null;
  };

  /** 生成不重名的名字 */
  U.uniqueName = function (base, taken) {
    const has = n => (taken instanceof Set ? taken.has(n) : taken.includes(n));
    if (!has(base)) return base;
    for (let i = 2; i < 1000; i++) if (!has(base + i)) return base + i;
    return base + Date.now();
  };

  /** Excel 工作表名：≤31 字符，不含 []:*?/\ */
  U.sanitizeSheetName = function (name) {
    let s = String(name || '结果').replace(/[\[\]:*?/\\]/g, '_').replace(/^'+|'+$/g, '').trim();
    if (!s) s = '结果';
    if (s.length > 31) s = s.slice(0, 31);
    return s;
  };

  U.formatNumber = function (v, fmt) {
    if (typeof v !== 'number' || !isFinite(v)) return String(v);
    if (fmt === 'percent') return (Math.round(v * 10000) / 100).toFixed(2).replace(/\.?0+$/, '') + '%';
    if (fmt === 'integer') return String(Math.round(v));
    if (fmt === 'money') return v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (Number.isInteger(v)) return String(v);
    const r = Math.round(v * 1e6) / 1e6;
    return String(r);
  };

  U.debounce = function (fn, ms) {
    let t = null;
    return function () {
      const args = arguments, self = this;
      clearTimeout(t);
      t = setTimeout(() => fn.apply(self, args), ms);
    };
  };

  U.sleep = ms => new Promise(r => setTimeout(r, ms));

  EFA.U = U;
})(window.EFA = window.EFA || {});
