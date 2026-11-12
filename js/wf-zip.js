/* 工作流配置器：把配置文件放进“Excel 工作流工具”压缩包，生成一个可以直接发给别人的完整工具包。
 * 做法：原压缩包里的文件原样保留（不解压、不重新压缩），只把 configs/ 里的旧配置去掉、追加本工作流的配置，再重写目录。 */
(function (EFA) {
  'use strict';
  const Z = {};

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  Z.crc32 = function (bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  };

  function fail(msg) { const e = new Error(msg); e.kind = 'user'; throw e; }

  /** 读取压缩包目录 → { entries: [{ name, cdStart, cdEnd, flags }], cdOffset, buf } */
  Z.readDirectory = function (buf) {
    const u8 = new Uint8Array(buf);
    const dv = new DataView(buf);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 22 - 65535); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) fail('这不是有效的 zip 压缩包。');
    const total = dv.getUint16(eocd + 10, true);
    const cdSize = dv.getUint32(eocd + 12, true);
    const cdOffset = dv.getUint32(eocd + 16, true);
    if (total === 0xFFFF || cdOffset === 0xFFFFFFFF) fail('压缩包太大（ZIP64 格式），暂不支持。');
    if (cdOffset + cdSize > eocd) fail('压缩包已损坏。');
    const dec = new TextDecoder('utf-8');
    const latin = new TextDecoder('latin1');
    const entries = [];
    let p = cdOffset;
    for (let i = 0; i < total; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) fail('压缩包目录已损坏。');
      const flags = dv.getUint16(p + 8, true);
      const nLen = dv.getUint16(p + 28, true), xLen = dv.getUint16(p + 30, true), cLen = dv.getUint16(p + 32, true);
      const nameBytes = u8.subarray(p + 46, p + 46 + nLen);
      const name = (flags & 0x800) ? dec.decode(nameBytes) : latin.decode(nameBytes);
      const end = p + 46 + nLen + xLen + cLen;
      entries.push({ name: name.replace(/\\/g, '/'), cdStart: p, cdEnd: end, flags });
      p = end;
    }
    return { entries, cdOffset, buf, u8 };
  };

  /** 找出工具包根目录（含 run.bat、process.py、python/python.exe 的那一层） */
  Z.toolkitRoot = function (dir) {
    const names = new Set(dir.entries.map(e => e.name));
    for (const e of dir.entries) {
      if (!/(^|\/)process\.py$/.test(e.name)) continue;
      const prefix = e.name.slice(0, e.name.length - 'process.py'.length);
      if (names.has(prefix + 'run.bat') && names.has(prefix + 'python/python.exe')) return prefix;
    }
    return null;
  };

  /** 检查压缩包是否为工作流工具包；返回 { root, version } */
  Z.inspect = function (buf) {
    const dir = Z.readDirectory(buf);
    const root = Z.toolkitRoot(dir);
    if (root === null) fail('这个压缩包不是“Excel 工作流工具”（里面没有 run.bat、process.py 和 python 文件夹）。请选择工具包的 zip 文件。');
    const m = /v(\d+\.\d+\.\d+)/.exec(dir.entries.map(e => e.name).find(n => /版本/.test(n)) || '');
    return { root, entries: dir.entries.length, version: m ? m[1] : '' };
  };

  function dosTime(d) {
    const time = (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2);
    const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    return { time, date };
  }

  /**
   * 生成新的工具包：去掉原 configs/ 下的 .json，加入 files（[{ path: 'configs/x.json', text }]，相对工具包根目录）。
   * 返回 Blob。
   */
  Z.build = function (buf, files) {
    const dir = Z.readDirectory(buf);
    const root = Z.toolkitRoot(dir);
    if (root === null) fail('这个压缩包不是“Excel 工作流工具”。');
    const enc = new TextEncoder();
    const now = dosTime(new Date());
    const addNames = new Set(files.map(f => root + f.path));
    const keep = dir.entries.filter(e => {
      if (addNames.has(e.name)) return false;
      const rel = e.name.slice(root.length);
      return !(e.name.startsWith(root) && /^configs\/[^/]+\.json$/i.test(rel));
    });
    const parts = [dir.u8.subarray(0, dir.cdOffset)];
    let offset = dir.cdOffset;
    const newCd = [];
    files.forEach(f => {
      const name = enc.encode(root + f.path);
      const data = enc.encode(f.text);
      const crc = Z.crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0x800, true);          // 文件名为 UTF-8
      lh.setUint16(8, 0, true);              // 不压缩
      lh.setUint16(10, now.time, true);
      lh.setUint16(12, now.date, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true);
      lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true);
      lh.setUint16(28, 0, true);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true);
      cd.setUint16(4, 20, true);
      cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x800, true);
      cd.setUint16(10, 0, true);
      cd.setUint16(12, now.time, true);
      cd.setUint16(14, now.date, true);
      cd.setUint32(16, crc, true);
      cd.setUint32(20, data.length, true);
      cd.setUint32(24, data.length, true);
      cd.setUint16(28, name.length, true);
      cd.setUint32(42, offset, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      newCd.push(new Uint8Array(cd.buffer), name);
      offset += 30 + name.length + data.length;
    });
    const cdStart = offset;
    let cdSize = 0;
    keep.forEach(e => { const b = dir.u8.subarray(e.cdStart, e.cdEnd); parts.push(b); cdSize += b.length; });
    newCd.forEach(b => { parts.push(b); cdSize += b.length; });
    const count = keep.length + files.length;
    if (count > 0xFFFF || cdStart + cdSize > 0xFFFFFFFF) fail('压缩包太大，无法写入。');
    const eocd = new DataView(new ArrayBuffer(22));
    eocd.setUint32(0, 0x06054b50, true);
    eocd.setUint16(8, count, true);
    eocd.setUint16(10, count, true);
    eocd.setUint32(12, cdSize, true);
    eocd.setUint32(16, cdStart, true);
    parts.push(new Uint8Array(eocd.buffer));
    return new Blob(parts, { type: 'application/zip' });
  };

  /* ---------- 记住工具包（浏览器本地，IndexedDB），下次不用重新选择 ---------- */
  const DB = 'efa-workflow', STORE = 'kv';
  function open() {
    return new Promise((resolve, reject) => {
      const rq = indexedDB.open(DB, 1);
      rq.onupgradeneeded = () => rq.result.createObjectStore(STORE);
      rq.onsuccess = () => resolve(rq.result);
      rq.onerror = () => reject(rq.error);
    });
  }
  Z.saveToolkit = async function (rec) {
    try {
      const db = await open();
      await new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).put(rec, 'toolkit');
        tx.oncomplete = resolve; tx.onerror = () => reject(tx.error);
      });
      return true;
    } catch (e) { return false; }
  };
  Z.loadToolkit = async function () {
    try {
      const db = await open();
      return await new Promise((resolve) => {
        const rq = db.transaction(STORE).objectStore(STORE).get('toolkit');
        rq.onsuccess = () => resolve(rq.result || null);
        rq.onerror = () => resolve(null);
      });
    } catch (e) { return null; }
  };
  Z.forgetToolkit = async function () {
    try {
      const db = await open();
      db.transaction(STORE, 'readwrite').objectStore(STORE).delete('toolkit');
    } catch (e) { /* ignore */ }
  };

  EFA.WFZip = Z;
})(window.EFA = window.EFA || {});
