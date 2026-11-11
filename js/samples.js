/* 示例数据（全部为虚构数据，仅用于体验） */
(function (EFA) {
  'use strict';
  const U = EFA.U;

  // 固定种子的伪随机，保证每次生成的示例一致
  function rng(seed) {
    let s = seed >>> 0;
    return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
  }

  const SURNAMES = '王李张刘陈杨黄赵周吴徐孙马朱胡郭何林高罗'.split('');
  const GIVEN = ['伟', '芳', '娜', '敏', '静', '磊', '洋', '艳', '勇', '军', '杰', '娟', '涛', '明', '超', '秀英', '霞', '平', '刚', '桂英', '晨', '欣怡', '子轩', '雨涵', '浩然'];

  function people(r, n) {
    const used = new Set();
    const out = [];
    while (out.length < n) {
      const name = SURNAMES[Math.floor(r() * SURNAMES.length)] + GIVEN[Math.floor(r() * GIVEN.length)];
      if (used.has(name)) continue;
      used.add(name);
      out.push(name);
    }
    return out;
  }

  /** 生成示例工作簿（返回 xlsx 二进制，走与上传文件完全相同的流程） */
  function buildSampleWorkbook() {
    const r = rng(20240501);
    const wb = XLSX.utils.book_new();

    // 员工信息
    const depts = ['销售部', '市场部', '技术部', '财务部', '人事部'];
    const names = people(r, 24);
    const staff = [['工号', '姓名', '性别', '部门', '身份证号', '入职日期', '基本工资', '手机号']];
    names.forEach((nm, i) => {
      const y = 1975 + Math.floor(r() * 26), m = 1 + Math.floor(r() * 12), d = 1 + Math.floor(r() * 28);
      const male = r() > 0.5;
      const seq = String(100 + Math.floor(r() * 800));
      const genderDigit = String((Math.floor(r() * 5) * 2) + (male ? 1 : 0));
      // 地区码使用不存在的 99 开头，确保不是真实号码
      const id = '990101' + y + String(m).padStart(2, '0') + String(d).padStart(2, '0') + seq + genderDigit;
      const hire = U.partsToSerial(2012 + Math.floor(r() * 12), 1 + Math.floor(r() * 12), 1 + Math.floor(r() * 28));
      const dept = depts[i % depts.length === 0 && i < 10 ? 0 : Math.floor(r() * depts.length)];
      staff.push(['E' + String(1001 + i), nm, male ? '男' : '女', dept, id, hire, 5000 + Math.round(r() * 120) * 100, '139' + String(10000000 + Math.floor(r() * 89999999))]);
    });
    const ws1 = XLSX.utils.aoa_to_sheet(staff);
    setDateFormat(ws1, 5, staff.length);

    // 销售明细
    const sellers = staff.slice(1).filter(s => s[3] === '销售部' || s[3] === '市场部').slice(0, 8);
    const products = [['笔记本电脑', 4999], ['显示器', 1299], ['机械键盘', 399], ['无线鼠标', 129], ['扩展坞', 259], ['耳机', 599]];
    const regions = ['华东', '华南', '华北', '西南'];
    const sales = [['日期', '订单号', '工号', '销售员', '区域', '产品', '单价', '数量', '销售额']];
    for (let i = 0; i < 60; i++) {
      const s = sellers[Math.floor(r() * sellers.length)];
      const p = products[Math.floor(r() * products.length)];
      const qty = 1 + Math.floor(r() * 12);
      const date = U.partsToSerial(2024, 1 + Math.floor(i / 10), 1 + Math.floor(r() * 28));
      sales.push([date, 'SO' + String(240001 + i), s[0], s[1], regions[Math.floor(r() * regions.length)], p[0], p[1], qty, p[1] * qty]);
    }
    sales.sort((a, b) => (typeof a[0] === 'number' && typeof b[0] === 'number' ? a[0] - b[0] : 0));
    const ws2 = XLSX.utils.aoa_to_sheet(sales);
    setDateFormat(ws2, 0, sales.length);

    // 成绩表
    const classes = ['一班', '二班', '三班'];
    const stu = people(rng(7), 18);
    const scores = [['学号', '姓名', '班级', '语文', '数学', '英语']];
    const r3 = rng(99);
    stu.forEach((nm, i) => {
      scores.push(['2024' + String(i + 1).padStart(3, '0'), nm, classes[i % 3], 55 + Math.floor(r3() * 45), 50 + Math.floor(r3() * 50), 58 + Math.floor(r3() * 42)]);
    });
    const ws3 = XLSX.utils.aoa_to_sheet(scores);

    XLSX.utils.book_append_sheet(wb, ws2, '销售明细');
    XLSX.utils.book_append_sheet(wb, ws1, '员工信息');
    XLSX.utils.book_append_sheet(wb, ws3, '成绩表');
    return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  }

  function setDateFormat(ws, col, rows) {
    for (let rr = 1; rr < rows; rr++) {
      const ref = XLSX.utils.encode_cell({ r: rr, c: col });
      if (ws[ref]) { ws[ref].t = 'n'; ws[ref].z = 'yyyy-mm-dd'; }
    }
  }

  EFA.Samples = { buildSampleWorkbook };
})(window.EFA = window.EFA || {});
