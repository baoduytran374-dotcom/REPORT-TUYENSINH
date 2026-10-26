/* BETMA ENGLISH — đọc file Excel tuyển sinh → dữ liệu dashboard.
 * Phần lõi chỉ làm việc trên mảng 2 chiều (aoa) nên chạy được cả trong trình duyệt lẫn Node (để kiểm thử).
 * Tiêu đề cột được nhận diện linh hoạt (không phân biệt hoa/thường, dấu tiếng Việt, khoảng trắng). */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.BetmaParser = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[đĐ]/g, 'd').toLowerCase().replace(/[^a-z0-9]+/g, '');

  const STUDENT_COLS = {
    id: ['mahv', 'mahocvien', 'mahocsinh', 'id', 'mshv'],
    name: ['hovaten', 'hoten', 'tenhocvien', 'hocvien'],
    reg: ['ngaydangky', 'ngaydk', 'ngayghidanh', 'ngaydangki'],
    kg: ['thangkhaigiang', 'thangkg', 'thangkhaigiangt'],
    cls: ['malop', 'lop', 'lophoc'],
    course: ['khoahoc', 'tenkhoahoc', 'khoa'],
    fee: ['hocphiapdung', 'hocphi', 'hocphisauuudai', 'hpapdung', 'hocphiphainop'],
    promo: ['uudai', 'chuongtrinhuudai', 'khuyenmai', 'uudaigiamgia'],
    tvv: ['tuvanvien', 'tvv', 'nguoituvan', 'tvvphutrach'],
    src: ['nguon', 'nguontuyensinh', 'nguonts'],
    st: ['tinhtrang', 'trangthai', 'tinhtranghocphi', 'trangthaihocphi'],
    note: ['ghichu', 'note'],
    paidTotal: ['dathanhtoan', 'tongthanhtoan', 'tongdathu']
  };
  const CLASS_COLS = {
    cls: ['malop', 'lop', 'lophoc'],
    sched: ['lichhoc', 'lich'],
    shift: ['cahoc', 'ca', 'giohoc'],
    teacher: ['giaovien', 'gv'],
    open: ['khaigiang', 'ngaykhaigiang', 'ngaykg']
  };
  const PAY_DATE = /^ngay(?:tt|thanhtoan)(\d+)$/;
  const PAY_AMT = /^sotien(?:tt|thanhtoan)?(\d+)$/;
  const STATUS = { dahoantat: 'Đã hoàn tất', chuahoantat: 'Chưa hoàn tất', candoisoat: 'Cần đối soát', chuadudulieu: 'Chưa đủ dữ liệu' };

  const pad = n => String(n).padStart(2, '0');
  const isoOf = (y, m, d) => y + '-' + pad(m) + '-' + pad(d);
  const validYMD = (y, m, d) => { const t = new Date(Date.UTC(y, m - 1, d)); return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d; };
  const serialToISO = n => { const t = new Date(Date.UTC(1899, 11, 30) + Math.floor(n) * 864e5); return isoOf(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); };
  const isoToSerial = s => { const [y, m, d] = s.split('-').map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5; };

  /** → {v: 'YYYY-MM-DD'|null, bad: bool} */
  function toDate(v) {
    if (v == null || v === '') return { v: null, bad: false };
    if (v instanceof Date && !isNaN(v)) return { v: isoOf(v.getFullYear(), v.getMonth() + 1, v.getDate()), bad: false };
    if (typeof v === 'number') return v > 20000 && v < 80000 ? { v: serialToISO(v), bad: false } : { v: null, bad: true };
    const s = String(v).trim(); if (!s) return { v: null, bad: false };
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
    if (m && validYMD(+m[1], +m[2], +m[3])) return { v: isoOf(+m[1], +m[2], +m[3]), bad: false };
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
    if (m) { let y = +m[3]; if (y < 100) y += 2000; if (validYMD(y, +m[2], +m[1])) return { v: isoOf(y, +m[2], +m[1]), bad: false }; }
    return { v: null, bad: true };
  }

  /** → {v: number|null, bad: bool} */
  function toMoney(v) {
    if (v == null || v === '') return { v: null, bad: false };
    if (typeof v === 'number') return isFinite(v) ? { v, bad: false } : { v: null, bad: true };
    let s = String(v).trim().replace(/[₫đĐ]|vnd|vnđ/gi, '').replace(/\s+/g, '');
    if (!s) return { v: null, bad: false };
    if (/^-?\d{1,3}([.,]\d{3})+$/.test(s)) s = s.replace(/[.,]/g, '');
    else s = s.replace(',', '.');
    const n = Number(s);
    return isFinite(n) && /^-?[\d.]+$/.test(s) ? { v: n, bad: false } : { v: null, bad: true };
  }

  const toKg = v => { if (v == null || v === '') return null; if (typeof v === 'number') return Math.round(v) || null; const m = String(v).match(/\d+/); return m ? +m[0] : null; };
  const txt = v => (v == null ? '' : String(v)).replace(/\s+/g, ' ').trim();

  function mapHeader(row, defs) {
    const col = {}, pays = {}; let score = 0;
    row.forEach((h, i) => {
      const k = norm(h); if (!k) return;
      for (const f in defs) if (col[f] == null && defs[f].includes(k)) { col[f] = i; score++; return; }
      let m = k.match(PAY_DATE); if (m) { (pays[+m[1]] = pays[+m[1]] || {}).d = i; return; }
      m = k.match(PAY_AMT); if (m) { (pays[+m[1]] = pays[+m[1]] || {}).a = i; }
    });
    return { col, pays, score };
  }
  function findHeader(aoa, defs, must) {
    let best = null;
    for (let r = 0; r < Math.min(aoa.length, 30); r++) {
      const h = mapHeader(aoa[r] || [], defs);
      if (must.every(f => h.col[f] != null) && h.score >= Math.min(3, must.length + 1) && (!best || h.score > best.score)) best = { ...h, row: r };
    }
    return best;
  }

  /** sheets: [{name, aoa}] → {ok, errors[], warnings[], data} */
  function parseSheets(sheets, opts) {
    opts = opts || {};
    const today = opts.today || (function () { const d = new Date(); return isoOf(d.getFullYear(), d.getMonth() + 1, d.getDate()); })();
    const errors = [], warnings = [], flags = [];
    // 1) chọn sheet học viên
    let cand = [];
    sheets.forEach(sh => { const h = findHeader(sh.aoa || [], STUDENT_COLS, ['name']); if (h) cand.push({ sh, h }); });
    if (!cand.length) return { ok: false, errors: ['Không tìm thấy sheet học viên. Cần một dòng tiêu đề có cột "Họ và tên" (và nên có "Mã HV", "Ngày đăng ký", "Học phí áp dụng"…). Hãy dùng file mẫu.'], warnings, data: null };
    cand.sort((a, b) => (norm(b.sh.name).includes('tong') - norm(a.sh.name).includes('tong')) || (b.h.score - a.h.score));
    const { sh, h } = cand[0]; const C = h.col;
    if (C.fee == null) warnings.push('Không thấy cột "Học phí áp dụng" — doanh thu và công nợ sẽ không tính được.');
    if (C.reg == null) warnings.push('Không thấy cột "Ngày đăng ký" — các báo cáo theo thời gian sẽ trống.');
    const payIdx = Object.keys(h.pays).map(Number).sort((a, b) => a - b).filter(k => h.pays[k].a != null);
    if (!payIdx.length && C.paidTotal == null) warnings.push('Không thấy cột thanh toán (Ngày TT 1 / Số tiền TT 1…) — thực thu = 0.');
    if (!payIdx.length && C.paidTotal != null) warnings.push('Chỉ có cột "Đã thanh toán" tổng: mỗi học viên được ghi 1 khoản, lấy ngày đăng ký làm ngày thu.');

    const students = [], seen = new Map(); let skipped = 0;
    for (let r = h.row + 1; r < sh.aoa.length; r++) {
      const row = sh.aoa[r] || [], g = f => (C[f] == null ? null : row[C[f]]);
      const name = txt(g('name')), rawId = txt(g('id'));
      if (!name && !rawId) { if (row.some(x => x != null && x !== '')) skipped++; continue; }
      const ref = sh.name + '!' + (r + 1);
      if (!name) { flags.push([ref, 'Thiếu họ tên (bỏ qua dòng)', rawId]); skipped++; continue; }
      const id = rawId || ('DONG' + (r + 1)); if (!rawId) flags.push([ref, 'Thiếu mã học viên (tự tạo mã)', id]);
      const dReg = toDate(g('reg')); if (dReg.bad) flags.push([id, 'Ngày đăng ký không đọc được', txt(g('reg'))]);
      const fee = toMoney(g('fee')); if (fee.bad) flags.push([id, 'Học phí không đọc được', txt(g('fee'))]);
      const st0 = STATUS[norm(g('st'))];
      if (txt(g('st')) && !st0) flags.push([id, 'Tình trạng không hợp lệ (sẽ tự tính)', txt(g('st'))]);
      const pays = [];
      payIdx.forEach(k => {
        const p = h.pays[k], a = toMoney(row[p.a]), d = p.d != null ? toDate(row[p.d]) : { v: null, bad: false };
        if (a.bad) flags.push([id, 'Số tiền thanh toán ' + k + ' không đọc được', txt(row[p.a])]);
        if (d.bad) flags.push([id, 'Ngày thanh toán ' + k + ' không đọc được', txt(row[p.d])]);
        if (a.v == null || a.v === 0) { if (d.v && a.v == null) flags.push([id, 'Có ngày thanh toán nhưng thiếu số tiền', 'TT ' + k]); return; }
        pays.push([d.v, a.v]);
      });
      if (!payIdx.length && C.paidTotal != null) { const t = toMoney(g('paidTotal')); if (t.v) pays.push([dReg.v, t.v]); }
      const paid = pays.reduce((x, p) => x + p[1], 0);
      let st = st0;
      if (!st) st = (fee.v == null || !dReg.v) ? 'Chưa đủ dữ liệu' : (fee.v - paid <= 0 ? 'Đã hoàn tất' : 'Chưa hoàn tất');
      const s = { id, name, reg: dReg.v, kg: toKg(g('kg')), cls: txt(g('cls')), course: txt(g('course')), fee: fee.v, promo: txt(g('promo')), tvv: txt(g('tvv')), src: txt(g('src')), pays, st, basis: sh.name, note: txt(g('note')) };
      // kiểm tra chất lượng
      if (!s.reg && !dReg.bad) flags.push([id, 'Thiếu ngày đăng ký', name]);
      if (!s.src) flags.push([id, 'Thiếu nguồn tuyển sinh', name]);
      if (!s.tvv) flags.push([id, 'Thiếu tư vấn viên', name]);
      if (!s.cls) flags.push([id, 'Chưa xếp lớp', name]);
      if (!s.course) flags.push([id, 'Thiếu khóa học', name]);
      if (s.fee == null && !fee.bad) flags.push([id, 'Thiếu học phí áp dụng', name]);
      if (s.reg && Math.abs(+s.reg.slice(0, 4) - +today.slice(0, 4)) > 1) flags.push([id, 'Năm đăng ký bất thường (giữ nguyên, cần xác nhận)', s.reg]);
      pays.forEach(p => {
        if (!p[0]) return;
        if (p[0] > today) flags.push([id, 'Ngày thanh toán nằm sau ngày hiện tại', p[0]]);
        if (s.reg) { const gap = (Date.parse(s.reg) - Date.parse(p[0])) / 864e5; if (gap > 3) flags.push([id, 'Thanh toán trước ngày đăng ký >3 ngày', 'ĐK ' + s.reg + ' / TT ' + p[0]]); }
      });
      if (s.st === 'Đã hoàn tất' && s.fee != null && s.fee - paid > 0) flags.push([id, 'Ghi "Đã hoàn tất" nhưng thanh toán < học phí', 'còn ' + (s.fee - paid)]);
      if (s.st === 'Chưa hoàn tất' && s.fee != null && s.fee - paid <= 0) flags.push([id, 'Ghi "Chưa hoàn tất" nhưng đã thu đủ', 'dư ' + (paid - s.fee)]);
      seen.set(id, (seen.get(id) || 0) + 1);
      students.push(s);
    }
    seen.forEach((n, id) => { if (n > 1) flags.push([id, 'Mã học viên trùng', 'xuất hiện ' + n + ' lần']); });
    if (!students.length) return { ok: false, errors: ['Sheet "' + sh.name + '" không có dòng học viên nào dưới dòng tiêu đề.'], warnings, data: null };
    if (skipped) warnings.push('Bỏ qua ' + skipped + ' dòng thiếu họ tên/mã.');

    // 2) sheet lớp (tùy chọn)
    let classes = null;
    for (const s2 of sheets) {
      if (s2 === sh) continue;
      const ch = findHeader(s2.aoa || [], CLASS_COLS, ['cls']);
      if (ch && ch.score >= 2) {
        classes = {};
        for (let r = ch.row + 1; r < s2.aoa.length; r++) {
          const row = s2.aoa[r] || [], g = f => (ch.col[f] == null ? '' : txt(row[ch.col[f]])), k = g('cls');
          if (k) classes[k] = { sched: g('sched'), shift: g('shift'), teacher: g('teacher'), open: g('open') };
        }
        break;
      }
    }
    const data = { students, classes: classes || {}, flags, today,
      meta: { version: 1, sheet: sh.name, sourceFile: opts.fileName || '', updatedAt: opts.now || new Date().toISOString() } };
    return { ok: true, errors, warnings, data, hasClassSheet: !!classes, skipped };
  }

  /** Tạo nội dung file mẫu (aoa) từ dữ liệu hiện tại. */
  function buildTemplate(data) {
    const S = data.students, mp = Math.max(6, ...S.map(s => s.pays.length));
    const head = ['Mã HV', 'Họ và tên', 'Ngày đăng ký', 'Tháng khai giảng', 'Mã lớp', 'Khóa học', 'Học phí áp dụng', 'Ưu đãi', 'Tư vấn viên', 'Nguồn tuyển sinh', 'Tình trạng', 'Ghi chú'];
    for (let k = 1; k <= mp; k++) head.push('Ngày TT ' + k, 'Số tiền TT ' + k);
    const nz = v => (v === '' || v == null ? null : v);
    const rows = S.map(s => {
      const r = [s.id, s.name, s.reg ? isoToSerial(s.reg) : null, s.kg, nz(s.cls), nz(s.course), s.fee, nz(s.promo), nz(s.tvv), nz(s.src), s.st, nz(s.note)];
      for (let k = 0; k < mp; k++) { const p = s.pays[k]; r.push(p && p[0] ? isoToSerial(p[0]) : null, p ? p[1] : null); }
      return r;
    });
    const dateCols = [2], moneyCols = [6];
    for (let k = 0; k < mp; k++) { dateCols.push(12 + 2 * k); moneyCols.push(13 + 2 * k); }
    const cls = [['Mã lớp', 'Lịch học', 'Ca học', 'Giáo viên', 'Khai giảng']].concat(Object.keys(data.classes || {}).map(k => { const c = data.classes[k]; return [k, c.sched || null, c.shift || null, c.teacher || null, c.open || null]; }));
    const guide = [
      ['HƯỚNG DẪN NHẬP FILE EXCEL CHO DASHBOARD BETMA ENGLISH'], [''],
      ['1. Sheet "TỔNG 2026": mỗi dòng là một học viên. Giữ nguyên dòng tiêu đề (tên cột không phân biệt hoa/thường và dấu).'],
      ['2. Cột bắt buộc: Họ và tên. Nên có: Mã HV, Ngày đăng ký, Mã lớp, Khóa học, Học phí áp dụng, Tư vấn viên, Nguồn tuyển sinh.'],
      ['3. Ngày nhập dạng dd/mm/yyyy hoặc kiểu Ngày của Excel. Tiền nhập là số (không cần ký hiệu ₫).'],
      ['4. Thanh toán: mỗi lần đóng một cặp cột "Ngày TT n" + "Số tiền TT n". Có thể thêm cột TT 7, TT 8… theo cùng quy tắc tên.'],
      ['5. Cột "Tình trạng" nhận: Đã hoàn tất / Chưa hoàn tất / Cần đối soát / Chưa đủ dữ liệu. Để trống thì dashboard tự tính (học phí − đã thu).'],
      ['6. Sheet "LỚP" (tùy chọn): Mã lớp, Lịch học, Ca học, Giáo viên, Khai giảng. Nếu bỏ sheet này, dashboard giữ thông tin lớp hiện có.'],
      ['7. File tải lên sẽ THAY THẾ toàn bộ dữ liệu hiện tại — luôn nhập đủ danh sách học viên, không chỉ phần mới.']
    ];
    return { sheets: [
      { name: 'TỔNG 2026', aoa: [head].concat(rows), dateCols, moneyCols, widths: [9, 28, 12, 10, 10, 14, 14, 24, 14, 20, 14, 28] },
      { name: 'LỚP', aoa: cls, dateCols: [], moneyCols: [], widths: [10, 10, 10, 16, 10] },
      { name: 'Hướng dẫn', aoa: guide, dateCols: [], moneyCols: [], widths: [120] }] };
  }

  return { parseSheets, buildTemplate, toDate, toMoney, norm, isoToSerial, serialToISO };
});
