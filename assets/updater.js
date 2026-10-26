/* BETMA ENGLISH — nạp dữ liệu từ GitHub, tải Excel mới, xem trước, xác nhận và đăng lên GitHub.
 * Bảo mật: KHÔNG có token nào trong source. Người có quyền ghi vào repo tự dán Fine-grained token khi đăng;
 * token chỉ nằm trong bộ nhớ của tab trình duyệt (hoặc sessionStorage nếu chọn "nhớ trong tab này"), chỉ gửi tới api.github.com. */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const cfg = () => window.BETMA_CONFIG || {};
  const App = () => window.BetmaApp;
  const E = s => App().esc(s);
  const SKEY = 'betma_gh_token';
  let pending = null;       // {data, fileName, res}
  let memToken = '';
  let form = { repo: '', branch: '' };   // giữ giá trị đã nhập khi phải nhập lại sau lỗi

  // ---------- thông báo trên đầu trang ----------
  function notice(html, kind) {
    const n = $('notice'); if (!n) return;
    if (!html) { n.hidden = true; n.innerHTML = ''; return; }
    n.className = 'notice ' + (kind || 'info'); n.innerHTML = html; n.hidden = false;
  }

  // ---------- nạp dữ liệu đã đăng ----------
  function loadScript(src) {
    return new Promise((ok, no) => { const s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = () => no(new Error('Không tải được ' + src)); document.head.appendChild(s); });
  }
  const valid = d => d && Array.isArray(d.students) && d.students.length && d.students.every(s => s && Array.isArray(s.pays));
  async function loadRemote() {
    const path = cfg().dataPath || 'data/data.json';
    try {
      const r = await fetch(path + (path.includes('?') ? '&' : '?') + 'v=' + Date.now(), { cache: 'no-store' });
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const d = await r.json(); if (!valid(d)) throw new Error('data.json không đúng định dạng');
      return { data: d, source: 'github' };
    } catch (e) {
      try { await loadScript('data/data.js'); if (valid(window.BETMA_DEFAULT_DATA)) return { data: window.BETMA_DEFAULT_DATA, source: 'fallback', err: e }; } catch (_) { /* bỏ qua */ }
      throw e;
    }
  }
  function show(res) {
    App().load(res.data);
    if (res.source === 'fallback') notice('Không đọc được <b>data/data.json</b> (' + E(res.err && res.err.message) + '). Đang hiển thị dữ liệu dự phòng đi kèm. Nếu mở bằng cách bấm đúp file index.html, hãy dùng GitHub Pages hoặc chạy máy chủ cục bộ để xem dữ liệu mới nhất.', 'warn');
    else notice('');
  }
  async function reloadRemote() { const r = await loadRemote(); show(r); return r; }

  // ---------- tiện ích ----------
  const sumPays = s => s.pays.reduce((x, p) => x + p[1], 0);
  function summarize(d) {
    const S = d.students, regs = S.map(s => s.reg).filter(Boolean).sort();
    return { n: S.length, fee: S.reduce((a, s) => a + (s.fee || 0), 0), paid: S.reduce((a, s) => a + sumPays(s), 0),
      debt: S.filter(s => s.st === 'Chưa hoàn tất').reduce((a, s) => a + ((s.fee || 0) - sumPays(s)), 0),
      nOwe: S.filter(s => s.st === 'Chưa hoàn tất').length, nRec: S.filter(s => s.st === 'Cần đối soát').length,
      last: regs[regs.length - 1] || null, nCls: new Set(S.map(s => s.cls).filter(Boolean)).size };
  }
  const FIELDS = ['id', 'name', 'reg', 'kg', 'cls', 'course', 'fee', 'promo', 'tvv', 'src', 'pays', 'st', 'basis', 'note'];
  function serialize(d) {   // mỗi học viên 1 dòng để Git so sánh dễ đọc
    const st = d.students.map(s => JSON.stringify(Object.fromEntries(FIELDS.map(k => [k, s[k]])))).join(',\n');
    return '{"meta":' + JSON.stringify(d.meta || {}) + ',"today":' + JSON.stringify(d.today) + ',"classes":' + JSON.stringify(d.classes || {}) +
      ',\n"students":[\n' + st + '\n],\n"flags":' + JSON.stringify(d.flags || []) + '}\n';
  }
  function b64(str) {
    const bytes = new TextEncoder().encode(str); let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  function detectRepo() {
    const c = cfg(); if (c.owner && c.repo) return { owner: c.owner, repo: c.repo, branch: c.branch || '' };
    const h = location.hostname;
    if (/\.github\.io$/i.test(h)) { const owner = h.split('.')[0], seg = location.pathname.split('/')[1]; return { owner, repo: seg && !/\.html?$/i.test(seg) ? seg : owner + '.github.io', branch: c.branch || '' }; }
    return { owner: '', repo: '', branch: c.branch || '' };
  }
  const needXlsx = () => { if (typeof XLSX === 'undefined') { alert('Chưa tải được thư viện đọc Excel (SheetJS). Kiểm tra kết nối mạng rồi tải lại trang.'); return true; } return false; };

  // ---------- file mẫu ----------
  function downloadTemplate() {
    if (needXlsx()) return;
    const t = BetmaParser.buildTemplate(App().data), wb = XLSX.utils.book_new();
    t.sheets.forEach(sh => {
      const ws = XLSX.utils.aoa_to_sheet(sh.aoa);
      for (let r = 1; r < sh.aoa.length; r++) {
        sh.dateCols.forEach(c => { const a = ws[XLSX.utils.encode_cell({ r, c })]; if (a) a.z = 'dd/mm/yyyy'; });
        sh.moneyCols.forEach(c => { const a = ws[XLSX.utils.encode_cell({ r, c })]; if (a) a.z = '#,##0'; });
      }
      ws['!cols'] = sh.widths.map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, ws, sh.name);
    });
    XLSX.writeFile(wb, 'BETMA_mau_du_lieu.xlsx');
  }

  // ---------- các bước trong hộp thoại ----------
  function frame(title, html) { $('mt').textContent = title; $('mb').innerHTML = html; $('md').classList.add('on'); }
  function stepPick(msg) {
    frame('Cập nhật dữ liệu từ Excel',
      '<div class="up-steps"><div><b>1.</b> Tải <a href="#" id="upTpl">file mẫu</a> (đã chứa dữ liệu hiện tại) hoặc dùng file Excel của bạn có cùng tên cột.</div>' +
      '<div><b>2.</b> Chọn file Excel đã chỉnh sửa — hệ thống đọc và cho xem trước kết quả.</div>' +
      '<div><b>3.</b> Xác nhận để cập nhật toàn bộ KPI, biểu đồ, bảng; chọn đăng lên GitHub để mọi người cùng thấy.</div></div>' +
      (msg ? '<div class="notice err">' + msg + '</div>' : '') +
      '<label class="up-file">Chọn file Excel (.xlsx, .xls)<input type="file" id="upFile" accept=".xlsx,.xls,.xlsm"></label>' +
      '<div class="na">File mới sẽ <b>thay thế toàn bộ</b> dữ liệu hiện tại, nên cần chứa đủ danh sách học viên.</div>');
    $('upTpl').onclick = e => { e.preventDefault(); downloadTemplate(); };
    $('upFile').onchange = e => { const f = e.target.files[0]; if (f) readFile(f); };
  }
  async function readFile(f) {
    if (needXlsx()) return;
    frame('Đang đọc file…', '<div class="na">Đang đọc ' + E(f.name) + '…</div>');
    try {
      const wb = XLSX.read(await f.arrayBuffer(), { type: 'array' });
      const sheets = wb.SheetNames.map(n => ({ name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }));
      const res = BetmaParser.parseSheets(sheets, { fileName: f.name });
      if (!res.ok) return stepPick('<b>Không đọc được file:</b> ' + res.errors.map(E).join(' '));
      if (!res.hasClassSheet) res.data.classes = App().data.classes || {};
      pending = { data: res.data, fileName: f.name, res };
      stepPreview();
    } catch (e) { stepPick('<b>Không đọc được file:</b> ' + E(e.message || e)); }
  }
  function stepPreview() {
    const { data, fileName, res } = pending, a = summarize(App().data), b = summarize(data), fm = App().fm, fn = App().fn, fd = App().fd;
    const dl = (x, y, money) => { const d = y - x; return d === 0 ? '—' : (d > 0 ? '+' : '−') + (money ? fm(Math.abs(d)) : fn(Math.abs(d))); };
    const row = (l, x, y, money, raw) => '<tr><td>' + l + '</td><td class="n">' + (raw ? x : money ? fm(x) : fn(x)) + '</td><td class="n">' + (raw ? y : money ? fm(y) : fn(y)) + '</td><td class="n">' + (raw ? '' : dl(x, y, money)) + '</td></tr>';
    const cnt = {}; data.flags.forEach(f => cnt[f[1]] = (cnt[f[1]] || 0) + 1);
    const top = Object.entries(cnt).sort((x, y) => y[1] - x[1]).slice(0, 6).map(([k, v]) => '<span class="tag t-g">' + E(k) + ': ' + v + '</span>').join(' ');
    frame('Xem trước dữ liệu mới',
      '<div class="na" style="padding-top:0">File: <b>' + E(fileName) + '</b> · sheet <b>' + E(data.meta.sheet) + '</b>' + (res.hasClassSheet ? ' · có sheet LỚP' : ' · không có sheet LỚP (giữ thông tin lớp hiện có)') + '</div>' +
      (res.warnings.length ? '<div class="notice warn">' + res.warnings.map(E).join('<br>') + '</div>' : '') +
      '<div class="tw"><table><thead><tr><th>Chỉ tiêu</th><th class="n">Hiện tại</th><th class="n">File mới</th><th class="n">Chênh lệch</th></tr></thead><tbody>' +
      row('Số học viên', a.n, b.n) + row('Tổng học phí đăng ký', a.fee, b.fee, 1) + row('Tổng thực thu lũy kế', a.paid, b.paid, 1) +
      row('Công nợ (HV chưa hoàn tất)', a.debt, b.debt, 1) + row('HV còn công nợ', a.nOwe, b.nOwe) + row('HV cần đối soát', a.nRec, b.nRec) + row('Số lớp', a.nCls, b.nCls) +
      row('Ngày đăng ký gần nhất', a.last ? fd(a.last) : '—', b.last ? fd(b.last) : '—', 0, 1) + '</tbody></table></div>' +
      '<div style="margin:10px 0 4px;font-size:12.5px;color:var(--mut)">Điểm cần lưu ý trong dữ liệu mới (' + data.flags.length + '):</div><div class="chips">' + (top || '<span class="na">Không có</span>') + '</div>' +
      '<div class="up-act"><button class="btn" id="upBack">‹ Chọn file khác</button><span class="sp"></span>' +
      '<button class="btn" id="upLocal">Chỉ xem trên máy này</button><button class="btn p" id="upPub">Xác nhận & đăng lên GitHub</button></div>' +
      '<div class="na">"Chỉ xem trên máy này": cập nhật dashboard ngay nhưng mất khi tải lại trang. "Đăng lên GitHub": lưu vĩnh viễn để mọi người truy cập đều thấy.</div>');
    $('upBack').onclick = () => stepPick();
    $('upLocal').onclick = () => { App().load(pending.data); notice('Đang xem dữ liệu <b>tạm thời</b> từ file <b>' + E(pending.fileName) + '</b> (chưa đăng lên GitHub). <a href="#" id="nbBack">Quay lại dữ liệu đã đăng</a>', 'warn'); const bk = $('nbBack'); if (bk) bk.onclick = async ev => { ev.preventDefault(); try { await reloadRemote(); } catch (e) { alert('Không tải được dữ liệu: ' + e.message); } }; };
    $('upPub').onclick = () => stepToken();
  }
  function stepToken(err) {
    const r = detectRepo(), remembered = sessionStorage.getItem(SKEY) || memToken || '';
    const repoVal = form.repo || (r.owner && r.repo ? r.owner + '/' + r.repo : ''), brVal = form.branch || r.branch;
    frame('Đăng dữ liệu lên GitHub',
      (err ? '<div class="notice err">' + err + '</div>' : '') +
      '<div class="na" style="padding-top:0">Dữ liệu sẽ được lưu vào <code>' + E(cfg().dataPath || 'data/data.json') + '</code> trong repo của bạn bằng một commit. Chỉ người có quyền ghi vào repo mới đăng được. <b>Token không nằm trong source code</b>; nó chỉ ở trong tab này và chỉ gửi tới api.github.com. <a href="#" id="upHelp">Cách tạo token</a></div>' +
      '<div class="up-form"><label>Repository (chủ/tên-repo)<input id="upRepo" value="' + E(repoVal) + '" placeholder="vd: betma-english/dashboard"></label>' +
      '<label>Nhánh (để trống = nhánh mặc định)<input id="upBranch" value="' + E(brVal) + '" placeholder="main"></label>' +
      '<label>GitHub token (Fine-grained, quyền Contents: Read and write)<input id="upTok" type="password" autocomplete="off" value="' + E(remembered) + '" placeholder="github_pat_…"></label>' +
      '<label class="up-chk"><input type="checkbox" id="upRem"' + (sessionStorage.getItem(SKEY) ? ' checked' : '') + '> Nhớ token trong tab này (xóa khi đóng tab)</label></div>' +
      '<div class="up-act"><button class="btn" id="upBack2">‹ Quay lại</button><span class="sp"></span><button class="btn p" id="upGo">Đăng lên GitHub</button></div>');
    $('upHelp').onclick = e => { e.preventDefault(); alert('Tạo token: GitHub → Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token.\n• Repository access: Only select repositories → chọn repo dashboard\n• Permissions → Repository permissions → Contents: Read and write\n• Expiration: nên đặt 30–90 ngày.\nSao chép token (bắt đầu bằng github_pat_) và dán vào ô bên dưới.'); };
    $('upBack2').onclick = stepPreview;
    $('upGo').onclick = async () => {
      const m = $('upRepo').value.trim().replace(/^https?:\/\/github\.com\//i, '').replace(/\.git$/i, '').replace(/\/$/, '').split('/'), tok = $('upTok').value.trim(), br = $('upBranch').value.trim();
      form = { repo: $('upRepo').value.trim(), branch: br };
      if (m.length < 2 || !m[0] || !m[1]) return stepToken('Nhập repo đúng dạng <b>chủ/tên-repo</b>.');
      if (!tok) return stepToken('Chưa nhập token.');
      memToken = tok; if ($('upRem').checked) sessionStorage.setItem(SKEY, tok); else sessionStorage.removeItem(SKEY);
      frame('Đang đăng lên GitHub…', '<div class="na">Đang ghi dữ liệu vào ' + E(m[0] + '/' + m[1]) + '…</div>');
      try { const info = await publish(pending.data, tok, { owner: m[0], repo: m[1], branch: br, path: cfg().dataPath || 'data/data.json' }); done(info); }
      catch (e) { stepToken(E(e.message || e)); }
    };
  }

  // ---------- đăng lên GitHub (REST API, Contents) ----------
  async function gh(url, tok, opt) {
    const r = await fetch(url, Object.assign({ cache: 'no-store' }, opt, { headers: Object.assign({ Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' }, (opt || {}).headers) }));
    return r;
  }
  const explain = (r, body) => r.status === 401 ? 'Token sai hoặc đã hết hạn (401).' :
    r.status === 403 ? 'Token không đủ quyền hoặc bị giới hạn (403). Cần quyền Contents: Read and write cho đúng repo.' :
    r.status === 404 ? 'Không thấy repo/nhánh, hoặc token chưa được cấp quyền cho repo này (404).' :
    'GitHub trả lỗi ' + r.status + (body && body.message ? ': ' + body.message : '') + '.';
  async function publish(data, tok, repo) {
    const branchIn = repo.branch || '';
    const api = 'https://api.github.com/repos/' + encodeURIComponent(repo.owner) + '/' + encodeURIComponent(repo.repo);
    let branch = branchIn;
    if (!branch) { const r = await gh(api, tok); const j = await r.json().catch(() => ({})); if (!r.ok) throw new Error(explain(r, j)); branch = j.default_branch || 'main'; }
    const url = api + '/contents/' + repo.path.split('/').map(encodeURIComponent).join('/');
    const content = b64(serialize(data));
    for (let attempt = 0; attempt < 2; attempt++) {
      let sha; const g = await gh(url + '?ref=' + encodeURIComponent(branch), tok);
      if (g.status === 200) sha = (await g.json()).sha; else if (g.status !== 404) throw new Error(explain(g, await g.json().catch(() => ({}))));
      const body = { message: 'Cập nhật dữ liệu dashboard từ ' + (pending && pending.fileName || 'Excel') + ' (' + data.students.length + ' học viên)', content, branch };
      if (sha) body.sha = sha;
      const p = await gh(url, tok, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      const j = await p.json().catch(() => ({}));
      if (p.ok) return { commit: j.commit, branch, repo: repo.owner + '/' + repo.repo };
      if ((p.status === 409 || p.status === 422) && attempt === 0) continue;   // ai đó vừa cập nhật → lấy lại sha và thử lại một lần
      throw new Error(explain(p, j));
    }
  }
  function done(info) {
    App().load(pending.data);
    const sha = info.commit && info.commit.sha ? info.commit.sha.slice(0, 7) : '', link = info.commit && info.commit.html_url;
    frame('Đã đăng lên GitHub', '<div class="notice ok">Đã lưu dữ liệu mới (' + pending.data.students.length + ' học viên) vào <b>' + E(info.repo) + '</b>, nhánh <b>' + E(info.branch) + '</b>' + (link ? ' — <a href="' + E(link) + '" target="_blank" rel="noopener">commit ' + E(sha) + '</a>' : '') + '.</div>' +
      '<div class="na">GitHub Pages cần khoảng 1–2 phút để phát hành bản mới; sau đó mọi người mở website đều thấy dữ liệu này (nếu vẫn thấy bản cũ, hãy tải lại trang bằng Ctrl+F5). Dashboard trên máy bạn đã hiển thị dữ liệu mới.</div>' +
      '<div class="up-act"><span class="sp"></span><button class="btn p" onclick="document.getElementById(\'md\').classList.remove(\'on\')">Xong</button></div>');
    notice('Đã đăng dữ liệu mới lên GitHub lúc ' + new Date().toLocaleTimeString('vi-VN') + '. Website sẽ cập nhật cho mọi người sau khoảng 1–2 phút.', 'ok');
    if (!sessionStorage.getItem(SKEY)) memToken = '';
    pending = null;
  }

  window.BetmaUpdater = { loadRemote, show, reloadRemote, open: () => stepPick(), serialize, publish, summarize };
})();
