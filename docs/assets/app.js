/* ================= Pallet Hub ================= */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const view = $('#view');
const S = { types: [], depts: [], stock: {}, dept: [], charts: [] };
const store = (fn, k, v) => { try { return fn === 'get' ? localStorage.getItem(k) : fn === 'del' ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch { return null; } };
const sessionGet = k => { try { return sessionStorage.getItem(k); } catch { return null; } };
const sessionSet = (k, v) => { try { sessionStorage.setItem(k, v); } catch {} };
/* เข้าสู่ระบบ: session token (จาก action login) เก็บใน localStorage — ไม่เก็บรหัสผ่าน
   ผู้ทำรายการในประวัติ/Log เซิร์ฟเวอร์บันทึกจากบัญชีที่เข้าสู่ระบบเอง (ไม่ได้มาจากหน้าเว็บ) */
const TOKEN_KEY = 'palletToken';
let authToken = store('get', TOKEN_KEY) || '';
S.me = null;

const STATUS = {
  available: { name: 'พร้อมใช้', icon: 'fa-circle-check', c: '#12b76a' },
  issued:    { name: 'เบิกไปใช้งาน', icon: 'fa-dolly', c: '#0ea5e9' },
  damaged:   { name: 'ชำรุด', icon: 'fa-triangle-exclamation', c: '#f04438' },
  repairing: { name: 'กำลังซ่อม', icon: 'fa-screwdriver-wrench', c: '#7a5af8' },
  scrapped:  { name: 'ตัดจำหน่าย', icon: 'fa-trash-can', c: '#98a2b3' },
};
const ACTIONS = {
  receive:      { name: 'รับเข้า', icon: 'fa-truck-ramp-box' },
  issue:        { name: 'เบิกจ่าย', icon: 'fa-dolly' },
  return:       { name: 'รับคืน', icon: 'fa-rotate-left' },
  damage:       { name: 'แจ้งชำรุด', icon: 'fa-triangle-exclamation' },
  repair_start: { name: 'ส่งซ่อม', icon: 'fa-screwdriver-wrench' },
  repair_done:  { name: 'ซ่อมเสร็จ', icon: 'fa-circle-check' },
  scrap:        { name: 'ตัดจำหน่าย', icon: 'fa-trash-can' },
};
const PAGES = {
  dashboard: ['แดชบอร์ด', 'ภาพรวมพาเลททั้งหมดแบบเรียลไทม์'],
  receive: ['รับเข้าพาเลท', 'เพิ่มพาเลทใหม่เข้าคลัง → สถานะ “พร้อมใช้”'],
  issue: ['เบิกจ่ายพาเลท', 'เลือกฝ่ายที่เบิก พร้อมวันที่และเวลา'],
  return: ['รับคืนพาเลท', 'รับพาเลทคืนจากฝ่าย — สภาพดี หรือ ชำรุด'],
  damage: ['แจ้งชำรุด', 'พบพาเลทเสียในคลัง → เปิดใบแจ้งซ่อม'],
  repair: ['งานซ่อมพาเลท', 'ติดตามสถานะ ชำรุด → กำลังซ่อม → ซ่อมเสร็จ'],
  stock: ['คงเหลือ', 'ยอดคงเหลือแยกตามประเภท ขนาด และสถานะ'],
  history: ['ประวัติเคลื่อนไหว', 'ทุกการเคลื่อนไหวของพาเลท ค้นหา / ส่งออก Excel'],
  logs: ['บันทึกประวัติ (Log)', 'ทุกการกระทำในระบบ — ใคร ทำอะไร เมื่อไร'],
  settings: ['ตั้งค่าฝ่าย', 'จัดการรายชื่อฝ่ายที่เบิกจ่ายพาเลท'],
  account: ['ตั้งค่าบัญชี', 'บัญชีผู้ใช้ รหัสผ่าน และสิทธิ์การใช้งาน'],
};

/* ---------- API (Google Apps Script web app) ---------- */
// อ่านข้อมูล = GET ?action=...&token=...  /  บันทึก = POST text/plain (JSON) — ไม่เกิด CORS preflight
// ทุกคำสั่งต้องมี session token (เข้าสู่ระบบ) — เซิร์ฟเวอร์ตอบ code 'AUTH' เมื่อยังไม่เข้าสู่ระบบ/หมดอายุ
const API_URL = String((window.PALLET_CONFIG || {}).apiUrl || '').trim();
const NO_API_MSG = 'ยังไม่ได้ตั้งค่า apiUrl — ใส่ URL ของ Google Apps Script (/exec) ในไฟล์ docs/config.js';
async function gasFetch(url, opt) {
  let j;
  try {
    const res = await fetch(url, { mode: 'cors', credentials: 'omit', cache: 'no-store', ...opt });
    j = await res.json();
  } catch { throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่ได้ (ตรวจสอบอินเทอร์เน็ต / Google Apps Script)'); }
  if (!j || !j.ok) {
    const e = new Error((j && j.error) || 'เกิดข้อผิดพลาด');
    e.passwordError = !!(j && j.passwordError);
    e.code = (j && j.code) || '';
    if (e.code === 'AUTH') authLost(); // token หมดอายุ / ถูกปิดบัญชี / เปลี่ยนรหัสผ่าน → กลับหน้าเข้าสู่ระบบ
    throw e;
  }
  return { ok: true, ...(j.data || {}) };
}
function gasGet(action, params = '') {
  if (!API_URL) return Promise.reject(new Error(NO_API_MSG));
  const u = new URL(API_URL, location.href);
  new URLSearchParams(params.replace(/^&/, '')).forEach((v, k) => u.searchParams.set(k, v));
  u.searchParams.set('action', action);
  u.searchParams.set('token', authToken);
  u.searchParams.set('_', Date.now());
  return gasFetch(u.toString(), { method: 'GET' });
}
function gasPost(payload) {
  if (!API_URL) return Promise.reject(new Error(NO_API_MSG));
  return gasFetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=UTF-8' }, body: JSON.stringify({ ...payload, token: authToken }) });
}
async function api(action, body, params = '') {
  if (!body) return gasGet(action, params);
  return gasPost({ ...body, action });
}
/* แก้ไข / ลบรายการที่บันทึกแล้ว: ใช้รหัสรีเซ็ตข้อมูล (PALLET_RESET_PASSWORD) เพิ่มจากการเข้าสู่ระบบ
   เก็บในหน่วยความจำเท่านั้นหลังตรวจผ่าน (หายเมื่อรีเฟรช/ปิดหน้า/ออกจากระบบ) และล้างทิ้งเมื่อเซิร์ฟเวอร์แจ้งรหัสผิด/ถูกล็อก */
let recordPassword = '';
async function recordApi(action, body) {
  const pw = await askPassword('record');
  try {
    return await gasPost({ ...body, action, resetPassword: pw });
  } catch (e) {
    if (e.passwordError) recordPassword = '';
    throw e;
  }
}
const PW_KIND = {
  record: {
    title: 'รหัสรีเซ็ตข้อมูล (แก้ไข / ลบรายการ)', label: 'รหัสรีเซ็ตข้อมูล', icon: 'fa-user-shield',
    desc: 'การแก้ไขหรือลบรายการที่บันทึกแล้วต้องใช้รหัสรีเซ็ตข้อมูล (ตั้งโดยผู้ดูแลระบบ — คนละรหัสกับรหัสเข้าสู่ระบบ · ถามครั้งเดียวต่อการเปิดหน้านี้)',
    cancel: 'ยกเลิก — ต้องใส่รหัสรีเซ็ตข้อมูลก่อนแก้ไข/ลบรายการ',
    get: () => recordPassword, set: v => { recordPassword = v; },
    verify: pw => gasPost({ action: 'verifyResetPassword', resetPassword: pw }),
  },
};
/* ถามรหัสครั้งเดียวต่อการเปิดหน้า แล้วตรวจกับเซิร์ฟเวอร์ (verifyResetPassword) */
function askPassword(kind = 'record') {
  const K = PW_KIND[kind];
  if (K.get()) return Promise.resolve(K.get());
  S.pwPending ||= {};
  if (S.pwPending[kind]) return S.pwPending[kind];
  if (!$('#pwModal').hidden) return Promise.reject(new Error('กรุณาใส่รหัสในหน้าต่างที่เปิดอยู่ก่อน'));
  S.pwPending[kind] = new Promise((resolve, reject) => {
    const box = $('#pwModal'), card = $('#pwCard');
    card.innerHTML = `<h3><i class="fa-solid ${K.icon}" style="color:var(--brand)"></i> ${K.title}</h3>
      <p style="color:var(--muted)">${K.desc}</p>
      <div class="field mt"><label>${K.label}</label><input type="password" id="pwInput" autocomplete="off"></div>
      <div id="pwErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
      <div class="modal-acts"><button class="btn btn-ghost" id="pwCancel">ยกเลิก</button><button class="btn btn-primary" id="pwOk"><i class="fa-solid fa-unlock"></i>ยืนยัน</button></div>`;
    box.hidden = false;
    const inp = $('#pwInput', card), ok = $('#pwOk', card), err = $('#pwErr', card);
    const done = () => { box.hidden = true; box.onclick = null; card.innerHTML = ''; };
    const cancel = () => { done(); reject(new Error(K.cancel)); };
    S.pwCancel = cancel;
    const submit = async () => {
      const pw = inp.value;
      if (!pw) { err.textContent = 'กรุณาใส่' + K.label; inp.focus(); return; }
      ok.disabled = true; ok.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> กำลังตรวจสอบ...'; err.textContent = '';
      try {
        await K.verify(pw);
        K.set(pw); done(); resolve(pw);
      } catch (e) {
        if (e.code === 'AUTH') { done(); reject(e); return; }
        err.textContent = e.message; inp.value = ''; inp.focus();
        ok.disabled = false; ok.innerHTML = '<i class="fa-solid fa-unlock"></i>ยืนยัน';
      }
    };
    ok.onclick = submit;
    $('#pwCancel', card).onclick = cancel;
    box.onclick = e => { if (e.target === box) cancel(); };
    inp.onkeydown = e => { e.stopPropagation(); if (e.key === 'Enter') submit(); if (e.key === 'Escape') cancel(); };
    inp.focus();
  }).finally(() => { S.pwPending[kind] = null; S.pwCancel = null; });
  return S.pwPending[kind];
}
/* ส่งออก Excel (CSV) — เซิร์ฟเวอร์สร้าง CSV แล้วดาวน์โหลดเป็นไฟล์ในเบราว์เซอร์ */
async function downloadCsv(action, params) {
  try {
    const r = await gasGet(action, params);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob(['\uFEFF' + r.csv], { type: 'text/csv;charset=utf-8' }));
    a.download = r.filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  } catch (e) { toast(e.message, 'err'); }
}
if (!API_URL) $('#cfgBanner').hidden = false;
async function boot() {
  const j = await api('bootstrap');
  Object.assign(S, { types: j.types, depts: j.departments, stock: j.stock, dept: j.dept, now: j.now });
}

/* ---------- helpers ---------- */
const fmt = n => Number(n || 0).toLocaleString('th-TH');
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const typeById = id => S.types.find(t => +t.id === +id);
const sizes = t => t.sizes.split(',');
const q = (t, size, st) => (S.stock[t]?.[size]?.[st]) || 0;
const qType = (t, st) => sizes(typeById(t)).reduce((a, s) => a + q(t, s, st), 0);
const qAll = st => S.types.reduce((a, t) => a + qType(t.id, st), 0);
const deptQty = (d, t, s) => +(S.dept.find(r => r.department === d && +r.type_id === +t && r.size === s)?.qty || 0);
const dtTH = s => { const d = new Date(s.replace(' ', 'T')); return d.toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: '2-digit' }) + ' ' + s.slice(11, 16) + ' น.'; };
const nowParts = () => { const d = new Date(); const p = n => String(n).padStart(2, '0'); return { date: `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`, time: `${p(d.getHours())}:${p(d.getMinutes())}` }; };

/* ---------- isometric pallet SVG ---------- */
function shade(hex, f) {
  const n = parseInt(hex.slice(1), 16);
  const ch = [n >> 16, (n >> 8) & 255, n & 255].map(v => Math.max(0, Math.min(255, Math.round(f < 0 ? v * (1 + f) : v + (255 - v) * f))));
  return '#' + ch.map(v => v.toString(16).padStart(2, '0')).join('');
}
function palletSVG(color, cls = 'pal') {
  const P = (x, y, z) => `${((x - y) * 0.866).toFixed(1)},${((x + y) * 0.5 - z).toFixed(1)}`;
  const box = (x, y, z, w, d, h) =>
    `<polygon fill="${shade(color, -0.32)}" points="${P(x, y + d, z)} ${P(x + w, y + d, z)} ${P(x + w, y + d, z + h)} ${P(x, y + d, z + h)}"/>` +
    `<polygon fill="${shade(color, -0.16)}" points="${P(x + w, y, z)} ${P(x + w, y + d, z)} ${P(x + w, y + d, z + h)} ${P(x + w, y, z + h)}"/>` +
    `<polygon fill="${shade(color, 0.12)}" points="${P(x, y, z + h)} ${P(x + w, y, z + h)} ${P(x + w, y + d, z + h)} ${P(x, y + d, z + h)}"/>`;
  let s = '';
  [0, 42, 84].forEach(y => s += box(0, y, 0, 100, 16, 4));
  const blocks = [];
  [0, 42, 84].forEach(x => [0, 42, 84].forEach(y => blocks.push([x, y])));
  blocks.sort((a, b) => a[0] + a[1] - b[0] - b[1]).forEach(([x, y]) => s += box(x, y, 4, 16, 16, 10));
  [0, 21.5, 43, 64.5, 86].forEach(y => s += box(0, y, 14, 100, 14, 4));
  return `<svg class="${cls}" viewBox="-92 -24 184 130" aria-hidden="true">${s}</svg>`;
}

/* ---------- UI effects ---------- */
document.addEventListener('pointerdown', e => {
  const b = e.target.closest('.btn, .chip, .dept, .ptype, .quick button');
  if (!b) return;
  const r = b.getBoundingClientRect(), d = Math.max(r.width, r.height);
  const sp = document.createElement('span');
  sp.className = 'ripple';
  Object.assign(sp.style, { width: d + 'px', height: d + 'px', left: e.clientX - r.left - d / 2 + 'px', top: e.clientY - r.top - d / 2 + 'px' });
  if (getComputedStyle(b).position === 'static') b.style.position = 'relative';
  b.style.overflow = 'hidden';
  b.appendChild(sp);
  setTimeout(() => sp.remove(), 700);
});

function toast(msg, type = 'ok') {
  const m = { ok: ['fa-circle-check', '#12b76a'], err: ['fa-circle-xmark', '#f04438'], info: ['fa-circle-info', '#E2231A'] }[type];
  const t = document.createElement('div');
  t.className = 'toast';
  t.style.setProperty('--c', m[1]);
  t.innerHTML = `<i class="fa-solid ${m[0]} fa-bounce" style="--fa-animation-iteration-count:1"></i><div>${esc(msg)}</div>`;
  $('#toasts').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 400); }, 3500);
}

function confetti() {
  const cv = $('#confetti'), cx = cv.getContext('2d');
  cv.width = innerWidth; cv.height = innerHeight;
  const cols = ['#E2231A', '#B5121B', '#F5B400', '#1FA83A', '#1E6FE0', '#ff7a6b', '#1d1d1f'];
  const ps = Array.from({ length: 160 }, () => ({
    x: innerWidth / 2 + (Math.random() - .5) * 200, y: innerHeight / 2.2,
    vx: (Math.random() - .5) * 18, vy: -Math.random() * 18 - 6, s: Math.random() * 8 + 5,
    r: Math.random() * 6, vr: (Math.random() - .5) * .4, c: cols[Math.random() * cols.length | 0],
  }));
  let f = 0;
  const run = S.confettiRun = (S.confettiRun || 0) + 1;
  const stop = () => cx.clearRect(0, 0, cv.width, cv.height);
  setTimeout(() => { if (S.confettiRun === run) { S.confettiRun++; stop(); } }, 2600); // กันค้างเมื่อแท็บถูกพัก
  (function loop() {
    if (S.confettiRun !== run) return;
    cx.clearRect(0, 0, cv.width, cv.height);
    ps.forEach(p => {
      p.vy += .45; p.vx *= .99; p.x += p.vx; p.y += p.vy; p.r += p.vr;
      cx.save(); cx.translate(p.x, p.y); cx.rotate(p.r); cx.fillStyle = p.c;
      cx.globalAlpha = Math.max(0, 1 - f / 140); cx.fillRect(-p.s / 2, -p.s / 4, p.s, p.s / 2); cx.restore();
    });
    if (++f < 140) requestAnimationFrame(loop); else cx.clearRect(0, 0, cv.width, cv.height);
  })();
}

function modal(html) {
  $('#modalCard').innerHTML = html;
  $('#modal').hidden = false;
  return $('#modalCard');
}
function closeModal() {
  if ($('#modal').hidden) return;
  $('#modal').hidden = true;
  const fn = S.onModalClose; S.onModalClose = null;
  if (fn) fn();
}
$('#modal').addEventListener('click', e => { if (e.target.id === 'modal' || e.target.closest('[data-close]')) closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

function countUp(root = view) {
  $$('[data-count]', root).forEach(el => {
    const to = +el.dataset.count, t0 = performance.now(), dur = 1100;
    const step = t => {
      const p = Math.min(1, (t - t0) / dur), e = 1 - Math.pow(1 - p, 4);
      el.textContent = fmt(Math.round(to * e));
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  });
}
function tilt(root = view) {
  $$('.tcard', root).forEach(c => {
    c.addEventListener('mousemove', e => {
      const r = c.getBoundingClientRect(), x = (e.clientX - r.left) / r.width - .5, y = (e.clientY - r.top) / r.height - .5;
      c.style.transform = `perspective(700px) rotateY(${x * 10}deg) rotateX(${-y * 10}deg) translateY(-6px)`;
    });
    c.addEventListener('mouseleave', () => c.style.transform = '');
  });
}
function killCharts() { S.charts.forEach(c => c.destroy()); S.charts = []; }

/* clock */
setInterval(tickClock, 1000); tickClock();
function tickClock() {
  const d = new Date();
  $('#clock').textContent = d.toLocaleTimeString('th-TH', { hour12: false });
  $('#clockDate').textContent = d.toLocaleDateString('th-TH', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
}
$('#menuBtn').onclick = () => $('#sidebar').classList.toggle('open');
$('#navOverlay').onclick = () => $('#sidebar').classList.remove('open');

/* stagger: release the animation once finished so hover / tilt transforms work */
document.addEventListener('animationend', e => {
  if (e.animationName === 'rise' && e.target.parentElement?.classList.contains('stagger')) {
    e.target.style.animation = 'none'; e.target.style.opacity = '1';
  }
});

/* ================= LOGIN / SESSION ================= */
const ROLE = { admin: ['ผู้ดูแลระบบ', '#E2231A', 'fa-user-shield'], user: ['ผู้ใช้งาน', '#0ea5e9', 'fa-user'] };
const roleTag = r => { const x = ROLE[r] || [r, '#667085', 'fa-user']; return `<span class="tag" style="--c:${x[1]}"><i class="fa-solid ${x[2]}"></i>${esc(x[0])}</span>`; };
const isAdmin = () => S.me?.role === 'admin';
function setMe(u) { S.me = u || null; paintUser(); }
function setToken(t) { authToken = t || ''; if (authToken) store('set', TOKEN_KEY, authToken); else store('del', TOKEN_KEY); }
function paintUser() {
  const u = S.me;
  $('#userName').textContent = u ? u.fullname || u.username : '—';
  $('#userAv').innerHTML = u ? esc((u.fullname || u.username).trim().charAt(0).toUpperCase()) : '<i class="fa-solid fa-user"></i>';
  $('#umName').textContent = u ? u.fullname || u.username : '';
  $('#umInfo').textContent = u ? `${u.username} · ${(ROLE[u.role] || [u.role])[0]}` : '';
  $('#userChip').title = u ? `เข้าสู่ระบบเป็น ${u.fullname} (${u.username})` : 'บัญชีผู้ใช้';
}
function showLogin(msg = '') {
  closeUserMenu();
  closeModal();
  S.pwCancel?.();
  const box = $('#loginScreen');
  box.classList.remove('checking');
  box.hidden = false;
  $('#lgErr').textContent = msg;
  $('#lgPass').value = '';
  const btn = $('#lgBtn'); btn.disabled = false; btn.innerHTML = '<i class="fa-solid fa-right-to-bracket"></i>เข้าสู่ระบบ';
  setTimeout(() => ($('#lgUser').value ? $('#lgPass') : $('#lgUser')).focus(), 50);
}
function hideLogin() { $('#loginScreen').hidden = true; $('#lgPass').value = ''; $('#lgErr').textContent = ''; }
/* เซิร์ฟเวอร์ตอบ AUTH (ยังไม่เข้าสู่ระบบ / เซสชันหมดอายุ / บัญชีถูกปิด / รหัสผ่านถูกเปลี่ยน) */
function authLost() {
  const had = !!authToken;
  setToken(''); setMe(null); recordPassword = '';
  showLogin(had ? 'เซสชันหมดอายุหรือถูกยกเลิก กรุณาเข้าสู่ระบบใหม่' : '');
}
$('#loginForm').onsubmit = async e => {
  e.preventDefault();
  const user = $('#lgUser').value.trim(), pass = $('#lgPass').value, btn = $('#lgBtn'), err = $('#lgErr');
  if (!user || !pass) { err.textContent = 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน'; return; }
  btn.disabled = true; btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> กำลังเข้าสู่ระบบ...'; err.textContent = '';
  try {
    const r = await gasPost({ action: 'login', username: user, password: pass });
    setToken(r.token); setMe(r.user); hideLogin();
    toast(`สวัสดีคุณ${r.user.fullname || r.user.username}`, 'info');
    if (r.initial) { toast('สร้างบัญชีผู้ดูแลระบบแล้ว — กรุณาเปลี่ยนรหัสผ่านที่ “ตั้งค่าบัญชี”', 'info'); location.hash = 'account'; }
    route();
  } catch (x) {
    err.textContent = x.message; $('#lgPass').value = ''; $('#lgPass').focus();
    btn.disabled = false; btn.innerHTML = '<i class="fa-solid fa-right-to-bracket"></i>เข้าสู่ระบบ';
  }
};
async function logout() {
  closeUserMenu();
  try { await gasPost({ action: 'logout' }); } catch {}
  setToken(''); setMe(null); recordPassword = '';
  view.innerHTML = '';
  showLogin('ออกจากระบบแล้ว');
}
function closeUserMenu() { $('#userMenu').hidden = true; $('#userChip').setAttribute('aria-expanded', 'false'); }
$('#userChip').onclick = e => {
  e.stopPropagation();
  const m = $('#userMenu');
  m.hidden = !m.hidden;
  $('#userChip').setAttribute('aria-expanded', String(!m.hidden));
};
$('#umAccount').onclick = () => { closeUserMenu(); location.hash = 'account'; };
$('#umLogout').onclick = logout;
document.addEventListener('click', e => { if (!e.target.closest('.user-wrap')) closeUserMenu(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeUserMenu(); });
/* เริ่มต้น: มี token → ตรวจกับเซิร์ฟเวอร์ (me) / ไม่มี → แสดงหน้าเข้าสู่ระบบ */
async function start() {
  if (!API_URL) return showLogin(NO_API_MSG);
  if (!authToken) return showLogin();
  try {
    const r = await gasGet('me');
    setMe(r.user); hideLogin(); route();
  } catch (e) {
    if (e.code !== 'AUTH') showLogin(e.message);
  }
}

/* ================= ROUTER ================= */
async function route() {
  if (!authToken || !S.me) return; // ยังไม่เข้าสู่ระบบ: หน้าเข้าสู่ระบบแสดงอยู่
  const page = (location.hash.slice(1) || 'dashboard').split('?')[0];
  const p = PAGES[page] ? page : 'dashboard';
  $$('.nav a').forEach(a => a.classList.toggle('active', a.dataset.page === p));
  $('#pageTitle').textContent = PAGES[p][0];
  $('#pageSub').textContent = PAGES[p][1];
  $('#sidebar').classList.remove('open');
  killCharts(); view.onclick = null;
  view.innerHTML = '<div class="spinner"></div>';
  view.style.animation = 'none'; view.offsetHeight; view.style.animation = '';
  try {
    await boot();
    updateBadge();
    await ({ dashboard, logs, receive: () => txForm('receive'), issue: () => txForm('issue'), return: () => txForm('return'), damage: () => txForm('damage'), repair, stock, history, settings, account })[p]();
  } catch (e) {
    if (e.code === 'AUTH') { view.innerHTML = ''; return; }
    view.innerHTML = `<div class="card empty"><i class="fa-solid fa-plug-circle-xmark"></i>${esc(e.message)}</div>`;
  }
}
addEventListener('hashchange', route);
async function updateBadge() {
  const n = qAll('damaged') + qAll('repairing');
  $('#repairBadge').textContent = n ? fmt(n) : '';
}

/* ================= DASHBOARD ================= */
async function dashboard() {
  const range = +(sessionGet('dashDays') || 7);
  const d = await api('dashboard', null, '&days=' + range);
  S.stock = d.stock; S.dept = d.dept;
  const sts = ['available', 'issued', 'damaged', 'repairing'];
  const total = sts.reduce((a, s) => a + qAll(s), 0);
  const today = Object.fromEntries(d.today.map(r => [r.action, +r.q]));
  const pct = s => total ? Math.round(qAll(s) / total * 100) : 0;
  const h = new Date().getHours();
  const greet = h < 12 ? ['สวัสดีตอนเช้า', 'fa-sun'] : h < 17 ? ['สวัสดีตอนบ่าย', 'fa-cloud-sun'] : ['สวัสดีตอนเย็น', 'fa-moon'];
  const rs = d.repairStat || {};
  const oldestH = rs.oldest ? Math.max(0, Math.round((Date.now() - new Date(rs.oldest.replace(' ', 'T'))) / 36e5)) : 0;
  const kpis = [
    ['พาเลทในระบบทั้งหมด', total, 'fa-pallet', '#1d1d1f', 'ไม่รวมตัดจำหน่าย'],
    ['พร้อมใช้งาน', qAll('available'), 'fa-circle-check', STATUS.available.c, `${pct('available')}% ของทั้งหมด`],
    ['เบิกไปใช้งาน', qAll('issued'), 'fa-dolly', STATUS.issued.c, `${pct('issued')}% · อยู่กับ ${new Set(S.dept.map(r => r.department)).size} ฝ่าย`],
    ['ชำรุด / รอซ่อม', qAll('damaged'), 'fa-triangle-exclamation', STATUS.damaged.c, `แจ้งวันนี้ ${fmt(today.damage)} ตัว`],
    ['กำลังซ่อม', qAll('repairing'), 'fa-screwdriver-wrench', STATUS.repairing.c, `ซ่อมเสร็จวันนี้ ${fmt(today.repair_done)} ตัว`],
  ];
  const fnode = (s, extra = '') => `<div class="fnode" style="--c:${STATUS[s].c}" data-go="${s === 'issued' ? 'stock' : s === 'available' ? 'issue' : 'repair'}">
      <em>${pct(s)}%</em><div class="fi"><i class="fa-solid ${STATUS[s].icon}"></i></div><b data-count="${qAll(s)}">0</b><span>${STATUS[s].name}${extra}</span></div>`;
  const conn = (a, label, label2, b) => `<div class="fconn"><small>${label}</small><div class="ln" style="--a:${a}"></div>${label2 ? `<div class="ln rev" style="--a:${b}"></div><small>${label2}</small>` : ''}</div>`;

  view.innerHTML = `
  <div class="reset-topbar"><button class="btn btn-bad" id="resetBtn">🗑 รีเซ็ตข้อมูล / Reset data</button></div>
  <div class="hero">
    <div class="leafs">${[12, 30, 48, 66, 84].map((l, i) => `<i class="fa-solid fa-leaf" style="left:${l}%;bottom:-20px;font-size:${18 + i * 6}px;animation-delay:${-i * 2.6}s"></i>`).join('')}</div>
    <div class="greet"><i class="fa-solid ${greet[1]}"></i>${greet[0]}${S.me ? ' คุณ' + esc(S.me.fullname || S.me.username) : ''} · ${new Date().toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}</div>
    <h2><span class="t-rm">RM</span> + <span class="t-pk">PK</span> + <span class="t-fg">FG</span> ระบบบริหารพาเลท</h2>
    <p>รับเข้า · เบิกจ่าย · ซ่อม · คงเหลือ</p>
    <div class="hero-stats">
      ${[['receive', 'รับเข้าวันนี้', 'fa-truck-ramp-box'], ['issue', 'เบิกจ่ายวันนี้', 'fa-dolly'], ['return', 'รับคืนวันนี้', 'fa-rotate-left'], ['damage', 'แจ้งชำรุดวันนี้', 'fa-heart-crack']]
        .map(([k, l, ic]) => `<div class="hs" data-go="${k === 'damage' ? 'repair' : k}"><small><i class="fa-solid ${ic}"></i>${l}</small><b data-count="${today[k] || 0}">0</b></div>`).join('')}
    </div>
    <div class="hero-scene">
      <div class="sparkles"><i class="fa-solid fa-star" style="top:18%;left:20%"></i><i class="fa-solid fa-star" style="top:40%;left:70%;animation-delay:-1s;font-size:10px"></i><i class="fa-solid fa-star" style="top:12%;left:88%;animation-delay:-2s;font-size:12px"></i></div>
      <div class="stack"><div class="box"></div><div class="box"></div><div class="box"></div>${palletSVG('#F5B400', 'pal-s')}</div>
      <div class="forklift">${forkliftSVG()}</div>
      <div class="road"></div>
    </div>
  </div>

  <div class="grid g-kpi mt stagger">
    ${kpis.map(([l, n, ic, c, tr]) => `
      <div class="card kpi" style="--c:${c}">
        <div class="ic"><i class="fa-solid ${ic}"></i></div>
        <div class="kpi-num" data-count="${n}">0</div>
        <div class="kpi-label">${l}</div><div class="trend">${tr}</div>
      </div>`).join('')}
  </div>

  <div class="sec-h"><i class="fa-solid fa-route" style="color:var(--brand)"></i>เส้นทางพาเลท <small>พาเลทแต่ละตัวอยู่กระบวนการไหน</small></div>
  <div class="card">
    <div class="flowmap">
      ${fnode('available')}${conn('#0ea5e9', 'เบิกจ่าย →', '← รับคืน', '#12b76a')}${fnode('issued')}${conn('#f04438', 'ชำรุด →')}${fnode('damaged')}${conn('#7a5af8', 'ส่งซ่อม →')}${fnode('repairing')}
    </div>
    <div class="flow-return"><i class="fa-solid fa-arrow-turn-up" style="color:#12b76a"></i><div class="lp"></div><span><i class="fa-solid fa-circle-check" style="color:#12b76a"></i> ซ่อมเสร็จ → กลับมาพร้อมใช้ (เดือนนี้ ${fmt(rs.done_m)} ตัว)</span></div>
  </div>

  <div class="sec-h"><i class="fa-solid fa-layer-group" style="color:var(--brand)"></i>แยกตามประเภทพาเลท <small>คลิกเพื่อดูคงเหลือ</small></div>
  <div class="grid g-types stagger">${S.types.map(typeCard).join('')}</div>

  <div class="sec-h"><i class="fa-solid fa-chart-column" style="color:var(--brand)"></i>วิเคราะห์ข้อมูล</div>
  <div class="grid g-2">
    <div class="card"><h3><i class="fa-solid fa-chart-line"></i>ความเคลื่อนไหว
      <div class="seg" id="seg">${[7, 14, 30].map(n => `<button data-days="${n}" class="${n === range ? 'on' : ''}">${n} วัน</button>`).join('')}</div></h3>
      <div class="chart-box"><canvas id="cLine"></canvas></div></div>
    <div class="card"><h3><i class="fa-solid fa-gauge-simple-high"></i>ตัวชี้วัด</h3>
      <div class="gauges">
        ${gauge('อัตราความพร้อมใช้', 'พร้อมใช้ ÷ ทั้งหมด', pct('available'), '#12b76a')}
        ${gauge('อัตราการใช้งาน', 'เบิกไปใช้ ÷ ทั้งหมด', pct('issued'), '#0ea5e9')}
        ${gauge('อัตราชำรุด', '(ชำรุด + ซ่อม) ÷ ทั้งหมด', total ? Math.round((qAll('damaged') + qAll('repairing')) / total * 100) : 0, '#E2231A')}
      </div></div>
  </div>
  <div class="grid g-3b mt">
    <div class="card"><h3><i class="fa-solid fa-chart-pie"></i>สถานะพาเลท</h3><div class="chart-box sm"><canvas id="cDonut"></canvas></div></div>
    <div class="card"><h3><i class="fa-solid fa-clock"></i>กิจกรรมรายชั่วโมงวันนี้<span class="sub">ตัว</span></h3><div class="chart-box sm"><canvas id="cHour"></canvas></div></div>
    <div class="card"><h3><i class="fa-solid fa-trophy"></i>ฝ่ายที่เบิกมากสุด<span class="sub">30 วัน</span></h3>
      <div class="leader">${d.topDept.length ? d.topDept.map(r => { const dp = S.depts.find(x => x.name === r.department) || {}; return `
        <div class="lrow" style="--c:${dp.color || '#E2231A'}"><span class="rk"></span><span class="li"><i class="fa-solid ${dp.icon || 'fa-building'}"></i></span>
          <div><div class="nm">${esc(r.department)}</div><div class="lb"><i data-w="${r.q / d.topDept[0].q * 100}%"></i></div></div><span class="qv">${fmt(r.q)}</span></div>`; }).join('')
        : '<div class="empty"><i class="fa-solid fa-trophy"></i>ยังไม่มีการเบิก</div>'}</div></div>
  </div>

  <div class="sec-h"><i class="fa-solid fa-screwdriver-wrench" style="color:var(--brand)"></i>งานซ่อม</div>
  <div class="insight stagger">
    ${[['ใบแจ้งซ่อมที่เปิดอยู่', fmt(rs.open_n) + ' ใบ', 'fa-folder-open', '#E2231A'],
       ['เวลาซ่อมเฉลี่ย', rs.avg_h != null ? fmt(rs.avg_h) + ' ชม.' : '—', 'fa-stopwatch', '#7a5af8'],
       ['ค้างนานสุด', rs.oldest ? (oldestH >= 24 ? Math.floor(oldestH / 24) + ' วัน' : oldestH + ' ชม.') : '—', 'fa-hourglass-half', '#f79009'],
       ['ซ่อมเสร็จ / ตัดจำหน่าย (เดือนนี้)', fmt(rs.done_m) + ' / ' + fmt(rs.scrap_m), 'fa-circle-check', '#12b76a']]
      .map(([l, v, ic, c]) => `<div class="ins" style="--c:${c}"><div class="ii"><i class="fa-solid ${ic}"></i></div><div><b>${v}</b><span>${l}</span></div></div>`).join('')}
  </div>

  <div class="grid g-2 mt">
    <div class="card"><h3><i class="fa-solid fa-building"></i>พาเลทที่อยู่กับแต่ละฝ่าย<span class="sub">ยังไม่คืน</span></h3><div class="chart-box"><canvas id="cDept"></canvas></div></div>
    <div class="card"><h3><i class="fa-solid fa-bolt"></i>รายการล่าสุด<a class="sub" href="#logs">บันทึกประวัติทั้งหมด →</a></h3>
      <div class="timeline">${d.recent.length ? d.recent.map(tlItem).join('') : '<div class="empty"><i class="fa-solid fa-inbox"></i>ยังไม่มีรายการ — เริ่มจาก “รับเข้า”</div>'}</div>
    </div>
  </div>`;
  countUp(); tilt();
  setTimeout(() => {
    $$('.sbar i, .lrow .lb i').forEach(i => i.style.width = i.dataset.w);
    $$('.gauge circle.v').forEach(c => c.style.strokeDashoffset = c.dataset.off);
  }, 150);
  $$('.tcard').forEach(c => c.onclick = () => location.hash = 'stock');
  $('#resetBtn').onclick = resetData;
  view.onclick = e => {
    const g = e.target.closest('[data-go]'); if (g) location.hash = g.dataset.go;
    const sb = e.target.closest('[data-days]'); if (sb) { sessionSet('dashDays', sb.dataset.days); route(); }
  };
  drawDashCharts(d);
}
function gauge(title, sub, v, c) {
  const R = 30, C = 2 * Math.PI * R;
  return `<div class="gauge"><div class="gr"><svg width="76" height="76"><circle cx="38" cy="38" r="${R}" stroke="#f1eeee" stroke-width="8" fill="none"/>
    <circle class="v" cx="38" cy="38" r="${R}" stroke="${c}" stroke-width="8" fill="none" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${C}" data-off="${C * (1 - v / 100)}"/></svg>
    <b style="color:${c}">${v}%</b></div><div><h5>${title}</h5><p>${sub}</p></div></div>`;
}
function typeCard(t) {
  const sts = ['available', 'issued', 'damaged', 'repairing'];
  const tot = sts.reduce((a, s) => a + qType(t.id, s), 0);
  return `<div class="card tcard" style="--c:${t.color}">
    <span class="ribbon">${esc(t.code)}</span>
    ${palletSVG(t.color)}
    <h4>${esc(t.name)}</h4><div class="desc">${esc(t.description)} · ขนาด ${sizes(t).join(' / ')} ม.</div>
    <div class="total mt" data-count="${tot}">0</div><small style="color:var(--muted)">ตัวในระบบ</small>
    <div class="sbar">${sts.map(s => `<i style="background:${STATUS[s].c}" data-w="${tot ? qType(t.id, s) / tot * 100 : 0}%"></i>`).join('')}</div>
    <div class="slegend">${sts.map(s => `<div><span class="dot" style="--c:${STATUS[s].c}"></span>${STATUS[s].name}<b>${fmt(qType(t.id, s))}</b></div>`).join('')}</div>
  </div>`;
}
function tlItem(m) {
  const a = ACTIONS[m.action] || { name: m.action, icon: 'fa-circle' };
  return `<div class="tl act-${m.action}">
    <div class="ai"><i class="fa-solid ${a.icon}"></i></div>
    <div><b>${a.name}</b> <span class="tag" style="--c:${m.color}">${esc(m.code)} · ${esc(m.size)}</span>
      <div class="meta">${dtTH(m.moved_at)}${m.department ? ' · ' + esc(m.department) : ''}${m.actor || m.person ? ' · <i class="fa-regular fa-user"></i> ' + esc(m.actor || m.person) : ''}</div></div>
    <div class="q">${m.action === 'issue' ? '−' : m.action === 'receive' ? '+' : ''}${fmt(m.qty)}</div>
  </div>`;
}
function forkliftSVG() {
  return `<svg viewBox="0 0 160 110" width="150"><g>
    <rect x="96" y="6" width="5" height="88" rx="2" fill="#2b2f3a"/><rect x="104" y="6" width="4" height="88" rx="2" fill="#3b4150"/>
    <rect x="104" y="80" width="54" height="5" rx="2" fill="#2b2f3a"/>
    <g transform="translate(108 46)"><rect width="46" height="30" rx="3" fill="#d9a35b" stroke="#a9733a" stroke-width="2"/><rect x="18" width="9" height="30" fill="rgba(255,255,255,.35)"/></g>
    <rect x="108" y="76" width="46" height="5" fill="#F5B400"/>
    <path d="M20 30 h46 l18 34 v24 h-70 z" fill="#E2231A"/><path d="M26 34 h36 l14 28 h-50 z" fill="#ffe1df" opacity=".85"/>
    <rect x="14" y="62" width="80" height="26" rx="6" fill="#B5121B"/><rect x="4" y="66" width="16" height="22" rx="4" fill="#2b2f3a"/>
    <rect x="18" y="24" width="52" height="6" rx="3" fill="#2b2f3a"/><rect x="20" y="24" width="4" height="40" fill="#2b2f3a"/>
    <circle class="wheel" cx="32" cy="92" r="13" fill="#20232b"/><circle cx="32" cy="92" r="5" fill="#9aa3b5"/>
    <circle class="wheel" cx="78" cy="94" r="11" fill="#20232b"/><circle cx="78" cy="94" r="4" fill="#9aa3b5"/>
    <rect class="wheel" x="27" y="87" width="10" height="2" fill="#555"/>
  </g></svg>`;
}
function drawDashCharts(d) {
  Chart.defaults.font.family = 'Prompt'; Chart.defaults.color = '#6b6b72';
  const grid = '#f1e9e9';
  const legend = { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8 } };
  const labels = Object.keys(d.days).map(k => new Date(k).toLocaleDateString('th-TH', { day: 'numeric', month: 'short' }));
  const series = [['receive', 'รับเข้า', '#12b76a'], ['issue', 'เบิกจ่าย', '#0ea5e9'], ['return', 'รับคืน', '#1d1d1f'], ['damage', 'แจ้งชำรุด', '#E2231A'], ['repair_done', 'ซ่อมเสร็จ', '#7a5af8']];
  const ctx = $('#cLine').getContext('2d');
  S.charts.push(new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: series.map(([k, n, c]) => {
      const g = ctx.createLinearGradient(0, 0, 0, 260); g.addColorStop(0, c + '55'); g.addColorStop(1, c + '00');
      return { label: n, data: Object.values(d.days).map(v => v[k]), borderColor: c, backgroundColor: g, fill: k === 'receive' || k === 'issue', tension: .42, borderWidth: 3, pointRadius: 3, pointHoverRadius: 7, pointBackgroundColor: '#fff' };
    }) },
    options: { maintainAspectRatio: false, interaction: { mode: 'index', intersect: false }, animation: { duration: 1400, easing: 'easeOutQuart' },
      plugins: { legend }, scales: { y: { beginAtZero: true, grid: { color: grid }, ticks: { precision: 0 } }, x: { grid: { display: false } } } },
  }));
  const sts = ['available', 'issued', 'damaged', 'repairing'];
  const vals = sts.map(qAll), any = vals.some(Boolean);
  S.charts.push(new Chart($('#cDonut'), {
    type: 'doughnut',
    data: { labels: sts.map(s => STATUS[s].name), datasets: [{ data: any ? vals : [1, 0, 0, 0], backgroundColor: any ? sts.map(s => STATUS[s].c) : ['#eee'], borderWidth: 4, borderColor: '#fff', hoverOffset: 14 }] },
    options: { maintainAspectRatio: false, cutout: '68%', animation: { animateRotate: true, duration: 1600 }, plugins: { legend, tooltip: { enabled: any } } },
    plugins: [{ id: 'center', afterDraw(c) {
      const { ctx, chartArea: a } = c; ctx.save(); ctx.textAlign = 'center';
      ctx.font = '700 28px Kanit'; ctx.fillStyle = '#1c1c1e'; ctx.fillText(fmt(vals.reduce((x, y) => x + y, 0)), (a.left + a.right) / 2, (a.top + a.bottom) / 2 + 6);
      ctx.font = '12px Prompt'; ctx.fillStyle = '#6b6b72'; ctx.fillText('ตัวทั้งหมด', (a.left + a.right) / 2, (a.top + a.bottom) / 2 + 24); ctx.restore();
    } }],
  }));
  const hrs = [...Array(24).keys()].filter(h => h >= 6 && h <= 22 || d.hours[h].in || d.hours[h].out);
  S.charts.push(new Chart($('#cHour'), {
    type: 'bar',
    data: { labels: hrs.map(h => String(h).padStart(2, '0') + ':00'), datasets: [
      { label: 'เข้าคลัง (รับเข้า/คืน/ซ่อมเสร็จ)', data: hrs.map(h => d.hours[h].in), backgroundColor: '#12b76a', borderRadius: 6 },
      { label: 'ออกคลัง (เบิก/ชำรุด)', data: hrs.map(h => -d.hours[h].out), backgroundColor: '#E2231A', borderRadius: 6 },
    ] },
    options: { maintainAspectRatio: false, animation: { duration: 1200, easing: 'easeOutBack' },
      plugins: { legend, tooltip: { callbacks: { label: c => `${c.dataset.label}: ${fmt(Math.abs(c.raw))}` } } },
      scales: { x: { stacked: true, grid: { display: false }, ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 9 } }, y: { stacked: true, grid: { color: grid }, ticks: { precision: 0, callback: v => Math.abs(v) } } } },
  }));
  const depts = [...new Set(d.dept.map(r => r.department))];
  S.charts.push(new Chart($('#cDept'), {
    type: 'bar',
    data: { labels: depts.length ? depts : ['ยังไม่มีการเบิก'], datasets: S.types.map(t => ({
      label: t.code + (t.tkey === 'PL' ? ' (พลาสติก)' : ''), backgroundColor: t.color, borderRadius: 8, borderSkipped: false,
      data: depts.map(dp => d.dept.filter(r => r.department === dp && +r.type_id === +t.id).reduce((a, r) => a + +r.qty, 0)),
    })) },
    options: { indexAxis: 'y', maintainAspectRatio: false, animation: { duration: 1300, easing: 'easeOutBack' },
      plugins: { legend }, scales: { x: { stacked: true, beginAtZero: true, grid: { color: grid }, ticks: { precision: 0 } }, y: { stacked: true, grid: { display: false } } } },
  }));
}

/* ================= AUDIT LOG ================= */
const LOGCAT = {
  pallet: ['รายการพาเลท', 'fa-pallet'], repair: ['งานซ่อม', 'fa-screwdriver-wrench'],
  setting: ['ตั้งค่าระบบ', 'fa-gear'], warn: ['ถูกปฏิเสธ', 'fa-ban'], account: ['บัญชีผู้ใช้', 'fa-user-shield'],
};
const LOGACT = {
  login: 'fa-right-to-bracket', logout: 'fa-right-from-bracket', login_failed: 'fa-user-lock', change_password: 'fa-key',
  user_create: 'fa-user-plus', user_update: 'fa-user-pen', user_reset_password: 'fa-key', user_toggle: 'fa-user-slash',
};
async function logs() {
  const n = nowParts();
  let cat = '';
  view.innerHTML = `
    <div class="log-kpi" id="logKpi"></div>
    <div class="card">
      <div class="filters">
        <div class="field"><label>ตั้งแต่วันที่</label><input type="date" id="lFrom" value="${n.date.slice(0, 8)}01"></div>
        <div class="field"><label>ถึงวันที่</label><input type="date" id="lTo" value="${n.date}"></div>
        <div class="field" style="flex:1;min-width:220px"><label>ค้นหา</label><input id="lQ" placeholder="เลขที่เอกสาร / ผู้ทำรายการ / ชื่อผู้ใช้ / ฝ่าย / รายละเอียด"></div>
        <button class="btn btn-primary" id="lGo"><i class="fa-solid fa-magnifying-glass"></i>ค้นหา</button>
        <button class="btn btn-ok" id="lCsv"><i class="fa-solid fa-file-excel"></i>Excel</button>
      </div>
      <div id="lRes"><div class="spinner"></div></div>
    </div>`;
  const params = (withCat = true) => '&' + new URLSearchParams({ from: $('#lFrom').value, to: $('#lTo').value, q: $('#lQ').value, cat: withCat ? cat : '' });
  const load = async () => {
    $('#lRes').innerHTML = '<div class="spinner"></div>';
    const { items } = await api('logs', null, params(false));
    const cnt = k => items.filter(i => i.category === k).length;
    $('#logKpi').innerHTML = [['', 'ทั้งหมด', 'fa-list', '#1d1d1f', items.length], ...Object.entries(LOGCAT).map(([k, [nm, ic]]) => [k, nm, ic, { pallet: '#0ea5e9', repair: '#7a5af8', setting: '#1d1d1f', warn: '#f79009', account: '#E2231A' }[k], cnt(k)])]
      .map(([k, nm, ic, c, v]) => `<div class="lk ${k === cat ? 'on' : ''}" data-cat="${k}" style="--c:${c}"><b>${fmt(v)}</b><span><i class="fa-solid ${ic}"></i>${nm}</span></div>`).join('');
    const list = cat ? items.filter(i => i.category === cat) : items;
    if (!list.length) { $('#lRes').innerHTML = '<div class="empty"><i class="fa-solid fa-clipboard"></i>ไม่พบบันทึกประวัติตามเงื่อนไข</div>'; return; }
    const groups = {};
    list.forEach(i => (groups[i.created_at.slice(0, 10)] ||= []).push(i));
    let k = 0;
    $('#lRes').innerHTML = Object.entries(groups).map(([day, arr]) => `
      <div class="logday"><h4><i class="fa-regular fa-calendar"></i>${new Date(day).toLocaleDateString('th-TH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })} · ${arr.length} รายการ</h4>
        <div class="logline">${arr.map(i => {
          const ac = ACTIONS[i.action], cc = LOGCAT[i.category] || ['', 'fa-circle'];
          return `<div class="logitem cat-${i.category}" style="animation-delay:${Math.min(k++, 20) * .03}s">
            <div class="ai"><i class="fa-solid ${ac?.icon || LOGACT[i.action] || cc[1]}"></i></div>
            <div class="tx"><p>${esc(i.detail)}</p>
              <div class="mt2"><span><i class="fa-solid fa-tag"></i> ${cc[0]}</span>${i.ref ? `<span><i class="fa-solid fa-hashtag"></i> ${esc(i.ref)}</span>` : ''}
              <span class="by" title="ผู้ทำรายการ / By"><i class="fa-solid fa-user-check"></i> ผู้ทำรายการ / By: <b>${i.actor ? esc(i.actor) : i.username ? esc(i.username) + ' <span class="legacy">(ยังไม่เข้าสู่ระบบ)</span>' : '<span class="legacy">ไม่ระบุ (ข้อมูลเดิม)</span>'}</b></span><span><i class="fa-solid fa-network-wired"></i> ${esc(i.ip)}</span></div></div>
            <div class="tm">${i.created_at.slice(11, 16)} น.</div></div>`;
        }).join('')}</div></div>`).join('');
  };
  view.onclick = e => { const c = e.target.closest('[data-cat]'); if (c) { cat = c.dataset.cat; load(); } };
  $('#lGo').onclick = load;
  $('#lQ').onkeydown = e => e.key === 'Enter' && load();
  $('#lCsv').onclick = () => downloadCsv('logs_export', params());
  load();
}

/* ================= TRANSACTION FORMS ================= */
// ชื่อที่เซิร์ฟเวอร์จะบันทึกเป็นผู้ทำรายการ: "ชื่อ (username)"
const meLabel = () => S.me ? (S.me.fullname ? `${S.me.fullname} (${S.me.username})` : S.me.username) : '';
const TX = {
  receive: { from: null, to: 'available', btn: 'btn-ok', icon: 'fa-truck-ramp-box', label: 'บันทึกรับเข้า', color: '#12b76a' },
  issue:   { from: 'available', to: 'issued', btn: 'btn-primary', icon: 'fa-dolly', label: 'ยืนยันเบิกจ่าย', color: '#0ea5e9', dept: true },
  return:  { from: 'issued', to: 'available', btn: 'btn-primary', icon: 'fa-rotate-left', label: 'บันทึกรับคืน', color: '#E2231A', dept: true },
  damage:  { from: 'available', to: 'damaged', btn: 'btn-bad', icon: 'fa-triangle-exclamation', label: 'แจ้งชำรุด', color: '#f04438' },
};
function txForm(kind) {
  const cfg = TX[kind];
  const f = { type_id: null, size: null, qty: 1, department: null, condition: 'good', ...nowParts() };
  const avail = (t, s) => kind === 'return' ? (f.department ? deptQty(f.department, t, s) : 0) : q(t, s, cfg.from);
  const availType = t => sizes(typeById(t)).reduce((a, s) => a + avail(t, s), 0);

  view.innerHTML = `
  <div class="wizard">
    <div class="card">
      ${cfg.dept ? `<div class="step" id="stDept"><div class="step-h"><span class="n">1</span>${kind === 'issue' ? 'ฝ่ายที่เบิกจ่าย' : 'ฝ่ายที่นำมาคืน'}</div>
        <div class="depts">${S.depts.map(d => `<button class="dept" data-d="${esc(d.name)}" style="--c:${d.color}"><i class="fa-solid ${d.icon}"></i><span>${esc(d.name)}</span>${kind === 'return' ? `<em>ถืออยู่ ${fmt(S.dept.filter(r => r.department === d.name).reduce((a, r) => a + +r.qty, 0))} ตัว</em>` : ''}</button>`).join('')}</div></div>` : ''}
      <div class="step" id="stType"><div class="step-h"><span class="n">${cfg.dept ? 2 : 1}</span>เลือกประเภทพาเลท</div><div class="pick-types" id="ptypes"></div></div>
      <div class="step" id="stSize"><div class="step-h"><span class="n">${cfg.dept ? 3 : 2}</span>ขนาด (ม.)</div><div class="chips" id="psizes"><span class="hint">เลือกประเภทก่อน</span></div></div>
      <div class="step" id="stQty"><div class="step-h"><span class="n">${cfg.dept ? 4 : 3}</span>จำนวน (ตัว)</div>
        <div class="stepper"><button data-q="-1"><i class="fa-solid fa-minus"></i></button><input id="qty" type="number" min="1" value="1"><button data-q="1"><i class="fa-solid fa-plus"></i></button></div>
        <div class="quick">${[5, 10, 20, 50, 100].map(n => `<button data-set="${n}">${n}</button>`).join('')}<button data-set="max">ทั้งหมด</button></div>
      </div>
      ${kind === 'return' ? `<div class="step"><div class="step-h"><span class="n">5</span>สภาพพาเลทที่คืน</div>
        <div class="cond"><button class="chip sel" data-cond="good" style="--c:#12b76a"><i class="fa-solid fa-circle-check"></i>สภาพดี → พร้อมใช้</button>
        <button class="chip" data-cond="damaged" style="--c:#f04438"><i class="fa-solid fa-heart-crack"></i>ชำรุด → ส่งซ่อม</button></div></div>` : ''}
      <div class="step"><div class="step-h"><span class="n"><i class="fa-solid fa-pen" style="font-size:12px"></i></span>รายละเอียด</div>
        <div class="fields">
          <div class="field"><label><i class="fa-regular fa-calendar"></i> วันที่</label><input type="date" id="fDate" value="${f.date}"></div>
          <div class="field"><label><i class="fa-regular fa-clock"></i> เวลา</label><input type="time" id="fTime" value="${f.time}"></div>
          <div class="field"><label><i class="fa-regular fa-user"></i> ${kind === 'issue' ? 'ชื่อผู้มาเบิก' : kind === 'receive' ? 'ชื่อผู้ส่ง / ผู้รับเข้า' : 'ชื่อผู้นำมาคืน / ผู้พบ'} (ไม่บังคับ)</label><input id="fPerson" placeholder="บันทึกเพิ่มเติมได้"></div>
          ${kind === 'damage' || kind === 'return' ? `<div class="field" id="causeBox" ${kind === 'return' ? 'hidden' : ''}><label><i class="fa-solid fa-heart-crack"></i> สาเหตุการชำรุด</label>
            <input id="fCause" list="causes" placeholder="เช่น ไม้หัก, ตะปูหลุด"><datalist id="causes"><option>ไม้หน้าพาเลทหัก</option><option>ไม้รองหัก</option><option>ตะปูหลุด/โผล่</option><option>ลูกบล็อกแตก</option><option>พลาสติกแตกร้าว</option><option>โดนรถโฟล์คลิฟท์ชน</option></datalist></div>` : ''}
          <div class="field full"><label><i class="fa-regular fa-note-sticky"></i> หมายเหตุ</label><input id="fNote" placeholder="${kind === 'receive' ? 'เช่น ซัพพลายเออร์ / เลขที่ใบส่งของ' : 'ไม่บังคับ'}"></div>
        </div>
      </div>
    </div>

    <div class="card summary">
      <div class="head" id="sHead"><h3><i class="fa-solid ${cfg.icon} fa-beat" style="--fa-animation-duration:2.4s"></i>${PAGES[kind][0]}</h3><div id="sPal">${palletSVG('#9aa3b5', 'big-pal')}</div></div>
      <div class="flow">${cfg.from ? `<span class="pill" style="--c:${STATUS[cfg.from].c}">${kind === 'return' ? 'อยู่ที่ฝ่าย' : STATUS[cfg.from].name}</span>` : '<span class="pill" style="--c:#1d1d1f">ภายนอก</span>'}
        <i class="fa-solid fa-angles-right"></i><span class="pill" id="flowTo" style="--c:${STATUS[cfg.to].c}">${STATUS[cfg.to].name}</span></div>
      <div id="sRows"></div>
      <button class="btn btn-lg ${cfg.btn}" id="submit"><i class="fa-solid ${cfg.icon}"></i>${cfg.label}</button>
      <div class="hint" id="sHint"></div>
    </div>
  </div>`;

  const renderTypes = () => {
    $('#ptypes').innerHTML = S.types.map(t => `<button class="ptype ${+f.type_id === +t.id ? 'sel' : ''}" data-t="${t.id}" style="--c:${t.color}">
      ${palletSVG(t.color)}<div class="code">${esc(t.code)}</div><div class="nm">${esc(t.name)}</div>
      ${kind !== 'receive' ? `<span class="avail">${kind === 'return' ? 'ถืออยู่' : 'คงเหลือ'} ${fmt(availType(t.id))}</span>` : ''}</button>`).join('');
  };
  const renderSizes = () => {
    const t = typeById(f.type_id);
    if (!t) return;
    $('#psizes').innerHTML = sizes(t).map(s => `<button class="chip ${f.size === s ? 'sel' : ''}" data-s="${s}" style="--c:${t.color}"><i class="fa-solid fa-ruler-combined"></i>${s} ม.${kind !== 'receive' ? ` <small>(${fmt(avail(t.id, s))})</small>` : ''}</button>`).join('');
  };
  const maxQty = () => (f.type_id && f.size && kind !== 'receive') ? avail(f.type_id, f.size) : 9999;
  const refresh = () => {
    const t = typeById(f.type_id);
    $('#stType').classList.toggle('done', !!t);
    $('#stSize').classList.toggle('done', !!f.size);
    $('#stQty').classList.toggle('done', f.qty > 0);
    $('#stDept')?.classList.toggle('done', !!f.department);
    if (t) { $('#sHead').style.setProperty('--sc', t.color); $('#sPal').innerHTML = palletSVG(t.color, 'big-pal'); }
    const to = kind === 'return' && f.condition === 'damaged' ? 'damaged' : cfg.to;
    const ft = $('#flowTo'); ft.style.setProperty('--c', STATUS[to].c); ft.textContent = STATUS[to].name;
    const rows = [
      ['ประเภท', t ? `<span class="tag" style="--c:${t.color}">${esc(t.code)}</span> ${esc(t.name)}` : '—'],
      ['ขนาด', f.size ? f.size + ' ม.' : '—'],
      ['จำนวน', `<span style="font-family:Kanit;font-size:20px">${fmt(f.qty)}</span> ตัว`],
    ];
    if (cfg.dept) rows.unshift(['ฝ่าย', f.department ? esc(f.department) : '—']);
    if (kind !== 'receive' && t && f.size) rows.push([kind === 'return' ? 'ฝ่ายถืออยู่' : 'คงเหลือพร้อมใช้', `${fmt(maxQty())} → <b style="color:${f.qty > maxQty() ? 'var(--bad)' : 'var(--ok)'}">${fmt(maxQty() - f.qty)}</b>`]);
    else if (kind === 'receive' && t && f.size) rows.push(['พร้อมใช้หลังรับ', `${fmt(q(t.id, f.size, 'available'))} → <b style="color:var(--ok)">${fmt(q(t.id, f.size, 'available') + f.qty)}</b>`]);
    rows.push(['วันที่ / เวลา', `${$('#fDate').value ? new Date($('#fDate').value).toLocaleDateString('th-TH', { day: 'numeric', month: 'short', year: 'numeric' }) : '—'} ${$('#fTime').value} น.`]);
    rows.push(['ผู้ทำรายการ', `<i class="fa-solid fa-user-check" style="color:var(--ok)"></i> ${esc(meLabel())}`]);
    $('#sRows').innerHTML = rows.map(([a, b]) => `<div class="srow"><span>${a}</span><b>${b}</b></div>`).join('');
    const miss = [cfg.dept && !f.department && 'ฝ่าย', !t && 'ประเภท', !f.size && 'ขนาด'].filter(Boolean);
    const over = kind !== 'receive' && t && f.size && f.qty > maxQty();
    $('#submit').disabled = miss.length > 0 || over || f.qty < 1;
    $('#sHint').innerHTML = miss.length ? `<i class="fa-solid fa-hand-pointer fa-bounce"></i> กรุณาเลือก ${miss.join(', ')}` : over ? `<span style="color:var(--bad)"><i class="fa-solid fa-circle-exclamation fa-shake"></i> จำนวนเกินที่มีอยู่</span>` : '<i class="fa-solid fa-circle-check" style="color:var(--ok)"></i> พร้อมบันทึก';
  };
  const setQty = n => { f.qty = Math.max(1, Math.min(9999, n | 0)); $('#qty').value = f.qty; refresh(); };

  renderTypes(); refresh();
  view.onclick = e => {
    const tEl = e.target.closest('[data-t]'), sEl = e.target.closest('[data-s]'), dEl = e.target.closest('[data-d]');
    const qb = e.target.closest('[data-q]'), qs = e.target.closest('[data-set]'), cEl = e.target.closest('[data-cond]');
    if (dEl) { f.department = dEl.dataset.d; $$('.dept').forEach(x => x.classList.toggle('sel', x === dEl)); renderTypes(); renderSizes(); }
    if (tEl) { f.type_id = +tEl.dataset.t; const ss = sizes(typeById(f.type_id)); f.size = ss.length === 1 ? ss[0] : (ss.includes(f.size) ? f.size : null); renderTypes(); renderSizes(); }
    if (sEl) { f.size = sEl.dataset.s; renderSizes(); }
    if (qb) setQty(f.qty + +qb.dataset.q);
    if (qs) setQty(qs.dataset.set === 'max' ? maxQty() : +qs.dataset.set);
    if (cEl) { f.condition = cEl.dataset.cond; $$('[data-cond]').forEach(x => x.classList.toggle('sel', x === cEl)); $('#causeBox').hidden = f.condition !== 'damaged'; }
    refresh();
  };
  $('#qty').oninput = e => setQty(+e.target.value);
  ['#fDate', '#fTime'].forEach(s => $(s).onchange = refresh);

  $('#submit').onclick = async () => {
    const btn = $('#submit');
    btn.disabled = true; btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> กำลังบันทึก...';
    try {
      const r = await api(kind, { ...f, date: $('#fDate').value, time: $('#fTime').value, person: $('#fPerson').value, note: $('#fNote').value, cause: $('#fCause')?.value || '' });
      confetti();
      const t = typeById(f.type_id);
      modal(`<div class="success-ic"><i class="fa-solid fa-check"></i></div>
        <h3 style="text-align:center">บันทึกสำเร็จ!</h3><p style="text-align:center;color:var(--muted)">${esc(r.message)}</p>
        <div class="receipt"><div class="doc">${esc(r.doc_no)}</div>
          <div class="srow"><span>รายการ</span><b>${PAGES[kind][0]}</b></div>
          ${f.department ? `<div class="srow"><span>ฝ่าย</span><b>${esc(f.department)}</b></div>` : ''}
          <div class="srow"><span>พาเลท</span><b>${esc(t.code)} · ${f.size} ม.</b></div>
          <div class="srow"><span>จำนวน</span><b>${fmt(f.qty)} ตัว</b></div>
          <div class="srow"><span>วันที่ / เวลา</span><b>${dtTH($('#fDate').value + ' ' + $('#fTime').value)}</b></div>
          <div class="srow"><span>ผู้ทำรายการ</span><b>${esc(meLabel())}</b></div></div>
        <div class="modal-acts"><button class="btn btn-ghost" data-close onclick="location.hash='dashboard'"><i class="fa-solid fa-gauge-high"></i>แดชบอร์ด</button>
        <button class="btn btn-primary" data-close><i class="fa-solid fa-plus"></i>ทำรายการต่อ</button></div>`);
      // ปิดหน้าต่างด้วยวิธีใดก็ตาม → โหลดฟอร์มใหม่ (ยอดคงเหลือ + วันเวลาปัจจุบัน)
      S.onModalClose = () => { if (location.hash.slice(1) === kind) route(); };
    } catch (e) {
      toast(e.message, 'err');
      btn.innerHTML = `<i class="fa-solid ${cfg.icon}"></i>${cfg.label}`;
      refresh();
    }
  };
}

/* ================= REPAIR ================= */
async function repair() {
  const { items } = await api('repairs');
  const cols = [
    ['damaged', 'ชำรุด / รอซ่อม', 'fa-heart-crack', 'c-damaged', items.filter(i => i.stage === 'damaged')],
    ['repairing', 'กำลังซ่อม', 'fa-screwdriver-wrench', 'c-repairing', items.filter(i => i.stage === 'repairing')],
    ['done', 'ซ่อมเสร็จ / ปิดงาน (30 วัน)', 'fa-circle-check', 'c-done', items.filter(i => i.stage === 'done' || i.stage === 'scrapped')],
  ];
  view.innerHTML = `
    <div style="display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap">
      <a href="#damage" class="btn btn-bad"><i class="fa-solid fa-triangle-exclamation"></i>แจ้งชำรุดจากคลัง</a>
      <a href="#return" class="btn btn-ghost"><i class="fa-solid fa-rotate-left"></i>รับคืนสภาพชำรุดจากฝ่าย</a>
    </div>
    <div class="kanban">${cols.map(([k, n, ic, cls, list]) => `
      <div class="col ${cls}"><div class="col-h"><span class="ci"><i class="fa-solid ${ic}"></i></span>${n}<span class="cnt">${fmt(list.reduce((a, i) => a + +i.qty, 0))} ตัว</span></div>
        ${list.length ? list.map((r, i) => repairCard(r, i)).join('') : `<div class="empty"><i class="fa-solid ${k === 'done' ? 'fa-mug-hot' : 'fa-face-smile'}"></i>${k === 'damaged' ? 'ไม่มีพาเลทชำรุด เยี่ยมมาก!' : 'ไม่มีรายการ'}</div>`}
      </div>`).join('')}</div>`;
  view.onclick = e => {
    const b = e.target.closest('[data-rp]');
    if (b) repairAction(items.find(i => +i.id === +b.dataset.id), b.dataset.rp);
    const ed = e.target.closest('[data-redit]'), dl = e.target.closest('[data-rdel]');
    if (ed) editRepair(items.find(i => +i.id === +ed.dataset.redit));
    if (dl) deleteRepair(items.find(i => +i.id === +dl.dataset.rdel));
  };
}
function repairCard(r, i) {
  const steps = [['damaged', 'ชำรุด', 'fa-heart-crack'], ['repairing', 'ซ่อม', 'fa-wrench'], ['done', r.stage === 'scrapped' ? 'ทิ้ง' : 'เสร็จ', r.stage === 'scrapped' ? 'fa-trash' : 'fa-check']];
  const lvl = { damaged: 0, repairing: 1, done: 2, scrapped: 2 }[r.stage];
  const stC = r.stage === 'scrapped' ? '#98a2b3' : ['#f04438', '#7a5af8', '#12b76a'][lvl];
  return `<div class="rcard" style="--tc:${r.color};animation-delay:${i * .06}s">
    <div class="top">${palletSVG(r.color)}<div><div class="tn">${esc(r.ticket_no)}</div><div class="tc">${esc(r.code)} · ${esc(r.size)} ม.${r.department ? ' · ' + esc(r.department) : ''}</div></div>
      <div class="qty">${fmt(r.qty)}<small> ตัว</small></div></div>
    <div class="track" style="--c:${stC}">${steps.map((s, k) => `${k ? `<div class="ln ${k <= lvl ? 'on' : ''}"></div>` : ''}<div class="s ${k <= lvl ? 'on' : ''} ${k === lvl && lvl < 2 ? 'cur' : ''}"><i class="fa-solid ${s[2]}"></i>${s[1]}</div>`).join('')}</div>
    ${r.cause ? `<div class="cause"><i class="fa-solid fa-heart-crack" style="color:var(--bad)"></i> ${esc(r.cause)}</div>` : ''}
    <div class="tc"><i class="fa-regular fa-clock"></i> แจ้ง ${dtTH(r.reported_at)}${r.reported_by ? ' · ' + esc(r.reported_by) : ''}</div>
    ${r.started_at ? `<div class="tc"><i class="fa-solid fa-wrench"></i> เริ่มซ่อม ${dtTH(r.started_at)}${r.repairer ? ' · ' + esc(r.repairer) : ''}</div>` : ''}
    ${r.finished_at ? `<div class="tc"><i class="fa-solid fa-flag-checkered"></i> ${r.stage === 'scrapped' ? 'ตัดจำหน่าย' : 'ซ่อมเสร็จ'} ${dtTH(r.finished_at)}</div>` : ''}
    ${r.updated_by ? `<div class="tc"><i class="fa-solid fa-user-check"></i> อัปเดตล่าสุดโดย ${esc(r.updated_by)}</div>` : ''}
    <div class="acts">
      ${r.stage === 'damaged' ? `<button class="btn btn-sm btn-fix" data-rp="repair_start" data-id="${r.id}"><i class="fa-solid fa-screwdriver-wrench"></i>ส่งซ่อม</button>` : ''}
      ${r.stage === 'repairing' ? `<button class="btn btn-sm btn-ok" data-rp="repair_done" data-id="${r.id}"><i class="fa-solid fa-circle-check"></i>ซ่อมเสร็จ</button>` : ''}
      ${r.stage === 'damaged' || r.stage === 'repairing' ? `<button class="btn btn-sm btn-ghost" data-rp="scrap" data-id="${r.id}"><i class="fa-solid fa-trash-can"></i>ซ่อมไม่ได้</button>` : ''}
      ${r.stage === 'done' ? `<span class="tag" style="--c:#12b76a"><i class="fa-solid fa-circle-check"></i>กลับเข้าคลังพร้อมใช้</span>` : ''}
      ${r.stage === 'scrapped' ? `<span class="tag" style="--c:#667085"><i class="fa-solid fa-trash-can"></i>ตัดจำหน่ายแล้ว</span>` : ''}
    </div>
    <div class="acts rec-acts"><button class="btn btn-sm btn-ghost rec-btn" data-redit="${r.id}" title="แก้ไขใบแจ้งซ่อม">✏️ แก้ไข</button><button class="btn btn-sm btn-ghost rec-btn" data-rdel="${r.id}" title="ลบใบแจ้งซ่อมทั้งชุด">🗑 ลบ</button></div></div>`;
}
function repairAction(r, act) {
  const meta = {
    repair_start: ['ส่งซ่อม', 'fa-screwdriver-wrench', 'btn-fix', 'ช่างผู้ซ่อม (ไม่ระบุ = ชื่อบัญชีของคุณ)', 'ชำรุด → กำลังซ่อม'],
    repair_done: ['ยืนยันซ่อมเสร็จ', 'fa-circle-check', 'btn-ok', 'ชื่อผู้ตรวจรับ (ไม่บังคับ)', 'กำลังซ่อม → พร้อมใช้'],
    scrap: ['ตัดจำหน่าย (ซ่อมไม่ได้)', 'fa-trash-can', 'btn-bad', 'ชื่อผู้อนุมัติ (ไม่บังคับ)', 'ออกจากระบบ'],
  }[act];
  const n = nowParts();
  const m = modal(`<h3><i class="fa-solid ${meta[1]}"></i> ${meta[0]}</h3>
    <p style="color:var(--muted)">${esc(r.ticket_no)} · ${esc(r.code)} ${esc(r.size)} ม. · ${fmt(r.qty)} ตัว — <b>${meta[4]}</b></p>
    <div style="text-align:center">${palletSVG(r.color, 'big-pal')}</div>
    <div class="fields mt">
      <div class="field"><label>วันที่</label><input type="date" id="mDate" value="${n.date}"></div>
      <div class="field"><label>เวลา</label><input type="time" id="mTime" value="${n.time}"></div>
      <div class="field full"><label>${meta[3]}</label><input id="mPerson" value="${act === 'repair_done' ? esc(r.repairer || '') : ''}"></div>
      <div class="field full"><label>หมายเหตุ</label><input id="mNote"></div>
    </div>
    <p class="hint" style="font-size:12.5px;margin-top:10px"><i class="fa-solid fa-user-check" style="color:var(--ok)"></i> ผู้ทำรายการ: <b>${esc(meLabel())}</b> (บันทึกจากบัญชีที่เข้าสู่ระบบ)</p>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn ${meta[2]}" id="mGo"><i class="fa-solid ${meta[1]}"></i>${meta[0]}</button></div>`);
  $('#mGo', m).onclick = async () => {
    try {
      const res = await api(act, { id: r.id, date: $('#mDate').value, time: $('#mTime').value, person: $('#mPerson').value, note: $('#mNote').value });
      closeModal();
      if (act === 'repair_done') confetti();
      toast(res.message);
      route();
    } catch (e) { toast(e.message, 'err'); }
  };
}

/* ================= STOCK ================= */
async function stock() {
  const sts = ['available', 'issued', 'damaged', 'repairing'];
  const sizeNum = s => s.split('x').reduce((a, n) => a * parseFloat(n), 1);
  const allSizes = [...new Set(S.types.flatMap(sizes))].sort((a, b) => sizeNum(a) - sizeNum(b));
  const rows = allSizes.flatMap(s => S.types.filter(t => sizes(t).includes(s))
    .map(t => ({ t, s, v: Object.fromEntries([...sts, 'scrapped'].map(x => [x, q(t.id, s, x)])) })));
  const sum = (k, list = rows) => list.reduce((a, r) => a + r.v[k], 0);
  const inSys = list => sts.reduce((a, k) => a + sum(k, list), 0);
  const szLabel = s => s.replace('x', ' x ');
  const R = 32, C = 2 * Math.PI * R;
  view.innerHTML = `
    ${allSizes.map(sz => { const grp = rows.filter(r => r.s === sz); return `
    <div class="size-h"><span class="size-pill"><i class="fa-solid fa-ruler-combined"></i>พาเลทขนาด ${szLabel(sz)} ม.</span>
      <small>${grp.length} ประเภท · ในระบบ ${fmt(inSys(grp))} ตัว · พร้อมใช้ ${fmt(sum('available', grp))} ตัว</small></div>
    <div class="stock-grid stagger">${grp.map(({ t, s, v }) => {
      const tot = sts.reduce((a, k) => a + v[k], 0), pct = tot ? v.available / tot : 0;
      return `<div class="card stile" style="--c:${t.color}">
        <div class="size-badge">${szLabel(s)} ม.</div>
        <div class="row1">${palletSVG(t.color)}<div><h4>${esc(t.code)}</h4><div class="sz">${esc(t.name)}<br>ขนาด ${s} ม.</div></div>
          <div class="ring"><svg width="74" height="74"><circle cx="37" cy="37" r="${R}" stroke="#e7ecf6" stroke-width="8" fill="none"/>
            <circle cx="37" cy="37" r="${R}" stroke="${t.color}" stroke-width="8" fill="none" stroke-linecap="round" stroke-dasharray="${C}" stroke-dashoffset="${C}" data-off="${C * (1 - pct)}"/></svg>
            <b>${Math.round(pct * 100)}%</b></div></div>
        <div class="cells">${sts.map(k => `<div class="cell" style="--c:${STATUS[k].c}"><b data-count="${v[k]}">0</b><span>${STATUS[k].name}</span></div>`).join('')}</div>
      </div>`;
    }).join('')}</div>`; }).join('')}

    <div class="card mt"><h3><i class="fa-solid fa-table"></i>ตารางคงเหลือ<span class="sub">หน่วย: ตัว</span></h3>
      <div class="tbl-wrap"><table><thead><tr><th>ลำดับ</th><th>รหัส</th><th>รายละเอียด</th><th>ขนาด/ม.</th>${sts.map(k => `<th class="num">${STATUS[k].name}</th>`).join('')}<th class="num">รวมในระบบ</th><th class="num">ตัดจำหน่าย</th></tr></thead>
      <tbody>${(() => { let i = 0; return allSizes.map(sz => { const grp = rows.filter(r => r.s === sz); return `
        <tr class="grp-row"><td colspan="${sts.length + 6}"><i class="fa-solid fa-ruler-combined"></i> พาเลทขนาด ${szLabel(sz)} ม.</td></tr>
        ${grp.map(({ t, s, v }) => `<tr><td>${++i}</td><td><span class="tcode" style="--c:${t.color}"><span class="sw"></span>${esc(t.code)}</span></td><td>${esc(t.name)}</td><td><b>${s}</b></td>
        ${sts.map(k => `<td class="num" style="color:${v[k] ? STATUS[k].c : '#c0c8d8'};font-weight:600">${fmt(v[k])}</td>`).join('')}<td class="num"><b>${fmt(sts.reduce((a, k) => a + v[k], 0))}</b></td><td class="num" style="color:#98a2b3">${fmt(v.scrapped)}</td></tr>`).join('')}
        <tr class="sub-row"><td colspan="4">รวมขนาด ${sz} ม.</td>${sts.map(k => `<td class="num">${fmt(sum(k, grp))}</td>`).join('')}<td class="num">${fmt(inSys(grp))}</td><td class="num">${fmt(sum('scrapped', grp))}</td></tr>`; }).join(''); })()}</tbody>
      <tfoot><tr><td colspan="4">รวมทั้งหมด</td>${sts.map(k => `<td class="num">${fmt(sum(k))}</td>`).join('')}<td class="num">${fmt(sts.reduce((a, k) => a + sum(k), 0))}</td><td class="num">${fmt(sum('scrapped'))}</td></tr></tfoot></table></div>
    </div>

    <div class="card mt"><h3><i class="fa-solid fa-building"></i>พาเลทที่อยู่กับฝ่ายต่าง ๆ (ยังไม่คืน)</h3>
      ${S.dept.length ? `<div class="tbl-wrap"><table><thead><tr><th>ฝ่าย</th><th>รหัส</th><th>ขนาด/ม.</th><th class="num">จำนวน</th></tr></thead><tbody>
      ${[...S.dept].sort((a, b) => sizeNum(a.size) - sizeNum(b.size) || typeById(a.type_id).sort - typeById(b.type_id).sort || a.department.localeCompare(b.department, 'th')).map(r => { const t = typeById(r.type_id); const d = S.depts.find(x => x.name === r.department); return `<tr><td><i class="fa-solid ${d?.icon || 'fa-building'}" style="color:${d?.color || '#E2231A'};width:20px"></i> ${esc(r.department)}</td><td><span class="tcode" style="--c:${t.color}"><span class="sw"></span>${esc(t.code)}</span> ${t.tkey === 'PL' ? '(พลาสติก)' : ''}</td><td>${esc(r.size)}</td><td class="num"><b>${fmt(r.qty)}</b></td></tr>`; }).join('')}</tbody></table></div>`
      : '<div class="empty"><i class="fa-solid fa-check-double"></i>ทุกฝ่ายคืนพาเลทครบแล้ว</div>'}
    </div>`;
  countUp();
  setTimeout(() => $$('[data-off]').forEach(c => c.style.strokeDashoffset = c.dataset.off), 120);
}

/* ================= HISTORY ================= */
async function history() {
  const n = nowParts();
  const first = n.date.slice(0, 8) + '01';
  view.innerHTML = `<div class="card">
    <div class="filters">
      <div class="field"><label>ตั้งแต่วันที่</label><input type="date" id="hFrom" value="${first}"></div>
      <div class="field"><label>ถึงวันที่</label><input type="date" id="hTo" value="${n.date}"></div>
      <div class="field"><label>ประเภท</label><select id="hType"><option value="">ทั้งหมด</option>${S.types.map(t => `<option value="${t.id}">${esc(t.code)} ${esc(t.name)}</option>`).join('')}</select></div>
      <div class="field"><label>รายการ</label><select id="hAct"><option value="">ทั้งหมด</option>${Object.entries(ACTIONS).map(([k, a]) => `<option value="${k}">${a.name}</option>`).join('')}</select></div>
      <div class="field"><label>ฝ่าย</label><select id="hDept"><option value="">ทั้งหมด</option>${S.depts.map(d => `<option>${esc(d.name)}</option>`).join('')}</select></div>
      <div class="field"><label>ค้นหา</label><input id="hQ" placeholder="เลขที่ / ผู้ทำรายการ / ชื่อ / หมายเหตุ"></div>
      <button class="btn btn-primary" id="hGo"><i class="fa-solid fa-magnifying-glass"></i>ค้นหา</button>
      <button class="btn btn-ok" id="hCsv"><i class="fa-solid fa-file-excel"></i>Excel</button>
    </div>
    <div id="hRes"><div class="spinner"></div></div></div>`;
  const params = () => '&' + new URLSearchParams({ from: $('#hFrom').value, to: $('#hTo').value, type: $('#hType').value, act: $('#hAct').value, dept: $('#hDept').value, q: $('#hQ').value });
  let items = [];
  const load = async () => {
    $('#hRes').innerHTML = '<div class="spinner"></div>';
    ({ items } = await api('history', null, params()));
    const tot = k => items.filter(i => i.action === k).reduce((a, i) => a + +i.qty, 0);
    $('#hRes').innerHTML = items.length ? `
      <div class="chips" style="margin-bottom:12px">${['receive', 'issue', 'return', 'damage', 'repair_done', 'scrap'].map(k => `<span class="tag act-${k}"><i class="fa-solid ${ACTIONS[k].icon}"></i>${ACTIONS[k].name} ${fmt(tot(k))}</span>`).join('')}</div>
      <div class="tbl-wrap" style="max-height:62vh"><table><thead><tr><th>เลขที่เอกสาร</th><th>วันที่ / เวลา</th><th>รายการ</th><th>พาเลท</th><th>ขนาด</th><th class="num">จำนวน</th><th>สถานะ</th><th>ฝ่าย</th><th>ผู้ทำรายการ / By</th><th>ชื่อที่ระบุ</th><th>หมายเหตุ</th><th>จัดการ</th></tr></thead><tbody>
      ${items.map(m => `<tr><td><b>${esc(m.doc_no)}</b></td><td>${dtTH(m.moved_at)}</td>
        <td><span class="tag act-${m.action}"><i class="fa-solid ${ACTIONS[m.action]?.icon}"></i>${ACTIONS[m.action]?.name || m.action}</span></td>
        <td><span class="tcode" style="--c:${m.color}"><span class="sw"></span>${esc(m.code)}</span></td><td>${esc(m.size)}</td><td class="num"><b>${fmt(m.qty)}</b></td>
        <td style="font-size:12.5px">${m.from_status ? STATUS[m.from_status].name : 'ภายนอก'} <i class="fa-solid fa-arrow-right" style="color:var(--muted);font-size:10px"></i> <b style="color:${STATUS[m.to_status]?.c}">${STATUS[m.to_status]?.name || ''}</b></td>
        <td>${esc(m.department || '—')}</td>
        <td class="by-cell">${m.actor ? `<i class="fa-solid fa-user-check" style="color:var(--ok)"></i> ${esc(m.actor)}` : '<span class="legacy" title="บันทึกก่อนมีระบบเข้าสู่ระบบ">— ข้อมูลเดิม</span>'}</td><td>${esc(m.person || '—')}</td><td style="max-width:240px;overflow:hidden;text-overflow:ellipsis">${esc(m.note || '')}</td>
        <td><div class="rec-acts"><button class="btn btn-sm btn-ghost rec-btn" data-medit="${m.id}" title="แก้ไขรายการ">✏️ แก้ไข</button><button class="btn btn-sm btn-ghost rec-btn" data-mdel="${m.id}" title="${m.repair_id ? 'ลบทั้งชุดงานซ่อม' : 'ลบรายการ'}">🗑 ลบ</button></div></td></tr>`).join('')}
      </tbody></table></div>` : '<div class="empty"><i class="fa-solid fa-magnifying-glass"></i>ไม่พบรายการตามเงื่อนไข</div>';
  };
  view.onclick = e => {
    const ed = e.target.closest('[data-medit]'), dl = e.target.closest('[data-mdel]');
    const m = ed || dl ? items.find(i => +i.id === +(ed || dl).dataset[ed ? 'medit' : 'mdel']) : null;
    if (!m) return;
    if (ed) editMovement(m); else deleteMovement(m);
  };
  $('#hGo').onclick = load;
  $('#hQ').onkeydown = e => e.key === 'Enter' && load();
  $('#hCsv').onclick = () => downloadCsv('export', params());
  load();
}

/* ================= EDIT / DELETE RECORDS (รหัสรีเซ็ตข้อมูล) =================
   เซิร์ฟเวอร์ตรวจย้อนทุกรายการตามลำดับเวลา — ถ้ายอดใดจะติดลบจะปฏิเสธและแจ้งเหตุผล (แสดงข้อความตามที่ได้รับ) */
const CHAIN_ACTS = ['damage', 'repair_start', 'repair_done', 'scrap'];
const deptEditable = m => m.action === 'issue' || m.action === 'return' || (m.action === 'damage' && m.from_status === 'issued');
function recordForm(html, onSave) {
  const m = modal(html + `<div id="eErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px;white-space:pre-line"></div>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-primary" id="eSave"><i class="fa-solid fa-floppy-disk"></i>บันทึกการแก้ไข</button></div>`);
  const go = $('#eSave', m), err = $('#eErr', m);
  go.onclick = async () => {
    if (go.disabled) return;
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> กำลังบันทึก...'; err.textContent = '';
    try {
      const r = await onSave(m);
      closeModal();
      toast(r.message, r.changed === false ? 'info' : 'ok');
      if (r.changed !== false) route(); // โหลดยอดคงเหลือ / ทุกหน้าใหม่
    } catch (e) {
      err.textContent = e.message; toast(e.message, 'err');
      go.disabled = false; go.innerHTML = '<i class="fa-solid fa-floppy-disk"></i>บันทึกการแก้ไข';
    }
  };
  return m;
}
function editMovement(m) {
  const a = ACTIONS[m.action] || { name: m.action, icon: 'fa-circle' };
  const chain = m.repair_id != null && CHAIN_ACTS.includes(m.action);
  const depts = [...new Set([...(m.department ? [m.department] : []), ...S.depts.map(d => d.name)])];
  recordForm(`<h3><i class="fa-solid fa-pen-to-square" style="color:var(--brand)"></i> แก้ไขรายการ ${esc(m.doc_no)}</h3>
    <p style="color:var(--muted)"><span class="tag act-${m.action}"><i class="fa-solid ${a.icon}"></i>${a.name}</span>
      <span class="tcode" style="--c:${m.color}"><span class="sw"></span>${esc(m.code)}</span> ${esc(m.size)} ม. ·
      ${m.from_status ? STATUS[m.from_status].name : 'ภายนอก'} → ${STATUS[m.to_status]?.name || ''}</p>
    <p class="hint" style="font-size:12.5px"><i class="fa-solid fa-circle-info"></i> เปลี่ยนรายการ / ประเภท / ขนาด / สถานะไม่ได้ — หากต้องการเปลี่ยน ให้ลบรายการนี้แล้วบันทึกใหม่</p>
    <p class="hint" style="font-size:12.5px"><i class="fa-solid fa-user-check" style="color:var(--ok)"></i> บันทึกโดย: <b>${esc(m.actor || 'ข้อมูลเดิม (ไม่ระบุ)')}</b> · การแก้ไขนี้จะบันทึกใน Log ในนาม <b>${esc(meLabel())}</b></p>
    ${chain ? `<p class="hint" style="font-size:12.5px;color:#7a5af8"><i class="fa-solid fa-link"></i> รายการนี้อยู่ในงานซ่อม <b>${esc(m.ticket_no || '')}</b> — แก้จำนวนจะปรับทุกรายการของใบนี้และใบแจ้งซ่อมพร้อมกัน · วันที่ต้องเรียง แจ้งชำรุด → ส่งซ่อม → ซ่อมเสร็จ/ตัดจำหน่าย</p>` : ''}
    <div class="fields mt">
      <div class="field"><label>จำนวน (ตัว)</label><input type="number" id="eQty" min="1" value="${+m.qty}"></div>
      ${deptEditable(m) ? `<div class="field"><label>ฝ่าย</label><select id="eDept">${depts.map(d => `<option ${d === m.department ? 'selected' : ''}>${esc(d)}</option>`).join('')}</select></div>` : '<div class="field"></div>'}
      <div class="field"><label>วันที่</label><input type="date" id="eDate" value="${m.moved_at.slice(0, 10)}"></div>
      <div class="field"><label>เวลา</label><input type="time" id="eTime" value="${m.moved_at.slice(11, 16)}"></div>
      <div class="field full"><label>ชื่อที่ระบุ (ไม่บังคับ)</label><input id="ePerson" value="${esc(m.person)}"></div>
      <div class="field full"><label>หมายเหตุ</label><input id="eNote" value="${esc(m.note)}"></div>
    </div>`, f => {
    const body = { id: m.id, doc_no: m.doc_no, qty: $('#eQty', f).value, date: $('#eDate', f).value, time: $('#eTime', f).value, person: $('#ePerson', f).value, note: $('#eNote', f).value };
    if ($('#eDept', f)) body.department = $('#eDept', f).value;
    return recordApi('movement_update', body);
  });
}
async function deleteMovement(m) {
  const a = ACTIONS[m.action]?.name || m.action;
  const chain = m.repair_id != null;
  const msg = chain ? [
    `ลบรายการ ${m.doc_no} (${a})?`, '',
    `รายการนี้เป็นส่วนหนึ่งของงานซ่อม ${m.ticket_no || ''}`,
    'ระบบจะลบทั้งชุดพร้อมกัน: ใบแจ้งซ่อม และทุกรายการเคลื่อนไหวของใบนี้ (แจ้งชำรุด / ส่งซ่อม / ซ่อมเสร็จ / ตัดจำหน่าย)', '',
    'ยอดคงเหลือจะคำนวณใหม่ · การลบย้อนกลับไม่ได้ ต้องการลบหรือไม่?',
  ] : [
    `ลบรายการ ${m.doc_no}?`, '',
    `${a} ${m.code} ${m.size} ม. จำนวน ${fmt(m.qty)} ตัว · ${dtTH(m.moved_at)}${m.department ? ' · ' + m.department : ''}`, '',
    'ยอดคงเหลือจะคำนวณใหม่ · การลบย้อนกลับไม่ได้ ต้องการลบหรือไม่?',
  ];
  try {
    await askPassword('record');
    if (!confirm(msg.join('\n'))) return;
    const r = await recordApi('movement_delete', { id: m.id, doc_no: m.doc_no });
    toast(r.message, 'info');
    route();
  } catch (e) { toast(e.message, 'err'); }
}
function editRepair(r) {
  recordForm(`<h3><i class="fa-solid fa-pen-to-square" style="color:var(--brand)"></i> แก้ไขใบแจ้งซ่อม ${esc(r.ticket_no)}</h3>
    <p style="color:var(--muted)">${esc(r.code)} · ${esc(r.size)} ม. · ${fmt(r.qty)} ตัว${r.department ? ' · ' + esc(r.department) : ''}</p>
    <p class="hint" style="font-size:12.5px"><i class="fa-solid fa-circle-info"></i> แก้ได้เฉพาะข้อความ — จำนวน / วันที่ แก้ที่รายการแจ้งชำรุดในหน้าประวัติเคลื่อนไหว (ปรับทั้งชุด) · สถานะเปลี่ยนด้วยปุ่มส่งซ่อม / ซ่อมเสร็จ</p>
    <p class="hint" style="font-size:12.5px"><i class="fa-solid fa-user-check" style="color:var(--ok)"></i> ผู้แจ้ง: <b>${esc(r.reported_by || 'ไม่ระบุ')}</b> (บันทึกจากบัญชีผู้ใช้ แก้ไขไม่ได้)</p>
    <div class="fields mt">
      <div class="field full"><label>สาเหตุการชำรุด</label><input id="eCause" value="${esc(r.cause)}"></div>
      <div class="field full"><label>ช่างผู้ซ่อม</label><input id="eFix" value="${esc(r.repairer)}"></div>
      <div class="field full"><label>หมายเหตุ</label><textarea id="eNote" rows="3">${esc(r.note)}</textarea></div>
    </div>`, f => recordApi('repair_update', { id: r.id, ticket_no: r.ticket_no, cause: $('#eCause', f).value, repairer: $('#eFix', f).value, note: $('#eNote', f).value }));
}
async function deleteRepair(r) {
  const msg = [
    `ลบใบแจ้งซ่อม ${r.ticket_no} ทั้งชุด?`, '',
    `${r.code} ${r.size} ม. จำนวน ${fmt(r.qty)} ตัว`,
    'ระบบจะลบใบแจ้งซ่อม และทุกรายการเคลื่อนไหวของใบนี้ (แจ้งชำรุด / ส่งซ่อม / ซ่อมเสร็จ / ตัดจำหน่าย) พร้อมกัน', '',
    'ยอดคงเหลือจะคำนวณใหม่ · การลบย้อนกลับไม่ได้ ต้องการลบหรือไม่?',
  ];
  try {
    await askPassword('record');
    if (!confirm(msg.join('\n'))) return;
    const res = await recordApi('repair_delete', { id: r.id, ticket_no: r.ticket_no });
    toast(res.message, 'info');
    route();
  } catch (e) { toast(e.message, 'err'); }
}

/* ================= SETTINGS ================= */
async function settings() {
  const icons = ['fa-industry', 'fa-box-open', 'fa-warehouse', 'fa-truck-fast', 'fa-seedling', 'fa-screwdriver-wrench', 'fa-clipboard-check', 'fa-flask', 'fa-oil-can', 'fa-people-carry-box', 'fa-store', 'fa-building'];
  const colors = ['#E2231A', '#1E6FE0', '#F5B400', '#1FA83A', '#E8590C', '#0CA678', '#7048E8', '#D6336C', '#0EA5E9', '#667085'];
  let ic = icons[0], col = colors[0];
  view.innerHTML = `<div class="grid g-2">
    <div class="card"><h3><i class="fa-solid fa-sitemap"></i>รายชื่อฝ่าย<span class="sub">${S.depts.length} ฝ่าย</span></h3>
      <div class="dept-list stagger">${S.depts.map(d => `<div class="ditem"><span class="di" style="--c:${d.color}"><i class="fa-solid ${d.icon}"></i></span><b>${esc(d.name)}</b><button data-del="${d.id}" title="ลบ"><i class="fa-solid fa-trash-can"></i></button></div>`).join('')}</div></div>
    <div class="card"><h3><i class="fa-solid fa-plus"></i>เพิ่มฝ่ายใหม่</h3>
      <div class="field"><label>ชื่อฝ่าย</label><input id="dName" placeholder="เช่น ฝ่ายจัดซื้อ"></div>
      <div class="field mt"><label>ไอคอน</label><div class="icon-pick">${icons.map(i => `<button data-ic="${i}" class="${i === ic ? 'sel' : ''}"><i class="fa-solid ${i}"></i></button>`).join('')}</div></div>
      <div class="field mt"><label>สี</label><div class="color-pick">${colors.map(c => `<button data-col="${c}" style="--c:${c}" class="${c === col ? 'sel' : ''}"></button>`).join('')}</div></div>
      <button class="btn btn-primary mt" id="dSave" style="width:100%"><i class="fa-solid fa-floppy-disk"></i>บันทึก</button>
    </div></div>`;
  view.onclick = async e => {
    const i = e.target.closest('[data-ic]'), c = e.target.closest('[data-col]'), d = e.target.closest('[data-del]');
    if (i) { ic = i.dataset.ic; $$('[data-ic]').forEach(x => x.classList.toggle('sel', x === i)); }
    if (c) { col = c.dataset.col; $$('[data-col]').forEach(x => x.classList.toggle('sel', x === c)); }
    if (d && confirm('ลบฝ่ายนี้ออกจากรายการเลือก? (ประวัติเดิมยังอยู่)')) {
      try { await api('dept_delete', { id: d.dataset.del }); toast('ลบฝ่ายแล้ว', 'info'); route(); }
      catch (err) { toast(err.message, 'err'); }
    }
  };
  $('#dSave').onclick = async () => {
    try { const r = await api('dept_save', { name: $('#dName').value, icon: ic, color: col }); toast(r.message); route(); }
    catch (e) { toast(e.message, 'err'); }
  };
}

/* ================= ACCOUNT (ตั้งค่าบัญชี) =================
   ทุกคน: ดูบัญชีของตนเอง + เปลี่ยนรหัสผ่าน · ผู้ดูแลระบบ: จัดการบัญชีผู้ใช้ทั้งหมด (เซิร์ฟเวอร์ตรวจสิทธิ์อีกชั้น) */
const dtOrDash = s => s ? dtTH(s) : '—';
async function account() {
  const { user } = await api('me');
  setMe(user);
  const me = S.me, admin = isAdmin();
  view.innerHTML = `<div class="grid g-2">
    <div class="card"><h3><i class="fa-solid fa-id-badge"></i>บัญชีของฉัน</h3>
      <div class="srow"><span>ชื่อผู้ใช้</span><b>${esc(me.username)}</b></div>
      <div class="srow"><span>ชื่อ - นามสกุล</span><b>${esc(me.fullname)}</b></div>
      <div class="srow"><span>สิทธิ์</span><b>${roleTag(me.role)}</b></div>
      <div class="srow"><span>เข้าสู่ระบบล่าสุด</span><b>${dtOrDash(me.last_login)}</b></div>
      <p class="hint" style="font-size:12.5px;margin-top:12px"><i class="fa-solid fa-circle-info"></i> ทุกรายการที่บันทึกจะแสดงผู้ทำรายการเป็น <b>${esc(meLabel())}</b></p>
      <button class="btn btn-ghost mt" id="aLogout"><i class="fa-solid fa-right-from-bracket"></i>ออกจากระบบ</button>
    </div>
    <div class="card"><h3><i class="fa-solid fa-key"></i>เปลี่ยนรหัสผ่าน</h3>
      <div class="field"><label>รหัสผ่านเดิม</label><input type="password" id="cpOld" autocomplete="current-password"></div>
      <div class="field mt"><label>รหัสผ่านใหม่ (อย่างน้อย 8 ตัวอักษร)</label><input type="password" id="cpNew" autocomplete="new-password"></div>
      <div class="field mt"><label>ยืนยันรหัสผ่านใหม่</label><input type="password" id="cpNew2" autocomplete="new-password"></div>
      <div id="cpErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
      <button class="btn btn-primary" id="cpGo" style="width:100%"><i class="fa-solid fa-floppy-disk"></i>เปลี่ยนรหัสผ่าน</button>
      <p class="hint" style="font-size:12.5px;margin-top:10px">เปลี่ยนแล้ว เครื่องอื่นที่เข้าสู่ระบบด้วยบัญชีนี้จะถูกออกจากระบบ</p>
    </div></div>
    ${admin ? `<div class="card mt"><h3><i class="fa-solid fa-users-gear"></i>จัดการบัญชีผู้ใช้<span class="sub" id="uCount"></span></h3><div id="uList"><div class="spinner"></div></div></div>
    <div class="card mt"><h3><i class="fa-solid fa-user-plus"></i>สร้างบัญชีใหม่</h3>
      <div class="fields">
        <div class="field"><label>ชื่อผู้ใช้ (a-z 0-9 . _ - · 3-30 ตัว)</label><input id="nuUser" autocomplete="off" autocapitalize="none" spellcheck="false" maxlength="30" placeholder="เช่น somchai"></div>
        <div class="field"><label>ชื่อ - นามสกุล</label><input id="nuName" maxlength="100" placeholder="เช่น สมชาย ใจดี"></div>
        <div class="field"><label>สิทธิ์</label><select id="nuRole"><option value="user">ผู้ใช้งาน</option><option value="admin">ผู้ดูแลระบบ</option></select></div>
        <div class="field"><label>รหัสผ่านเริ่มต้น (อย่างน้อย 8 ตัวอักษร)</label><input type="password" id="nuPass" autocomplete="new-password"></div>
      </div>
      <div id="nuErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
      <button class="btn btn-primary" id="nuGo"><i class="fa-solid fa-user-plus"></i>สร้างบัญชี</button>
    </div>` : ''}`;

  $('#aLogout').onclick = logout;
  const busy = (b, on, html) => { b.disabled = on; if (html) b.innerHTML = html; };
  $('#cpGo').onclick = async () => {
    const b = $('#cpGo'), e = $('#cpErr'), o = $('#cpOld').value, n = $('#cpNew').value, n2 = $('#cpNew2').value;
    e.textContent = '';
    if (!o || !n) { e.textContent = 'กรุณากรอกรหัสผ่านเดิมและรหัสผ่านใหม่'; return; }
    if (n.length < 8) { e.textContent = 'รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร'; return; }
    if (n !== n2) { e.textContent = 'ยืนยันรหัสผ่านใหม่ไม่ตรงกัน'; return; }
    busy(b, true, '<i class="fa-solid fa-spinner fa-spin"></i> กำลังบันทึก...');
    try {
      const r = await api('change_password', { old_password: o, new_password: n });
      setToken(r.token);
      ['#cpOld', '#cpNew', '#cpNew2'].forEach(s => $(s).value = '');
      toast(r.message);
    } catch (x) { e.textContent = x.message; }
    busy(b, false, '<i class="fa-solid fa-floppy-disk"></i>เปลี่ยนรหัสผ่าน');
  };
  if (!admin) return;

  let users = [];
  const loadUsers = async () => {
    const r = await api('users');
    users = r.users;
    $('#uCount').textContent = `${users.length} บัญชี · ใช้งาน ${users.filter(u => u.active).length}`;
    $('#uList').innerHTML = `<div class="tbl-wrap"><table><thead><tr><th>ชื่อผู้ใช้</th><th>ชื่อ - นามสกุล</th><th>สิทธิ์</th><th>สถานะ</th><th>เข้าสู่ระบบล่าสุด</th><th>จัดการ</th></tr></thead><tbody>
      ${users.map(u => `<tr class="${u.active ? '' : 'u-off'}"><td><b>${esc(u.username)}</b>${u.id === me.id ? ' <span class="tag" style="--c:#12b76a">คุณ</span>' : ''}</td><td>${esc(u.fullname)}</td><td>${roleTag(u.role)}</td>
        <td>${u.active ? '<span class="tag" style="--c:#12b76a"><i class="fa-solid fa-circle-check"></i>ใช้งาน</span>' : '<span class="tag" style="--c:#667085"><i class="fa-solid fa-ban"></i>ปิดใช้งาน</span>'}</td>
        <td style="white-space:nowrap">${dtOrDash(u.last_login)}</td>
        <td><div class="u-acts"><button class="btn btn-sm btn-ghost rec-btn" data-uedit="${u.id}"><i class="fa-solid fa-user-pen"></i>แก้ไข</button>
          <button class="btn btn-sm btn-ghost rec-btn" data-upw="${u.id}"><i class="fa-solid fa-key"></i>ตั้งรหัสใหม่</button>
          <button class="btn btn-sm btn-ghost rec-btn" data-utog="${u.id}"><i class="fa-solid ${u.active ? 'fa-user-slash' : 'fa-user-check'}"></i>${u.active ? 'ปิดใช้งาน' : 'เปิดใช้งาน'}</button></div></td></tr>`).join('')}
      </tbody></table></div>`;
  };
  const afterSelfChange = async () => { const r = await api('me'); setMe(r.user); if (!isAdmin()) return route(); };
  const userForm = (html, onSave) => {
    const m = modal(html + `<div id="uErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
      <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-primary" id="uSave"><i class="fa-solid fa-floppy-disk"></i>บันทึก</button></div>`);
    const go = $('#uSave', m);
    go.onclick = async () => {
      if (go.disabled) return;
      busy(go, true, '<i class="fa-solid fa-spinner fa-spin"></i> กำลังบันทึก...');
      try {
        const r = await onSave(m);
        closeModal(); toast(r.message, r.changed === false ? 'info' : 'ok');
        await loadUsers();
      } catch (x) {
        if (x.code === 'AUTH') return;
        $('#uErr', m).textContent = x.message;
        busy(go, false, '<i class="fa-solid fa-floppy-disk"></i>บันทึก');
      }
    };
    return m;
  };
  view.onclick = async e => {
    const ed = e.target.closest('[data-uedit]'), pw = e.target.closest('[data-upw]'), tg = e.target.closest('[data-utog]');
    const u = users.find(x => x.id === +(ed || pw || tg)?.dataset[ed ? 'uedit' : pw ? 'upw' : 'utog']);
    if (!u) return;
    if (ed) userForm(`<h3><i class="fa-solid fa-user-pen" style="color:var(--brand)"></i> แก้ไขบัญชี ${esc(u.username)}</h3>
        <div class="field mt"><label>ชื่อ - นามสกุล</label><input id="ueName" maxlength="100" value="${esc(u.fullname)}"></div>
        <div class="field mt"><label>สิทธิ์</label><select id="ueRole"><option value="user" ${u.role === 'user' ? 'selected' : ''}>ผู้ใช้งาน</option><option value="admin" ${u.role === 'admin' ? 'selected' : ''}>ผู้ดูแลระบบ</option></select></div>
        <div class="field mt"><label>สถานะ</label><select id="ueActive"><option value="1" ${u.active ? 'selected' : ''}>ใช้งาน</option><option value="0" ${u.active ? '' : 'selected'}>ปิดใช้งาน (ออกจากระบบทุกเครื่อง)</option></select></div>`,
      async m => {
        const r = await api('user_save', { id: u.id, fullname: $('#ueName', m).value, role: $('#ueRole', m).value, active: $('#ueActive', m).value });
        if (u.id === me.id) await afterSelfChange();
        return r;
      });
    if (pw) userForm(`<h3><i class="fa-solid fa-key" style="color:var(--brand)"></i> ตั้งรหัสผ่านใหม่ให้ ${esc(u.username)}</h3>
        <p style="color:var(--muted)">${esc(u.fullname)} จะถูกออกจากระบบทุกเครื่อง และต้องเข้าสู่ระบบด้วยรหัสผ่านใหม่</p>
        <div class="field mt"><label>รหัสผ่านใหม่ (อย่างน้อย 8 ตัวอักษร)</label><input type="password" id="upNew" autocomplete="new-password"></div>
        <div class="field mt"><label>ยืนยันรหัสผ่านใหม่</label><input type="password" id="upNew2" autocomplete="new-password"></div>`,
      async m => {
        const n = $('#upNew', m).value;
        if (n.length < 8) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 8 ตัวอักษร');
        if (n !== $('#upNew2', m).value) throw new Error('ยืนยันรหัสผ่านใหม่ไม่ตรงกัน');
        const r = await api('user_reset_password', { id: u.id, password: n });
        if (r.token) setToken(r.token); // ตั้งรหัสใหม่ให้ตัวเอง: เครื่องนี้ยังเข้าสู่ระบบอยู่
        return r;
      });
    if (tg) {
      if (!confirm(u.active ? `ปิดใช้งานบัญชี ${u.username} (${u.fullname})?\nผู้ใช้นี้จะถูกออกจากระบบทุกเครื่องและเข้าสู่ระบบไม่ได้จนกว่าจะเปิดใช้งานอีกครั้ง` : `เปิดใช้งานบัญชี ${u.username} (${u.fullname})?`)) return;
      try {
        const r = await api('user_toggle', { id: u.id });
        toast(r.message, 'info');
        if (u.id === me.id) return authLost();
        await loadUsers();
      } catch (x) { if (x.code !== 'AUTH') toast(x.message, 'err'); }
    }
  };
  $('#nuGo').onclick = async () => {
    const b = $('#nuGo'), e = $('#nuErr');
    const body = { username: $('#nuUser').value.trim().toLowerCase(), fullname: $('#nuName').value.trim(), role: $('#nuRole').value, password: $('#nuPass').value };
    e.textContent = '';
    if (!/^[a-z0-9._-]{3,30}$/.test(body.username)) { e.textContent = 'ชื่อผู้ใช้ต้องยาว 3-30 ตัว ใช้ได้เฉพาะ a-z 0-9 . _ -'; return; }
    if (!body.fullname) { e.textContent = 'กรุณากรอกชื่อ - นามสกุล'; return; }
    if (body.password.length < 8) { e.textContent = 'รหัสผ่านต้องมีอย่างน้อย 8 ตัวอักษร'; return; }
    busy(b, true, '<i class="fa-solid fa-spinner fa-spin"></i> กำลังสร้าง...');
    try {
      const r = await api('user_save', body);
      toast(r.message);
      ['#nuUser', '#nuName', '#nuPass'].forEach(s => $(s).value = '');
      await loadUsers();
    } catch (x) { e.textContent = x.message; }
    busy(b, false, '<i class="fa-solid fa-user-plus"></i>สร้างบัญชี');
  };
  await loadUsers();
}

/* รีเซ็ตข้อมูล: ถามรหัสรีเซ็ต (ใช้ครั้งเดียว ไม่เก็บไว้) → ยืนยัน → POST reset_data */
const RESET_CONFIRM = [
  'ยืนยันการรีเซ็ตข้อมูล?',
  '',
  'ระบบจะลบข้อมูลต่อไปนี้ทั้งหมด:',
  '• ประวัติเคลื่อนไหวพาเลท (movements) — ยอดคงเหลือทุกประเภทจะกลับเป็น 0',
  '• ใบแจ้งซ่อม / งานซ่อม (repairs)',
  '• บันทึกประวัติ (audit_logs)',
  '',
  'ข้อมูลที่ยังเก็บไว้: ประเภทพาเลท และรายชื่อฝ่าย',
  '',
  'การลบนี้ย้อนกลับไม่ได้ ต้องการดำเนินการต่อหรือไม่?',
].join('\n');
function resetData() {
  const m = modal(`<h3><i class="fa-solid fa-trash-can" style="color:var(--bad)"></i> รีเซ็ตข้อมูล / Reset data</h3>
    <p style="color:var(--muted)">ใส่รหัสรีเซ็ตข้อมูล (ตั้งโดยผู้ดูแลระบบ — คนละรหัสกับรหัสเข้าสู่ระบบ)</p>
    <div class="field mt"><label>รหัสรีเซ็ตข้อมูล</label><input type="password" id="rPw" autocomplete="off"></div>
    <div id="rErr" style="color:var(--bad);min-height:20px;margin-top:8px;font-size:13px"></div>
    <div class="modal-acts"><button class="btn btn-ghost" data-close>ยกเลิก</button><button class="btn btn-bad" id="rGo"><i class="fa-solid fa-trash-can"></i>รีเซ็ตข้อมูล</button></div>`);
  const inp = $('#rPw', m), go = $('#rGo', m), err = $('#rErr', m);
  const idle = () => { go.disabled = false; go.innerHTML = '<i class="fa-solid fa-trash-can"></i>รีเซ็ตข้อมูล'; };
  const submit = async () => {
    if (go.disabled) return;
    let pw = inp.value;
    if (!pw) { err.textContent = 'กรุณาใส่รหัสรีเซ็ตข้อมูล'; inp.focus(); return; }
    if (!confirm(RESET_CONFIRM)) { inp.focus(); return; }
    go.disabled = true; go.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> กำลังรีเซ็ต...'; err.textContent = '';
    try {
      const r = await gasPost({ action: 'reset_data', resetPassword: pw });
      pw = ''; inp.value = '';
      closeModal();
      const c = r.removed || {};
      toast(`รีเซ็ตข้อมูลแล้ว — ลบประวัติเคลื่อนไหว ${fmt(c.movements)} · งานซ่อม ${fmt(c.repairs)} · Log ${fmt(c.audit_logs)} แถว`);
      route(); // โหลดข้อมูล (bootstrap) และหน้าปัจจุบันใหม่ทั้งหมด
    } catch (e) {
      pw = ''; inp.value = '';
      err.textContent = e.message; idle(); inp.focus();
    }
  };
  go.onclick = submit;
  inp.onkeydown = e => { if (e.key === 'Enter') submit(); };
  inp.focus();
}

start();
