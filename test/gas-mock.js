/*
 * In-memory mocks of the Apps Script services used by apps-script/Code.gs,
 * plus an end-to-end test of every API action through doGet/doPost.
 *
 *   node test/gas-mock.js
 */
"use strict";
const vm = require("vm");
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const crypto = require("crypto");
const { Buffer } = require("buffer"); // explicit: test/pallet-dev.js evaluates this file in a bare vm context

const ROOT = path.join(__dirname, "..");

/* ===================== mocks ===================== */
function createGas() {
  const state = { props: {}, cache: {}, cachePuts: [], sleeps: 0, locks: 0, unlocks: 0, lockFail: false, spreadsheets: {}, seq: 0, formatCalls: 0, created: 0 };

  class Range {
    constructor(sheet, row, col, nr, nc) {
      if (row < 1 || col < 1 || nr < 1 || nc < 1) throw new Error("Range out of bounds");
      if (row + nr - 1 > sheet.maxRows) throw new Error("Range exceeds sheet rows: " + (row + nr - 1) + " > " + sheet.maxRows);
      if (col + nc - 1 > sheet.maxCols) throw new Error("Range exceeds sheet columns: " + (col + nc - 1) + " > " + sheet.maxCols);
      Object.assign(this, { sheet, row, col, nr, nc });
    }
    getValues() {
      const out = [];
      for (let r = 0; r < this.nr; r++) {
        const src = this.sheet.data[this.row - 1 + r] || [];
        const line = [];
        for (let c = 0; c < this.nc; c++) {
          const v = src[this.col - 1 + c];
          line.push(v === undefined || v === null ? "" : v);
        }
        out.push(line);
      }
      return out;
    }
    setValues(values) {
      assert.strictEqual(values.length, this.nr, "setValues row count");
      values.forEach((line, r) => {
        assert.strictEqual(line.length, this.nc, "setValues column count");
        const idx = this.row - 1 + r;
        while (this.sheet.data.length <= idx) this.sheet.data.push([]);
        line.forEach((v, c) => {
          if (typeof v === "string" && v.charAt(0) === "'") v = v.slice(1); // Sheets consumes the apostrophe prefix (stored as text)
          else if (typeof v === "string" && v.charAt(0) === "=") throw new Error("Formula written to sheet: " + v);
          this.sheet.data[idx][this.col - 1 + c] = v;
        });
      });
      return this;
    }
    setNumberFormats(f) {
      assert.strictEqual(f.length, this.nr, "setNumberFormats rows");
      f.forEach(line => assert.strictEqual(line.length, this.nc, "setNumberFormats cols"));
      state.formatCalls++;
      return this;
    }
    setFontWeight() { return this; }
    clearContent() {
      state.clearCalls = (state.clearCalls || 0) + 1;
      for (let r = 0; r < this.nr; r++) {
        const line = this.sheet.data[this.row - 1 + r];
        if (line) for (let c = 0; c < this.nc; c++) if (this.col - 1 + c < line.length) line[this.col - 1 + c] = "";
      }
      return this;
    }
  }

  class Sheet {
    constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; this.maxCols = 26; }
    getName() { return this.name; }
    setName(n) { this.name = n; return this; }
    getLastRow() {
      for (let r = this.data.length - 1; r >= 0; r--) {
        if ((this.data[r] || []).some(v => v !== "" && v !== undefined && v !== null)) return r + 1;
      }
      return 0;
    }
    getLastColumn() { return this.data.reduce((m, row) => Math.max(m, row.length), 0); }
    getMaxRows() { return this.maxRows; }
    insertRowsAfter(after, n) { this.maxRows += n; return this; }
    getMaxColumns() { return this.maxCols; }
    insertColumnsAfter(after, n) { this.maxCols += n; state.colInserts = (state.colInserts || 0) + 1; return this; }
    getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    setFrozenRows(n) { this.frozen = n; return this; }
    getFrozenRows() { return this.frozen || 0; }
    deleteRows(pos, n) {
      if (pos < 1 || n < 1 || pos + n - 1 > this.maxRows) throw new Error("deleteRows out of bounds");
      if (this.maxRows - n <= (this.frozen || 0)) throw new Error("Sorry, it is not possible to delete all non-frozen rows.");
      state.deleteCalls = (state.deleteCalls || []).concat([[this.name, pos, n]]);
      this.data.splice(pos - 1, n);
      this.maxRows -= n;
      return this;
    }
    deleteRow(pos) { return this.deleteRows(pos, 1); }
  }

  class Spreadsheet {
    constructor(name) { this.id = "ss" + (++state.seq); this.name = name; this.sheets = [new Sheet("Sheet1")]; this.tz = null; }
    getId() { return this.id; }
    getUrl() { return "https://docs.google.com/spreadsheets/d/" + this.id + "/edit"; }
    getSheets() { return this.sheets.slice(); }
    getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
    insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
    setSpreadsheetTimeZone(tz) { this.tz = tz; }
  }

  const gas = {
    SpreadsheetApp: {
      create(name) { const ss = new Spreadsheet(name); state.spreadsheets[ss.id] = ss; state.created++; return ss; },
      openById(id) { const ss = state.spreadsheets[id]; if (!ss) throw new Error("Spreadsheet not found: " + id); return ss; }
    },
    PropertiesService: {
      getScriptProperties() {
        return {
          getProperty: k => (k in state.props ? state.props[k] : null),
          setProperty: (k, v) => { state.props[k] = String(v); },
          setProperties: (o) => { Object.assign(state.props, o); },
          deleteProperty: k => { delete state.props[k]; }
        };
      }
    },
    CacheService: {
      getScriptCache() {
        return {
          get: k => (k in state.cache ? state.cache[k] : null),
          put: (k, v, ttl) => { state.cache[k] = String(v); state.cachePuts.push(ttl); },
          remove: k => { delete state.cache[k]; }
        };
      }
    },
    LockService: {
      getScriptLock() {
        return {
          waitLock() { if (state.lockFail) throw new Error("Lock timeout"); state.locks++; },
          tryLock() { state.locks++; return true; },
          releaseLock() { state.unlocks++; }
        };
      }
    },
    Utilities: {
      sleep(ms) { state.sleeps += ms; },
      DigestAlgorithm: { SHA_256: "SHA_256" },
      Charset: { UTF_8: "UTF_8" },
      getUuid() { return crypto.randomUUID(); },
      // Like Apps Script: string (charset) or byte array in, signed byte array out.
      computeDigest(alg, value) {
        assert.strictEqual(alg, "SHA_256");
        const buf = typeof value === "string" ? Buffer.from(value, "utf8") : Buffer.from(value.map(b => b & 255));
        return Array.from(crypto.createHash("sha256").update(buf).digest(), b => (b > 127 ? b - 256 : b));
      }
    },
    ContentService: {
      MimeType: { JSON: "application/json", JAVASCRIPT: "application/javascript", CSV: "text/csv", TEXT: "text/plain" },
      createTextOutput(text) {
        return { text, mime: null, setMimeType(m) { this.mime = m; return this; }, getContent() { return this.text; }, getMimeType() { return this.mime; } };
      }
    },
    Logger: { log() {} },
    console: { log() {}, error: console.error, warn: console.warn }
  };
  return { gas, state };
}

function loadCode(gas) {
  const ctx = vm.createContext(Object.assign({}, gas));
  const src = fs.readFileSync(path.join(ROOT, "apps-script", "Code.gs"), "utf8");
  vm.runInContext(src, ctx, { filename: "Code.gs" });
  return ctx;
}

/* ===================== test runner ===================== */
let passed = 0, failed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log("  ok   " + name); }
  catch (e) { failed++; console.log("  FAIL " + name + "\n       " + String(e && e.message || e).split("\n").join("\n       ")); }
}

const { gas, state } = createGas();
const ctx = loadCode(gas);
const T = { tester: "", admin: "", editor: "" }; // session tokens
// get/postRaw send the Tester session token unless tok is given (null = no token).
const get = (action, params = {}, tok = T.tester) => {
  const out = ctx.doGet({ parameter: Object.assign({ action }, tok ? { token: tok } : {}, params) });
  assert.strictEqual(out.getMimeType(), "application/json");
  return JSON.parse(out.getContent());
};
const postRaw = (payload, tok = T.tester) =>
  JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(Object.assign(tok ? { token: tok } : {}, payload)) } }).getContent());
const PW = "s3cret-Pallet";      // Tester's login password
const ADMIN_PW = "init-Admin-77"; // PALLET_INITIAL_ADMIN_PASSWORD
const EDITOR_PW = "editor-Pass-1";
const TESTER = "Tester (tester)", EDITOR = "Editor (editor)", ADMIN = "ผู้ดูแลระบบ (admin)";
// Writes: the client also sends a fake actor that the server must ignore.
const post = (action, body = {}, tok = T.tester) => postRaw(Object.assign({}, body, { action, actor: "Spoofed Name" }), tok);
const login = (username, password) => postRaw({ action: "login", username, password }, null);
const okData = (res) => { assert.strictEqual(res.ok, true, "expected ok, got: " + JSON.stringify(res)); return res.data; };
const err = (res, re) => {
  assert.strictEqual(res.ok, false, "expected error, got: " + JSON.stringify(res));
  if (re) assert.ok(re.test(res.error), "error '" + res.error + "' !~ " + re);
  return res;
};
const authErr = (res) => { err(res); assert.strictEqual(res.code, "AUTH"); assert.strictEqual(res.error, "กรุณาเข้าสู่ระบบ"); return res; };
// Forget login / reset-password failure counters ("15 minutes later") but keep sessions.
const clearFails = () => Object.keys(state.cache).forEach(k => { if (/FAILURES/.test(k)) delete state.cache[k]; });

const bkk = new Date(Date.now() + 7 * 3600 * 1000);
const p2 = n => String(n).padStart(2, "0");
const TODAY = `${bkk.getUTCFullYear()}-${p2(bkk.getUTCMonth() + 1)}-${p2(bkk.getUTCDate())}`;
const YMD = TODAY.slice(2).replace(/-/g, "");
const HOUR = p2(bkk.getUTCHours());
const AT = { date: TODAY, time: HOUR + ":00" };
const sheet = n => state.spreadsheets[state.props.PALLET_SPREADSHEET_ID].getSheetByName(n);
const dataRows = n => sheet(n).getLastRow() - 1;
const stockOf = (stock, type, size, st) => ((stock[type] || {})[size] || {})[st] || 0;
const lastLog = () => okData(get("logs")).items[0];
const userRows = () => { const d = sheet("users").data; return d.slice(1).filter(r => r.some(v => v !== "" && v != null)).map(r => Object.fromEntries(d[0].map((k, i) => [k, r[i]]))); };

const ALL_READS = ["bootstrap", "dashboard", "repairs", "history", "export", "logs", "logs_export", "me", "users"];
const ALL_POSTS = ["receive", "issue", "return", "damage", "repair_start", "repair_done", "scrap", "dept_save", "dept_delete",
  "reset_data", "movement_update", "movement_delete", "repair_update", "repair_delete", "verifyResetPassword",
  "logout", "change_password", "user_save", "user_reset_password", "user_toggle"];

console.log("Pallet Hub GAS backend tests (today " + TODAY + " Asia/Bangkok)");

/* ---------- setup ---------- */
test("API before setup: login reports setupSystem error, other actions need login first", () => {
  err(login("admin", "x"), /setupSystem/);
  authErr(get("bootstrap", {}, null));
});

test("setupSystem creates spreadsheet, sheets (incl. users), headers, seed data", () => {
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, true);
  assert.ok(state.props.PALLET_SPREADSHEET_ID);
  const ss = state.spreadsheets[state.props.PALLET_SPREADSHEET_ID];
  assert.strictEqual(ss.tz, "Asia/Bangkok");
  assert.deepStrictEqual(ss.getSheets().map(s => s.getName()), ["pallet_types", "departments", "repairs", "movements", "audit_logs", "users"]);
  assert.deepStrictEqual(sheet("movements").data[0], ["id", "doc_no", "action", "type_id", "size", "qty", "from_status", "to_status", "department", "person", "note", "repair_id", "moved_at", "created_at", "actor", "username"]);
  assert.deepStrictEqual(sheet("audit_logs").data[0], ["id", "category", "action", "ref", "detail", "actor", "ip", "created_at", "username"]);
  assert.deepStrictEqual(sheet("users").data[0], ["id", "username", "password_hash", "salt", "fullname", "role", "active", "created_at", "updated_at", "last_login"]);
  assert.strictEqual(sheet("users").getFrozenRows(), 1);
  assert.strictEqual(dataRows("pallet_types"), 4);
  assert.strictEqual(dataRows("departments"), 8);
  assert.strictEqual(dataRows("users"), 0);
  assert.strictEqual(r.users, 0);
  assert.strictEqual(r.initialAdminPasswordConfigured, false);
});

test("setupSystem is idempotent (no second spreadsheet, no duplicate seed)", () => {
  const id = state.props.PALLET_SPREADSHEET_ID;
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(state.props.PALLET_SPREADSHEET_ID, id);
  assert.strictEqual(state.created, 1);
  assert.strictEqual(dataRows("pallet_types"), 4);
  assert.strictEqual(dataRows("departments"), 8);
  assert.strictEqual(sheet("movements").data[0].length, 16);
});

/* ---------- authentication ---------- */
test("every action refuses without a token / with an unknown or malformed token (code AUTH), nothing written", () => {
  const before = JSON.stringify(state.spreadsheets[state.props.PALLET_SPREADSHEET_ID].sheets.map(s => s.data));
  const bad = ["", "abc", "f".repeat(64), "<script>", "F".repeat(64)];
  for (const a of ALL_READS) {
    authErr(get(a, {}, null));
    for (const t of bad) authErr(get(a, { token: t }, null));
    authErr(postRaw({ action: a }, null));
  }
  for (const a of ALL_POSTS) {
    authErr(postRaw({ action: a, resetPassword: "x", type_id: 1, size: "1.2x1.2", qty: 1, name: "X", id: 1 }, null));
    for (const t of bad) authErr(postRaw({ action: a, token: t }, null));
  }
  assert.strictEqual(JSON.stringify(state.spreadsheets[state.props.PALLET_SPREADSHEET_ID].sheets.map(s => s.data)), before);
});

test("unknown action / POST-only action via GET are rejected", () => {
  err(get("nope"), /^Unknown action$/);
  err(postRaw({ action: "nope" }), /^Unknown action$/);
  for (const a of ["receive", "login", "reset_data", "user_save", "logout", "verifyResetPassword"]) err(get(a), /POST/);
  err(get("verifyPassword"), /^Unknown action$/); // the old shared-password check is gone
  assert.strictEqual(okData(get(undefined)).service, "Pallet Hub API");
});

test("invalid JSON body is rejected", () => {
  const r = JSON.parse(ctx.doPost({ postData: { contents: "{bad" } }).getContent());
  err(r, /JSON/);
});

test("initial admin: refused with a Thai message while PALLET_INITIAL_ADMIN_PASSWORD is unset", () => {
  const r = err(login("admin", "anything"), /PALLET_INITIAL_ADMIN_PASSWORD/);
  assert.ok(/ยังไม่มีบัญชีผู้ใช้/.test(r.error));
  assert.strictEqual(dataRows("users"), 0);
  err(login("", ""), /กรุณากรอกชื่อผู้ใช้และรหัสผ่าน/);
});

test("initial admin: wrong password or another username creates nothing, costs 1 s and is logged", () => {
  state.props.PALLET_INITIAL_ADMIN_PASSWORD = ADMIN_PW;
  const sleeps = state.sleeps;
  const r = err(login("admin", "wrong"), /^ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง$/);
  assert.strictEqual(r.code, "LOGIN_FAILED");
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(login("root", ADMIN_PW), /^ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง$/); // only "admin" can be bootstrapped
  assert.strictEqual(dataRows("users"), 0);
  const logs = sheet("audit_logs").data.slice(1);
  assert.strictEqual(logs.length, 2);
  const h = sheet("audit_logs").data[0];
  const row = Object.fromEntries(h.map((k, i) => [k, logs[0][i]]));
  assert.strictEqual(row.category, "account");
  assert.strictEqual(row.action, "login_failed");
  assert.strictEqual(row.username, "admin");
  assert.strictEqual(row.actor, "");
  assert.ok(!JSON.stringify(sheet("audit_logs").data).includes("wrong")); // passwords never logged
  clearFails();
});

test("initial admin: right property password creates the admin (hashed) and logs in", () => {
  const d = okData(login(" Admin ", ADMIN_PW));
  assert.ok(/^[a-f0-9]{64}$/.test(d.token));
  assert.strictEqual(d.expires_in, 21600);
  assert.strictEqual(d.initial, true);
  assert.deepStrictEqual([d.user.username, d.user.fullname, d.user.role, d.user.active], ["admin", "ผู้ดูแลระบบ", "admin", true]);
  assert.ok(!("password_hash" in d.user) && !("salt" in d.user));
  T.admin = d.token;
  const u = userRows();
  assert.strictEqual(u.length, 1);
  assert.strictEqual(u[0].role, "admin");
  assert.strictEqual(u[0].active, 1);
  assert.ok(/^[a-f0-9]{64}$/.test(u[0].password_hash) && /^[a-f0-9]{32}$/.test(u[0].salt));
  assert.ok(!JSON.stringify(sheet("users").data).includes(ADMIN_PW));
  assert.strictEqual(state.cache["PALLET_SESSION_" + d.token] !== undefined, true);
  const me = okData(get("me", {}, T.admin)).user;
  assert.strictEqual(me.username, "admin");
  const logs = okData(get("logs", { cat: "account" }, T.admin)).items;
  assert.deepStrictEqual(logs.slice(0, 2).map(l => [l.action, l.actor, l.username]), [["login", ADMIN, "admin"], ["user_create", ADMIN, "admin"]]);
  // a second "initial" login does not create another admin
  T.admin2 = okData(login("admin", ADMIN_PW)).token;
  assert.strictEqual(dataRows("users"), 1);
});

test("admin creates accounts; username/password/role validation; usernames unique case-insensitively", () => {
  const save = b => post("user_save", b, T.admin);
  err(save({ username: "ab", fullname: "X", password: "12345678" }), /3-30/);
  err(save({ username: "bad name", fullname: "X", password: "12345678" }), /3-30/);
  err(save({ username: "x".repeat(31), fullname: "X", password: "12345678" }), /3-30/);
  err(save({ username: "ok.user", fullname: "X", password: "1234567" }), /อย่างน้อย 8/);
  err(save({ username: "ok.user", fullname: "  ", password: "12345678" }), /ชื่อ-นามสกุล/);
  err(save({ username: "ok.user", fullname: "X", password: "12345678", role: "root" }), /บทบาท/);
  err(save({ username: "ADMIN", fullname: "X", password: "12345678" }), /มีอยู่แล้ว/);
  const d = okData(save({ username: "Tester", fullname: "Tester", password: PW }));
  assert.deepStrictEqual([d.user.username, d.user.role, d.user.active], ["tester", "user", true]);
  okData(save({ username: "editor", fullname: "Editor", password: EDITOR_PW }));
  err(save({ username: "TESTER", fullname: "Dup", password: "12345678" }), /มีอยู่แล้ว/);
  const list = okData(get("users", {}, T.admin));
  assert.deepStrictEqual(list.users.map(u => u.username), ["admin", "tester", "editor"]);
  assert.ok(list.users.every(u => !("password_hash" in u) && !("salt" in u)));
  const created = okData(get("logs", { cat: "account" }, T.admin)).items.filter(l => l.action === "user_create");
  assert.ok(created.some(l => l.ref === "tester" && l.actor === ADMIN && l.detail.includes("tester (Tester)")));
});

test("login: right password returns token + user, updates last_login, logged in audit", () => {
  const d = okData(login("TESTER", PW));
  T.tester = d.token;
  assert.deepStrictEqual(Object.keys(d.user).sort(), ["active", "created_at", "fullname", "id", "last_login", "role", "updated_at", "username"]);
  assert.strictEqual(d.user.username, "tester");
  assert.strictEqual(d.initial, false);
  assert.strictEqual(userRows().find(u => u.username === "tester").last_login.slice(0, 10), TODAY);
  T.editor = okData(login("editor", EDITOR_PW)).token;
  assert.notStrictEqual(T.tester, T.editor);
  assert.strictEqual(okData(get("me")).user.fullname, "Tester");
  const l = okData(get("logs", { cat: "account" }, T.admin)).items[0];
  assert.deepStrictEqual([l.action, l.actor, l.username], ["login", EDITOR, "editor"]);
});

test("login: wrong password refused (1 s), audit row; 10 failures lock the username for 15 min", () => {
  const sleeps = state.sleeps;
  err(login("editor", "nope"), /^ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง$/);
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(login("nobody", "nope"), /^ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง$/); // unknown user: same message
  const l = okData(get("logs", { cat: "account" }, T.admin)).items[0];
  assert.deepStrictEqual([l.action, l.username], ["login_failed", "nobody"]);
  state.cachePuts = [];
  for (let i = 1; i < 10; i++) err(login("Editor", "guess" + i), /ไม่ถูกต้อง/);
  assert.ok(state.cachePuts.every(t => t === 900));
  const r = err(login("editor", EDITOR_PW), /ระงับชั่วคราว 15 นาที/);
  assert.strictEqual(r.code, "LOCKED");
  okData(login("tester", PW)); // other usernames are not affected
  assert.ok(okData(get("logs", { cat: "account", q: "ระงับ" }, T.admin)).items.length >= 1);
  okData(get("bootstrap", {}, T.editor)); // existing sessions keep working
  clearFails(); // 15 minutes later
  okData(login("editor", EDITOR_PW));
});

test("normal users cannot call admin actions (FORBIDDEN), nothing changes", () => {
  const before = JSON.stringify(sheet("users").data);
  for (const [a, b] of [["users", {}], ["user_save", { username: "evil", fullname: "E", password: "12345678", role: "admin" }],
    ["user_save", { id: 2, fullname: "T", role: "admin" }], ["user_reset_password", { id: 1, password: "12345678" }], ["user_toggle", { id: 1 }]]) {
    const r = err(post(a, b), /เฉพาะผู้ดูแลระบบ/);
    assert.strictEqual(r.code, "FORBIDDEN");
  }
  const r = err(get("users"), /เฉพาะผู้ดูแลระบบ/);
  assert.strictEqual(r.code, "FORBIDDEN");
  assert.strictEqual(JSON.stringify(sheet("users").data), before);
});

/* ---------- receive ---------- */
test("receive: creates RC doc, stock available, server records the logged-in user (client actor ignored)", () => {
  const d = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 50, person: "  สมชาย ใจดี  ", note: "PO-123", username: "admin" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0001");
  assert.strictEqual(d.id, 1);
  assert.strictEqual(d.message, "รับเข้า 50 ตัว เรียบร้อย");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 50);
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.deepStrictEqual([m.actor, m.username, m.person], [TESTER, "tester", "สมชาย ใจดี"]);
  const log = lastLog();
  assert.strictEqual(log.category, "pallet");
  assert.strictEqual(log.action, "receive");
  assert.strictEqual(log.ref, d.doc_no);
  assert.strictEqual(log.actor, TESTER);
  assert.strictEqual(log.username, "tester");
  assert.strictEqual(log.ip, "web");
  assert.strictEqual(log.detail, `รับเข้า RM (พาเลทสำหรับใส่ RM) ขนาด 1.2x1.2 ม. จำนวน 50 ตัว [ภายนอก → พร้อมใช้] · เวลาทำรายการ ${TODAY.slice(8, 10)}/${TODAY.slice(5, 7)}/${TODAY.slice(0, 4)} ${HOUR}:00 · ชื่อที่ระบุ: สมชาย ใจดี · PO-123`);
});

test("receive: second doc number increments; actor is the session user even without a person", () => {
  const d = okData(post("receive", Object.assign({ type_id: 2, size: "1.1x1.1", qty: "30", person: "" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0002");
  assert.strictEqual(lastLog().actor, TESTER);
  assert.ok(!JSON.stringify(sheet("movements").data).includes("Spoofed"));
  assert.ok(!JSON.stringify(sheet("audit_logs").data).includes("Spoofed"));
});

test("receive: validation errors are rejected and logged as warn", () => {
  const n = dataRows("movements");
  err(post("receive", Object.assign({ type_id: 99, size: "1.2x1.2", qty: 1 }, AT)), /^ไม่พบประเภทพาเลท$/);
  err(post("receive", Object.assign({ type_id: 1, size: "1.1x1.1", qty: 1 }, AT)), /^กรุณาเลือกขนาดพาเลท$/);
  err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 0 }, AT)), /^จำนวนต้องมากกว่า 0$/);
  err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: -3 }, AT)), /^จำนวนต้องมากกว่า 0$/);
  err(post("receive", { type_id: 1, size: "1.2x1.2", qty: 1, date: "2026/01/01", time: "10:00" }), /^วันที่\/เวลาไม่ถูกต้อง$/);
  err(post("receive", { type_id: 1, size: "1.2x1.2", qty: 1, date: "", time: "10:00" }), /^วันที่\/เวลาไม่ถูกต้อง$/);
  err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1, person: "x".repeat(101) }, AT)), /ยาวเกิน 100/);
  assert.strictEqual(dataRows("movements"), n);
  const log = lastLog();
  assert.strictEqual(log.category, "warn");
  assert.strictEqual(log.action, "receive");
  assert.strictEqual(log.ref, "");
  const warns = okData(get("logs", { cat: "warn" })).items;
  assert.strictEqual(warns.length, 7);
  assert.ok(warns.some(w => w.detail === "ปฏิเสธรับเข้า: จำนวนต้องมากกว่า 0"));
});

test("receive: default date/time = now (Bangkok), overflow dates normalised like PHP", () => {
  const d = okData(post("receive", { type_id: 3, size: "1.1x1.1", qty: 1 }));
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.strictEqual(m.moved_at.slice(0, 10), TODAY);
  assert.ok(/:00$/.test(m.moved_at));
  const d2 = okData(post("receive", { type_id: 3, size: "1.1x1.1", qty: 1, date: "2025-02-30", time: "08:15:59" }));
  const m2 = okData(get("history", { q: d2.doc_no })).items[0];
  assert.strictEqual(m2.moved_at, "2025-03-02 08:15:00");
});

/* ---------- issue ---------- */
test("issue: requires department, enough stock; moves available -> issued", () => {
  err(post("issue", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 5, department: "  " }, AT)), /^กรุณาเลือกฝ่ายที่เบิก$/);
  err(post("issue", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 60, department: "ฝ่ายผลิต" }, AT)), /^พาเลทพร้อมใช้ไม่พอ \(คงเหลือ 50 ตัว\)$/);
  const d = okData(post("issue", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 20, department: "ฝ่ายผลิต", person: "A" }, AT)));
  assert.strictEqual(d.doc_no, "IS-" + YMD + "-0001");
  assert.strictEqual(d.message, "เบิกจ่ายให้ ฝ่ายผลิต 20 ตัว เรียบร้อย");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 30);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 20);
  assert.deepStrictEqual(b.dept, [{ department: "ฝ่ายผลิต", type_id: 1, size: "1.2x1.2", qty: 20 }]);
  assert.ok(lastLog().detail.includes("[พร้อมใช้ → เบิกไปใช้งาน] ฝ่าย: ฝ่ายผลิต"));
});

/* ---------- return ---------- */
test("return: dept required, cannot exceed what the dept holds", () => {
  err(post("return", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1 }, AT)), /^กรุณาเลือกฝ่ายที่คืน$/);
  err(post("return", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 25, department: "ฝ่ายผลิต" }, AT)), /^ฝ่ายผลิต ถือพาเลทนี้อยู่ 20 ตัว$/);
  err(post("return", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายบรรจุ" }, AT)), /^ฝ่ายบรรจุ ถือพาเลทนี้อยู่ 0 ตัว$/);
});

test("return good: issued -> available", () => {
  const d = okData(post("return", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 5, department: "ฝ่ายผลิต", condition: "good" }, AT)));
  assert.strictEqual(d.doc_no, "RT-" + YMD + "-0001");
  assert.strictEqual(d.message, "รับคืนจาก ฝ่ายผลิต 5 ตัว เรียบร้อย");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 35);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 15);
});

let ticket1, ticket2;
test("return damaged: opens repair ticket, issued -> damaged, DM doc", () => {
  const d = okData(post("return", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 3, department: "ฝ่ายผลิต", condition: "damaged", cause: "ไม้หัก", person: "B", note: "n1" }, AT)));
  assert.strictEqual(d.doc_no, "DM-" + YMD + "-0001");
  assert.strictEqual(d.message, "รับคืนชำรุด 3 ตัว — เปิดใบแจ้งซ่อมแล้ว");
  const items = okData(get("repairs")).items;
  assert.strictEqual(items.length, 1);
  ticket1 = items[0];
  assert.strictEqual(ticket1.ticket_no, "RPR-" + YMD + "-0001");
  assert.strictEqual(ticket1.stage, "damaged");
  assert.strictEqual(ticket1.source, "issued");
  assert.strictEqual(ticket1.department, "ฝ่ายผลิต");
  assert.strictEqual(ticket1.cause, "ไม้หัก");
  assert.strictEqual(ticket1.reported_by, TESTER); // the logged-in user, not the typed person "B"
  assert.strictEqual(ticket1.reported_username, "tester");
  assert.strictEqual(ticket1.note, "n1");
  assert.strictEqual(ticket1.started_at, null);
  assert.strictEqual(ticket1.finished_at, null);
  assert.strictEqual(ticket1.code, "RM");
  assert.strictEqual(ticket1.type_name, "พาเลทสำหรับใส่ RM");
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.strictEqual(m.note, "คืนสภาพชำรุด: ไม้หัก");
  assert.strictEqual(m.repair_id, ticket1.id);
  assert.strictEqual(m.from_status, "issued");
  assert.strictEqual(m.to_status, "damaged");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 12);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "damaged"), 3);
  assert.deepStrictEqual(b.dept, [{ department: "ฝ่ายผลิต", type_id: 1, size: "1.2x1.2", qty: 12 }]);
  assert.strictEqual(lastLog().category, "repair");
});

/* ---------- damage ---------- */
test("damage: available -> damaged with ticket; over-stock rejected", () => {
  err(post("damage", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 999 }, AT)), /^พาเลทพร้อมใช้มีเพียง 35 ตัว$/);
  const d = okData(post("damage", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 2, cause: "ตะปูหลุด", person: "C" }, AT)));
  assert.strictEqual(d.doc_no, "DM-" + YMD + "-0002");
  assert.strictEqual(d.message, "แจ้งชำรุด 2 ตัว เรียบร้อย");
  const items = okData(get("repairs")).items;
  assert.deepStrictEqual(items.map(i => i.id), [2, 1]); // id DESC
  ticket2 = items[0];
  assert.strictEqual(ticket2.ticket_no, "RPR-" + YMD + "-0002");
  assert.strictEqual(ticket2.source, "available");
  assert.strictEqual(ticket2.department, null);
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.strictEqual(m.department, null);
  assert.strictEqual(m.note, "ตะปูหลุด");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 33);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "damaged"), 5);
});

/* ---------- repair workflow ---------- */
test("invalid transitions are rejected and logged (repair_done/scrap on wrong stage, missing ticket)", () => {
  err(post("repair_done", Object.assign({ id: ticket1.id }, AT)), /^ใบนี้ไม่ได้อยู่ระหว่างซ่อม$/);
  assert.strictEqual(lastLog().detail, "ปฏิเสธซ่อมเสร็จ: ใบนี้ไม่ได้อยู่ระหว่างซ่อม");
  err(post("repair_start", Object.assign({ id: 999 }, AT)), /^ไม่พบใบแจ้งซ่อม$/);
  err(post("scrap", Object.assign({ id: "abc" }, AT)), /^ไม่พบใบแจ้งซ่อม$/);
});

test("repair_start: damaged -> repairing (repairer, started_at), cannot start twice", () => {
  const d = okData(post("repair_start", { id: String(ticket1.id), date: TODAY, time: HOUR + ":00", person: "ช่างหนึ่ง" }));
  assert.strictEqual(d.doc_no, "RP-" + YMD + "-0001");
  assert.strictEqual(d.message, "ส่งซ่อม " + ticket1.ticket_no + " แล้ว");
  const t = okData(get("repairs")).items.find(i => i.id === ticket1.id);
  assert.strictEqual(t.stage, "repairing");
  assert.strictEqual(t.repairer, "ช่างหนึ่ง");
  assert.strictEqual(t.started_at, `${TODAY} ${HOUR}:00:00`);
  assert.deepStrictEqual([t.updated_by, t.updated_username], [TESTER, "tester"]);
  assert.strictEqual(okData(get("history", { q: d.doc_no })).items[0].actor, TESTER);
  err(post("repair_start", Object.assign({ id: ticket1.id }, AT)), /^ใบนี้ไม่ได้อยู่สถานะชำรุด$/);
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "damaged"), 2);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "repairing"), 3);
});

test("repair_done: repairing -> available, note appended, person falls back to repairer", () => {
  const d = okData(post("repair_done", { id: ticket1.id, date: TODAY, time: HOUR + ":30", person: "  ", note: "เปลี่ยนไม้" }));
  assert.strictEqual(d.doc_no, "RD-" + YMD + "-0001");
  assert.strictEqual(d.message, "ซ่อมเสร็จ 3 ตัว กลับเข้าคลังพร้อมใช้");
  const t = okData(get("repairs")).items.find(i => i.id === ticket1.id);
  assert.strictEqual(t.stage, "done");
  assert.strictEqual(t.finished_at, `${TODAY} ${HOUR}:30:00`);
  assert.strictEqual(t.note, "n1\nซ่อมเสร็จ: เปลี่ยนไม้");
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.strictEqual(m.person, "ช่างหนึ่ง");
  assert.strictEqual(m.note, ticket1.ticket_no);
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 36);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "repairing"), 0);
});

test("scrap: done ticket rejected; damaged ticket -> scrapped", () => {
  err(post("scrap", Object.assign({ id: ticket1.id }, AT)), /^ไม่สามารถตัดจำหน่ายใบนี้ได้$/);
  const d = okData(post("scrap", Object.assign({ id: ticket2.id, note: "แตกหมด" }, AT)));
  assert.strictEqual(d.doc_no, "SC-" + YMD + "-0001");
  assert.strictEqual(d.message, "ตัดจำหน่าย 2 ตัว แล้ว");
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.strictEqual(m.from_status, "damaged");
  assert.strictEqual(m.note, ticket2.ticket_no + " แตกหมด");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "damaged"), 0);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "scrapped"), 2);
  err(post("scrap", Object.assign({ id: ticket2.id }, AT)), /^ไม่สามารถตัดจำหน่ายใบนี้ได้$/);
});

test("scrap from repairing stage records from_status=repairing", () => {
  okData(post("damage", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 1 }, AT)));
  const t = okData(get("repairs")).items[0];
  okData(post("repair_start", Object.assign({ id: t.id }, AT)));
  const d = okData(post("scrap", Object.assign({ id: t.id }, AT)));
  assert.strictEqual(okData(get("history", { q: d.doc_no })).items[0].from_status, "repairing");
});

/* ---------- dashboard ---------- */
test("dashboard: structure, day buckets, today, hours, topDept, repairStat, logsToday", () => {
  for (const [days, n] of [[7, 7], [14, 14], [30, 30], [99, 7], ["x", 7]]) {
    const d = okData(get("dashboard", { days }));
    assert.strictEqual(Object.keys(d.days).length, n, "days=" + days);
    assert.strictEqual(Object.keys(d.days).pop(), TODAY);
  }
  const d = okData(get("dashboard", { days: 7 }));
  assert.deepStrictEqual(Object.keys(d.days[TODAY]), ["receive", "issue", "return", "damage", "repair_done"]);
  // receives today: 50 + 30 + 1 (default now) + 0 (the 2025 one is out of range)
  assert.strictEqual(d.days[TODAY].receive, 81);
  assert.strictEqual(d.days[TODAY].issue, 20);
  assert.strictEqual(d.days[TODAY].return, 5);
  assert.strictEqual(d.days[TODAY].damage, 6);
  assert.strictEqual(d.days[TODAY].repair_done, 3);
  const today = Object.fromEntries(d.today.map(r => [r.action, r]));
  assert.strictEqual(today.receive.q, 81);
  assert.strictEqual(today.receive.n, 3);
  assert.strictEqual(today.scrap.q, 3);
  assert.strictEqual(d.hours.length, 24);
  const h = +HOUR;
  assert.ok(d.hours[h].in >= 50 + 30 + 5 + 3);
  assert.strictEqual(d.hours[h].out, 20 + 6);
  assert.deepStrictEqual(d.topDept, [{ department: "ฝ่ายผลิต", q: 20, n: 1 }]);
  assert.ok(d.recent.length <= 12 && d.recent.length > 0);
  assert.ok(d.recent[0].code && d.recent[0].color && d.recent[0].tkey);
  for (let i = 1; i < d.recent.length; i++) {
    const a = d.recent[i - 1], b = d.recent[i];
    assert.ok(a.moved_at > b.moved_at || (a.moved_at === b.moved_at && a.id > b.id), "recent order");
  }
  assert.strictEqual(d.repairStat.avg_h, 0.5);
  assert.strictEqual(d.repairStat.open_n, 0);
  assert.strictEqual(d.repairStat.done_m, 3);
  assert.strictEqual(d.repairStat.scrap_m, 3);
  assert.strictEqual(d.repairStat.oldest, null);
  const stages = Object.fromEntries(d.repairs.map(r => [r.stage, r]));
  assert.deepStrictEqual(stages.done, { stage: "done", n: 1, q: 3 });
  assert.deepStrictEqual(stages.scrapped, { stage: "scrapped", n: 2, q: 3 });
  assert.strictEqual(d.logsToday, dataRows("audit_logs"));
  assert.deepStrictEqual(d.stock, okData(get("bootstrap")).stock);
});

test("dashboard repairStat: open tickets and oldest", () => {
  okData(post("damage", { type_id: 1, size: "1.2x1.2", qty: 1, date: "2025-01-01", time: "07:00" }));
  const rs = okData(get("dashboard")).repairStat;
  assert.strictEqual(rs.open_n, 1);
  assert.strictEqual(rs.oldest, "2025-01-01 07:00:00");
});

/* ---------- history / export ---------- */
test("history filters: act, dept, type, q (LIKE), date range, order", () => {
  const all = okData(get("history")).items;
  assert.strictEqual(all.length, dataRows("movements"));
  assert.ok(all[0].type_name && all[0].code);
  assert.strictEqual(okData(get("history", { act: "issue" })).items.length, 1);
  assert.strictEqual(okData(get("history", { dept: "ฝ่ายผลิต" })).items.length, 3);
  assert.strictEqual(okData(get("history", { type: "2" })).items.length, 4);
  assert.strictEqual(okData(get("history", { type: "0" })).items.length, all.length); // PHP empty("0")
  assert.strictEqual(okData(get("history", { q: "rc-" })).items.length, 4); // case-insensitive LIKE
  assert.strictEqual(okData(get("history", { q: "PO_123" })).items.length, 1); // _ wildcard
  assert.strictEqual(okData(get("history", { q: "สมชาย" })).items.length, 1);
  assert.strictEqual(okData(get("history", { from: TODAY, to: TODAY })).items.length, all.length - 2);
  assert.strictEqual(okData(get("history", { from: "2025-03-01", to: "2025-03-31" })).items.length, 1);
  assert.strictEqual(okData(get("history", { from: "2099-01-01" })).items.length, 0);
});

test("export: CSV with PHP fputcsv quoting", () => {
  const d = okData(get("export", { act: "receive" }));
  assert.ok(/^pallet_history_\d{8}_\d{6}\.csv$/.test(d.filename));
  const lines = d.csv.split("\n");
  assert.strictEqual(lines[0], "เลขที่เอกสาร,วันที่,เวลา,รายการ,รหัส,ประเภท,ขนาด,จำนวน,ฝ่าย,\"ผู้ทำรายการ (บัญชี)\",ชื่อผู้ใช้,ชื่อที่ระบุ,หมายเหตุ");
  assert.strictEqual(lines.length, 1 + 4 + 1);
  const row = lines.find(l => l.startsWith("RC-" + YMD + "-0001"));
  assert.strictEqual(row, `RC-${YMD}-0001,${TODAY},${HOUR}:00,รับเข้า,RM,"พาเลทสำหรับใส่ RM",1.2x1.2,50,,"Tester (tester)",tester,"สมชาย ใจดี",PO-123`);
  assert.strictEqual(ctx.csvLine_(['a"b', "c\\\"d", null, 5, "x,y"]), '"a""b","c\\"d",,5,"x,y"\n');
});

/* ---------- logs ---------- */
test("logs: categories, filters, order and CSV export", () => {
  const all = okData(get("logs")).items;
  assert.strictEqual(all.length, dataRows("audit_logs"));
  for (let i = 1; i < all.length; i++) assert.ok(all[i - 1].id > all[i].id);
  const cats = new Set(all.map(i => i.category));
  ["pallet", "repair", "warn"].forEach(c => assert.ok(cats.has(c), c));
  assert.ok(okData(get("logs", { q: "Tester" })).items.length > 0);
  assert.ok(okData(get("logs")).items.every(l => l.category === "account" || (l.actor === TESTER && l.username === "tester")));
  assert.strictEqual(okData(get("logs", { from: "2099-01-01" })).items.length, 0);
  const x = okData(get("logs_export", { cat: "warn" }));
  assert.ok(/^pallet_log_\d{8}_\d{6}\.csv$/.test(x.filename));
  const lines = x.csv.trim().split("\n");
  assert.strictEqual(lines[0], "ลำดับ,วันที่,เวลา,หมวด,เลขที่อ้างอิง,รายละเอียด,ผู้ทำรายการ,ชื่อผู้ใช้,IP");
  assert.ok(lines.slice(1).every(l => l.includes(',"Tester (tester)",tester,web')));
  assert.ok(lines.slice(1).every(l => l.includes(",ถูกปฏิเสธ,") && l.endsWith(",web")));
});

/* ---------- departments ---------- */
test("dept_save: add new, update existing (case-insensitive), validation", () => {
  err(post("dept_save", { name: "  " }), /^กรุณาระบุชื่อฝ่าย$/);
  const warnBefore = okData(get("logs", { cat: "warn" })).items.length;
  assert.strictEqual(okData(get("logs", { cat: "warn" })).items.length, warnBefore); // settings errors not logged as warn
  const r = okData(post("dept_save", { name: "ฝ่ายจัดซื้อ", icon: "fa-store", color: "#1E6FE0" }));
  assert.strictEqual(r.message, "เพิ่ม ฝ่ายจัดซื้อ แล้ว");
  let deps = okData(get("bootstrap")).departments;
  assert.strictEqual(deps.length, 9);
  assert.deepStrictEqual(deps[8], { id: 9, name: "ฝ่ายจัดซื้อ", icon: "fa-store", color: "#1E6FE0", active: 1 });
  const log = lastLog();
  assert.strictEqual(log.category, "setting");
  assert.strictEqual(log.detail, "เพิ่ม/แก้ไขฝ่าย: ฝ่ายจัดซื้อ");
  assert.strictEqual(log.actor, TESTER);
  okData(post("dept_save", { name: "Office" }));
  okData(post("dept_save", { name: "OFFICE", icon: "", color: "#0EA5E9" }));
  deps = okData(get("bootstrap")).departments;
  const office = deps.filter(d => d.name.toLowerCase() === "office");
  assert.strictEqual(office.length, 1);
  assert.deepStrictEqual([office[0].name, office[0].icon, office[0].color], ["Office", "fa-building", "#0EA5E9"]);
});

test("dept_delete: deactivates, history kept, re-save reactivates", () => {
  err(post("dept_delete", { id: 999 }), /^ไม่พบฝ่าย$/);
  const r = okData(post("dept_delete", { id: "1" }));
  assert.strictEqual(r.message, "ลบฝ่ายแล้ว");
  assert.strictEqual(lastLog().detail, "ลบฝ่าย: ฝ่ายผลิต");
  const b = okData(get("bootstrap"));
  assert.ok(!b.departments.some(d => d.id === 1));
  assert.strictEqual(b.dept[0].department, "ฝ่ายผลิต"); // outstanding pallets still tracked
  okData(post("dept_save", { name: "ฝ่ายผลิต", icon: "fa-industry", color: "#E2231A" }));
  assert.ok(okData(get("bootstrap")).departments.some(d => d.id === 1));
  okData(post("dept_delete", { id: 1 })); // deleting an inactive dept still works (PHP selects by id only)
  okData(post("dept_delete", { id: 1 }));
});

/* ---------- storage details ---------- */
test("formula-looking text is stored safely and read back unchanged", () => {
  const d = okData(post("receive", Object.assign({ type_id: 4, size: "1.2x1.2", qty: 1, note: "=HYPERLINK(\"x\")" }, AT)));
  assert.strictEqual(okData(get("history", { q: d.doc_no })).items[0].note, "=HYPERLINK(\"x\")");
});

test("Date objects in sheet cells are read back as Bangkok 'YYYY-MM-DD HH:MM:SS'", () => {
  const s = sheet("movements");
  const col = s.data[0].indexOf("moved_at");
  const row = s.data.findIndex(r => r[1] === "RC-" + YMD + "-0001");
  const orig = s.data[row][col];
  s.data[row][col] = new Date(Date.UTC(2026, 0, 2, 3, 4, 5) - 7 * 3600 * 1000);
  assert.strictEqual(okData(get("history", { q: "RC-" + YMD + "-0001" })).items[0].moved_at, "2026-01-02 03:04:05");
  s.data[row][col] = orig;
});

test("rows appended in batches with number formats; sheet grows past maxRows", () => {
  const s = sheet("audit_logs");
  s.maxRows = s.getLastRow(); // full sheet
  const before = state.formatCalls;
  okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1 }, AT)));
  assert.ok(s.maxRows > s.getLastRow() - 1);
  assert.strictEqual(state.formatCalls - before, 2); // movements + audit_logs
});

test("ids are max+1 and stable after manual row deletion", () => {
  const s = sheet("departments");
  const n = s.data.length;
  okData(post("dept_save", { name: "Temp" }));
  const id = okData(get("bootstrap")).departments.find(d => d.name === "Temp").id;
  assert.strictEqual(id, n); // header + n-1 rows -> next id n
});

/* ---------- lock & lockout ---------- */
test("every write acquires and releases the script lock; lock timeout returns error", () => {
  assert.ok(state.locks > 0);
  assert.strictEqual(state.locks, state.unlocks);
  state.lockFail = true;
  err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1 }, AT)), /ลองใหม่/);
  state.lockFail = false;
});

test("reads work with a session (GET and POST)", () => {
  for (const a of ["bootstrap", "dashboard", "repairs", "history", "export", "logs", "logs_export", "me"]) {
    okData(get(a));
    okData(postRaw({ action: a }));
  }
});

/* ---------- reset data ---------- */
const RPW = "reset-Only-9";
const reset = (pw, extra = {}, tok = T.admin) => postRaw(Object.assign({ action: "reset_data", resetPassword: pw, actor: "Spoofed" }, extra), tok);
const snapshot = () => ["movements", "repairs", "audit_logs", "pallet_types", "departments"].map(dataRows);

test("reset_data refused while PALLET_RESET_PASSWORD is unset (Thai admin message); GET refused", () => {
  assert.ok(!("PALLET_RESET_PASSWORD" in state.props));
  const before = snapshot();
  const r = err(reset(PW, { password: PW }), /PALLET_RESET_PASSWORD/);
  assert.ok(/ยังไม่ได้ตั้งรหัสรีเซ็ตข้อมูล/.test(r.error));
  assert.strictEqual(r.passwordError, true);
  err(get("reset_data"), /POST/);
  assert.deepStrictEqual(snapshot(), before);
});

test("reset_data with wrong password (or the action password) is refused, slept, own lockout key", () => {
  state.props.PALLET_RESET_PASSWORD = RPW;
  clearFails();
  const before = snapshot();
  const sleeps = state.sleeps;
  const r = err(reset("wrong"), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(reset(PW, { password: PW }), /รหัสไม่ถูกต้อง/); // a login password is not the reset password
  err(postRaw({ action: "reset_data", password: RPW }), /รหัสไม่ถูกต้อง/); // must come in resetPassword
  assert.strictEqual(state.cache.PALLET_RESET_FAILURES, "3");
  assert.deepStrictEqual(snapshot(), before);
  clearFails();
});

test("reset_data lockout: >10 wrong in 15 min refuses even the right reset password; normal writes unaffected", () => {
  clearFails();
  state.cachePuts = [];
  for (let i = 0; i < 10; i++) err(reset("guess" + i), /รหัสไม่ถูกต้อง/);
  assert.ok(state.cachePuts.every(t => t === 900));
  const before = snapshot();
  const r = err(reset(RPW), /หลายครั้งเกินไป/);
  assert.strictEqual(r.passwordError, true);
  assert.deepStrictEqual(snapshot(), before);
  okData(get("bootstrap")); // normal use unaffected
  clearFails(); // 15 minutes later
});

let resetResult;
test("reset_data success: clears movements/repairs/audit_logs (one clear per sheet), keeps types & departments, no action password needed", () => {
  const [mv, rp, lg, ty, dp] = snapshot();
  assert.ok(mv > 0 && rp > 0 && lg > 0);
  const before = okData(get("bootstrap"));
  const headers = ["movements", "repairs", "audit_logs"].map(n => sheet(n).data[0].slice());
  const locks = state.locks, clears = state.clearCalls || 0;
  resetResult = okData(reset(RPW)); // payload has no "password" field
  assert.deepStrictEqual(resetResult.removed, { movements: mv, repairs: rp, audit_logs: lg });
  assert.ok(/รีเซ็ตข้อมูลแล้ว/.test(resetResult.message));
  assert.strictEqual(state.clearCalls - clears, 3);
  assert.strictEqual(state.locks - locks, 1);
  assert.strictEqual(state.locks, state.unlocks);
  assert.deepStrictEqual(snapshot(), [0, 0, 1, ty, dp]); // audit_logs holds only the reset entry
  assert.deepStrictEqual(["movements", "repairs", "audit_logs"].map(n => sheet(n).data[0]), headers);
  const b = okData(get("bootstrap"));
  assert.deepStrictEqual(b.departments, before.departments);
  assert.deepStrictEqual(b.types, before.types);
});

test("after reset: audit log contains exactly the reset entry", () => {
  const items = okData(get("logs")).items;
  assert.strictEqual(items.length, 1);
  const l = items[0];
  assert.strictEqual(l.id, 1);
  assert.strictEqual(l.category, "setting");
  assert.strictEqual(l.action, "รีเซ็ตข้อมูล");
  assert.strictEqual(l.actor, ADMIN);
  assert.strictEqual(l.username, "admin");
  assert.strictEqual(l.ref, "");
  ["movements", "repairs", "audit_logs"].forEach(n => assert.ok(l.detail.includes(n), n));
  assert.ok(l.detail.includes("(" + resetResult.removed.movements + " แถว)"));
  assert.strictEqual(sheet("audit_logs").getLastRow(), 2); // written right under the header
});

test("after reset: stock, dept, repairs, history and dashboard are all zero/empty", () => {
  const b = okData(get("bootstrap"));
  assert.deepStrictEqual(b.stock, {});
  assert.deepStrictEqual(b.dept, []);
  assert.deepStrictEqual(okData(get("repairs")).items, []);
  assert.deepStrictEqual(okData(get("history")).items, []);
  const d = okData(get("dashboard"));
  assert.deepStrictEqual(d.stock, {});
  assert.deepStrictEqual(d.dept, []);
  assert.deepStrictEqual(d.today, []);
  assert.deepStrictEqual(d.recent, []);
  assert.deepStrictEqual(d.repairs, []);
  assert.deepStrictEqual(d.topDept, []);
  assert.ok(Object.values(d.days).every(v => Object.values(v).every(n => n === 0)));
  assert.ok(d.hours.every(h => h.in === 0 && h.out === 0));
  assert.deepStrictEqual(d.repairStat, { avg_h: null, open_n: 0, done_m: 0, scrap_m: 0, oldest: null });
  assert.strictEqual(d.logsToday, 1);
});

test("after reset: numbering restarts (RC/IS/DM/RPR -0001, ids from 1) and flows work", () => {
  const d = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 10, person: "P" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0001");
  assert.strictEqual(d.id, 1);
  assert.strictEqual(sheet("movements").getLastRow(), 2);
  assert.strictEqual(okData(post("issue", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 4, department: "ฝ่ายบรรจุ" }, AT))).doc_no, "IS-" + YMD + "-0001");
  assert.strictEqual(okData(post("damage", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1, cause: "x" }, AT))).doc_no, "DM-" + YMD + "-0001");
  const t = okData(get("repairs")).items;
  assert.strictEqual(t.length, 1);
  assert.strictEqual(t[0].id, 1);
  assert.strictEqual(t[0].ticket_no, "RPR-" + YMD + "-0001");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 5);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 4);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "damaged"), 1);
  const logs = okData(get("logs")).items;
  assert.deepStrictEqual(logs.map(l => l.id), [4, 3, 2, 1]);
  assert.strictEqual(logs[3].action, "รีเซ็ตข้อมูล");
});

test("reset_data on already-empty sheets reports zero rows removed", () => {
  okData(reset(RPW));
  const r = okData(reset(RPW));
  assert.deepStrictEqual(r.removed, { movements: 0, repairs: 0, audit_logs: 1 });
  assert.strictEqual(okData(get("logs")).items.length, 1);
});

/* ---------- edit / delete records (reset password) ---------- */
const rec = (action, body = {}, pw = RPW) => postRaw(Object.assign({}, body, { action, resetPassword: pw, actor: "Spoofed" }), T.editor);
const hist = () => okData(get("history")).items;
const byDoc = doc => hist().find(m => m.doc_no === doc);
const rowsOf = n => { const d = sheet(n).data; const h = d[0]; return d.slice(1).filter(r => r.some(v => v !== "" && v != null)).map(r => Object.fromEntries(h.map((k, i) => [k, r[i]]))); };
const ticketOf = id => rowsOf("repairs").find(r => r.id === id);
const sheetDump = () => JSON.stringify(["movements", "repairs", "audit_logs"].map(n => sheet(n).data));
const st = () => okData(get("bootstrap"));
const doc = (p, n) => `${p}-${YMD}-${String(n).padStart(4, "0")}`;
const at = (date, time) => ({ date, time });
const D = {};

test("edit/delete setup: RM + PK movements and repair chains", () => {
  okData(reset(RPW));
  clearFails();
  D.r1 = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 10 }, at("2026-01-10", "08:00"))));
  D.i1 = okData(post("issue", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 6, department: "ฝ่ายผลิต" }, at("2026-01-10", "09:00"))));
  D.r2 = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 5 }, at("2026-01-11", "08:00"))));
  D.t1 = okData(post("return", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 2, department: "ฝ่ายผลิต" }, at("2026-01-12", "08:00"))));
  // PK 1.1x1.1
  D.p1 = okData(post("receive", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 20 }, at("2026-02-01", "08:00"))));
  D.p2 = okData(post("issue", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 10, department: "ฝ่ายบรรจุ" }, at("2026-02-01", "09:00"))));
  D.xDm = okData(post("return", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 4, department: "ฝ่ายบรรจุ", condition: "damaged", cause: "หัก" }, at("2026-02-02", "08:00"))));
  D.x = rowsOf("repairs").pop();
  D.xRp = okData(post("repair_start", Object.assign({ id: D.x.id }, at("2026-02-02", "09:00"))));
  D.xRd = okData(post("repair_done", Object.assign({ id: D.x.id }, at("2026-02-02", "10:00"))));
  D.p3 = okData(post("issue", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 14, department: "ฝ่ายผลิต" }, at("2026-02-03", "08:00"))));
  D.p4 = okData(post("receive", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 10 }, at("2026-02-04", "08:00"))));
  D.zDm = okData(post("damage", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 3, cause: "z" }, at("2026-02-04", "08:30"))));
  D.z = rowsOf("repairs").pop();
  D.yDm = okData(post("damage", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 3, cause: "y" }, at("2026-02-04", "09:00"))));
  D.y = rowsOf("repairs").pop();
  D.yRp = okData(post("repair_start", Object.assign({ id: D.y.id, person: "ช่าง" }, at("2026-02-04", "10:00"))));
  const b = st();
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 11);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 4);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "available"), 4);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "damaged"), 3);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "repairing"), 3);
  assert.strictEqual(byDoc(D.xRd.doc_no).ticket_no, D.x.ticket_no); // history exposes the chain's ticket
  assert.strictEqual(byDoc(D.r1.doc_no).ticket_no, undefined);
});

test("edit/delete refused while PALLET_RESET_PASSWORD is unset; GET refused", () => {
  const saved = state.props.PALLET_RESET_PASSWORD;
  delete state.props.PALLET_RESET_PASSWORD;
  const before = sheetDump();
  for (const a of ["movement_update", "movement_delete", "repair_update", "repair_delete"]) {
    const r = err(rec(a, { id: D.r1.id, qty: 1 }), /PALLET_RESET_PASSWORD/);
    assert.strictEqual(r.passwordError, true);
  }
  err(postRaw({ action: "verifyResetPassword", resetPassword: "x" }), /ยังไม่ได้ตั้งรหัสรีเซ็ตข้อมูล/);
  for (const a of ["movement_update", "movement_delete", "repair_update", "repair_delete", "verifyResetPassword"]) err(get(a), /POST/);
  assert.strictEqual(sheetDump(), before);
  state.props.PALLET_RESET_PASSWORD = saved;
});

test("edit/delete with wrong password (or the action password) refused, slept, NOT logged", () => {
  clearFails();
  const before = sheetDump();
  const sleeps = state.sleeps;
  const r = err(rec("movement_delete", { id: D.r2.id }, "wrong"), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(rec("movement_update", { id: D.r2.id, qty: 1 }, PW), /รหัสไม่ถูกต้อง/); // action password is not enough
  err(postRaw({ action: "repair_update", password: RPW, id: D.x.id, cause: "x" }), /รหัสไม่ถูกต้อง/); // must be resetPassword
  err(postRaw({ action: "verifyResetPassword", resetPassword: "nope" }), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(state.cache.PALLET_RESET_FAILURES, "4");
  assert.strictEqual(sheetDump(), before);
  assert.deepStrictEqual(okData(postRaw({ action: "verifyResetPassword", resetPassword: RPW })), { valid: true });
  clearFails();
});

test("edit/delete lockout is shared with reset_data", () => {
  clearFails();
  for (let i = 0; i < 5; i++) err(reset("guess" + i), /รหัสไม่ถูกต้อง/);
  for (let i = 0; i < 5; i++) err(rec("movement_delete", { id: D.r2.id }, "guess" + i), /รหัสไม่ถูกต้อง/);
  const before = sheetDump();
  err(rec("movement_delete", { id: D.r2.id }), /หลายครั้งเกินไป/);
  err(rec("repair_update", { id: D.x.id, cause: "x" }), /หลายครั้งเกินไป/);
  err(postRaw({ action: "verifyResetPassword", resetPassword: RPW }), /หลายครั้งเกินไป/);
  err(reset(RPW), /หลายครั้งเกินไป/);
  assert.strictEqual(sheetDump(), before);
  okData(login("tester", PW)); // login counter unaffected
  clearFails(); // 15 minutes later
});

test("delete that would make stock negative is rejected with Thai reason, nothing changed, warn logged", () => {
  const before = sheetDump();
  const logs = dataRows("audit_logs");
  const r = err(rec("movement_delete", { id: D.r1.id, doc_no: D.r1.doc_no }));
  assert.strictEqual(r.error, `ลบไม่ได้: ยอดพร้อมใช้ RM 1.2x1.2 จะติดลบ (-6) ณ วันที่ 10/01/2026 09:00 (${D.i1.doc_no})`);
  assert.ok(!r.passwordError);
  const after = JSON.parse(sheetDump()), was = JSON.parse(before);
  assert.deepStrictEqual(after.slice(0, 2), was.slice(0, 2)); // movements + repairs untouched
  assert.strictEqual(dataRows("audit_logs"), logs + 1); // + one warn row
  const log = lastLog();
  assert.strictEqual(log.category, "warn");
  assert.strictEqual(log.action, "movement_delete");
  assert.strictEqual(log.detail, "ปฏิเสธลบรายการ: " + r.error);
  assert.strictEqual(log.actor, EDITOR);
  // deleting the issue would make the "issued" bucket negative when the return happens
  err(rec("movement_delete", { id: D.i1.id }), /^ลบไม่ได้: ยอดเบิกไปใช้งาน RM 1\.2x1\.2 จะติดลบ \(-2\) ณ วันที่ 12\/01\/2026 08:00/);
});

test("delete simple movement: row removed, stock recomputed, audit row, lock used", () => {
  const n = dataRows("movements");
  const locks = state.locks;
  const calls = (state.deleteCalls || []).length;
  const row = sheet("movements").data.findIndex(r => r[1] === D.r2.doc_no) + 1;
  const d = okData(rec("movement_delete", { id: D.r2.id, doc_no: D.r2.doc_no }));
  assert.deepStrictEqual(d.removed, { movements: 1, repairs: 0 });
  assert.strictEqual(d.message, "ลบรายการ " + D.r2.doc_no + " แล้ว");
  assert.strictEqual(state.locks - locks, 1);
  assert.strictEqual(state.locks, state.unlocks);
  assert.deepStrictEqual(state.deleteCalls.slice(calls), [["movements", row, 1]]);
  assert.strictEqual(dataRows("movements"), n - 1);
  assert.ok(!byDoc(D.r2.doc_no));
  const b = st();
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 6);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 4);
  const log = lastLog();
  assert.strictEqual(log.category, "pallet");
  assert.strictEqual(log.action, "ลบรายการ");
  assert.strictEqual(log.ref, D.r2.doc_no);
  assert.strictEqual(log.actor, EDITOR);
  assert.ok(log.detail.startsWith(`ลบรายการ ${D.r2.doc_no}: รับเข้า RM (พาเลทสำหรับใส่ RM) ขนาด 1.2x1.2 ม. จำนวน 5 ตัว · วันที่ 11/01/2026 08:00`), log.detail);
  // stale id / doc_no
  err(rec("movement_delete", { id: D.r2.id, doc_no: D.r2.doc_no }), /ไม่พบรายการเคลื่อนไหว/);
  err(rec("movement_delete", { id: D.r1.id, doc_no: "RC-000000-0009" }), /ไม่พบรายการเคลื่อนไหว/);
});

test("after delete: new entries get fresh ids and never reuse a deleted doc number", () => {
  const maxId = Math.max(...rowsOf("movements").map(r => r.id));
  const d = okData(post("receive", Object.assign({ type_id: 4, size: "1.2x1.2", qty: 1 }, AT)));
  assert.strictEqual(d.doc_no, doc("RC", 5)); // RC-0002 was deleted: not handed out again
  assert.strictEqual(d.id, maxId + 1);
  const docs = hist().map(m => m.doc_no);
  assert.strictEqual(new Set(docs).size, docs.length);
  okData(rec("movement_delete", { id: d.id, doc_no: d.doc_no })); // highest id + newest number deleted
  const d2 = okData(post("receive", Object.assign({ type_id: 4, size: "1.2x1.2", qty: 1 }, AT)));
  assert.strictEqual(d2.doc_no, doc("RC", 6));
  okData(rec("movement_delete", { id: d2.id, doc_no: d2.doc_no }));
});

test("edit qty: valid change updates stock + audit before→after; negative replay rejected", () => {
  const d = okData(rec("movement_update", { id: D.i1.id, doc_no: D.i1.doc_no, qty: 5 }));
  assert.strictEqual(d.changed, true);
  assert.strictEqual(d.message, "แก้ไขรายการ " + D.i1.doc_no + " แล้ว");
  assert.strictEqual(byDoc(D.i1.doc_no).qty, 5);
  let b = st();
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 7);
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "issued"), 3);
  assert.deepStrictEqual(b.dept.find(r => r.department === "ฝ่ายผลิต" && r.type_id === 1), { department: "ฝ่ายผลิต", type_id: 1, size: "1.2x1.2", qty: 3 });
  const log = lastLog();
  assert.strictEqual(log.category, "pallet");
  assert.strictEqual(log.action, "แก้ไขรายการ");
  assert.strictEqual(log.ref, D.i1.doc_no);
  assert.strictEqual(log.detail, `แก้ไขรายการ ${D.i1.doc_no} (เบิกจ่าย RM ขนาด 1.2x1.2): จำนวน 6 → 5`);
  const before = sheetDump();
  err(rec("movement_update", { id: D.i1.id, qty: 20 }), /^แก้ไขไม่ได้: ยอดพร้อมใช้ RM 1\.2x1\.2 จะติดลบ \(-10\) ณ วันที่ 10\/01\/2026 09:00/);
  err(rec("movement_update", { id: D.i1.id, qty: 1 }), /^แก้ไขไม่ได้: ยอดเบิกไปใช้งาน RM 1\.2x1\.2 จะติดลบ \(-1\) ณ วันที่ 12\/01\/2026 08:00/);
  err(rec("movement_update", { id: D.i1.id, qty: 0 }), /^จำนวนต้องมากกว่า 0$/);
  const a = JSON.parse(sheetDump()), w = JSON.parse(before);
  assert.deepStrictEqual(a.slice(0, 2), w.slice(0, 2));
  assert.strictEqual(lastLog().detail, "ปฏิเสธแก้ไขรายการ: จำนวนต้องมากกว่า 0");
});

test("edit date: reorder causing negative is rejected; valid date change saved", () => {
  const before = JSON.parse(sheetDump()).slice(0, 2);
  err(rec("movement_update", { id: D.i1.id, date: "2026-01-09", time: "08:00" }),
    /^แก้ไขไม่ได้: ยอดพร้อมใช้ RM 1\.2x1\.2 จะติดลบ \(-5\) ณ วันที่ 09\/01\/2026 08:00/);
  err(rec("movement_update", { id: D.t1.id, date: "2026-01-10", time: "08:30" }), // return before the issue
    /^แก้ไขไม่ได้: ยอดเบิกไปใช้งาน RM 1\.2x1\.2 จะติดลบ \(-2\)/);
  err(rec("movement_update", { id: D.t1.id, date: "2026/01/10", time: "08:30" }), /^วันที่\/เวลาไม่ถูกต้อง$/);
  assert.deepStrictEqual(JSON.parse(sheetDump()).slice(0, 2), before);
  okData(rec("movement_update", { id: D.t1.id, date: "2026-01-13", time: "10:30" }));
  assert.strictEqual(byDoc(D.t1.doc_no).moved_at, "2026-01-13 10:30:00");
  assert.strictEqual(lastLog().detail, `แก้ไขรายการ ${D.t1.doc_no} (รับคืน RM ขนาด 1.2x1.2): วันที่/เวลา 12/01/2026 08:00 → 13/01/2026 10:30`);
});

test("edit department / person / note; department checks; immutable fields; no-op", () => {
  err(rec("movement_update", { id: D.t1.id, department: "ฝ่ายบรรจุ" }),
    /^แก้ไขไม่ได้: ยอดพาเลท RM 1\.2x1\.2 ที่ ฝ่ายบรรจุ ถืออยู่ จะติดลบ \(-2\) ณ วันที่ 13\/01\/2026 10:30/);
  err(rec("movement_update", { id: D.t1.id, department: " " }), /^กรุณาเลือกฝ่าย$/);
  err(rec("movement_update", { id: D.r1.id, department: "ฝ่ายผลิต" }), /ไม่มีฝ่ายให้แก้ไข/);
  okData(rec("movement_update", { id: D.r1.id, department: "" })); // receive has no department: blank is fine
  okData(rec("movement_update", { id: D.p3.id, department: "ฝ่ายคลังสินค้า" }));
  const b = st();
  assert.ok(b.dept.some(r => r.department === "ฝ่ายคลังสินค้า" && r.type_id === 2 && r.qty === 14));
  assert.ok(!b.dept.some(r => r.department === "ฝ่ายผลิต" && r.type_id === 2));
  okData(rec("movement_update", { id: D.r1.id, person: " สมหญิง ", note: "แก้ PO" }));
  const m = byDoc(D.r1.doc_no);
  assert.deepStrictEqual([m.person, m.note, m.qty], ["สมหญิง", "แก้ PO", 10]);
  assert.strictEqual(lastLog().detail, `แก้ไขรายการ ${D.r1.doc_no} (รับเข้า RM ขนาด 1.2x1.2): ชื่อที่ระบุ "" → "สมหญิง"; หมายเหตุ "" → "แก้ PO"`);
  assert.strictEqual(lastLog().actor, EDITOR);
  assert.strictEqual(m.actor, TESTER); // who recorded the movement is never edited
  okData(rec("movement_update", { id: D.r1.id, actor: "Hacker", username: "admin" })); // ignored fields
  assert.deepStrictEqual([byDoc(D.r1.doc_no).actor, byDoc(D.r1.doc_no).username], [TESTER, "tester"]);
  for (const f of [{ type_id: 2 }, { size: "1.1x1.1" }, { from_status: "damaged" }, { to_status: "issued" }]) {
    err(rec("movement_update", Object.assign({ id: D.r1.id }, f)), /ลบรายการนี้แล้วบันทึกใหม่/);
  }
  okData(rec("movement_update", { id: D.r1.id, type_id: 1, size: "1.2x1.2", to_status: "available" })); // same values are fine
  const logs = dataRows("audit_logs");
  const same = okData(rec("movement_update", { id: D.r1.id, qty: 10, person: "สมหญิง", date: "2026-01-10", time: "08:00" }));
  assert.deepStrictEqual(same, { changed: false, message: "ไม่มีการเปลี่ยนแปลง" });
  assert.strictEqual(dataRows("audit_logs"), logs);
  err(rec("movement_update", { id: D.r1.id, note: "x".repeat(256) }), /ยาวเกิน 255/);
});

test("repair chain: date out of chain order is rejected even when stock allows it", () => {
  // Z keeps 3 pallets "damaged" at 08:30, so moving Y's repair_start to 08:45 (before Y's damage at 09:00) keeps stock >= 0
  const before = JSON.parse(sheetDump()).slice(0, 2);
  err(rec("movement_update", { id: D.yRp.id, date: "2026-02-04", time: "08:45" }),
    new RegExp(`^แก้ไขไม่ได้: ลำดับเวลาของงานซ่อม ${D.y.ticket_no} ไม่ถูกต้อง`));
  err(rec("movement_update", { id: D.yRp.id, date: "2026-02-04", time: "08:00" }),
    /^แก้ไขไม่ได้: ยอดชำรุด PK 1\.1x1\.1 จะติดลบ \(-3\) ณ วันที่ 04\/02\/2026 08:00/);
  assert.deepStrictEqual(JSON.parse(sheetDump()).slice(0, 2), before);
  okData(rec("movement_update", { id: D.yRp.id, date: "2026-02-04", time: "11:00" }));
  assert.strictEqual(ticketOf(D.y.id).started_at, "2026-02-04 11:00:00");
  okData(rec("movement_update", { id: D.yDm.id, date: "2026-02-04", time: "09:15" }));
  assert.strictEqual(ticketOf(D.y.id).reported_at, "2026-02-04 09:15:00");
});

test("repair chain: qty edit propagates to every chain movement and the ticket; too large rejected", () => {
  const d = okData(rec("movement_update", { id: D.yRp.id, qty: 4 })); // edit via the repair_start row
  assert.strictEqual(d.changed, true);
  assert.deepStrictEqual([byDoc(D.yDm.doc_no).qty, byDoc(D.yRp.doc_no).qty, ticketOf(D.y.id).qty], [4, 4, 4]);
  const b = st();
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "repairing"), 4);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "available"), 3);
  assert.ok(lastLog().detail.includes(`จำนวน 3 → 4 · ปรับจำนวนทั้งชุดงานซ่อม ${D.y.ticket_no} (2 รายการ + ใบแจ้งซ่อม)`), lastLog().detail);
  const before = JSON.parse(sheetDump()).slice(0, 2);
  err(rec("movement_update", { id: D.yDm.id, qty: 50 }), /^แก้ไขไม่ได้: ยอดพร้อมใช้ PK 1\.1x1\.1 จะติดลบ/);
  assert.deepStrictEqual(JSON.parse(sheetDump()).slice(0, 2), before);
  // department of a damage-from-issued chain movement is synced to the ticket, and checked
  err(rec("movement_update", { id: D.xDm.id, department: "ฝ่ายผลิต" }), /^แก้ไขไม่ได้: ยอดพาเลท PK 1\.1x1\.1 ที่ ฝ่ายผลิต ถืออยู่ จะติดลบ \(-4\) ณ วันที่ 02\/02\/2026 08:00/);
  err(rec("movement_update", { id: D.zDm.id, department: "ฝ่ายผลิต" }), /ไม่มีฝ่ายให้แก้ไข/); // damage from stock has no dept
});

test("repair_update: cause / reported_by / repairer / note only; audit before→after", () => {
  err(rec("repair_update", { id: D.x.id, reported_by: "ผู้แจ้งปลอม" }), /บันทึกจากบัญชีผู้ใช้โดยอัตโนมัติ/);
  err(rec("repair_update", { id: D.x.id, updated_by: "x" }), /แก้ไขไม่ได้/);
  const d = okData(rec("repair_update", { id: D.x.id, ticket_no: D.x.ticket_no, cause: "ไม้หักสองแผ่น", repairer: "ช่างสอง", reported_by: TESTER, note: "บันทึก\nบรรทัดสอง" }));
  assert.strictEqual(d.message, "แก้ไขใบแจ้งซ่อม " + D.x.ticket_no + " แล้ว");
  const t = ticketOf(D.x.id);
  assert.deepStrictEqual([t.cause, t.repairer, t.reported_by, t.note, t.qty, t.stage], ["ไม้หักสองแผ่น", "ช่างสอง", TESTER, "บันทึก\nบรรทัดสอง", 4, "done"]);
  assert.deepStrictEqual([t.updated_by, t.updated_username], [EDITOR, "editor"]);
  const log = lastLog();
  assert.strictEqual(log.category, "repair");
  assert.strictEqual(log.action, "แก้ไขใบแจ้งซ่อม");
  assert.strictEqual(log.ref, D.x.ticket_no);
  assert.strictEqual(log.actor, EDITOR);
  assert.ok(log.detail.startsWith(`แก้ไขใบแจ้งซ่อม ${D.x.ticket_no}: สาเหตุการชำรุด "หัก" → "ไม้หักสองแผ่น"; ช่างผู้ซ่อม "Tester" → "ช่างสอง"`), log.detail);
  err(rec("repair_update", { id: D.x.id, qty: 9 }), /แก้ไขจำนวน\/สถานะ/);
  err(rec("repair_update", { id: D.x.id, stage: "damaged" }), /แก้ไขจำนวน\/สถานะ/);
  err(rec("repair_update", { id: D.x.id, note: "x".repeat(1001) }), /ยาวเกิน 1000/);
  err(rec("repair_update", { id: D.x.id, ticket_no: "RPR-000000-0001", cause: "a" }), /ไม่พบใบแจ้งซ่อม/);
  assert.deepStrictEqual(okData(rec("repair_update", { id: D.x.id, cause: "ไม้หักสองแผ่น" })).changed, false);
  assert.strictEqual(lastLog().detail, "ปฏิเสธแก้ไขใบแจ้งซ่อม: ไม่พบใบแจ้งซ่อมนี้ (อาจถูกลบไปแล้ว) กรุณาโหลดหน้าใหม่");
});

test("delete a chain that would make stock negative is rejected (repaired pallets were re-issued)", () => {
  const before = JSON.parse(sheetDump()).slice(0, 2);
  err(rec("movement_delete", { id: D.xRd.id, doc_no: D.xRd.doc_no }),
    new RegExp(`^ลบไม่ได้: ยอดพร้อมใช้ PK 1\\.1x1\\.1 จะติดลบ \\(-4\\) ณ วันที่ 03/02/2026 08:00 \\(${D.p3.doc_no}\\)`));
  err(rec("repair_delete", { id: D.x.id }), /^ลบไม่ได้: ยอดพร้อมใช้ PK/);
  assert.deepStrictEqual(JSON.parse(sheetDump()).slice(0, 2), before);
});

test("delete a repair-chain movement removes the whole chain and its ticket", () => {
  const nMv = dataRows("movements"), nRp = dataRows("repairs");
  const d = okData(rec("movement_delete", { id: D.yDm.id, doc_no: D.yDm.doc_no }));
  assert.deepStrictEqual(d.removed, { movements: 2, repairs: 1 });
  assert.strictEqual(d.message, `ลบงานซ่อม ${D.y.ticket_no} ทั้งชุดแล้ว (ใบแจ้งซ่อม + 2 รายการเคลื่อนไหว)`);
  assert.strictEqual(dataRows("movements"), nMv - 2);
  assert.strictEqual(dataRows("repairs"), nRp - 1);
  assert.ok(!byDoc(D.yDm.doc_no) && !byDoc(D.yRp.doc_no));
  assert.ok(!ticketOf(D.y.id));
  assert.ok(!hist().some(m => m.repair_id === D.y.id));
  const b = st();
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "repairing"), 0);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "damaged"), 3);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "available"), 7);
  const log = lastLog();
  assert.strictEqual(log.category, "pallet");
  assert.strictEqual(log.action, "ลบรายการ");
  assert.strictEqual(log.ref, D.y.ticket_no);
  assert.ok(log.detail.startsWith(`ลบรายการทั้งชุดงานซ่อม ${D.y.ticket_no} (PK ขนาด 1.1x1.1): ใบแจ้งซ่อม + 2 รายการ — ${D.yDm.doc_no} แจ้งชำรุด 4 ตัว`), log.detail);
  // repair_delete from the repairs page (ticket Z: single damage movement)
  const r = okData(rec("repair_delete", { id: D.z.id, ticket_no: D.z.ticket_no }));
  assert.deepStrictEqual(r.removed, { movements: 1, repairs: 1 });
  assert.strictEqual(stockOf(st().stock, 2, "1.1x1.1", "damaged"), 0);
  // new tickets / docs don't reuse deleted numbers
  okData(post("damage", Object.assign({ type_id: 2, size: "1.1x1.1", qty: 1 }, AT)));
  const t = rowsOf("repairs").pop();
  assert.strictEqual(t.ticket_no, "RPR-" + YMD + "-0004");
  assert.strictEqual(byDoc(doc("DM", 4)).repair_id, t.id);
  okData(rec("repair_delete", { id: t.id }));
});

test("deleting the last data rows keeps a spare sheet row (Sheets cannot delete all non-frozen rows)", () => {
  okData(rec("movement_delete", { id: D.p3.id })); // frees the re-issued pallets so chain X can go
  const s = sheet("repairs");
  assert.strictEqual(dataRows("repairs"), 1);
  s.maxRows = s.getLastRow(); // header + 1 row, nothing spare
  const d = okData(rec("repair_delete", { id: D.x.id, ticket_no: D.x.ticket_no }));
  assert.deepStrictEqual(d.removed, { movements: 3, repairs: 1 });
  assert.strictEqual(dataRows("repairs"), 0);
  assert.ok(s.maxRows >= 2);
  assert.deepStrictEqual(okData(get("repairs")).items, []);
  const b = st();
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "available"), 20);
  assert.strictEqual(stockOf(b.stock, 2, "1.1x1.1", "issued"), 10);
  // every movement row still parses to a consistent ledger (replay of all data is non-negative)
  const bal = {};
  hist().slice().reverse().forEach(m => {
    if (m.to_status) bal[m.type_id + m.size + m.to_status] = (bal[m.type_id + m.size + m.to_status] || 0) + m.qty;
    if (m.from_status) bal[m.type_id + m.size + m.from_status] = (bal[m.type_id + m.size + m.from_status] || 0) - m.qty;
    assert.ok(Object.values(bal).every(v => v >= 0), "negative at " + m.doc_no);
  });
  assert.strictEqual(state.locks, state.unlocks);
});

/* ---------- attribution & account management ---------- */
test("actor is decided by the server: fake person/actor/reported_by from the client never becomes the actor", () => {
  okData(postRaw(Object.assign({ action: "receive", type_id: 3, size: "1.2x1.2", qty: 5, actor: "Hacker", username: "admin", person: "ใครก็ได้" }, AT), T.editor));
  const d = okData(post("damage", Object.assign({ type_id: 3, size: "1.2x1.2", qty: 1, cause: "แตก", person: "Fake", reported_by: "Fake", reported_username: "admin", actor: "Fake" }, AT), T.editor));
  const m = byDoc(d.doc_no);
  assert.deepStrictEqual([m.actor, m.username, m.person], [EDITOR, "editor", "Fake"]);
  const t = rowsOf("repairs").pop();
  assert.deepStrictEqual([t.reported_by, t.reported_username], [EDITOR, "editor"]);
  const rp = okData(post("repair_start", Object.assign({ id: t.id, actor: "Fake" }, AT)));
  const t2 = ticketOf(t.id);
  assert.deepStrictEqual([t2.reported_by, t2.updated_by, t2.updated_username, t2.repairer], [EDITOR, TESTER, "tester", "Tester"]);
  assert.strictEqual(byDoc(rp.doc_no).actor, TESTER);
  const logs = okData(get("logs")).items.slice(0, 3);
  assert.deepStrictEqual(logs.map(l => l.actor), [TESTER, EDITOR, EDITOR]);
  assert.ok(!JSON.stringify(state.spreadsheets[state.props.PALLET_SPREADSHEET_ID].sheets.map(x => x.data)).includes("Hacker"));
  okData(rec("repair_delete", { id: t.id })); // tidy up
});

test("change_password: wrong old password refused (1 s), min 8, success rotates sessions, audit row", () => {
  const second = okData(login("tester", PW)).token; // the same user on another device
  const sleeps = state.sleeps;
  err(post("change_password", { old_password: "nope", new_password: "newPass-123" }), /รหัสผ่านเดิมไม่ถูกต้อง/);
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(post("change_password", { old_password: PW, new_password: "short" }), /อย่างน้อย 8/);
  const d = okData(post("change_password", { old_password: PW, new_password: "newPass-123" }));
  assert.ok(/^[a-f0-9]{64}$/.test(d.token) && d.token !== T.tester);
  authErr(get("bootstrap"));            // this browser's old token
  authErr(get("bootstrap", {}, second)); // and the other device
  T.tester = d.token;
  okData(get("bootstrap"));
  err(login("tester", PW), /ไม่ถูกต้อง/);
  T.tester = okData(login("tester", "newPass-123")).token;
  const l = okData(get("logs", { cat: "account" }, T.admin)).items.find(x => x.action === "change_password");
  assert.deepStrictEqual([l.actor, l.username, l.ref], [TESTER, "tester", "tester"]);
  assert.ok(!JSON.stringify(sheet("audit_logs").data).includes("newPass-123"));
  clearFails();
});

test("admin reset password ends that user's sessions; own reset keeps the admin logged in", () => {
  const tester = userRows().find(u => u.username === "tester");
  err(post("user_reset_password", { id: tester.id, password: "1234567" }, T.admin), /อย่างน้อย 8/);
  err(post("user_reset_password", { id: 999, password: "12345678" }, T.admin), /ไม่พบผู้ใช้/);
  okData(post("user_reset_password", { id: tester.id, password: PW }, T.admin));
  authErr(get("bootstrap"));
  err(login("tester", "newPass-123"), /ไม่ถูกต้อง/);
  T.tester = okData(login("tester", PW)).token;
  const own = okData(post("user_reset_password", { id: 1, password: "admin-New-99" }, T.admin));
  assert.ok(own.token);
  authErr(get("me", {}, T.admin));
  authErr(get("me", {}, T.admin2));
  T.admin = own.token;
  assert.strictEqual(okData(get("me", {}, T.admin)).user.username, "admin");
  const l = okData(get("logs", { cat: "account" }, T.admin)).items.find(x => x.action === "user_reset_password" && x.ref === "tester");
  assert.strictEqual(l.actor, ADMIN);
  clearFails();
});

test("initial admin password is not used once accounts exist", () => {
  err(login("admin", ADMIN_PW), /ไม่ถูกต้อง/); // admin password was changed above
  assert.strictEqual(userRows().filter(u => u.username === "admin").length, 1);
  delete state.props.PALLET_INITIAL_ADMIN_PASSWORD;
  T.admin = okData(login("admin", "admin-New-99")).token;
  clearFails();
});

test("user_save update / user_toggle: disabling ends sessions and blocks login; re-enable works", () => {
  const editor = userRows().find(u => u.username === "editor");
  const save = b => post("user_save", b, T.admin);
  const r = okData(save({ id: editor.id, fullname: "Editor", role: "user", active: 1 }));
  assert.strictEqual(r.changed, false);
  okData(save({ id: editor.id, fullname: "Editor", active: 0 }));
  authErr(get("bootstrap", {}, T.editor));
  err(login("editor", EDITOR_PW), /ไม่ถูกต้อง/);
  okData(post("user_toggle", { id: editor.id }, T.admin)); // enable again
  assert.strictEqual(userRows().find(u => u.username === "editor").active, 1);
  T.editor = okData(login("editor", EDITOR_PW)).token;
  okData(post("user_toggle", { id: editor.id }, T.admin)); // disable
  authErr(get("bootstrap", {}, T.editor));
  okData(post("user_toggle", { id: editor.id }, T.admin)); // enable
  T.editor = okData(login("editor", EDITOR_PW)).token;
  const acts = okData(get("logs", { cat: "account" }, T.admin)).items.filter(l => l.ref === "editor").map(l => l.action);
  assert.ok(acts.includes("user_update") && acts.includes("user_toggle"));
  // rename + role change are read fresh on every request
  okData(save({ id: editor.id, fullname: "Editor Two", role: "admin" }));
  okData(get("users", {}, T.editor)); // now an admin
  okData(save({ id: editor.id, fullname: "Editor", role: "user" }));
  err(get("users", {}, T.editor), /เฉพาะผู้ดูแลระบบ/);
  const upd = okData(get("logs", { cat: "account" }, T.admin)).items.find(l => l.action === "user_update");
  assert.ok(upd.detail.includes("บทบาท ผู้ดูแลระบบ → ผู้ใช้งาน"), upd.detail);
  clearFails();
});

test("last active admin cannot be disabled or demoted (also not by themselves)", () => {
  const before = JSON.stringify(sheet("users").data);
  err(post("user_toggle", { id: 1 }, T.admin), /อย่างน้อย 1 คน/);
  err(post("user_save", { id: 1, fullname: "ผู้ดูแลระบบ", role: "user" }, T.admin), /อย่างน้อย 1 คน/);
  err(post("user_save", { id: 1, fullname: "ผู้ดูแลระบบ", active: 0 }, T.admin), /อย่างน้อย 1 คน/);
  assert.strictEqual(JSON.stringify(sheet("users").data), before);
  // with a second admin, the first may step down
  const editor = userRows().find(u => u.username === "editor");
  okData(post("user_save", { id: editor.id, fullname: "Editor", role: "admin" }, T.admin));
  okData(post("user_save", { id: 1, fullname: "ผู้ดูแลระบบ", role: "user" }, T.admin));
  err(get("users", {}, T.admin), /เฉพาะผู้ดูแลระบบ/);
  err(post("user_toggle", { id: editor.id }, T.editor), /อย่างน้อย 1 คน/); // editor is now the last admin
  okData(post("user_save", { id: 1, fullname: "ผู้ดูแลระบบ", role: "admin" }, T.editor));
  okData(post("user_save", { id: editor.id, fullname: "Editor", role: "user" }, T.admin));
  // a disabled admin does not count
  okData(post("user_save", { username: "boss2", fullname: "Boss", role: "admin", password: "boss-Pass-1", active: 0 }, T.admin));
  err(post("user_save", { id: 1, fullname: "ผู้ดูแลระบบ", role: "user" }, T.admin), /อย่างน้อย 1 คน/);
});

test("logout ends the session and is logged", () => {
  const tok = okData(login("tester", PW)).token;
  const d = okData(post("logout", {}, tok));
  assert.strictEqual(d.message, "ออกจากระบบแล้ว");
  authErr(get("bootstrap", {}, tok));
  authErr(post("logout", {}, tok));
  okData(get("bootstrap")); // the other session of the same user is unaffected
  const l = okData(get("logs", { cat: "account" }, T.admin)).items[0];
  assert.deepStrictEqual([l.action, l.actor, l.username], ["logout", TESTER, "tester"]);
  assert.strictEqual(state.locks, state.unlocks);
});

/* ---------- migration of an existing (pre-login) spreadsheet ---------- */
const LEGACY_HEADERS = {
  pallet_types: ["id", "tkey", "code", "name", "short", "description", "color", "sizes", "sort"],
  departments: ["id", "name", "icon", "color", "active"],
  repairs: ["id", "ticket_no", "type_id", "size", "qty", "stage", "source", "department", "cause", "reported_by", "repairer", "reported_at", "started_at", "finished_at", "note"],
  movements: ["id", "doc_no", "action", "type_id", "size", "qty", "from_status", "to_status", "department", "person", "note", "repair_id", "moved_at", "created_at"],
  audit_logs: ["id", "category", "action", "ref", "detail", "actor", "ip", "created_at"]
};
function legacyInstance() {
  const g = createGas();
  const c = loadCode(g.gas);
  const ss = g.gas.SpreadsheetApp.create("Pallet Hub Database");
  ss.sheets = [];
  g.state.props.PALLET_SPREADSHEET_ID = ss.getId();
  for (const [name, h] of Object.entries(LEGACY_HEADERS)) {
    const sh = ss.insertSheet(name);
    sh.data.push(h.slice());
    sh.maxCols = h.length; // a tight sheet: new columns need insertColumnsAfter
    sh.frozen = 1;
  }
  ss.getSheetByName("pallet_types").data.push([1, "RM", "RM", "พาเลท RM", "RM", "วัตถุดิบ", "#1E6FE0", "1.2x1.2", 1]);
  ss.getSheetByName("departments").data.push([1, "ฝ่ายผลิต", "fa-industry", "#E2231A", 1]);
  ss.getSheetByName("movements").data.push([1, "RC-260101-0001", "receive", 1, "1.2x1.2", 10, "", "available", "", "คนเก่า", "ของเดิม", "", "2026-01-01 08:00:00", "2026-01-01 08:00:00"]);
  ss.getSheetByName("audit_logs").data.push([1, "pallet", "receive", "RC-260101-0001", "รับเข้าเดิม", "คนเก่า", "web", "2026-01-01 08:00:00"]);
  const G = (action, tok, params = {}) => JSON.parse(c.doGet({ parameter: Object.assign({ action, token: tok }, params) }).getContent());
  const P = (payload) => JSON.parse(c.doPost({ postData: { contents: JSON.stringify(payload) } }).getContent());
  return { g, c, ss, G, P, sh: n => ss.getSheetByName(n) };
}

test("existing spreadsheet without users sheet: first login creates it; old sheets gain new columns; old data intact", () => {
  const L = legacyInstance();
  L.g.state.props.PALLET_INITIAL_ADMIN_PASSWORD = "legacy-Admin-1";
  assert.strictEqual(L.sh("users"), null);
  const tok = okData(L.P({ action: "login", username: "admin", password: "legacy-Admin-1" })).token;
  assert.ok(L.sh("users"));
  assert.deepStrictEqual(L.sh("users").data[0], ["id", "username", "password_hash", "salt", "fullname", "role", "active", "created_at", "updated_at", "last_login"]);
  assert.strictEqual(L.sh("users").getLastRow(), 2);
  assert.deepStrictEqual(L.sh("audit_logs").data[0], LEGACY_HEADERS.audit_logs.concat(["username"]));
  // reads upgrade movements; the legacy row reads its new columns as empty
  const h = okData(L.G("history", tok)).items;
  assert.deepStrictEqual(L.sh("movements").data[0], LEGACY_HEADERS.movements.concat(["actor", "username"]));
  assert.deepStrictEqual([h[0].doc_no, h[0].person, h[0].actor, h[0].username, h[0].from_status], ["RC-260101-0001", "คนเก่า", "", "", null]);
  assert.strictEqual(L.sh("movements").data[1][9], "คนเก่า");
  const logs = okData(L.G("logs", tok)).items;
  assert.deepStrictEqual(logs.find(l => l.id === 1), { id: 1, category: "pallet", action: "receive", ref: "RC-260101-0001", detail: "รับเข้าเดิม", actor: "คนเก่า", ip: "web", created_at: "2026-01-01 08:00:00", username: "" });
  // writes work and attribute the user; repairs gains its columns
  okData(L.P({ action: "receive", token: tok, type_id: 1, size: "1.2x1.2", qty: 2 }));
  okData(L.P({ action: "damage", token: tok, type_id: 1, size: "1.2x1.2", qty: 1, cause: "x" }));
  assert.deepStrictEqual(L.sh("repairs").data[0], LEGACY_HEADERS.repairs.concat(["reported_username", "updated_by", "updated_username", "updated_at"]));
  const t = okData(L.G("repairs", tok)).items[0];
  assert.deepStrictEqual([t.reported_by, t.reported_username], [ADMIN, "admin"]);
  const st = okData(L.G("bootstrap", tok)).stock;
  assert.strictEqual(stockOf(st, 1, "1.2x1.2", "available"), 11);
  assert.ok(L.g.state.colInserts >= 3);
  assert.ok(okData(L.G("history", tok)).items.slice(0, 2).every(m => m.actor === ADMIN));
  // setupSystem afterwards is harmless and keeps the account
  const r = L.c.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(r.users, 1);
  okData(L.G("me", tok));
});

test("existing spreadsheet upgraded by re-running setupSystem (users sheet + columns added, data kept)", () => {
  const L = legacyInstance();
  const r = L.c.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(r.users, 0);
  assert.deepStrictEqual(L.ss.getSheets().map(s => s.getName()), ["pallet_types", "departments", "repairs", "movements", "audit_logs", "users"]);
  assert.deepStrictEqual(L.sh("movements").data[0], LEGACY_HEADERS.movements.concat(["actor", "username"]));
  assert.deepStrictEqual(L.sh("audit_logs").data[0], LEGACY_HEADERS.audit_logs.concat(["username"]));
  assert.strictEqual(L.sh("movements").data[1][1], "RC-260101-0001");
  assert.strictEqual(L.sh("pallet_types").getLastRow(), 2); // no re-seeding over existing data
  assert.strictEqual(r.initialAdminPasswordConfigured, false);
  const e = err(L.P({ action: "login", username: "admin", password: "x" }), /PALLET_INITIAL_ADMIN_PASSWORD/);
  assert.strictEqual(e.code, "SETUP");
});

/* ---------- frontend syntax ---------- */
test("docs JS files parse (new Function)", () => {
  for (const f of ["docs/config.js", "docs/assets/app.js"]) {
    const src = fs.readFileSync(path.join(ROOT, f), "utf8");
    new Function(src); // throws SyntaxError on bad code
  }
  const cfg = {};
  new Function("window", fs.readFileSync(path.join(ROOT, "docs/config.js"), "utf8"))(cfg);
  assert.deepStrictEqual(Object.keys(cfg.PALLET_CONFIG), ["apiUrl"]);
  assert.ok(cfg.PALLET_CONFIG.apiUrl === "" || /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(cfg.PALLET_CONFIG.apiUrl));
  const app = fs.readFileSync(path.join(ROOT, "docs/assets/app.js"), "utf8");
  assert.ok(!/api\.php|X-User|localStorage\.setItem\([^)]*[Pp]ass/.test(app));
  assert.ok(app.includes("action: 'reset_data', resetPassword"));
  assert.ok(!/(local|session)Storage\.setItem\([^)]*[Rr]eset|sessionSet\([^)]*[Rr]eset/.test(app)); // reset password never stored
  assert.ok(!/(local|session)Storage\.setItem\([^)]*[Rr]ecord|sessionSet\([^)]*[Rr]ecord|store\('set', *'[^']*[Pp]ass/.test(app));
  for (const a of ["movement_update", "movement_delete", "repair_update", "repair_delete", "verifyResetPassword"]) assert.ok(app.includes(`'${a}'`), a);
  assert.ok(app.includes("resetPassword: pw }"));
  assert.ok(/data-medit=.*✏️ แก้ไข/.test(app) && /data-mdel=.*🗑 ลบ/.test(app) && /data-redit=.*data-rdel=/.test(app));
  const html = fs.readFileSync(path.join(ROOT, "docs/index.html"), "utf8");
  assert.ok(html.includes('assets/app.js?v=15"'));
  assert.ok(html.includes('id="loginScreen"') && html.includes('data-page="account"'));
  assert.ok(!/verifyPassword'|actionPassword|palletUser/.test(app)); // old shared password / typed-name features removed
  for (const a of ["login", "logout", "me", "change_password", "users", "user_save", "user_reset_password", "user_toggle"]) assert.ok(app.includes(`'${a}'`), a);
  assert.ok(/code === 'AUTH'/.test(app));
  assert.ok(app.includes("ผู้ทำรายการ / By"));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
