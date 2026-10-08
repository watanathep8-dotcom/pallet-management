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
  const state = { props: {}, cache: {}, cachePuts: [], cachePutAlls: [], cacheGetAlls: 0, flushes: 0, cacheDown: false, sheetReads: {}, rowReads: 0, events: [], sleeps: 0, locks: 0, unlocks: 0, lockFail: false, spreadsheets: {}, seq: 0, formatCalls: 0, created: 0, fetches: [], fetchStatus: 202, fetchError: null, logs: [] };

  class Range {
    constructor(sheet, row, col, nr, nc) {
      if (row < 1 || col < 1 || nr < 1 || nc < 1) throw new Error("Range out of bounds");
      if (row + nr - 1 > sheet.maxRows) throw new Error("Range exceeds sheet rows: " + (row + nr - 1) + " > " + sheet.maxRows);
      if (col + nc - 1 > sheet.maxCols) throw new Error("Range exceeds sheet columns: " + (col + nc - 1) + " > " + sheet.maxCols);
      Object.assign(this, { sheet, row, col, nr, nc });
    }
    getValues() {
      if (this.nr === 1 && this.row > 1) state.rowReads++;
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
      state.events.push("write:" + this.sheet.name);
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
    getDataRange() {
      state.sheetReads[this.name] = (state.sheetReads[this.name] || 0) + 1;
      return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
    }
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
      openById(id) { const ss = state.spreadsheets[id]; if (!ss) throw new Error("Spreadsheet not found: " + id); state.opens = (state.opens || 0) + 1; return ss; },
      flush() { state.flushes++; state.events.push("flush"); }
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
          remove: k => { delete state.cache[k]; },
          // Batch calls (read cache / data versions); tracked apart from put().
          getAll: keys => { if (state.cacheDown) throw new Error("Cache down"); state.cacheGetAlls++; const o = {}; keys.forEach(k => { if (k in state.cache) o[k] = state.cache[k]; }); return o; },
          putAll: (map, ttl) => {
            if (state.cacheDown) throw new Error("Cache down");
            Object.keys(map).forEach(k => {
              if (/^PALLET_V_/.test(k)) state.events.push("version:" + k.slice(9));
              const v = String(map[k]);
              assert.ok(k.length <= 250, "cache key too long: " + k);
              assert.ok(Buffer.byteLength(v, "utf8") <= 100 * 1024, "cache value over 100 KB: " + k);
              state.cache[k] = v;
            });
            state.cachePutAlls.push(ttl);
          },
          removeAll: keys => { keys.forEach(k => { delete state.cache[k]; }); }
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
    // Everything Code.gs logs is kept in state.logs (checked for leaked secrets).
    console: {
      log(...a) { state.logs.push(["log", a.join(" ")]); },
      warn(...a) { state.logs.push(["warn", a.join(" ")]); },
      error(...a) { state.logs.push(["error", a.join(" ")]); console.error(...a); }
    },
    // Teams webhook: never touches the network. Records each call (and whether the
    // script lock was held at that moment); answers state.fetchStatus or throws
    // state.fetchError.
    UrlFetchApp: {
      fetch(url, opts) {
        state.fetches.push({ url, opts, lockHeld: state.locks !== state.unlocks });
        if (state.fetchError) throw new Error(state.fetchError);
        const code = state.fetchStatus;
        return { getResponseCode: () => code, getContentText: () => "" };
      }
    },
    // Project triggers (keep-warm). state.triggers: [{handler, everyMinutes, id}]
    ScriptApp: {
      getProjectTriggers() {
        return (state.triggers || []).map(t => ({ getHandlerFunction: () => t.handler, getUniqueId: () => t.id, _t: t }));
      },
      deleteTrigger(tr) { state.triggers = (state.triggers || []).filter(t => t !== tr._t); },
      newTrigger(handler) {
        const t = { handler, id: "trg" + (++state.seq) };
        const b = {
          timeBased() { t.timeBased = true; return b; },
          everyMinutes(n) { assert.ok([1, 5, 10, 15, 30].includes(n), "everyMinutes(" + n + ")"); t.everyMinutes = n; return b; },
          create() { assert.ok(t.timeBased, "time-based"); (state.triggers = state.triggers || []).push(t); return { getUniqueId: () => t.id }; }
        };
        return b;
      }
    }
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
// No login: the name typed in the page header is sent as "actor" with every write.
const T = { tester: "Tester", admin: "Admin", editor: "Editor" };
const RPW = "reset-Only-9"; // PALLET_RESET_PASSWORD (also guards the audit log)
const LOG_ACTIONS = ["logs", "logs_export"];
// get(): GET read. The audit log is POST-only behind the reset password, so
// logs / logs_export are sent as POST with the current reset password.
const get = (action, params = {}) => {
  if (LOG_ACTIONS.includes(action)) return postRaw(Object.assign({ action, resetPassword: state.props.PALLET_RESET_PASSWORD }, params), null);
  const out = ctx.doGet({ parameter: Object.assign({ action }, params) });
  assert.strictEqual(out.getMimeType(), "application/json");
  return JSON.parse(out.getContent());
};
// postRaw(payload, who): who = the typed name sent as actor (null = none).
const postRaw = (payload, who = T.tester) =>
  JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(Object.assign(who ? { actor: who } : {}, payload)) } }).getContent());
const PW = "s3cret-Pallet"; // just some other password (never the reset password)
const TESTER = "Tester", EDITOR = "Editor", ADMIN = "Admin";
// Department settings sit behind the reset password: tests send it by default.
const DEPT_ACTIONS = ["dept_save", "dept_delete"];
const post = (action, body = {}, who = T.tester) => postRaw(Object.assign(DEPT_ACTIONS.includes(action) ? { resetPassword: RPW } : {}, body, { action }), who);
const okData = (res) => { assert.strictEqual(res.ok, true, "expected ok, got: " + JSON.stringify(res)); return res.data; };
const err = (res, re) => {
  assert.strictEqual(res.ok, false, "expected error, got: " + JSON.stringify(res));
  if (re) assert.ok(re.test(res.error), "error '" + res.error + "' !~ " + re);
  return res;
};
// Forget reset-password failure counters ("15 minutes later").
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

console.log("Pallet Hub GAS backend tests (today " + TODAY + " Asia/Bangkok)");

/* ---------- setup ---------- */
test("API before setup: reads and writes report the setupSystem error", () => {
  err(get("bootstrap"), /setupSystem/);
  err(post("receive", { type_id: 1, size: "1.2x1.2", qty: 1 }), /setupSystem/);
});

test("setupSystem creates spreadsheet, sheets, headers, seed data (no users sheet)", () => {
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, true);
  assert.ok(state.props.PALLET_SPREADSHEET_ID);
  const ss = state.spreadsheets[state.props.PALLET_SPREADSHEET_ID];
  assert.strictEqual(ss.tz, "Asia/Bangkok");
  assert.deepStrictEqual(ss.getSheets().map(s => s.getName()), ["pallet_types", "departments", "repairs", "movements", "audit_logs"]);
  assert.deepStrictEqual(sheet("movements").data[0], ["id", "doc_no", "action", "type_id", "size", "qty", "from_status", "to_status", "department", "person", "note", "repair_id", "moved_at", "created_at", "actor", "username"]);
  assert.deepStrictEqual(sheet("audit_logs").data[0], ["id", "category", "action", "ref", "detail", "actor", "ip", "created_at", "username"]);
  assert.strictEqual(dataRows("pallet_types"), 4);
  assert.strictEqual(dataRows("departments"), 8);
  assert.ok(!("users" in r) && !("initialAdminPasswordConfigured" in r));
  assert.strictEqual(r.resetPasswordConfigured, false);
  state.props.PALLET_RESET_PASSWORD = RPW;
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

/* ---------- no login ---------- */
test("no login: reads and normal writes work without any password or session (old token fields ignored)", () => {
  for (const a of ["bootstrap", "dashboard", "repairs", "history", "export"]) {
    okData(get(a));
    okData(postRaw({ action: a }, null));
    okData(get(a, { token: "f".repeat(64) }));
  }
  const deps = okData(get("bootstrap")).departments;
  const d = okData(postRaw({ action: "dept_save", name: "ฝ่ายผลิต", icon: "fa-industry", color: "#E2231A", token: "garbage", password: "x", actor: "Tester", resetPassword: RPW }, null));
  assert.strictEqual(d.message, "เพิ่ม ฝ่ายผลิต แล้ว");
  assert.deepStrictEqual(okData(get("bootstrap")).departments, deps);
  err(postRaw({ action: "dept_delete", id: 999, actor: "Tester", resetPassword: RPW }, null), /^ไม่พบฝ่าย$/);
  assert.ok(!sheet("users"), "no users sheet is created");
});

test("unknown / removed actions are rejected; POST-only actions via GET are rejected", () => {
  err(get("nope"), /^Unknown action$/);
  err(postRaw({ action: "nope" }), /^Unknown action$/);
  for (const a of ["login", "logout", "me", "users", "change_password", "user_save", "user_reset_password", "user_toggle", "verifyPassword"]) {
    err(get(a), /^Unknown action$/);
    err(postRaw({ action: a, username: "admin", password: "x" }), /^Unknown action$/);
  }
  for (const a of ["receive", "reset_data", "movement_delete", "verifyResetPassword", "logs", "logs_export"]) {
    const out = JSON.parse(ctx.doGet({ parameter: { action: a, resetPassword: RPW } }).getContent());
    err(out, /POST/);
  }
  assert.strictEqual(okData(get(undefined)).service, "Pallet Hub API");
});

test("invalid JSON body is rejected", () => {
  const r = JSON.parse(ctx.doPost({ postData: { contents: "{bad" } }).getContent());
  err(r, /JSON/);
});

test("audit log (logs / logs_export) needs the reset password: POST only, wrong = 1 s + shared lockout", () => {
  clearFails();
  const before = JSON.stringify(sheet("audit_logs").data);
  for (const a of ["logs", "logs_export"]) {
    okData(postRaw({ action: a, resetPassword: RPW }, null));
    for (const pw of [undefined, "", "wrong", PW]) {
      const sleeps = state.sleeps;
      const r = err(postRaw({ action: a, resetPassword: pw }, null), /รหัส/);
      assert.strictEqual(r.passwordError, true);
      if (pw) assert.strictEqual(state.sleeps - sleeps, 1000);
    }
    err(postRaw({ action: a, password: RPW }, null), /รหัสไม่ถูกต้อง/); // must come in resetPassword
  }
  assert.strictEqual(JSON.stringify(sheet("audit_logs").data), before); // reading / refusing never writes the log
  // the same lockout counter as reset_data and record edits
  assert.ok(Number(state.cache.PALLET_RESET_FAILURES) >= 6);
  for (let i = 0; i < 10; i++) postRaw({ action: "logs", resetPassword: "guess" + i }, null);
  err(postRaw({ action: "logs", resetPassword: RPW }, null), /หลายครั้งเกินไป/);
  err(postRaw({ action: "reset_data", resetPassword: RPW }), /หลายครั้งเกินไป/);
  okData(get("bootstrap")); // normal use unaffected
  clearFails();
  // reset password not configured: Thai admin message
  delete state.props.PALLET_RESET_PASSWORD;
  err(postRaw({ action: "logs", resetPassword: RPW }, null), /PALLET_RESET_PASSWORD/);
  state.props.PALLET_RESET_PASSWORD = RPW;
  // never inside a batch (the password would have to travel in a URL)
  err(get("batch", { reads: JSON.stringify(["logs"]) }), /^Unknown action$/);
  clearFails();
});

/* ---------- receive ---------- */
test("receive: creates RC doc, stock available, records the typed name (actor) and the person", () => {
  const d = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 50, person: "  สมชาย ใจดี  ", note: "PO-123", username: "admin" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0001");
  assert.strictEqual(d.id, 1);
  assert.strictEqual(d.message, "รับเข้า 50 ตัว เรียบร้อย");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 50);
  const m = okData(get("history", { q: d.doc_no })).items[0];
  assert.deepStrictEqual([m.actor, m.username, m.person], [TESTER, "", "สมชาย ใจดี"]);
  const log = lastLog();
  assert.strictEqual(log.category, "pallet");
  assert.strictEqual(log.action, "receive");
  assert.strictEqual(log.ref, d.doc_no);
  assert.strictEqual(log.actor, TESTER);
  assert.strictEqual(log.username, "");
  assert.strictEqual(log.ip, "web");
  assert.strictEqual(log.detail, `รับเข้า RM (พาเลทสำหรับใส่ RM) ขนาด 1.2x1.2 ม. จำนวน 50 ตัว [ภายนอก → พร้อมใช้] · เวลาทำรายการ ${TODAY.slice(8, 10)}/${TODAY.slice(5, 7)}/${TODAY.slice(0, 4)} ${HOUR}:00 · ชื่อที่ระบุ: สมชาย ใจดี · PO-123`);
});

test("receive: second doc number increments; actor is the typed name even without a person", () => {
  const d = okData(post("receive", Object.assign({ type_id: 2, size: "1.1x1.1", qty: "30", person: "" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0002");
  assert.strictEqual(lastLog().actor, TESTER);
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
  assert.strictEqual(ticket1.reported_by, TESTER); // the typed name (actor), not the person "B"
  assert.strictEqual(ticket1.reported_username, "");
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
  assert.deepStrictEqual([t.updated_by, t.updated_username], [TESTER, ""]);
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
  assert.strictEqual(row, `RC-${YMD}-0001,${TODAY},${HOUR}:00,รับเข้า,RM,"พาเลทสำหรับใส่ RM",1.2x1.2,50,,Tester,,"สมชาย ใจดี",PO-123`);
  assert.strictEqual(ctx.csvLine_(['a"b', "c\\\"d", null, 5, "x,y"]), '"a""b","c\\"d",,5,"x,y"\n');
});

/* ---------- logs ---------- */
test("logs: categories, filters, order and CSV export", () => {
  const all = okData(get("logs", {})).items;
  assert.strictEqual(all.length, dataRows("audit_logs"));
  for (let i = 1; i < all.length; i++) assert.ok(all[i - 1].id > all[i].id);
  const cats = new Set(all.map(i => i.category));
  ["pallet", "repair", "warn"].forEach(c => assert.ok(cats.has(c), c));
  assert.ok(okData(get("logs", { q: "Tester" })).items.length > 0);
  assert.ok(okData(get("logs")).items.every(l => (l.actor === TESTER || l.actor === "") && l.username === ""));
  assert.strictEqual(okData(get("logs", { from: "2099-01-01" })).items.length, 0);
  const x = okData(get("logs_export", { cat: "warn" }));
  assert.ok(/^pallet_log_\d{8}_\d{6}\.csv$/.test(x.filename));
  const lines = x.csv.trim().split("\n");
  assert.strictEqual(lines[0], "ลำดับ,วันที่,เวลา,หมวด,เลขที่อ้างอิง,รายละเอียด,ผู้ทำรายการ,ชื่อผู้ใช้,IP");
  assert.ok(lines.slice(1).every(l => l.includes(",Tester,,web")));
  assert.ok(lines.slice(1).every(l => l.includes(",ถูกปฏิเสธ,") && l.endsWith(",web")));
});

/* ---------- departments ---------- */
test("dept_save: add new, update existing (case-insensitive), validation", () => {
  // without / with a wrong reset password: refused, nothing saved
  const depsBefore = dataRows("departments");
  err(postRaw({ action: "dept_save", name: "ไม่มีรหัส", actor: "Tester" }, null));
  err(postRaw({ action: "dept_save", name: "รหัสผิด", actor: "Tester", resetPassword: "nope" }, null), /รหัส/);
  err(postRaw({ action: "dept_delete", id: 1, actor: "Tester" }, null));
  assert.strictEqual(dataRows("departments"), depsBefore);
  clearFails();
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
  ctx.clearReadCache(); // edited by hand (not through the API): drop the read cache as documented
  assert.strictEqual(okData(get("history", { q: "RC-" + YMD + "-0001" })).items[0].moved_at, "2026-01-02 03:04:05");
  s.data[row][col] = orig;
  ctx.clearReadCache();
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

/* ---------- reset data ---------- */
const reset = (pw, extra = {}, who = T.admin) => postRaw(Object.assign({ action: "reset_data", resetPassword: pw }, extra), who);
const snapshot = () => ["movements", "repairs", "audit_logs", "pallet_types", "departments"].map(dataRows);

test("reset_data refused while PALLET_RESET_PASSWORD is unset (Thai admin message); GET refused", () => {
  delete state.props.PALLET_RESET_PASSWORD;
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
  err(reset(PW, { password: PW }), /รหัสไม่ถูกต้อง/); // another password is not the reset password
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
  const items = okData(get("logs", {})).items;
  assert.strictEqual(items.length, 1);
  const l = items[0];
  assert.strictEqual(l.id, 1);
  assert.strictEqual(l.category, "setting");
  assert.strictEqual(l.action, "รีเซ็ตข้อมูล");
  assert.strictEqual(l.actor, ADMIN);
  assert.strictEqual(l.username, "");
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
  const logs = okData(get("logs", {})).items;
  assert.deepStrictEqual(logs.map(l => l.id), [4, 3, 2, 1]);
  assert.strictEqual(logs[3].action, "รีเซ็ตข้อมูล");
});

test("reset_data on already-empty sheets reports zero rows removed", () => {
  okData(reset(RPW));
  const r = okData(reset(RPW));
  assert.deepStrictEqual(r.removed, { movements: 0, repairs: 0, audit_logs: 1 });
  assert.strictEqual(okData(get("logs", {})).items.length, 1);
});

/* ---------- edit / delete records (reset password) ---------- */
const rec = (action, body = {}, pw = RPW) => postRaw(Object.assign({}, body, { action, resetPassword: pw }), T.editor);
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
  assert.deepStrictEqual([byDoc(D.r1.doc_no).actor, byDoc(D.r1.doc_no).username], [TESTER, ""]);
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
  err(rec("repair_update", { id: D.x.id, reported_by: "ผู้แจ้งปลอม" }), /บันทึกจากชื่อผู้ใช้งานโดยอัตโนมัติ/);
  err(rec("repair_update", { id: D.x.id, updated_by: "x" }), /แก้ไขไม่ได้/);
  const d = okData(rec("repair_update", { id: D.x.id, ticket_no: D.x.ticket_no, cause: "ไม้หักสองแผ่น", repairer: "ช่างสอง", reported_by: TESTER, note: "บันทึก\nบรรทัดสอง" }));
  assert.strictEqual(d.message, "แก้ไขใบแจ้งซ่อม " + D.x.ticket_no + " แล้ว");
  const t = ticketOf(D.x.id);
  assert.deepStrictEqual([t.cause, t.repairer, t.reported_by, t.note, t.qty, t.stage], ["ไม้หักสองแผ่น", "ช่างสอง", TESTER, "บันทึก\nบรรทัดสอง", 4, "done"]);
  assert.deepStrictEqual([t.updated_by, t.updated_username], [EDITOR, ""]);
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

/* ---------- attribution (name typed in the page header) ---------- */
test("actor = the typed name (trimmed, max 100, required); person / reported_by from the client never replace it", () => {
  const d = okData(postRaw(Object.assign({ action: "receive", type_id: 3, size: "1.2x1.2", qty: 5, username: "admin", person: "ใครก็ได้" }, AT), "  สมชาย ใจดี  "));
  let m = byDoc(d.doc_no);
  assert.deepStrictEqual([m.actor, m.username, m.person], ["สมชาย ใจดี", "", "ใครก็ได้"]);
  assert.strictEqual(lastLog().actor, "สมชาย ใจดี");
  const dm = okData(post("damage", Object.assign({ type_id: 3, size: "1.2x1.2", qty: 1, cause: "แตก", person: "Fake", reported_by: "Fake", reported_username: "admin" }, AT), T.editor));
  const t = rowsOf("repairs").pop();
  assert.deepStrictEqual([t.reported_by, t.reported_username], [EDITOR, ""]);
  assert.deepStrictEqual([byDoc(dm.doc_no).actor, byDoc(dm.doc_no).person], [EDITOR, "Fake"]);
  // repair_start without a repairer name: the typed name becomes the repairer
  const rp = okData(post("repair_start", Object.assign({ id: t.id }, AT)));
  const t2 = ticketOf(t.id);
  assert.deepStrictEqual([t2.reported_by, t2.updated_by, t2.updated_username, t2.repairer], [EDITOR, TESTER, "", TESTER]);
  assert.strictEqual(byDoc(rp.doc_no).actor, TESTER);
  // no name typed (missing, blank or not a string): refused, nothing written; long names are cut at 100 characters
  const before = dataRows("movements");
  for (const actor of [undefined, "   ", { x: 1 }]) {
    const body = Object.assign({ action: "receive", type_id: 3, size: "1.2x1.2", qty: 1 }, AT);
    if (actor !== undefined) body.actor = actor;
    const r = err(postRaw(body, null), /กรุณาระบุชื่อผู้ใช้งาน/);
    assert.strictEqual(r.code, "NAME_REQUIRED");
  }
  assert.strictEqual(dataRows("movements"), before, "no movement written without a name");
  const n3 = okData(post("receive", Object.assign({ type_id: 3, size: "1.2x1.2", qty: 1 }, AT), "ก".repeat(150)));
  assert.strictEqual(byDoc(n3.doc_no).actor, "ก".repeat(100));
  // refused writes are logged with the typed name too
  err(post("issue", Object.assign({ type_id: 3, size: "1.2x1.2", qty: 99999, department: "ฝ่ายผลิต" }, AT), T.editor));
  assert.deepStrictEqual([lastLog().category, lastLog().actor], ["warn", EDITOR]);
  okData(rec("repair_delete", { id: t.id })); // tidy up
});

test("rows recorded during the login period keep their stored names", () => {
  const s = sheet("movements");
  const h = s.data[0];
  const id = Math.max(...rowsOf("movements").map(r => r.id)) + 1;
  const row = h.map(k => ({ id, doc_no: "RC-260101-0999", action: "receive", type_id: 4, size: "1.2x1.2", qty: 1, to_status: "available", person: "", note: "",
    moved_at: "2026-01-01 08:00:00", created_at: "2026-01-01 08:00:00", actor: "ผู้ดูแลระบบ (admin)", username: "admin" }[k] ?? ""));
  s.data.push(row);
  s.maxRows = Math.max(s.maxRows, s.data.length + 1);
  const lg = sheet("audit_logs");
  lg.maxRows = Math.max(lg.maxRows, lg.data.length + 2);
  lg.data.push(lg.data[0].map(k => ({ id: 99999, category: "account", action: "login", ref: "admin", detail: "เข้าสู่ระบบ", actor: "ผู้ดูแลระบบ (admin)", ip: "web", created_at: "2026-01-01 08:00:00", username: "admin" }[k] ?? "")));
  ctx.clearReadCache(); // edited by hand
  const m = byDoc("RC-260101-0999");
  assert.deepStrictEqual([m.actor, m.username], ["ผู้ดูแลระบบ (admin)", "admin"]);
  const l = okData(get("logs", { cat: "account" })).items[0];
  assert.deepStrictEqual([l.actor, l.username, l.detail], ["ผู้ดูแลระบบ (admin)", "admin", "เข้าสู่ระบบ"]);
  assert.ok(okData(get("logs_export", { cat: "account" })).csv.includes("บัญชีผู้ใช้"));
  // still editable with the reset password; the stored actor is kept
  okData(rec("movement_update", { id: m.id, doc_no: m.doc_no, note: "แก้ไข" }));
  assert.deepStrictEqual([byDoc("RC-260101-0999").actor, byDoc("RC-260101-0999").note], ["ผู้ดูแลระบบ (admin)", "แก้ไข"]);
  assert.strictEqual(lastLog().actor, EDITOR);
});

/* ---------- batch reads, reads returned with writes, read cache ---------- */
const CDate = vm.runInContext("Date", ctx);
// Runs fn with the server clock stopped (time-dependent reads compare equal).
const frozen = fn => { const real = CDate.now; const t = real(); CDate.now = () => t; try { return fn(); } finally { CDate.now = real; } };
// The same request answered without the read cache (CacheService failing -> live sheets).
const live = fn => { state.cacheDown = true; try { return fn(); } finally { state.cacheDown = false; } };
const enc = reads => ({ reads: JSON.stringify(reads) });
const vers = () => Object.fromEntries(Object.keys(state.cache).filter(k => /^PALLET_V_/.test(k)).map(k => [k.slice(9), state.cache[k]]));
// Every read the pages use, as [action, params].
const PAGE_READS = () => [
  ["bootstrap", {}], ["dashboard", { days: "7" }], ["dashboard", { days: "30" }], ["repairs", {}],
  ["history", { from: TODAY.slice(0, 8) + "01", to: TODAY, type: "", act: "", dept: "", q: "" }], ["history", {}],
  ["history", { q: "RC", act: "receive" }], ["logs", { from: TODAY.slice(0, 8) + "01", to: TODAY, q: "", cat: "" }],
  ["logs", { cat: "warn" }], ["export", {}], ["logs_export", {}]
];
// Cached answers (cache warmed by the first pass) must equal the live answers.
const assertReadsFresh = label => frozen(() => {
  for (const [a, p] of PAGE_READS()) {
    const want = live(() => okData(get(a, p)));
    assert.deepStrictEqual(okData(get(a, p)), want, label + ": " + a + " " + JSON.stringify(p));
    assert.deepStrictEqual(okData(get(a, p)), want, label + " (2nd): " + a);
  }
});

test("batch: one request returns exactly the single reads' data (GET and POST), in order", () => frozen(() => {
  const reads = [
    { action: "bootstrap" }, { action: "dashboard", days: "14" }, { action: "repairs" },
    { action: "history", q: "RC", from: TODAY.slice(0, 8) + "01", to: TODAY }, { action: "history" }
  ];
  const single = reads.map(r => okData(get(r.action, Object.fromEntries(Object.entries(r).filter(([k]) => k !== "action")))));
  assert.deepStrictEqual(okData(get("batch", enc(reads))).results, single);
  assert.deepStrictEqual(okData(postRaw({ action: "batch", reads }, null)).results, single);
  // plain action names work for reads without params
  assert.deepStrictEqual(okData(get("batch", enc(["bootstrap", "repairs"]))).results, [okData(get("bootstrap")), okData(get("repairs"))]);
}));

test("batch: bad lists refused; logs / exports / writes never inside a batch; no lock, nothing written", () => {
  const locks = state.locks, before = JSON.stringify(state.spreadsheets[state.props.PALLET_SPREADSHEET_ID].sheets.map(x => x.data));
  err(get("batch", {}), /reads ไม่ถูกต้อง/);
  err(get("batch", { reads: "not json" }), /reads ไม่ถูกต้อง/);
  err(get("batch", enc([])), /reads ไม่ถูกต้อง/);
  err(get("batch", enc({ action: "bootstrap" })), /reads ไม่ถูกต้อง/);
  err(get("batch", enc(Array(7).fill("bootstrap"))), /reads ไม่ถูกต้อง/);
  err(get("batch", enc([["bootstrap"]])), /reads ไม่ถูกต้อง/);
  for (const a of ["batch", "export", "logs", "logs_export", "receive", "reset_data", "me", "users", "nope"]) err(get("batch", enc([a])), /^Unknown action$/);
  assert.strictEqual(state.locks, locks); // reads (single or batch) never take the script lock
  assert.strictEqual(JSON.stringify(state.spreadsheets[state.props.PALLET_SPREADSHEET_ID].sheets.map(x => x.data)), before);
});

test("read cache: cached reads equal live reads; a warm read opens the spreadsheet but reads no sheet", () => {
  assertReadsFresh("warm");
  state.sheetReads = {};
  frozen(() => okData(get("batch", enc(["bootstrap", { action: "dashboard", days: "7" }, "repairs", "history"]))));
  assert.deepStrictEqual(state.sheetReads, {}); // everything from the cache
  // bootstrap alone comes from the derived cache: one getAll (versions) + one get
  const gets = state.cacheGetAlls, opens = state.opens;
  okData(get("bootstrap"));
  assert.deepStrictEqual(state.sheetReads, {});
  assert.strictEqual(state.cacheGetAlls - gets, 1);
  assert.strictEqual(state.opens, opens); // not even openById
  assert.ok(Object.keys(state.cache).some(k => /^PALLET_RC_movements_/.test(k)));
});

test("read cache: big sheets are chunked (<= 100 KB per value) or not cached; cache failures fall back to the sheets", () => {
  const chunk = ctx.RC_CHUNK, max = ctx.RC_MAX_CHUNKS;
  try {
    ctx.RC_CHUNK = 700; // force several chunks per sheet
    ctx.clearReadCache();
    assertReadsFresh("chunked");
    assert.ok(Object.keys(state.cache).some(k => /^PALLET_RC_audit_logs_.*_\d+$/.test(k)), "audit_logs stored in chunks");
    // a missing chunk is a cache miss (never partial data)
    Object.keys(state.cache).filter(k => /^PALLET_RC_movements_.*_1$/.test(k)).forEach(k => delete state.cache[k]);
    assertReadsFresh("chunk evicted");
    ctx.RC_MAX_CHUNKS = 2; // too big -> read live every time, still correct
    ctx.clearReadCache();
    assertReadsFresh("too big");
    // versions evicted -> new versions, old copies are never used again
    Object.keys(state.cache).filter(k => /^PALLET_V_/.test(k)).forEach(k => delete state.cache[k]);
    assertReadsFresh("versions evicted");
    state.cacheDown = true; // CacheService unavailable: reads and writes still work
    try {
      okData(get("dashboard"));
      okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1 }, AT)));
    } finally { state.cacheDown = false; }
    assertReadsFresh("after cache outage");
  } finally {
    ctx.RC_CHUNK = chunk; ctx.RC_MAX_CHUNKS = max;
  }
});

test("read cache invalidation: after every kind of write the next reads are fresh; only changed sheets get new versions", () => {
  state.props.PALLET_RESET_PASSWORD = RPW;
  clearFails();
  const recA = (action, body) => postRaw(Object.assign({}, body, { action, resetPassword: RPW }), T.admin);
  const step = (label, fn, mustBump, mustKeep = []) => {
    assertReadsFresh("before " + label); // cache warm
    const v0 = vers();
    state.events = [];
    fn();
    const v1 = vers();
    for (const t of mustBump) assert.notStrictEqual(v1[t], v0[t], label + ": version of " + t + " not bumped");
    for (const t of mustKeep) assert.strictEqual(v1[t], v0[t], label + ": version of " + t + " bumped needlessly");
    // the version is announced only after the sheet writes were flushed
    const lastWrite = state.events.map(e => /^write:/.test(e)).lastIndexOf(true);
    const firstVer = state.events.findIndex(e => /^version:/.test(e));
    if (lastWrite !== -1) assert.ok(firstVer > state.events.lastIndexOf("flush") && state.events.lastIndexOf("flush") > lastWrite, label + ": " + state.events.join(","));
    assertReadsFresh("after " + label);
  };
  const base = { type_id: 2, size: "1.1x1.1" };
  let ticket;
  step("receive", () => okData(post("receive", Object.assign({}, base, { qty: 30 }, AT))), ["movements", "audit_logs"], ["pallet_types", "departments", "repairs"]);
  step("issue", () => okData(post("issue", Object.assign({}, base, { qty: 10, department: "ฝ่ายผลิต" }, AT))), ["movements", "audit_logs"], ["repairs"]);
  step("return good", () => okData(post("return", Object.assign({}, base, { qty: 2, department: "ฝ่ายผลิต" }, AT))), ["movements", "audit_logs"]);
  step("return damaged", () => okData(post("return", Object.assign({}, base, { qty: 2, department: "ฝ่ายผลิต", condition: "damaged", cause: "หัก" }, AT))), ["movements", "repairs", "audit_logs"]);
  step("damage", () => okData(post("damage", Object.assign({}, base, { qty: 3, cause: "แตก" }, AT))), ["movements", "repairs", "audit_logs"]);
  ticket = okData(get("repairs")).items.find(r => r.stage === "damaged" && r.qty === 3);
  step("repair_start", () => okData(post("repair_start", Object.assign({ id: ticket.id }, AT))), ["movements", "repairs"]);
  step("repair_done", () => okData(post("repair_done", Object.assign({ id: ticket.id, note: "ok" }, AT))), ["movements", "repairs"]);
  const t2 = okData(get("repairs")).items.find(r => r.stage === "damaged");
  step("scrap", () => okData(post("scrap", Object.assign({ id: t2.id }, AT))), ["movements", "repairs"]);
  step("refused write (warn row)", () => err(post("issue", Object.assign({}, base, { qty: 99999, department: "ฝ่ายผลิต" }, AT))), ["audit_logs"], ["movements", "repairs"]);
  step("dept_save", () => okData(post("dept_save", { name: "ฝ่ายแคช", icon: "fa-store", color: "#0EA5E9" })), ["departments", "audit_logs"], ["movements"]);
  const dep = okData(get("bootstrap")).departments.find(d => d.name === "ฝ่ายแคช");
  step("dept_delete", () => okData(post("dept_delete", { id: dep.id })), ["departments", "audit_logs"], ["movements"]);
  const rc = hist().find(m => m.action === "receive" && m.qty === 30);
  step("movement_update", () => okData(recA("movement_update", { id: rc.id, doc_no: rc.doc_no, qty: 31, note: "แก้" })), ["movements", "audit_logs"], ["departments"]);
  const chainMove = hist().find(m => m.repair_id === ticket.id && m.action === "damage");
  step("movement_update (chain qty)", () => okData(recA("movement_update", { id: chainMove.id, doc_no: chainMove.doc_no, qty: 4 })), ["movements", "repairs", "audit_logs"]);
  step("repair_update", () => okData(recA("repair_update", { id: ticket.id, cause: "แตกมาก" })), ["repairs", "audit_logs"], ["movements"]);
  step("repair_delete (chain)", () => okData(recA("repair_delete", { id: ticket.id })), ["movements", "repairs", "audit_logs"]);
  const rt = hist().find(m => m.action === "return");
  step("movement_delete", () => okData(recA("movement_delete", { id: rt.id, doc_no: rt.doc_no })), ["movements", "audit_logs"]);
  step("refused edit (wrong reset password)", () => err(postRaw({ action: "movement_delete", id: rc.id, resetPassword: "nope" })), [], ["movements", "audit_logs", "repairs"]);
  clearFails();
  step("reset_data", () => okData(recA("reset_data", {})), ["movements", "repairs", "audit_logs"], ["pallet_types", "departments"]);
  assert.deepStrictEqual(okData(get("history")).items, []);
  step("receive after reset", () => okData(post("receive", Object.assign({}, base, { qty: 5 }, AT))), ["movements", "audit_logs"]);
  step("clearReadCache()", () => ctx.clearReadCache(), ["pallet_types", "departments", "repairs", "movements", "audit_logs"]);
});

test("writes can return reads computed after the write (same data as a fresh read)", () => frozen(() => {
  const reads = [{ action: "bootstrap" }, { action: "dashboard", days: "7" }];
  const r = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 4, reads }, AT)));
  assert.ok(/^RC-/.test(r.doc_no));
  assert.deepStrictEqual(r.reads, [okData(get("bootstrap")), okData(get("dashboard", { days: "7" }))]);
  assert.deepStrictEqual(r.reads, live(() => [okData(get("bootstrap")), okData(get("dashboard", { days: "7" }))]));
  // JSON text and plain names are accepted too
  assert.deepStrictEqual(okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1, reads: JSON.stringify(["repairs"]) }, AT))).reads,
    [okData(get("repairs"))]);
  // without "reads" the result is exactly as before (no extra key)
  assert.deepStrictEqual(Object.keys(okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1 }, AT)))), ["doc_no", "id", "message"]);
  // reset-password actions
  const m = hist()[0];
  const u = okData(postRaw({ action: "movement_update", id: m.id, doc_no: m.doc_no, note: "x", resetPassword: RPW, reads: ["bootstrap", "history"] }));
  assert.deepStrictEqual(u.reads, [okData(get("bootstrap")), okData(get("history"))]);
  const d = okData(post("dept_save", { name: "ฝ่ายหลังบันทึก", reads: ["bootstrap"] }));
  assert.ok(d.reads[0].departments.some(x => x.name === "ฝ่ายหลังบันทึก"));
}));

test("reads with writes: invalid list refuses the write before anything is written; refused writes return no reads", () => {
  const n = dataRows("movements"), logsN = dataRows("audit_logs");
  for (const bad of ["nope", [], ["export"], ["logs"], [{ action: "batch" }], Array(7).fill("bootstrap")]) {
    err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1, reads: bad }, AT)));
  }
  assert.strictEqual(dataRows("movements"), n);
  assert.strictEqual(dataRows("audit_logs"), logsN);
  const e = err(post("issue", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 999999, department: "ฝ่ายผลิต", reads: ["bootstrap"] }, AT)));
  assert.strictEqual(e.reads, undefined);
  const w = err(postRaw({ action: "reset_data", resetPassword: "nope", reads: ["bootstrap"] }));
  assert.strictEqual(w.reads, undefined);
  clearFails();
});

test("updating rows reuses the values read under the lock (no per-row re-read)", () => {
  const m = hist().find(x => x.action === "receive" && x.repair_id == null);
  const rows = state.rowReads;
  okData(postRaw({ action: "movement_update", id: m.id, doc_no: m.doc_no, note: "no reread", resetPassword: RPW }));
  assert.strictEqual(state.rowReads, rows);
  assert.strictEqual(byDoc(m.doc_no).note, "no reread");
  const raw = rowsOf("movements").find(x => x.id === m.id);
  assert.strictEqual(raw.note, "no reread");
  assert.strictEqual(raw.qty, m.qty);
});

/* ---------- existing spreadsheets (pre-login and login versions) ---------- */
const LEGACY_HEADERS = {
  pallet_types: ["id", "tkey", "code", "name", "short", "description", "color", "sizes", "sort"],
  departments: ["id", "name", "icon", "color", "active"],
  repairs: ["id", "ticket_no", "type_id", "size", "qty", "stage", "source", "department", "cause", "reported_by", "repairer", "reported_at", "started_at", "finished_at", "note"],
  movements: ["id", "doc_no", "action", "type_id", "size", "qty", "from_status", "to_status", "department", "person", "note", "repair_id", "moved_at", "created_at"],
  audit_logs: ["id", "category", "action", "ref", "detail", "actor", "ip", "created_at"]
};
const LOGIN_HEADERS = {
  repairs: LEGACY_HEADERS.repairs.concat(["reported_username", "updated_by", "updated_username", "updated_at"]),
  movements: LEGACY_HEADERS.movements.concat(["actor", "username"]),
  audit_logs: LEGACY_HEADERS.audit_logs.concat(["username"]),
  users: ["id", "username", "password_hash", "salt", "fullname", "role", "active", "created_at", "updated_at", "last_login"]
};
// version "pre": sheets from before the login version; "login": sheets written by the login version.
function legacyInstance(version) {
  const g = createGas();
  const c = loadCode(g.gas);
  const ss = g.gas.SpreadsheetApp.create("Pallet Hub Database");
  ss.sheets = [];
  g.state.props.PALLET_SPREADSHEET_ID = ss.getId();
  g.state.props.PALLET_RESET_PASSWORD = RPW;
  const headers = version === "login" ? Object.assign({}, LEGACY_HEADERS, LOGIN_HEADERS) : LEGACY_HEADERS;
  for (const [name, h] of Object.entries(headers)) {
    const sh = ss.insertSheet(name);
    sh.data.push(h.slice());
    sh.maxCols = h.length; // a tight sheet: new columns need insertColumnsAfter
    sh.frozen = 1;
  }
  const login = version === "login";
  ss.getSheetByName("pallet_types").data.push([1, "RM", "RM", "พาเลท RM", "RM", "วัตถุดิบ", "#1E6FE0", "1.2x1.2", 1]);
  ss.getSheetByName("departments").data.push([1, "ฝ่ายผลิต", "fa-industry", "#E2231A", 1]);
  ss.getSheetByName("movements").data.push([1, "RC-260101-0001", "receive", 1, "1.2x1.2", 10, "", "available", "", "คนเก่า", "ของเดิม", "", "2026-01-01 08:00:00", "2026-01-01 08:00:00"]
    .concat(login ? ["ผู้ดูแลระบบ (admin)", "admin"] : []));
  ss.getSheetByName("audit_logs").data.push([1, "pallet", "receive", "RC-260101-0001", "รับเข้าเดิม", login ? "ผู้ดูแลระบบ (admin)" : "คนเก่า", "web", "2026-01-01 08:00:00"]
    .concat(login ? ["admin"] : []));
  if (login) ss.getSheetByName("users").data.push([1, "admin", "a".repeat(64), "b".repeat(32), "ผู้ดูแลระบบ", "admin", 1, "2026-01-01 07:00:00", "", "2026-01-01 07:59:00"]);
  const G = (action, params = {}) => JSON.parse(c.doGet({ parameter: Object.assign({ action }, params) }).getContent());
  const P = (payload) => JSON.parse(c.doPost({ postData: { contents: JSON.stringify(payload) } }).getContent());
  return { g, c, ss, G, P, sh: n => ss.getSheetByName(n) };
}

test("pre-login spreadsheet: old sheets gain the new columns on first use; old data intact; no users sheet created", () => {
  const L = legacyInstance("pre");
  const h = okData(L.G("history")).items;
  assert.deepStrictEqual(L.sh("movements").data[0], LOGIN_HEADERS.movements);
  assert.deepStrictEqual([h[0].doc_no, h[0].person, h[0].actor, h[0].username, h[0].from_status], ["RC-260101-0001", "คนเก่า", "", "", null]);
  const logs = okData(L.P({ action: "logs", resetPassword: RPW })).items;
  assert.deepStrictEqual(L.sh("audit_logs").data[0], LOGIN_HEADERS.audit_logs);
  assert.deepStrictEqual(logs.find(l => l.id === 1), { id: 1, category: "pallet", action: "receive", ref: "RC-260101-0001", detail: "รับเข้าเดิม", actor: "คนเก่า", ip: "web", created_at: "2026-01-01 08:00:00", username: "" });
  okData(L.P({ action: "receive", actor: "ใหม่", type_id: 1, size: "1.2x1.2", qty: 2 }));
  okData(L.P({ action: "damage", actor: "ใหม่", type_id: 1, size: "1.2x1.2", qty: 1, cause: "x" }));
  assert.deepStrictEqual(L.sh("repairs").data[0], LOGIN_HEADERS.repairs);
  const t = okData(L.G("repairs")).items[0];
  assert.deepStrictEqual([t.reported_by, t.reported_username], ["ใหม่", ""]);
  assert.strictEqual(stockOf(okData(L.G("bootstrap")).stock, 1, "1.2x1.2", "available"), 11);
  assert.ok(L.g.state.colInserts >= 3);
  assert.strictEqual(L.sh("users"), null);
  const r = L.c.setupSystem();
  assert.strictEqual(r.created, false);
  assert.deepStrictEqual(L.ss.getSheets().map(s => s.getName()), ["pallet_types", "departments", "repairs", "movements", "audit_logs"]);
  assert.strictEqual(L.sh("pallet_types").getLastRow(), 2); // no re-seeding over existing data
});

test("login-version spreadsheet: extra username columns tolerated (empty for new rows), users sheet never touched", () => {
  const L = legacyInstance("login");
  const users = JSON.stringify(L.sh("users").data);
  const old = okData(L.G("history")).items[0];
  assert.deepStrictEqual([old.actor, old.username, old.person], ["ผู้ดูแลระบบ (admin)", "admin", "คนเก่า"]);
  const d = okData(L.P({ action: "receive", actor: "สมชาย", type_id: 1, size: "1.2x1.2", qty: 3 }));
  const m = okData(L.G("history", { q: d.doc_no })).items[0];
  assert.deepStrictEqual([m.actor, m.username], ["สมชาย", ""]);
  const row = L.sh("movements").data.find(r => r[1] === d.doc_no);
  assert.deepStrictEqual(row.slice(-2), ["สมชาย", ""]);
  okData(L.P({ action: "damage", actor: "สมชาย", type_id: 1, size: "1.2x1.2", qty: 1, cause: "x" }));
  const t = okData(L.G("repairs")).items[0];
  assert.deepStrictEqual([t.reported_by, t.reported_username, t.updated_by, t.updated_username], ["สมชาย", "", "", ""]);
  const logs = okData(L.P({ action: "logs", resetPassword: RPW })).items;
  assert.deepStrictEqual(logs.map(l => [l.actor, l.username]), [["สมชาย", ""], ["สมชาย", ""], ["ผู้ดูแลระบบ (admin)", "admin"]]);
  okData(L.P({ action: "reset_data", resetPassword: RPW, actor: "สมชาย" }));
  const r = L.c.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(L.g.state.colInserts || 0, 0); // nothing to add
  assert.strictEqual(JSON.stringify(L.sh("users").data), users); // left exactly as it was
  assert.deepStrictEqual(L.ss.getSheets().map(s => s.getName()), ["pallet_types", "departments", "repairs", "movements", "audit_logs", "users"]);
  err(L.P({ action: "login", username: "admin", password: "x" }), /^Unknown action$/);
});

/* ---------- keep-warm trigger ---------- */
// A fresh, set-up instance (own spreadsheet / cache / triggers).
function freshInstance() {
  const g = createGas();
  const c = loadCode(g.gas);
  c.setupSystem();
  g.state.props.PALLET_RESET_PASSWORD = RPW;
  const G = (action, params = {}) => JSON.parse(c.doGet({ parameter: Object.assign({ action }, params) }).getContent());
  const P = (payload, who = T.tester) => JSON.parse(c.doPost({ postData: { contents: JSON.stringify(Object.assign(who ? { actor: who } : {}, payload)) } }).getContent());
  const ss = g.state.spreadsheets[g.state.props.PALLET_SPREADSHEET_ID];
  const dump = () => JSON.stringify(ss.sheets.map(x => x.data));
  const verOf = () => Object.fromEntries(Object.keys(g.state.cache).filter(k => /^PALLET_V_/.test(k)).map(k => [k, g.state.cache[k]]));
  return { g, c, G, P, ss, dump, verOf, rows: n => ss.getSheetByName(n).data.slice(1) };
}

test("keepWarm: refreshes the read caches (live sheet reads, longer TTL) without writing rows or bumping versions", () => {
  const F = freshInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 9 }, AT)));
  okData(F.G("bootstrap")); // versions exist
  const data = F.dump(), v0 = F.verOf(), locks = F.g.state.locks;
  F.g.state.sheetReads = {}; F.g.state.cachePutAlls = []; F.g.state.events = [];
  const r = F.c.keepWarm();
  assert.strictEqual(r.warmed, true);
  assert.strictEqual(F.dump(), data, "no sheet data written");
  assert.deepStrictEqual(F.verOf(), v0, "no version bumped");
  assert.ok(!F.g.state.events.some(e => /^write:|^flush$/.test(e)), F.g.state.events.join(","));
  assert.strictEqual(F.g.state.locks, locks, "no script lock");
  // every data sheet was read live (so hand edits are picked up and the TTL restarts) ...
  for (const n of ["pallet_types", "departments", "repairs", "movements", "audit_logs"]) assert.strictEqual(F.g.state.sheetReads[n], 1, n);
  // ... and stored with a TTL that outlives the 10-minute trigger period
  assert.ok(F.g.state.cachePutAlls.length > 0 && F.g.state.cachePutAlls.every(t => t > 600), JSON.stringify(F.g.state.cachePutAlls));
  // the page reads after a warm run come from the cache only
  F.g.state.sheetReads = {};
  frozenIn(F, () => {
    okData(F.G("batch", { reads: JSON.stringify(["bootstrap", { action: "dashboard", days: "7" }, "repairs", "history"]) }));
  });
  assert.deepStrictEqual(F.g.state.sheetReads, {});
  // and they equal live reads
  const want = (() => { F.g.state.cacheDown = true; try { return okData(F.G("dashboard")); } finally { F.g.state.cacheDown = false; } })();
  assert.deepStrictEqual(okData(F.G("dashboard")).stock, want.stock);
  // a hand edit is visible after the next warm run (no clearReadCache needed)
  const mv = F.ss.getSheetByName("movements");
  mv.data[1][mv.data[0].indexOf("qty")] = 7;
  F.c.keepWarm();
  assert.strictEqual(okData(F.G("bootstrap")).stock[1]["1.2x1.2"].available, 7);
});

test("keepWarm: not set up / cache down -> returns an error object instead of throwing; nothing written", () => {
  const g = createGas();
  const c = loadCode(g.gas);
  const r = c.keepWarm();
  assert.strictEqual(r.warmed, false);
  assert.ok(/setupSystem/.test(r.error));
  const F = freshInstance();
  const data = F.dump();
  F.g.state.cacheDown = true;
  try { assert.strictEqual(F.c.keepWarm().warmed, true); } finally { F.g.state.cacheDown = false; }
  assert.strictEqual(F.dump(), data);
});

test("installKeepWarmTrigger: idempotent, one 10-minute trigger, other triggers untouched; removeKeepWarmTrigger", () => {
  const F = freshInstance();
  F.g.state.triggers = [{ handler: "somethingElse", id: "other", timeBased: true, everyMinutes: 30 }];
  const a = F.c.installKeepWarmTrigger();
  assert.strictEqual(a.installed, true);
  F.c.installKeepWarmTrigger();
  F.c.installKeepWarmTrigger();
  const kw = F.g.state.triggers.filter(t => t.handler === "keepWarm");
  assert.strictEqual(kw.length, 1);
  assert.strictEqual(kw[0].everyMinutes, 10);
  assert.ok(F.g.state.triggers.some(t => t.id === "other"));
  // installing also warms the cache right away
  F.g.state.sheetReads = {};
  frozenIn(F, () => okData(F.G("bootstrap")));
  assert.deepStrictEqual(F.g.state.sheetReads, {});
  assert.strictEqual(F.c.removeKeepWarmTrigger().removed, 1);
  assert.strictEqual(F.c.removeKeepWarmTrigger().removed, 0);
  assert.deepStrictEqual(F.g.state.triggers.map(t => t.id), ["other"]);
  // not reachable through the web app
  err(F.G("keepWarm"), /^Unknown action$/);
  err(F.P({ action: "keepWarm" }), /^Unknown action$/);
  err(F.P({ action: "installKeepWarmTrigger" }), /^Unknown action$/);
});
function frozenIn(F, fn) {
  const D = vm.runInContext("Date", F.c);
  const real = D.now; const t = real(); D.now = () => t;
  try { return fn(); } finally { D.now = real; }
}

/* ---------- bug fixes ---------- */
test("repair steps cannot be dated before the previous step of their chain (chain would become uneditable)", () => {
  const F = freshInstance();
  okData(F.P({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 10, date: "2026-03-01", time: "08:00" }));
  okData(F.P({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 2, cause: "x", date: "2026-03-01", time: "10:00" }));
  const t = okData(F.G("repairs")).items[0];
  const n = F.rows("movements").length;
  err(F.P({ action: "repair_start", id: t.id, date: "2026-03-01", time: "09:59" }), new RegExp("ก่อนขั้นตอนก่อนหน้าของงานซ่อม " + t.ticket_no + " \\(01/03/2026 10:00\\)"));
  err(F.P({ action: "scrap", id: t.id, date: "2026-02-28", time: "23:00" }), /ก่อนขั้นตอนก่อนหน้า/);
  assert.strictEqual(F.rows("movements").length, n);
  assert.strictEqual(okData(F.G("repairs")).items[0].stage, "damaged");
  okData(F.P({ action: "repair_start", id: t.id, date: "2026-03-01", time: "10:00" })); // same minute is fine
  err(F.P({ action: "repair_done", id: t.id, date: "2026-03-01", time: "09:00" }), /ก่อนขั้นตอนก่อนหน้า/);
  err(F.P({ action: "scrap", id: t.id, date: "2026-03-01", time: "09:30" }), /ก่อนขั้นตอนก่อนหน้า/);
  okData(F.P({ action: "repair_done", id: t.id, date: "2026-03-02", time: "08:00" }));
  // the chain stays editable (date edit of the done step within order works)
  const done = okData(F.G("history", { act: "repair_done" })).items[0];
  okData(F.P({ action: "movement_update", id: done.id, doc_no: done.doc_no, date: "2026-03-02", time: "09:00", resetPassword: RPW }));
});

test("repair actions check ticket_no when sent (a stale page never acts on a ticket that reused the id)", () => {
  const F = freshInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 10 }, AT)));
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, cause: "a" }, AT)));
  const a = okData(F.G("repairs")).items[0];
  okData(F.P({ action: "repair_delete", id: a.id, ticket_no: a.ticket_no, resetPassword: RPW }));
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 3, cause: "b" }, AT)));
  const b = okData(F.G("repairs")).items[0];
  assert.strictEqual(b.id, a.id); // the id was handed out again
  for (const act of ["repair_start", "scrap"]) {
    err(F.P(Object.assign({ action: act, id: a.id, ticket_no: a.ticket_no }, AT)), /ไม่พบใบแจ้งซ่อมนี้ \(อาจถูกลบไปแล้ว\)/);
  }
  assert.strictEqual(okData(F.G("repairs")).items[0].stage, "damaged");
  okData(F.P(Object.assign({ action: "repair_start", id: b.id, ticket_no: b.ticket_no }, AT)));
  err(F.P(Object.assign({ action: "repair_done", id: a.id, ticket_no: a.ticket_no }, AT)), /ไม่พบใบแจ้งซ่อมนี้/);
  okData(F.P(Object.assign({ action: "repair_done", id: b.id, ticket_no: b.ticket_no }, AT)));
  err(F.P(Object.assign({ action: "repair_start", id: 999 }, AT)), /^ไม่พบใบแจ้งซ่อม$/); // without ticket_no: as before
});

test("a new repair ticket's note is limited like ticket edits (1000 characters)", () => {
  const F = freshInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 10 }, AT)));
  const n = F.rows("repairs").length;
  err(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, cause: "a", note: "x".repeat(1001) }, AT)), /หมายเหตุยาวเกิน 1000/);
  okData(F.P(Object.assign({ action: "issue", type_id: 1, size: "1.2x1.2", qty: 2, department: "ฝ่ายผลิต" }, AT)));
  err(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายผลิต", condition: "damaged", cause: "a", note: "x".repeat(1001) }, AT)), /หมายเหตุยาวเกิน 1000/);
  assert.strictEqual(F.rows("repairs").length, n);
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, cause: "a", note: "  " + "x".repeat(1000) + "  " }, AT)));
});

/* ---------- partial repair steps (qty -> split ticket) ---------- */
function splitInstance() {
  const F = freshInstance();
  const objs = n => { const d = F.ss.getSheetByName(n).data; return d.slice(1).filter(r => r.some(v => v !== "" && v != null)).map(r => Object.fromEntries(d[0].map((k, i) => [k, r[i]]))); };
  F.mv = () => objs("movements");
  F.tk = () => objs("repairs");
  F.tkNo = no => F.tk().find(t => t.ticket_no === no);
  F.chainOf = id => F.mv().filter(m => m.repair_id === id).sort((a, b) => a.moved_at < b.moved_at ? -1 : a.moved_at > b.moved_at ? 1 : a.id - b.id);
  F.stock = st => stockOf(okData(F.G("bootstrap")).stock, 1, "1.2x1.2", st);
  F.deptQ = () => (okData(F.G("bootstrap")).dept.find(d => d.department === "ฝ่ายผลิต") || { qty: 0 }).qty;
  F.R = (action, body = {}) => F.P(Object.assign({}, body, { action, resetPassword: RPW }));
  F.logs = () => okData(F.P({ action: "logs", resetPassword: RPW }, null)).items;
  // every ticket's chain has the ticket's qty, and the sequence fits its stage
  F.chainsOk = () => F.tk().forEach(t => {
    const ch = F.chainOf(t.id);
    assert.ok(ch.length > 0, t.ticket_no);
    ch.forEach(m => assert.strictEqual(m.qty, t.qty, `${t.ticket_no} ${m.doc_no}`));
    const seq = ch.map(m => m.action).join(",");
    const ok = { damaged: ["damage"], repairing: ["damage,repair_start"], done: ["damage,repair_start,repair_done"], scrapped: ["damage,scrap", "damage,repair_start,scrap"] }[t.stage];
    assert.ok(ok.includes(seq), `${t.ticket_no} ${t.stage}: ${seq}`);
  });
  const d = (date, time) => ({ date, time });
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 30 }, d("2026-03-01", "08:00"))));
  okData(F.P(Object.assign({ action: "issue", type_id: 1, size: "1.2x1.2", qty: 25, department: "ฝ่ายผลิต" }, d("2026-03-01", "08:30"))));
  okData(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 21, department: "ฝ่ายผลิต", condition: "damaged", cause: "ไม้หัก", person: "คนคืน", note: "n0" }, d("2026-03-01", "09:00"))));
  F.A = F.tk()[0];
  return F;
}

test("partial repair: invalid qty rejected (Thai), nothing written; qty never checked past the stage", () => {
  const F = splitInstance();
  const A = F.A;
  const data = () => JSON.stringify(["movements", "repairs"].map(n => F.ss.getSheetByName(n).data));
  const before = data();
  const go = (act, qty) => F.P({ action: act, id: A.id, ticket_no: A.ticket_no, qty, date: "2026-03-01", time: "10:00" });
  for (const act of ["repair_start", "scrap"]) {
    err(go(act, "abc"), /^จำนวนต้องเป็นตัวเลขจำนวนเต็ม$/);
    err(go(act, "1.5"), /^จำนวนต้องเป็นตัวเลขจำนวนเต็ม$/);
    err(go(act, 2.5), /^จำนวนต้องเป็นตัวเลขจำนวนเต็ม$/);
    err(go(act, 0), /^จำนวนต้องมากกว่า 0$/);
    err(go(act, "-3"), /^จำนวนต้องมากกว่า 0$/);
    err(go(act, 22), new RegExp(`^ใบแจ้งซ่อม ${A.ticket_no} มีเพียง 21 ตัว$`));
  }
  err(go("repair_done", 3), /^ใบนี้ไม่ได้อยู่ระหว่างซ่อม$/);
  err(F.P({ action: "repair_start", id: A.id, qty: 5, date: "2026-03-01", time: "08:59" }), /ก่อนขั้นตอนก่อนหน้า/);
  assert.strictEqual(data(), before);
  assert.strictEqual(F.logs()[0].category, "warn");
  assert.strictEqual(F.tk().length, 1);
});

test("partial repair: send 6 of 21, finish 4 of 6, scrap 2 of the rest -> split tickets, totals as if moved directly", () => {
  const F = splitInstance();
  const A = F.A;
  const A2 = A.ticket_no + "-2", A3 = A.ticket_no + "-3", A4 = A.ticket_no + "-4";
  assert.strictEqual(A.qty, 21);
  assert.strictEqual(F.stock("damaged"), 21);
  assert.strictEqual(F.deptQ(), 4);

  // --- send 6 of 21 to repair
  const s = okData(F.P({ action: "repair_start", id: A.id, ticket_no: A.ticket_no, qty: "6", person: "ช่างหนึ่ง", date: "2026-03-01", time: "10:00" }));
  assert.strictEqual(s.message, `ส่งซ่อม 6 ตัว จาก ${A.ticket_no} แล้ว (แยกเป็นใบ ${A2})`);
  let a = F.tkNo(A.ticket_no), b = F.tkNo(A2);
  assert.deepStrictEqual([a.qty, a.stage, a.started_at, a.repairer], [15, "damaged", "", ""]);
  assert.deepStrictEqual([b.qty, b.stage, b.started_at, b.repairer], [6, "repairing", "2026-03-01 10:00:00", "ช่างหนึ่ง"]);
  // the new ticket carries the original report
  for (const k of ["type_id", "size", "source", "department", "cause", "reported_by", "reported_at"]) assert.strictEqual(b[k], a[k], k);
  assert.strictEqual(b.note, `n0\nแยกจาก ${A.ticket_no} (6 จาก 21 ตัว)`);
  assert.strictEqual(a.note, `n0\nแยก 6 ตัว ไปใบ ${A2} (เหลือ 15 ตัว)`);
  assert.strictEqual(b.updated_by, TESTER);
  // chain copies: same moved_at / statuses / department / person / note / actor, new doc no
  const [dmA] = F.chainOf(a.id);
  const [dmB, rpB] = F.chainOf(b.id);
  assert.strictEqual(dmA.qty, 15);
  assert.strictEqual(dmB.doc_no, "DM-" + YMD + "-0002");
  for (const k of ["action", "type_id", "size", "from_status", "to_status", "department", "person", "note", "moved_at", "actor"]) assert.strictEqual(dmB[k], dmA[k], k);
  assert.deepStrictEqual([dmB.qty, rpB.action, rpB.qty, rpB.note, rpB.person, rpB.doc_no], [6, "repair_start", 6, A2, "ช่างหนึ่ง", s.doc_no]);
  assert.deepStrictEqual([F.stock("damaged"), F.stock("repairing"), F.stock("available"), F.stock("issued"), F.deptQ()], [15, 6, 5, 4, 4]);
  // audit: the split, then the step
  const [l1, l2] = F.logs();
  assert.deepStrictEqual([l1.action, l1.ref], ["repair_start", s.doc_no]);
  assert.deepStrictEqual([l2.category, l2.action, l2.ref], ["repair", "แยกใบแจ้งซ่อม", A2]);
  assert.strictEqual(l2.detail, `แยกใบแจ้งซ่อม ${A.ticket_no} → ${A2} (RM ขนาด 1.2x1.2): 6 จาก 21 ตัว · ใบเดิมเหลือ 15 ตัว · คัดลอกรายการ DM-${YMD}-0002 (จาก DM-${YMD}-0001)`);
  F.chainsOk();
  // repairs page shows both tickets in their columns
  const items = okData(F.G("repairs")).items;
  assert.deepStrictEqual(items.filter(i => i.stage === "damaged").map(i => i.qty), [15]);
  assert.deepStrictEqual(items.filter(i => i.stage === "repairing").map(i => i.qty), [6]);

  // --- finish 4 of the 6 (root numbering: -3, not -2-2)
  const dn = okData(F.P({ action: "repair_done", id: b.id, ticket_no: A2, qty: 4, note: "เปลี่ยนไม้", date: "2026-03-01", time: "11:00" }));
  assert.strictEqual(dn.message, `ซ่อมเสร็จ 4 ตัว กลับเข้าคลังพร้อมใช้ (แยกจาก ${A2} เป็นใบ ${A3})`);
  b = F.tkNo(A2);
  const c = F.tkNo(A3);
  assert.deepStrictEqual([b.qty, b.stage], [2, "repairing"]);
  assert.deepStrictEqual([c.qty, c.stage, c.started_at, c.finished_at, c.repairer], [4, "done", "2026-03-01 10:00:00", "2026-03-01 11:00:00", "ช่างหนึ่ง"]);
  assert.strictEqual(c.note, `n0\nแยกจาก ${A.ticket_no} (6 จาก 21 ตัว)\nแยกจาก ${A2} (4 จาก 6 ตัว)\nซ่อมเสร็จ: เปลี่ยนไม้`);
  const chC = F.chainOf(c.id);
  assert.deepStrictEqual(chC.map(m => [m.action, m.qty]), [["damage", 4], ["repair_start", 4], ["repair_done", 4]]);
  assert.strictEqual(chC[1].note, A3); // RP note re-pointed at the new ticket
  assert.strictEqual(chC[1].moved_at, "2026-03-01 10:00:00");
  assert.deepStrictEqual(F.chainOf(b.id).map(m => m.qty), [2, 2]);
  assert.deepStrictEqual([F.stock("damaged"), F.stock("repairing"), F.stock("available"), F.deptQ()], [15, 2, 9, 4]);
  F.chainsOk();

  // --- scrap 2 of the 15 still damaged
  const sc = okData(F.P({ action: "scrap", id: A.id, ticket_no: A.ticket_no, qty: 2, note: "แตก", date: "2026-03-01", time: "12:00" }));
  assert.strictEqual(sc.message, `ตัดจำหน่าย 2 ตัว แล้ว (แยกจาก ${A.ticket_no} เป็นใบ ${A4})`);
  const d4 = F.tkNo(A4);
  assert.deepStrictEqual([F.tkNo(A.ticket_no).qty, F.tkNo(A.ticket_no).stage, d4.qty, d4.stage], [13, "damaged", 2, "scrapped"]);
  assert.deepStrictEqual(F.chainOf(d4.id).map(m => [m.action, m.qty, m.from_status, m.note]), [["damage", 2, "issued", "คืนสภาพชำรุด: ไม้หัก"], ["scrap", 2, "damaged", A4 + " แตก"]]);
  assert.deepStrictEqual([F.stock("damaged"), F.stock("repairing"), F.stock("available"), F.stock("scrapped"), F.deptQ()], [13, 2, 9, 2, 4]);
  F.chainsOk();
  // history exposes each copy's own ticket; ticket count by stage
  assert.strictEqual(okData(F.G("history", { q: "DM-" + YMD + "-0002" })).items[0].ticket_no, A2);
  assert.strictEqual(F.tk().length, 4);

  // --- full qty (explicit or missing) behaves exactly as before: no split
  const full = okData(F.P({ action: "scrap", id: b.id, ticket_no: A2, qty: "2", date: "2026-03-01", time: "13:00" }));
  assert.strictEqual(full.message, "ตัดจำหน่าย 2 ตัว แล้ว");
  const st0 = okData(F.P({ action: "repair_start", id: A.id, ticket_no: A.ticket_no, qty: "", date: "2026-03-01", time: "13:00" }));
  assert.strictEqual(st0.message, `ส่งซ่อม ${A.ticket_no} แล้ว`);
  assert.strictEqual(F.tk().length, 4);
  assert.notStrictEqual(F.logs()[1].action, "แยกใบแจ้งซ่อม");
  assert.deepStrictEqual([F.stock("damaged"), F.stock("repairing"), F.stock("available"), F.stock("scrapped")], [0, 13, 9, 4]);
  F.chainsOk();
});

test("partial repair: split chains stay editable / deletable as one unit; next split numbers never reuse deleted ones", () => {
  const F = splitInstance();
  const A = F.A;
  okData(F.P({ action: "repair_start", id: A.id, ticket_no: A.ticket_no, qty: 6, date: "2026-03-01", time: "10:00" }));
  const b = F.tkNo(A.ticket_no + "-2");
  okData(F.P({ action: "repair_done", id: b.id, qty: 4, date: "2026-03-01", time: "11:00" }));
  const c = F.tkNo(A.ticket_no + "-3");
  // qty edit on the copied damage row of -3 propagates to its whole chain only
  const [dmC] = F.chainOf(c.id);
  const e = okData(F.R("movement_update", { id: dmC.id, doc_no: dmC.doc_no, qty: 5 }));
  assert.strictEqual(e.changed, true);
  assert.deepStrictEqual(F.chainOf(c.id).map(m => m.qty), [5, 5, 5]);
  assert.strictEqual(F.tkNo(c.ticket_no).qty, 5);
  assert.deepStrictEqual([F.tkNo(A.ticket_no).qty, F.tkNo(b.ticket_no).qty], [15, 2]);
  assert.deepStrictEqual([F.stock("damaged"), F.stock("repairing"), F.stock("available"), F.deptQ()], [15, 2, 10, 3]);
  F.chainsOk();
  // the original's qty edit still works (15 -> 14)
  const [dmA] = F.chainOf(A.id);
  okData(F.R("movement_update", { id: dmA.id, doc_no: dmA.doc_no, qty: 14 }));
  assert.strictEqual(F.tkNo(A.ticket_no).qty, 14);
  F.chainsOk();
  // delete the -2 chain (via its repair_start row): its 2 movements + ticket only
  const rpB = F.chainOf(b.id)[1];
  const del = okData(F.R("movement_delete", { id: rpB.id, doc_no: rpB.doc_no }));
  assert.deepStrictEqual(del.removed, { movements: 2, repairs: 1 });
  assert.ok(!F.tkNo(b.ticket_no));
  assert.deepStrictEqual([F.stock("damaged"), F.stock("repairing"), F.stock("available"), F.stock("issued"), F.deptQ()], [14, 0, 10, 6, 6]);
  F.chainsOk();
  // repair_delete of a split ticket too
  okData(F.R("repair_delete", { id: c.id, ticket_no: c.ticket_no }));
  assert.deepStrictEqual(F.tk().map(t => t.ticket_no), [A.ticket_no]);
  // a new split never reuses -2 / -3 (still in the audit log)
  okData(F.P({ action: "scrap", id: A.id, qty: 1, date: "2026-03-01", time: "12:00" }));
  assert.ok(F.tkNo(A.ticket_no + "-4"));
  assert.deepStrictEqual([F.stock("damaged"), F.stock("scrapped")], [13, 1]);
  F.chainsOk();
  // a later ordinary ticket number is not affected by the suffixes
  okData(F.P({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, cause: "x", date: "2026-03-01", time: "13:00" }));
  assert.strictEqual(F.tk().pop().ticket_no, "RPR-" + YMD + "-0002");
});

/* ---------- Microsoft Teams notifications ---------- */
// A Workflows webhook URL carries its signature in the query string: it must never leak.
const HOOK = "https://prod-00.southeastasia.logic.azure.com:443/workflows/0a1b2c/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers&sv=1.0&sig=SECRET-SIG-123";
const SITE = "https://watanathep8-dotcom.github.io/pallet-management/";
const card = f => JSON.parse(f.opts.payload);
const factsOf = p => Object.fromEntries(p.attachments[0].content.body.find(b => b.type === "FactSet").facts.map(x => [x.title, x.value]));
const titleOf = p => p.attachments[0].content.body[1].text;
const linkOf = p => p.attachments[0].content.actions[0].url;
const DTH = `${TODAY.slice(8, 10)}/${TODAY.slice(5, 7)}/${TODAY.slice(0, 4)} ${HOUR}:00 น.`;
// Teams-enabled instance; every response text is kept so leaks can be checked.
function teamsInstance(hook = HOOK) {
  const F = freshInstance();
  if (hook !== null) F.g.state.props.TEAMS_WEBHOOK_URL = hook;
  F.texts = [];
  const P0 = F.P;
  F.P = (payload, who) => { const r = P0(payload, who); F.texts.push(JSON.stringify(r)); return r; };
  F.R = (action, body = {}) => F.P(Object.assign({}, body, { action, resetPassword: RPW }));
  // fetches made by fn()
  F.sent = fn => { const n = F.g.state.fetches.length; fn(); return F.g.state.fetches.slice(n); };
  F.one = fn => { const f = F.sent(fn); assert.strictEqual(f.length, 1, "expected one Teams fetch, got " + f.length); return card(f[0]); };
  F.noLeak = () => {
    const all = F.texts.join("\n") + "\n" + F.g.state.logs.map(l => l[1]).join("\n") + "\n" + F.dump();
    assert.ok(!all.includes(HOOK) && !all.includes("SECRET-SIG") && !all.includes("logic.azure.com"), "webhook URL leaked");
  };
  return F;
}

test("Teams: TEAMS_WEBHOOK_URL unset or blank -> no fetch for any event, writes work as before", () => {
  for (const hook of [null, "", "   "]) {
    const F = teamsInstance(hook);
    const f = F.sent(() => {
      okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 10 }, AT)));
      okData(F.P(Object.assign({ action: "issue", type_id: 1, size: "1.2x1.2", qty: 4, department: "ฝ่ายผลิต" }, AT)));
      okData(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายผลิต", condition: "good" }, AT)));
      okData(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายผลิต", condition: "damaged", cause: "y" }, AT)));
      okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 2, cause: "x" }, AT)));
    });
    assert.strictEqual(f.length, 0, "hook " + JSON.stringify(hook));
    assert.strictEqual(okData(F.G("repairs")).items.length, 2);
    assert.ok(!F.g.state.logs.some(l => /Teams/.test(l[1])), "silently skipped");
  }
});

test("Teams: return (good / damaged) -> exactly one card each: doc no, type/size, qty, department, resulting status (+ ticket), note, who, when", () => {
  const F = teamsInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 20 }, AT)));
  okData(F.P(Object.assign({ action: "issue", type_id: 1, size: "1.2x1.2", qty: 5, department: "ฝ่ายผลิต" }, AT)));

  let d;
  const f = F.sent(() => { d = okData(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 2, department: "ฝ่ายผลิต", condition: "good", person: "C", note: "คืนหลังกะ\nดึก" }, AT), "สมชาย")); });
  assert.strictEqual(f.length, 1);
  assert.strictEqual(f[0].url, HOOK);
  assert.deepStrictEqual(Object.assign({}, f[0].opts, { payload: undefined }), { method: "post", contentType: "application/json", payload: undefined, muteHttpExceptions: true });
  assert.strictEqual(f[0].lockHeld, false, "sent after the lock was released");
  const p = card(f[0]);
  assert.strictEqual(p.type, "message");
  assert.strictEqual(p.attachments.length, 1);
  const a = p.attachments[0];
  assert.strictEqual(a.contentType, "application/vnd.microsoft.card.adaptive");
  assert.strictEqual(a.contentUrl, null);
  assert.strictEqual(a.content.$schema, "http://adaptivecards.io/schemas/adaptive-card.json");
  assert.strictEqual(a.content.type, "AdaptiveCard");
  assert.strictEqual(a.content.version, "1.4");
  assert.strictEqual(a.content.body[1].color, "Attention");
  assert.strictEqual(a.content.body.length, 4);
  assert.ok(/^แจ้งเมื่อ \d\d\/\d\d\/\d{4} \d\d:\d\d น\. \(เวลาไทย\)$/.test(a.content.body[2].text), a.content.body[2].text);
  assert.deepStrictEqual(a.content.actions, [{ type: "Action.OpenUrl", title: "เปิด Pallet Hub", url: SITE + "#history" }]);
  assert.strictEqual(titleOf(p), "รับคืน " + d.doc_no);
  assert.deepStrictEqual(factsOf(p), {
    "เลขที่เอกสาร": d.doc_no, "ประเภทพาเลท": "RM — พาเลทสำหรับใส่ RM", "ขนาด": "1.2x1.2 ม.", "จำนวน": "2 ตัว",
    "รับคืนจากฝ่าย": "ฝ่ายผลิต", "สถานะหลังรับคืน": "พร้อมใช้", "หมายเหตุ": "คืนหลังกะ ดึก",
    "ชื่อที่ระบุ": "C", "ผู้ทำรายการ": "สมชาย", "วันที่/เวลา": DTH
  });

  // damaged return: ONE card (the return card, naming the opened ticket) - no separate damage card
  let d2;
  const p2 = F.one(() => { d2 = okData(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายผลิต", condition: "damaged", cause: "ตะปูหลุด", note: "ขาหัก 1 ขา" }, AT))); });
  const t = okData(F.G("repairs")).items[0];
  assert.strictEqual(titleOf(p2), "รับคืน (ชำรุด) " + d2.doc_no);
  assert.strictEqual(linkOf(p2), SITE + "#history");
  assert.deepStrictEqual(factsOf(p2), {
    "เลขที่เอกสาร": d2.doc_no, "ประเภทพาเลท": "RM — พาเลทสำหรับใส่ RM", "ขนาด": "1.2x1.2 ม.", "จำนวน": "1 ตัว",
    "รับคืนจากฝ่าย": "ฝ่ายผลิต", "สถานะหลังรับคืน": "ชำรุด — เปิดใบแจ้งซ่อม " + t.ticket_no, "สาเหตุ": "ตะปูหลุด",
    "หมายเหตุ": "ขาหัก 1 ขา", "ผู้ทำรายการ": T.tester, "วันที่/เวลา": DTH
  });
  // a long (ticket) note is flattened and cut
  const p3 = F.one(() => okData(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายผลิต", condition: "damaged", cause: "x", note: "a\tb " + "ก".repeat(400) }, AT))));
  const n3 = factsOf(p3)["หมายเหตุ"];
  assert.ok(n3.startsWith("a b กกก") && n3.endsWith("…") && Array.from(n3).length === 300, n3);
  assert.ok(factsOf(p3)["สถานะหลังรับคืน"].startsWith("ชำรุด — เปิดใบแจ้งซ่อม "));
  F.noLeak();
});

test("Teams: damage from stock -> one card: ticket, type/size, qty, cause, who, when; links #repair", () => {
  const F = teamsInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 2, size: "1.1x1.1", qty: 20 }, AT)));
  let d;
  const p = F.one(() => { d = okData(F.P(Object.assign({ action: "damage", type_id: 2, size: "1.1x1.1", qty: 3, cause: "ไม้หัก", person: "C" }, AT), "สมชาย")); });
  const t = okData(F.G("repairs")).items[0];
  assert.strictEqual(titleOf(p), "แจ้งชำรุด " + t.ticket_no);
  assert.strictEqual(linkOf(p), SITE + "#repair");
  assert.strictEqual(p.attachments[0].content.body[1].color, "Attention");
  assert.deepStrictEqual(factsOf(p), {
    "เลขที่ใบแจ้งซ่อม": t.ticket_no, "ประเภทพาเลท": "PK — พาเลทสำหรับใส่ PK", "ขนาด": "1.1x1.1 ม.", "จำนวน": "3 ตัว",
    "สาเหตุ": "ไม้หัก", "เลขที่เอกสาร": d.doc_no, "ชื่อที่ระบุ": "C", "ผู้ทำรายการ": "สมชาย", "วันที่/เวลา": DTH
  });
  // overridable site URL; a non-web value falls back to the default
  F.g.state.props.PALLET_SITE_URL = "https://example.test/hub/#old";
  assert.strictEqual(linkOf(F.one(() => okData(F.P(Object.assign({ action: "damage", type_id: 2, size: "1.1x1.1", qty: 1 }, AT))))), "https://example.test/hub/#repair");
  F.g.state.props.PALLET_SITE_URL = "javascript:alert(1)";
  const p2 = F.one(() => okData(F.P(Object.assign({ action: "damage", type_id: 2, size: "1.1x1.1", qty: 1 }, AT))));
  assert.strictEqual(linkOf(p2), SITE + "#repair");
  assert.strictEqual(factsOf(p2)["สาเหตุ"], "-");
  F.noLeak();
});

test("Teams: receive / issue / repair_start / repair_done / scrap / record edits & deletes / reset / department changes send nothing", () => {
  const F = teamsInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 20 }, AT)));
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 3, cause: "a" }, AT)));
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 2, cause: "b" }, AT)));
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, cause: "c" }, AT)));
  const [t3, t2, t1] = okData(F.G("repairs")).items;
  const before = F.g.state.fetches.length;
  const f = F.sent(() => {
    const rc = okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 10 }, AT)));
    const rc2 = okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 3 }, AT)));
    okData(F.P(Object.assign({ action: "issue", type_id: 1, size: "1.2x1.2", qty: 5, department: "ฝ่ายผลิต" }, AT)));
    okData(F.P(Object.assign({ action: "repair_start", id: t1.id, person: "ช่าง" }, AT)));
    okData(F.P(Object.assign({ action: "repair_done", id: t1.id, note: "ok" }, AT)));
    okData(F.P(Object.assign({ action: "scrap", id: t2.id, note: "แตก" }, AT)));
    okData(F.R("movement_update", { id: rc.id, qty: 12, note: "แก้" }));
    okData(F.R("movement_delete", { id: rc2.id }));
    okData(F.R("repair_update", { id: t3.id, cause: "ccc" }));
    okData(F.R("repair_delete", { id: t3.id }));
    okData(F.P({ action: "dept_save", name: "ฝ่ายใหม่", resetPassword: RPW }));
    okData(F.R("reset_data"));
  });
  assert.strictEqual(before, 3);
  assert.strictEqual(f.length, 0);
  F.noLeak();
});

test("Teams: rejected writes (validation, wrong password, no name, lock timeout) send nothing", () => {
  const F = teamsInstance();
  okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 5 }, AT)));
  okData(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, cause: "a" }, AT)));
  const t = okData(F.G("repairs")).items[0];
  const m = okData(F.G("history", { act: "receive" })).items[0];
  const f = F.sent(() => {
    err(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "ฝ่ายผลิต" }, AT)), /ถือพาเลทนี้อยู่ 0 ตัว/);
    err(F.P(Object.assign({ action: "return", type_id: 1, size: "1.2x1.2", qty: 1, department: "" }, AT)), /กรุณาเลือกฝ่ายที่คืน/);
    err(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 99 }, AT)), /มีเพียง/);
    err(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1, note: "x".repeat(1001) }, AT)), /ยาวเกิน/);
    err(F.P(Object.assign({ action: "repair_done", id: t.id }, AT)), /ไม่ได้อยู่ระหว่างซ่อม/);
    err(F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1 }, AT), null), /กรุณาระบุชื่อ/);
    err(F.P({ action: "movement_delete", id: m.id, resetPassword: "wrong" }), /รหัสไม่ถูกต้อง/);
    err(F.R("movement_delete", { id: m.id }), /^ลบไม่ได้/);
    err(F.R("movement_update", { id: m.id, qty: 0 }), /มากกว่า 0/);
    err(F.R("repair_update", { id: t.id, qty: 5 }), /แก้ไขจำนวน/);
    F.g.state.lockFail = true;
    try { err(F.P(Object.assign({ action: "scrap", id: t.id }, AT)), /ลองใหม่/); } finally { F.g.state.lockFail = false; }
    // a write that fails while saving (sheet error during flush) is not announced either
    const mv = F.ss.getSheetByName("movements");
    const orig = mv.getRange;
    mv.getRange = function () { throw new Error("Service Spreadsheets failed"); };
    try { err(F.P(Object.assign({ action: "scrap", id: t.id }, AT))); } finally { mv.getRange = orig; }
  });
  assert.strictEqual(f.length, 0);
  assert.strictEqual(F.g.state.locks, F.g.state.unlocks);
  F.noLeak();
});

test("Teams: HTTP errors / fetch exceptions never change the response or the saved data; logged with status, never the URL", () => {
  const run = (setup) => {
    const F = teamsInstance();
    setup(F.g.state);
    frozenIn(F, () => {
      okData(F.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 9 }, AT)));
      F.res = F.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 2, cause: "x", reads: ["repairs"] }, AT));
    });
    return F;
  };
  const ok = run(() => {});
  const base = ok.res;
  okData(base);
  assert.deepStrictEqual(Object.keys(base.data), ["doc_no", "id", "message", "reads"]);
  for (const [label, setup, logRe] of [
    ["HTTP 500", s => { s.fetchStatus = 500; }, /^Teams notification failed: HTTP 500$/],
    ["HTTP 404", s => { s.fetchStatus = 404; }, /^Teams notification failed: HTTP 404$/],
    ["exception with the URL in its message", s => { s.fetchError = "Address unavailable: " + HOOK; }, /^Teams notification failed: Address unavailable: \[webhook\]$/],
    ["exception with another URL", s => { s.fetchError = "DNS error: https://other.example/x?sig=1"; }, /^Teams notification failed: DNS error: \[url\]$/]
  ]) {
    const F = run(setup);
    assert.strictEqual(F.g.state.fetches.length, 1, label);
    assert.deepStrictEqual(F.res, base, label + ": response unchanged");
    const noClock = x => x.replace(/\d{4}-\d\d-\d\d \d\d:\d\d:\d\d/g, "<dt>"); // created_at differs between instances
    assert.strictEqual(noClock(F.dump()), noClock(ok.dump()), label + ": data unchanged");
    const warns = F.g.state.logs.filter(l => l[0] === "warn").map(l => l[1]);
    assert.strictEqual(warns.length, 1, label);
    assert.ok(logRe.test(warns[0]), label + ": " + warns[0]);
    assert.strictEqual(F.g.state.locks, F.g.state.unlocks);
    F.noLeak();
  }
  // a non-https webhook value is refused (warned, without the value) and not called
  const H = teamsInstance("http://insecure.example/hook?sig=SECRET-SIG-9");
  okData(H.P(Object.assign({ action: "receive", type_id: 1, size: "1.2x1.2", qty: 9 }, AT)));
  okData(H.P(Object.assign({ action: "damage", type_id: 1, size: "1.2x1.2", qty: 1 }, AT)));
  assert.strictEqual(H.g.state.fetches.length, 0);
  assert.ok(H.g.state.logs.some(l => /TEAMS_WEBHOOK_URL must be an https/.test(l[1])));
  assert.ok(!H.g.state.logs.some(l => /SECRET-SIG|insecure/.test(l[1])));
});

test("testTeamsNotification: sends one sample card, returns / logs the HTTP status (never the URL); not a web action", () => {
  const off = teamsInstance(null);
  const r0 = off.c.testTeamsNotification();
  assert.deepStrictEqual([r0.sent, r0.status], [false, null]);
  assert.ok(r0.message.includes("TEAMS_WEBHOOK_URL"));
  assert.strictEqual(off.g.state.fetches.length, 0);

  const F = teamsInstance();
  const r = F.c.testTeamsNotification();
  assert.deepStrictEqual([r.sent, r.status], [true, 202]);
  assert.strictEqual(F.g.state.fetches.length, 1);
  const p = card(F.g.state.fetches[0]);
  assert.strictEqual(titleOf(p), "ทดสอบการแจ้งเตือน");
  assert.strictEqual(p.attachments[0].content.type, "AdaptiveCard");
  assert.strictEqual(linkOf(p), SITE);
  assert.ok(F.g.state.logs.some(l => l[0] === "log" && l[1].includes('"status":202')));
  F.g.state.fetchStatus = 400;
  const bad = F.c.testTeamsNotification();
  assert.deepStrictEqual([bad.sent, bad.status], [false, 400]);
  assert.ok(bad.message.includes("HTTP 400"));
  F.g.state.fetchStatus = 202;
  F.texts.push(JSON.stringify([r0, r, bad]));
  for (const a of ["testTeamsNotification", "notifyTeams_", "sendTeams_"]) err(F.P({ action: a }), /^Unknown action$/);
  err(F.G("testTeamsNotification"), /^Unknown action$/);
  assert.strictEqual(F.g.state.fetches.length, 2);
  F.noLeak();
});

/* ---------- frontend ---------- */
// Regression checks for frontend fixes (behaviour verified in a browser against test/pallet-dev.js).
test("frontend fixes: no double submit, qty field, deleted departments on return, history errors, repaint rules", () => {
  const app = fs.readFileSync(path.join(ROOT, "docs/assets/app.js"), "utf8");
  const fnSrc = name => { const i = app.indexOf("function " + name + "("); assert.ok(i !== -1, name); return app.slice(i, app.indexOf("\nfunction ", i + 10)); };
  const tx = fnSrc("txForm");
  // the bubbling view.onclick -> refresh() used to re-enable the submit button right after it was disabled
  assert.ok(/\$\('#submit'\)\.disabled = saving \|\|/.test(tx), "refresh keeps the button disabled while saving");
  assert.ok(/if \(saving\) return;\s*saving = true;/.test(tx), "submit ignores clicks while saving");
  assert.ok(/catch \(e\) \{\s*toast\(e\.message, 'err'\);\s*saving = false;/.test(tx), "an error re-enables saving");
  // emptied qty field is not forced back to "1" while typing
  assert.ok(/\$\('#qty'\)\.oninput = e => \{ if \(e\.target\.value === ''\)/.test(tx));
  // return form: departments still holding pallets are listed even after they were deleted
  assert.ok(/deptList = kind === 'return'[\s\S]*S\.dept\.map\(r => r\.department\)/.test(tx) && tx.includes("${deptList.map(d =>"));
  // repair step buttons: one request per click, ticket_no sent with the id
  const ra = fnSrc("repairAction");
  assert.ok(/if \(go\.disabled\) return;\s*go\.disabled = true;/.test(ra) && /ticket_no: r\.ticket_no/.test(ra) && /go\.disabled = false/.test(ra));
  // history search: errors are shown, stale answers dropped
  const hi = fnSrc("history");
  assert.ok(/try \{\s*const r = await api\('history'/.test(hi) && /catch \(e\) \{[\s\S]*box\.innerHTML = `<div class="empty">/.test(hi) && /seq !== loadSeq/.test(hi));
  // background refresh only keeps a stale screen on pages with user input
  assert.ok(/S\.dirty && INPUT_PAGES\.includes\(p\)/.test(app));
  assert.ok(/INPUT_PAGES = \['receive', 'issue', 'return', 'damage', 'history', 'settings'\]/.test(app));
});

test("frontend: repair step dialog has a qty input (1..ticket qty, default all, only when qty > 1) and sends qty", () => {
  const app = fs.readFileSync(path.join(ROOT, "docs/assets/app.js"), "utf8");
  const i = app.indexOf("function repairAction(");
  const src = app.slice(i, app.indexOf("\nfunction ", i + 10));
  assert.ok(src.includes("const max = +r.qty;"));
  assert.ok(src.includes("${max > 1 ? `"));
  assert.ok(/id="mQty" type="number"[^>]*min="1" max="\$\{max\}"[^>]*value="\$\{max\}"/.test(src));
  assert.ok(src.includes("จำนวน (สูงสุด ${fmt(max)} ตัว)"));
  assert.ok(/writeApi\(act, \{ id: r\.id, ticket_no: r\.ticket_no, qty: qVal\(\),/.test(src));
  assert.ok(/go\.onclick = async \(\) => \{\s*if \(!qOk\(\)\) \{.*return; \}/.test(src), "invalid qty never sent");
  assert.ok(!/on(keydown|keyup|keypress|input) = [^;]*&&/.test(src), "no handler that can return false");
});

test("docs JS files parse (new Function); no login UI; reset password never stored; cache-buster", () => {
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
  for (const a of ["movement_update", "movement_delete", "repair_update", "repair_delete", "verifyResetPassword", "logs_export", "batch"]) assert.ok(app.includes(`'${a}'`), a);
  assert.ok(app.includes("resetPassword: pw }"));
  assert.ok(/data-medit=.*✏️ แก้ไข/.test(app) && /data-mdel=.*🗑 ลบ/.test(app) && /data-redit=.*data-rdel=/.test(app));
  // no login / accounts
  for (const a of ["'login'", "'logout'", "'change_password'", "'user_save'", "'me'", "authToken"]) assert.ok(!app.includes(a), a);
  assert.ok(app.includes("store('get', 'palletUser')") && /actor: S\.user/.test(app)); // typed name sent as actor
  assert.ok(!/palletRC[^\n]*logs/.test(app));
  assert.ok(app.includes("ผู้ทำรายการ / By"));
  const html = fs.readFileSync(path.join(ROOT, "docs/index.html"), "utf8");
  assert.ok(html.includes('assets/app.js?v=23"'));
  assert.ok(!html.includes('id="loginScreen"') && !html.includes('data-page="account"') && !html.includes("umLogout"));
  assert.ok(html.includes('id="userChip"') && html.includes('data-page="logs"'));
  assert.ok(/<link rel="preconnect" href="https:\/\/script\.google\.com"/.test(html) && /script\.googleusercontent\.com/.test(html));
  assert.ok(!html.includes("ราชบุรี"));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
