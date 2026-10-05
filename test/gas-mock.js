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

const ROOT = path.join(__dirname, "..");

/* ===================== mocks ===================== */
function createGas() {
  const state = { props: {}, cache: {}, cachePuts: [], sleeps: 0, locks: 0, unlocks: 0, lockFail: false, spreadsheets: {}, seq: 0, formatCalls: 0, created: 0 };

  class Range {
    constructor(sheet, row, col, nr, nc) {
      if (row < 1 || col < 1 || nr < 1 || nc < 1) throw new Error("Range out of bounds");
      if (row + nr - 1 > sheet.maxRows) throw new Error("Range exceeds sheet rows: " + (row + nr - 1) + " > " + sheet.maxRows);
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
  }

  class Sheet {
    constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; }
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
    getRange(r, c, nr = 1, nc = 1) { return new Range(this, r, c, nr, nc); }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    setFrozenRows() { return this; }
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
    Utilities: { sleep(ms) { state.sleeps += ms; } },
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
const get = (action, params = {}) => {
  const out = ctx.doGet({ parameter: Object.assign({ action }, params) });
  assert.strictEqual(out.getMimeType(), "application/json");
  return JSON.parse(out.getContent());
};
const postRaw = (payload) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(payload) } }).getContent());
const PW = "s3cret-Pallet";
const post = (action, body = {}) => postRaw(Object.assign({}, body, { action, password: PW, actor: "Tester" }));
const okData = (res) => { assert.strictEqual(res.ok, true, "expected ok, got: " + JSON.stringify(res)); return res.data; };
const err = (res, re) => {
  assert.strictEqual(res.ok, false, "expected error, got: " + JSON.stringify(res));
  if (re) assert.ok(re.test(res.error), "error '" + res.error + "' !~ " + re);
  return res;
};

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
test("API before setup reports setupSystem error", () => {
  err(get("bootstrap"), /setupSystem/);
});

test("setupSystem creates spreadsheet, sheets, headers, seed data", () => {
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, true);
  assert.ok(state.props.PALLET_SPREADSHEET_ID);
  const ss = state.spreadsheets[state.props.PALLET_SPREADSHEET_ID];
  assert.strictEqual(ss.tz, "Asia/Bangkok");
  assert.deepStrictEqual(ss.getSheets().map(s => s.getName()), ["pallet_types", "departments", "repairs", "movements", "audit_logs"]);
  assert.deepStrictEqual(sheet("movements").data[0], ["id", "doc_no", "action", "type_id", "size", "qty", "from_status", "to_status", "department", "person", "note", "repair_id", "moved_at", "created_at"]);
  assert.strictEqual(dataRows("pallet_types"), 4);
  assert.strictEqual(dataRows("departments"), 8);
  assert.strictEqual(r.passwordConfigured, false);
});

test("setupSystem is idempotent (no second spreadsheet, no duplicate seed)", () => {
  const id = state.props.PALLET_SPREADSHEET_ID;
  const r = ctx.setupSystem();
  assert.strictEqual(r.created, false);
  assert.strictEqual(state.props.PALLET_SPREADSHEET_ID, id);
  assert.strictEqual(state.created, 1);
  assert.strictEqual(dataRows("pallet_types"), 4);
  assert.strictEqual(dataRows("departments"), 8);
});

/* ---------- reads ---------- */
let types;
test("bootstrap returns types (sorted), active departments, empty stock, now", () => {
  const d = okData(get("bootstrap"));
  types = d.types;
  assert.deepStrictEqual(d.types.map(t => t.tkey), ["RM", "PK", "FG", "PL"]);
  assert.deepStrictEqual(d.types[1], { id: 2, tkey: "PK", code: "PK", name: "พาเลทสำหรับใส่ PK", short: "ผลิตภัณฑ์ (PK)", description: "ขวด + ฝา + แกลลอน ฯ", color: "#F5B400", sizes: "1.1x1.1,1.2x1.2", sort: 2 });
  assert.strictEqual(d.departments.length, 8);
  assert.deepStrictEqual(d.departments[0], { id: 1, name: "ฝ่ายผลิต", icon: "fa-industry", color: "#E2231A", active: 1 });
  assert.deepStrictEqual(d.stock, {});
  assert.deepStrictEqual(d.dept, []);
  assert.strictEqual(d.now.date, TODAY);
  assert.ok(/^\d{2}:\d{2}$/.test(d.now.time));
});

test("unknown action / write action via GET are rejected", () => {
  err(get("nope"), /^Unknown action$/);
  err(postRaw({ action: "nope", password: PW }), /^Unknown action$/);
  err(get("receive"), /POST/);
  err(get("verifyPassword"), /POST/);
  assert.strictEqual(okData(get(undefined)).service, "Pallet Hub API");
});

test("invalid JSON body is rejected", () => {
  const r = JSON.parse(ctx.doPost({ postData: { contents: "{bad" } }).getContent());
  err(r, /JSON/);
});

/* ---------- password ---------- */
test("writes refused while PALLET_ACTION_PASSWORD is unset (Thai admin message)", () => {
  const r = err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 5 }, AT)), /PALLET_ACTION_PASSWORD/);
  assert.strictEqual(r.passwordError, true);
  err(postRaw({ action: "verifyPassword", password: "x" }), /ยังไม่ได้ตั้งรหัสผ่าน/);
  assert.strictEqual(dataRows("movements"), 0);
});

test("verifyPassword: wrong -> error + 1s sleep, right -> valid", () => {
  state.props.PALLET_ACTION_PASSWORD = PW;
  const before = state.sleeps;
  const r = err(postRaw({ action: "verifyPassword", password: "wrong" }), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.sleeps - before, 1000);
  assert.deepStrictEqual(okData(postRaw({ action: "verifyPassword", password: PW })), { valid: true });
  err(postRaw({ action: "verifyPassword" }), /รหัสไม่ถูกต้อง/);
});

test("write with wrong/missing password is refused and NOT logged", () => {
  const logsBefore = dataRows("audit_logs");
  err(postRaw(Object.assign({ action: "receive", password: "bad", type_id: 1, size: "1.2x1.2", qty: 5 }, AT)), /รหัสไม่ถูกต้อง/);
  err(postRaw(Object.assign({ action: "dept_save", name: "X" })), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(dataRows("movements"), 0);
  assert.strictEqual(dataRows("audit_logs"), logsBefore);
  state.cache = {}; // reset failure counter for the rest of the tests
});

/* ---------- receive ---------- */
test("receive: creates RC doc, stock available, audit log with actor/ip", () => {
  const d = okData(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 50, person: "  สมชาย ใจดี  ", note: "PO-123" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0001");
  assert.strictEqual(d.id, 1);
  assert.strictEqual(d.message, "รับเข้า 50 ตัว เรียบร้อย");
  const b = okData(get("bootstrap"));
  assert.strictEqual(stockOf(b.stock, 1, "1.2x1.2", "available"), 50);
  const log = lastLog();
  assert.strictEqual(log.category, "pallet");
  assert.strictEqual(log.action, "receive");
  assert.strictEqual(log.ref, d.doc_no);
  assert.strictEqual(log.actor, "สมชาย ใจดี");
  assert.strictEqual(log.ip, "web");
  assert.strictEqual(log.detail, `รับเข้า RM (พาเลทสำหรับใส่ RM) ขนาด 1.2x1.2 ม. จำนวน 50 ตัว [ภายนอก → พร้อมใช้] · เวลาทำรายการ ${TODAY.slice(8, 10)}/${TODAY.slice(5, 7)}/${TODAY.slice(0, 4)} ${HOUR}:00 · PO-123`);
});

test("receive: second doc number increments; actor falls back to header actor", () => {
  const d = okData(post("receive", Object.assign({ type_id: 2, size: "1.1x1.1", qty: "30", person: "" }, AT)));
  assert.strictEqual(d.doc_no, "RC-" + YMD + "-0002");
  assert.strictEqual(lastLog().actor, "Tester");
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
  assert.strictEqual(ticket1.reported_by, "B");
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
  assert.strictEqual(lines[0], "เลขที่เอกสาร,วันที่,เวลา,รายการ,รหัส,ประเภท,ขนาด,จำนวน,ฝ่าย,ผู้ทำรายการ,หมายเหตุ");
  assert.strictEqual(lines.length, 1 + 4 + 1);
  const row = lines.find(l => l.startsWith("RC-" + YMD + "-0001"));
  assert.strictEqual(row, `RC-${YMD}-0001,${TODAY},${HOUR}:00,รับเข้า,RM,"พาเลทสำหรับใส่ RM",1.2x1.2,50,,"สมชาย ใจดี",PO-123`);
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
  assert.strictEqual(okData(get("logs", { from: "2099-01-01" })).items.length, 0);
  const x = okData(get("logs_export", { cat: "warn" }));
  assert.ok(/^pallet_log_\d{8}_\d{6}\.csv$/.test(x.filename));
  const lines = x.csv.trim().split("\n");
  assert.strictEqual(lines[0], "ลำดับ,วันที่,เวลา,หมวด,เลขที่อ้างอิง,รายละเอียด,ผู้ทำรายการ,IP");
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
  assert.strictEqual(log.actor, "Tester");
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

test("lockout: >10 wrong attempts within 15 min refuses even the right password", () => {
  state.cache = {};
  state.cachePuts = [];
  for (let i = 0; i < 10; i++) err(postRaw({ action: "verifyPassword", password: "guess" + i }), /รหัสไม่ถูกต้อง/);
  assert.ok(state.cachePuts.every(t => t === 900));
  const r = err(postRaw({ action: "verifyPassword", password: PW }), /หลายครั้งเกินไป/);
  assert.strictEqual(r.passwordError, true);
  const n = dataRows("movements");
  err(post("receive", Object.assign({ type_id: 1, size: "1.2x1.2", qty: 1 }, AT)), /หลายครั้งเกินไป/);
  assert.strictEqual(dataRows("movements"), n);
  state.cache = {}; // 15 minutes later (cache entry expired)
  okData(postRaw({ action: "verifyPassword", password: PW }));
});

test("reads work without any password", () => {
  for (const a of ["bootstrap", "dashboard", "repairs", "history", "export", "logs", "logs_export"]) okData(get(a));
});

/* ---------- frontend syntax ---------- */
test("docs JS files parse (new Function)", () => {
  for (const f of ["docs/config.js", "docs/assets/app.js"]) {
    const src = fs.readFileSync(path.join(ROOT, f), "utf8");
    new Function(src); // throws SyntaxError on bad code
  }
  const cfg = {};
  new Function("window", fs.readFileSync(path.join(ROOT, "docs/config.js"), "utf8"))(cfg);
  assert.deepStrictEqual(cfg.PALLET_CONFIG, { apiUrl: "" });
  const app = fs.readFileSync(path.join(ROOT, "docs/assets/app.js"), "utf8");
  assert.ok(!/api\.php|X-User|localStorage\.setItem\([^)]*[Pp]ass/.test(app));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
