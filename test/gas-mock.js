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

/* ---------- reset data ---------- */
const RPW = "reset-Only-9";
const reset = (pw, extra = {}) => postRaw(Object.assign({ action: "reset_data", resetPassword: pw, actor: "Admin" }, extra));
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
  state.cache = {};
  const before = snapshot();
  const sleeps = state.sleeps;
  const r = err(reset("wrong"), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(reset(PW, { password: PW }), /รหัสไม่ถูกต้อง/); // action password is not the reset password
  err(postRaw({ action: "reset_data", password: RPW }), /รหัสไม่ถูกต้อง/); // must come in resetPassword
  assert.strictEqual(state.cache.PALLET_RESET_FAILURES, "3");
  assert.ok(!("PALLET_PASSWORD_FAILURES" in state.cache));
  assert.deepStrictEqual(snapshot(), before);
  state.cache = {};
});

test("reset_data lockout: >10 wrong in 15 min refuses even the right reset password; normal writes unaffected", () => {
  state.cache = {};
  state.cachePuts = [];
  for (let i = 0; i < 10; i++) err(reset("guess" + i), /รหัสไม่ถูกต้อง/);
  assert.ok(state.cachePuts.every(t => t === 900));
  const before = snapshot();
  const r = err(reset(RPW), /หลายครั้งเกินไป/);
  assert.strictEqual(r.passwordError, true);
  assert.deepStrictEqual(snapshot(), before);
  okData(postRaw({ action: "verifyPassword", password: PW })); // separate counter
  state.cache = {}; // 15 minutes later
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
  assert.strictEqual(l.actor, "Admin");
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
const rec = (action, body = {}, pw = RPW) => postRaw(Object.assign({}, body, { action, resetPassword: pw, actor: "Editor" }));
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
  state.cache = {};
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
  state.cache = {};
  const before = sheetDump();
  const sleeps = state.sleeps;
  const r = err(rec("movement_delete", { id: D.r2.id }, "wrong"), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(r.passwordError, true);
  assert.strictEqual(state.sleeps - sleeps, 1000);
  err(rec("movement_update", { id: D.r2.id, qty: 1 }, PW), /รหัสไม่ถูกต้อง/); // action password is not enough
  err(postRaw({ action: "repair_update", password: RPW, id: D.x.id, cause: "x" }), /รหัสไม่ถูกต้อง/); // must be resetPassword
  err(postRaw({ action: "verifyResetPassword", resetPassword: "nope" }), /รหัสไม่ถูกต้อง/);
  assert.strictEqual(state.cache.PALLET_RESET_FAILURES, "4");
  assert.ok(!("PALLET_PASSWORD_FAILURES" in state.cache));
  assert.strictEqual(sheetDump(), before);
  assert.deepStrictEqual(okData(postRaw({ action: "verifyResetPassword", resetPassword: RPW })), { valid: true });
  state.cache = {};
});

test("edit/delete lockout is shared with reset_data", () => {
  state.cache = {};
  for (let i = 0; i < 5; i++) err(reset("guess" + i), /รหัสไม่ถูกต้อง/);
  for (let i = 0; i < 5; i++) err(rec("movement_delete", { id: D.r2.id }, "guess" + i), /รหัสไม่ถูกต้อง/);
  const before = sheetDump();
  err(rec("movement_delete", { id: D.r2.id }), /หลายครั้งเกินไป/);
  err(rec("repair_update", { id: D.x.id, cause: "x" }), /หลายครั้งเกินไป/);
  err(postRaw({ action: "verifyResetPassword", resetPassword: RPW }), /หลายครั้งเกินไป/);
  err(reset(RPW), /หลายครั้งเกินไป/);
  assert.strictEqual(sheetDump(), before);
  okData(postRaw({ action: "verifyPassword", password: PW })); // normal password counter unaffected
  state.cache = {}; // 15 minutes later
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
  assert.strictEqual(log.actor, "Editor");
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
  assert.strictEqual(log.actor, "Editor");
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
  assert.strictEqual(lastLog().detail, `แก้ไขรายการ ${D.r1.doc_no} (รับเข้า RM ขนาด 1.2x1.2): ผู้ทำรายการ "" → "สมหญิง"; หมายเหตุ "" → "แก้ PO"`);
  assert.strictEqual(lastLog().actor, "Editor");
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
  const d = okData(rec("repair_update", { id: D.x.id, ticket_no: D.x.ticket_no, cause: "ไม้หักสองแผ่น", repairer: "ช่างสอง", reported_by: "ผู้แจ้ง", note: "บันทึก\nบรรทัดสอง" }));
  assert.strictEqual(d.message, "แก้ไขใบแจ้งซ่อม " + D.x.ticket_no + " แล้ว");
  const t = ticketOf(D.x.id);
  assert.deepStrictEqual([t.cause, t.repairer, t.reported_by, t.note, t.qty, t.stage], ["ไม้หักสองแผ่น", "ช่างสอง", "ผู้แจ้ง", "บันทึก\nบรรทัดสอง", 4, "done"]);
  const log = lastLog();
  assert.strictEqual(log.category, "repair");
  assert.strictEqual(log.action, "แก้ไขใบแจ้งซ่อม");
  assert.strictEqual(log.ref, D.x.ticket_no);
  assert.ok(log.detail.startsWith(`แก้ไขใบแจ้งซ่อม ${D.x.ticket_no}: สาเหตุการชำรุด "หัก" → "ไม้หักสองแผ่น"; ผู้แจ้ง "" → "ผู้แจ้ง"; ช่างผู้ซ่อม "" → "ช่างสอง"`), log.detail);
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
  assert.ok(html.includes('assets/app.js?v=14"'));
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
