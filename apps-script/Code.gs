/**
 * Pallet Hub backend for Google Apps Script (port of the PHP/MySQL api.php).
 *
 * Setup (once): run setupSystem() from the Apps Script editor. It creates the
 * "Pallet Hub Database" spreadsheet (one sheet per table, header row on top),
 * seeds the pallet types and departments, and stores the spreadsheet ID in
 * Script Properties (PALLET_SPREADSHEET_ID). Then set the shared password in
 * Project Settings > Script properties > PALLET_ACTION_PASSWORD and deploy as
 * a web app (Execute as: me, Who has access: Anyone).
 *
 * Optional: set Script property PALLET_RESET_PASSWORD (a separate password)
 * to enable the "reset_data" action, which wipes movements, repairs and
 * audit_logs (pallet types and departments are kept).
 *
 * Transport:
 *   GET  ?action=<read action>&...params        -> {ok:true,data} | {ok:false,error}
 *   POST text/plain JSON {action, password, actor, ...fields}
 *
 * Every write action requires the shared password. Reading is public.
 */

var SPREADSHEET_ID_PROPERTY = "PALLET_SPREADSHEET_ID";
var SPREADSHEET_NAME = "Pallet Hub Database";

// The shared password is stored ONLY in Script Properties (Project Settings >
// Script properties > PALLET_ACTION_PASSWORD) so it never appears in source.
// Repeated wrong attempts temporarily lock the check to slow down guessing.
var ACTION_PASSWORD_PROPERTY = "PALLET_ACTION_PASSWORD";
var PASSWORD_FAIL_CACHE_KEY = "PALLET_PASSWORD_FAILURES";
var PASSWORD_MAX_FAILURES = 10;
var PASSWORD_LOCK_SECONDS = 900;

// Separate password (Script Property only) for wiping test data, with its own
// lockout counter so it cannot be guessed via, or lock out, normal writes.
var RESET_ACTION = "reset_data";
var RESET_PASSWORD_PROPERTY = "PALLET_RESET_PASSWORD";
var RESET_FAIL_CACHE_KEY = "PALLET_RESET_FAILURES";
var RESET_TABLES = ["movements", "repairs", "audit_logs"]; // pallet_types/departments are kept

// Asia/Bangkok has no daylight saving time, so a fixed +07:00 offset gives the
// same wall-clock values as PHP's date_default_timezone_set('Asia/Bangkok').
var TZ_OFFSET_MS = 7 * 60 * 60 * 1000;

/* Column definitions (same columns as the MySQL tables in config.php).
 * Types: int, str, dt ("YYYY-MM-DD HH:MM:SS"). A trailing "?" marks columns
 * that are NULL in MySQL when empty (returned to the client as null). */
var SCHEMA = {
  pallet_types: [
    ["id", "int"], ["tkey", "str"], ["code", "str"], ["name", "str"], ["short", "str"],
    ["description", "str"], ["color", "str"], ["sizes", "str"], ["sort", "int"]
  ],
  departments: [
    ["id", "int"], ["name", "str"], ["icon", "str"], ["color", "str"], ["active", "int"]
  ],
  repairs: [
    ["id", "int"], ["ticket_no", "str"], ["type_id", "int"], ["size", "str"], ["qty", "int"],
    ["stage", "str"], ["source", "str"], ["department", "str?"], ["cause", "str"],
    ["reported_by", "str"], ["repairer", "str"], ["reported_at", "dt"], ["started_at", "dt?"],
    ["finished_at", "dt?"], ["note", "str"]
  ],
  movements: [
    ["id", "int"], ["doc_no", "str"], ["action", "str"], ["type_id", "int"], ["size", "str"],
    ["qty", "int"], ["from_status", "str?"], ["to_status", "str?"], ["department", "str?"],
    ["person", "str"], ["note", "str"], ["repair_id", "int?"], ["moved_at", "dt"], ["created_at", "dt"]
  ],
  audit_logs: [
    ["id", "int"], ["category", "str"], ["action", "str"], ["ref", "str"], ["detail", "str"],
    ["actor", "str"], ["ip", "str"], ["created_at", "dt"]
  ]
};
var TABLE_ORDER = ["pallet_types", "departments", "repairs", "movements", "audit_logs"];

var SEED_TYPES = [
  ["RM", "RM", "พาเลทสำหรับใส่ RM", "วัตถุดิบ (RM)", "วัตถุดิบ", "#1E6FE0", "1.2x1.2", 1],
  ["PK", "PK", "พาเลทสำหรับใส่ PK", "ผลิตภัณฑ์ (PK)", "ขวด + ฝา + แกลลอน ฯ", "#F5B400", "1.1x1.1,1.2x1.2", 2],
  ["FG", "FG", "พาเลทสำหรับใส่ FG", "สินค้าสำเร็จรูป (FG)", "สินค้าสำเร็จรูป", "#1FA83A", "1.2x1.2,1.1x1.1", 3],
  ["PL", "PK", "พาเลทพลาสติก (ใส่ลัง)", "พาเลทพลาสติก", "ใส่ลัง", "#7C8798", "1.2x1.2", 4]
];
var SEED_DEPARTMENTS = [
  ["ฝ่ายผลิต", "fa-industry", "#E2231A"],
  ["ฝ่ายบรรจุ", "fa-box-open", "#F5B400"],
  ["ฝ่ายคลังสินค้า", "fa-warehouse", "#1FA83A"],
  ["ฝ่ายจัดส่ง / ขนส่ง", "fa-truck-fast", "#E8590C"],
  ["ฝ่ายวัตถุดิบ", "fa-seedling", "#0CA678"],
  ["ฝ่ายซ่อมบำรุง", "fa-screwdriver-wrench", "#7048E8"],
  ["ฝ่ายควบคุมคุณภาพ (QC)", "fa-clipboard-check", "#D6336C"],
  ["ฝ่ายสิ่งแวดล้อม", "fa-leaf", "#2F9E44"]
];

var PREFIX = {
  receive: "RC", issue: "IS", "return": "RT", damage: "DM",
  repair_start: "RP", repair_done: "RD", scrap: "SC"
};
var ACT_NAME = {
  receive: "รับเข้า", issue: "เบิกจ่าย", "return": "รับคืน", damage: "แจ้งชำรุด",
  repair_start: "ส่งซ่อม", repair_done: "ซ่อมเสร็จ", scrap: "ตัดจำหน่าย"
};
var STATUS_NAME = {
  available: "พร้อมใช้", issued: "เบิกไปใช้งาน", damaged: "ชำรุด", repairing: "กำลังซ่อม", scrapped: "ตัดจำหน่าย"
};
var LOG_CAT_NAME = { pallet: "รายการพาเลท", repair: "งานซ่อม", setting: "ตั้งค่า", warn: "ถูกปฏิเสธ" };

var READ_ACTIONS = ["bootstrap", "dashboard", "repairs", "history", "export", "logs", "logs_export"];
var WRITE_ACTIONS = ["receive", "issue", "return", "damage", "repair_start", "repair_done", "scrap", "dept_save", "dept_delete"];

// Length limits of the MySQL VARCHAR columns that store user text.
var MAX_PERSON = 100;
var MAX_TEXT = 255;
var MAX_REPAIR_NOTE = 1000;

/* ===================== web app entry points ===================== */

function doGet(e) {
  var params = (e && e.parameter) || {};
  var action = String(params.action || "");
  try {
    if (!action) return jsonResponse_({ ok: true, data: { service: "Pallet Hub API" } });
    if (WRITE_ACTIONS.indexOf(action) !== -1 || action === "verifyPassword" || action === RESET_ACTION) {
      throw new Error("คำสั่งนี้ต้องส่งแบบ POST");
    }
    resetRequest_("");
    return jsonResponse_({ ok: true, data: handleRead_(action, params) });
  } catch (error) {
    return errorResponse_(error);
  }
}

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  } catch (parseError) {
    return jsonResponse_({ ok: false, error: "ข้อมูลที่ส่งมาไม่ใช่ JSON ที่ถูกต้อง" });
  }
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) payload = {};
  var action = String(payload.action || "");

  try {
    if (action === "verifyPassword") {
      return jsonResponse_({ ok: true, data: verifyActionPassword(payload.password) });
    }
    if (action === RESET_ACTION) {
      // Uses ONLY the reset password (PALLET_ACTION_PASSWORD is not required).
      assertResetPassword_(payload.resetPassword);
    } else {
      if (WRITE_ACTIONS.indexOf(action) === -1) {
        if (READ_ACTIONS.indexOf(action) !== -1) {
          resetRequest_("");
          return jsonResponse_({ ok: true, data: handleRead_(action, payload) });
        }
        throw new Error("Unknown action");
      }
      // Checked before taking the lock so password guessing (with its 1 s delay)
      // never blocks legitimate writers. Wrong passwords are not written to the
      // audit log: unauthenticated requests must not be able to write the sheet.
      assertActionPassword_(payload.password);
    }
  } catch (error) {
    return errorResponse_(error);
  }
  return runWrite_(action, payload);
}

function runWrite_(action, input) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockError) {
    return jsonResponse_({ ok: false, error: "ระบบกำลังบันทึกรายการอื่นอยู่ กรุณาลองใหม่อีกครั้ง" });
  }
  var actor = mbSubstr_(phpTrim_(str_(input.actor)), 0, MAX_PERSON);
  try {
    resetRequest_(actor);
    try {
      var result = handleWrite_(action, input);
      flush_();
      return jsonResponse_({ ok: true, data: result });
    } catch (error) {
      // Nothing buffered by the failed action is written (acts as a rollback).
      var message = errorMessage_(error);
      if (ACT_NAME[action]) {
        try {
          resetRequest_(actor);
          audit_("warn", action, "ปฏิเสธ" + ACT_NAME[action] + ": " + message, "", phpTrim_(safeStr_(input.person)));
          flush_();
        } catch (ignored) {}
      }
      return errorResponse_(error);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ===================== password ===================== */

function verifyActionPassword(password) {
  assertActionPassword_(password);
  return { valid: true };
}

function assertActionPassword_(password) {
  checkPassword_(password, ACTION_PASSWORD_PROPERTY, PASSWORD_FAIL_CACHE_KEY,
    "ยังไม่ได้ตั้งรหัสผ่านใน Script Properties (" + ACTION_PASSWORD_PROPERTY +
    ") กรุณาให้ผู้ดูแลระบบตั้งค่าที่ Project Settings > Script properties ก่อนบันทึกข้อมูล");
}

function assertResetPassword_(password) {
  checkPassword_(password, RESET_PASSWORD_PROPERTY, RESET_FAIL_CACHE_KEY,
    "ยังไม่ได้ตั้งรหัสรีเซ็ตข้อมูลใน Script Properties (" + RESET_PASSWORD_PROPERTY +
    ") กรุณาให้ผู้ดูแลระบบตั้งค่าที่ Project Settings > Script properties ก่อนใช้งานรีเซ็ตข้อมูล");
}

// Shared check: property must be set; >10 wrong attempts per 15 min (counted
// under failKey) locks the check; every wrong attempt costs 1 s.
function checkPassword_(password, property, failKey, missingMessage) {
  var expected = PropertiesService.getScriptProperties().getProperty(property) || "";
  if (!expected) throw passwordError_(missingMessage);

  var cache = CacheService.getScriptCache();
  var failures = Number(cache.get(failKey) || 0);
  if (failures >= PASSWORD_MAX_FAILURES) {
    throw passwordError_("ใส่รหัสผิดหลายครั้งเกินไป กรุณารอ 15 นาที / Too many wrong attempts. Try again in 15 minutes.");
  }

  if (String(password == null ? "" : password) !== expected) {
    cache.put(failKey, String(failures + 1), PASSWORD_LOCK_SECONDS);
    Utilities.sleep(1000);
    throw passwordError_("รหัสไม่ถูกต้อง / Incorrect password.");
  }
}

function passwordError_(message) {
  var error = new Error(message);
  error.passwordError = true;
  return error;
}

/* ===================== router ===================== */

function handleRead_(action, p) {
  switch (action) {
    case "bootstrap": return actionBootstrap_();
    case "dashboard": return actionDashboard_(p);
    case "repairs": return actionRepairs_();
    case "history": return { items: queryHistory_(p) };
    case "export": return exportHistory_(queryHistory_(p));
    case "logs": return { items: queryLogs_(p) };
    case "logs_export": return exportLogs_(queryLogs_(p));
  }
  fail_("Unknown action");
}

function handleWrite_(action, input) {
  switch (action) {
    case "receive": return actionReceive_(input);
    case "issue": return actionIssue_(input);
    case "return": return actionReturn_(input);
    case "damage": return actionDamage_(input);
    case "repair_start": return actionRepairStart_(input);
    case "repair_done": return actionRepairDone_(input);
    case "scrap": return actionScrap_(input);
    case "dept_save": return actionDeptSave_(input);
    case "dept_delete": return actionDeptDelete_(input);
    case RESET_ACTION: return actionResetData_();
  }
  fail_("Unknown action");
}

/* ===================== read actions ===================== */

function actionBootstrap_() {
  var now = nowParts_();
  return {
    types: sortBy_(table_("pallet_types").rows.slice(), function (a, b) { return a.sort - b.sort || a.id - b.id; }).map(plain_),
    departments: table_("departments").rows
      .filter(function (d) { return d.active === 1; })
      .sort(function (a, b) { return a.id - b.id; })
      .map(plain_),
    stock: stockMap_(),
    dept: deptOutstanding_(),
    now: { date: now.date, time: now.time }
  };
}

function actionDashboard_(p) {
  var n = phpInt_(p.days == null ? 7 : p.days);
  if ([7, 14, 30].indexOf(n) === -1) n = 7;
  var now = nowParts_();
  var today = now.date;
  var moves = table_("movements").rows;
  var repairs = table_("repairs").rows;

  var days = {};
  for (var i = n - 1; i >= 0; i--) {
    days[shiftDate_(today, -i)] = { receive: 0, issue: 0, "return": 0, damage: 0, repair_done: 0 };
  }
  var start = shiftDate_(today, -(n - 1));
  var hours = [];
  for (var h = 0; h < 24; h++) hours.push({ "in": 0, out: 0 });
  var todayAgg = {};
  var todayOrder = [];
  var since30 = shiftDateTime_(now.datetime, -30);
  var topAgg = {};
  var topOrder = [];

  moves.forEach(function (m) {
    var d = m.moved_at.slice(0, 10);
    if (m.moved_at >= start && days[d] && days[d][m.action] !== undefined) days[d][m.action] += m.qty;
    if (d === today) {
      var hr = parseInt(m.moved_at.slice(11, 13), 10) || 0;
      if (["receive", "return", "repair_done"].indexOf(m.action) !== -1) hours[hr]["in"] += m.qty;
      if (["issue", "damage"].indexOf(m.action) !== -1) hours[hr].out += m.qty;
      if (!todayAgg[m.action]) { todayAgg[m.action] = { action: m.action, q: 0, n: 0 }; todayOrder.push(m.action); }
      todayAgg[m.action].q += m.qty;
      todayAgg[m.action].n += 1;
    }
    if (m.action === "issue" && m.moved_at >= since30) {
      var key = m.department == null ? "\u0000" : m.department;
      if (!topAgg[key]) { topAgg[key] = { department: m.department, q: 0, n: 0 }; topOrder.push(key); }
      topAgg[key].q += m.qty;
      topAgg[key].n += 1;
    }
  });

  var stageAgg = {};
  var stageOrder = [];
  repairs.forEach(function (r) {
    if (!stageAgg[r.stage]) { stageAgg[r.stage] = { stage: r.stage, n: 0, q: 0 }; stageOrder.push(r.stage); }
    stageAgg[r.stage].n += 1;
    stageAgg[r.stage].q += r.qty;
  });

  var typeMap = typeMap_();
  var recent = sortBy_(moves.slice(), byMovedDesc_).slice(0, 12).map(function (m) {
    var t = typeMap[m.type_id];
    if (!t) return null;
    var o = plain_(m);
    o.tkey = t.tkey; o.color = t.color; o.code = t.code;
    return o;
  }).filter(Boolean);

  var topDept = sortBy_(topOrder.map(function (k) { return topAgg[k]; }), function (a, b) { return b.q - a.q; }).slice(0, 6);

  return {
    stock: stockMap_(),
    dept: deptOutstanding_(),
    days: days,
    today: todayOrder.sort().map(function (k) { return todayAgg[k]; }),
    repairs: stageOrder.map(function (k) { return stageAgg[k]; }),
    recent: recent,
    hours: hours,
    topDept: topDept,
    repairStat: repairStat_(repairs, now),
    logsToday: table_("audit_logs").rows.filter(function (l) { return l.created_at.slice(0, 10) === today; }).length
  };
}

function repairStat_(repairs, now) {
  var monthStart = now.date.slice(0, 8) + "01";
  var minutes = [];
  var openN = 0, doneM = 0, scrapM = 0, oldest = null;
  repairs.forEach(function (r) {
    if (r.stage === "done" && r.finished_at != null) {
      var diff = (parseDt_(r.finished_at) - parseDt_(r.reported_at)) / 60000;
      minutes.push(diff < 0 ? Math.ceil(diff) : Math.floor(diff)); // TIMESTAMPDIFF truncates
    }
    if (r.stage === "damaged" || r.stage === "repairing") {
      openN += 1;
      if (oldest === null || r.reported_at < oldest) oldest = r.reported_at;
    }
    if (r.stage === "done" && r.finished_at != null && r.finished_at >= monthStart) doneM += r.qty;
    if (r.stage === "scrapped" && r.finished_at != null && r.finished_at >= monthStart) scrapM += r.qty;
  });
  var avgH = null;
  if (minutes.length) {
    var avg = minutes.reduce(function (a, b) { return a + b; }, 0) / minutes.length / 60;
    avgH = (avg < 0 ? -1 : 1) * Math.round(Math.abs(avg) * 10 + 1e-9) / 10;
  }
  return { avg_h: avgH, open_n: openN, done_m: doneM, scrap_m: scrapM, oldest: oldest };
}

function actionRepairs_() {
  var since30 = shiftDateTime_(nowParts_().datetime, -30);
  var typeMap = typeMap_();
  var items = table_("repairs").rows
    .filter(function (r) {
      return r.stage === "damaged" || r.stage === "repairing" || (r.finished_at != null && r.finished_at >= since30);
    })
    .sort(function (a, b) { return b.id - a.id; })
    .map(function (r) {
      var t = typeMap[r.type_id];
      if (!t) return null;
      var o = plain_(r);
      o.tkey = t.tkey; o.code = t.code; o.color = t.color; o.type_name = t.name;
      return o;
    })
    .filter(Boolean);
  return { items: items };
}

function queryHistory_(p) {
  var typeMap = typeMap_();
  var from = phpEmpty_(p.from) ? null : str_(p.from) + " 00:00:00";
  var to = phpEmpty_(p.to) ? null : str_(p.to) + " 23:59:59";
  var type = phpEmpty_(p.type) ? null : phpInt_(p.type);
  var act = phpEmpty_(p.act) ? null : str_(p.act);
  var dept = phpEmpty_(p.dept) ? null : str_(p.dept);
  var q = phpEmpty_(p.q) ? null : likeMatcher_(str_(p.q));
  var rows = table_("movements").rows.filter(function (m) {
    if (!typeMap[m.type_id]) return false;
    if (from !== null && !(m.moved_at >= from)) return false;
    if (to !== null && !(m.moved_at <= to)) return false;
    if (type !== null && m.type_id !== type) return false;
    if (act !== null && !eqCi_(m.action, act)) return false;
    if (dept !== null && !eqCi_(m.department, dept)) return false;
    if (q && !(q(m.doc_no) || q(m.person) || q(m.note))) return false;
    return true;
  });
  return sortBy_(rows, byMovedDesc_).slice(0, 2000).map(function (m) {
    var t = typeMap[m.type_id];
    var o = plain_(m);
    o.tkey = t.tkey; o.code = t.code; o.color = t.color; o.type_name = t.name;
    return o;
  });
}

function exportHistory_(items) {
  var lines = [csvLine_(["เลขที่เอกสาร", "วันที่", "เวลา", "รายการ", "รหัส", "ประเภท", "ขนาด", "จำนวน", "ฝ่าย", "ผู้ทำรายการ", "หมายเหตุ"])];
  items.forEach(function (r) {
    lines.push(csvLine_([r.doc_no, r.moved_at.slice(0, 10), r.moved_at.slice(11, 16), ACT_NAME[r.action] || r.action,
      r.code, r.type_name, r.size, r.qty, r.department, r.person, r.note]));
  });
  return { filename: "pallet_history_" + stamp_() + ".csv", csv: lines.join("") };
}

function queryLogs_(p) {
  var from = phpEmpty_(p.from) ? null : str_(p.from) + " 00:00:00";
  var to = phpEmpty_(p.to) ? null : str_(p.to) + " 23:59:59";
  var cat = phpEmpty_(p.cat) ? null : str_(p.cat);
  var q = phpEmpty_(p.q) ? null : likeMatcher_(str_(p.q));
  return table_("audit_logs").rows.filter(function (l) {
    if (from !== null && !(l.created_at >= from)) return false;
    if (to !== null && !(l.created_at <= to)) return false;
    if (cat !== null && !eqCi_(l.category, cat)) return false;
    if (q && !(q(l.detail) || q(l.ref) || q(l.actor))) return false;
    return true;
  }).sort(function (a, b) { return b.id - a.id; }).slice(0, 3000).map(plain_);
}

function exportLogs_(items) {
  var lines = [csvLine_(["ลำดับ", "วันที่", "เวลา", "หมวด", "เลขที่อ้างอิง", "รายละเอียด", "ผู้ทำรายการ", "IP"])];
  items.forEach(function (r) {
    lines.push(csvLine_([r.id, r.created_at.slice(0, 10), r.created_at.slice(11, 19), LOG_CAT_NAME[r.category] || r.category,
      r.ref, r.detail, r.actor, r.ip]));
  });
  return { filename: "pallet_log_" + stamp_() + ".csv", csv: lines.join("") };
}

/* ===================== write actions ===================== */

function actionReceive_(input) {
  var b = baseInput_(input);
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  var note = textIn_(input, "note", MAX_TEXT, "หมายเหตุ");
  var at = moment_(input);
  var r = move_("receive", b.type, b.size, b.qty, null, "available", { person: person, note: note, moved_at: at });
  r.message = "รับเข้า " + b.qty + " ตัว เรียบร้อย";
  return r;
}

function actionIssue_(input) {
  var b = baseInput_(input);
  var dept = phpTrim_(safeStr_(input.department));
  if (dept === "") fail_("กรุณาเลือกฝ่ายที่เบิก");
  var have = qtyOf_(b.type, b.size, "available");
  if (b.qty > have) fail_("พาเลทพร้อมใช้ไม่พอ (คงเหลือ " + have + " ตัว)");
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  var note = textIn_(input, "note", MAX_TEXT, "หมายเหตุ");
  var r = move_("issue", b.type, b.size, b.qty, "available", "issued",
    { department: dept, person: person, note: note, moved_at: moment_(input) });
  r.message = "เบิกจ่ายให้ " + dept + " " + b.qty + " ตัว เรียบร้อย";
  return r;
}

function actionReturn_(input) {
  var b = baseInput_(input);
  var dept = phpTrim_(safeStr_(input.department));
  if (dept === "") fail_("กรุณาเลือกฝ่ายที่คืน");
  var have = deptQty_(dept, b.type, b.size);
  if (b.qty > have) fail_(dept + " ถือพาเลทนี้อยู่ " + have + " ตัว");
  var at = moment_(input);
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  var r, msg;
  if ((input.condition == null ? "good" : input.condition) === "damaged") {
    var cause = textIn_(input, "cause", MAX_TEXT, "สาเหตุการชำรุด");
    var rid = createTicket_(b.type, b.size, b.qty, "issued", dept, input, at);
    r = move_("damage", b.type, b.size, b.qty, "issued", "damaged",
      { department: dept, person: person, note: "คืนสภาพชำรุด: " + cause, repair_id: rid, moved_at: at });
    msg = "รับคืนชำรุด " + b.qty + " ตัว — เปิดใบแจ้งซ่อมแล้ว";
  } else {
    var note = textIn_(input, "note", MAX_TEXT, "หมายเหตุ");
    r = move_("return", b.type, b.size, b.qty, "issued", "available",
      { department: dept, person: person, note: note, moved_at: at });
    msg = "รับคืนจาก " + dept + " " + b.qty + " ตัว เรียบร้อย";
  }
  r.message = msg;
  return r;
}

function actionDamage_(input) {
  var b = baseInput_(input);
  var have = qtyOf_(b.type, b.size, "available");
  if (b.qty > have) fail_("พาเลทพร้อมใช้มีเพียง " + have + " ตัว");
  var at = moment_(input);
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  var cause = textIn_(input, "cause", MAX_TEXT, "สาเหตุการชำรุด");
  var rid = createTicket_(b.type, b.size, b.qty, "available", null, input, at);
  var r = move_("damage", b.type, b.size, b.qty, "available", "damaged",
    { person: person, note: cause, repair_id: rid, moved_at: at });
  r.message = "แจ้งชำรุด " + b.qty + " ตัว เรียบร้อย";
  return r;
}

function actionRepairStart_(input) {
  var rp = getRepair_(phpInt_(input.id));
  if (rp.stage !== "damaged") fail_("ใบนี้ไม่ได้อยู่สถานะชำรุด");
  var at = moment_(input);
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  update_("repairs", rp, { stage: "repairing", started_at: at, repairer: person });
  var r = move_("repair_start", rp.type_id, rp.size, rp.qty, "damaged", "repairing",
    { person: person, note: rp.ticket_no, repair_id: rp.id, moved_at: at });
  r.message = "ส่งซ่อม " + rp.ticket_no + " แล้ว";
  return r;
}

function actionRepairDone_(input) {
  var rp = getRepair_(phpInt_(input.id));
  if (rp.stage !== "repairing") fail_("ใบนี้ไม่ได้อยู่ระหว่างซ่อม");
  var at = moment_(input);
  var rawNote = safeStr_(input.note);
  if (mbLen_(rawNote) > MAX_REPAIR_NOTE) fail_("หมายเหตุยาวเกิน " + MAX_REPAIR_NOTE + " ตัวอักษร");
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  update_("repairs", rp, {
    stage: "done",
    finished_at: at,
    note: (rp.note == null ? "" : rp.note) + (phpTruthy_(rawNote) ? "\nซ่อมเสร็จ: " + rawNote : "")
  });
  var r = move_("repair_done", rp.type_id, rp.size, rp.qty, "repairing", "available",
    { person: phpTruthy_(person) ? person : rp.repairer, note: rp.ticket_no, repair_id: rp.id, moved_at: at });
  r.message = "ซ่อมเสร็จ " + rp.qty + " ตัว กลับเข้าคลังพร้อมใช้";
  return r;
}

function actionScrap_(input) {
  var rp = getRepair_(phpInt_(input.id));
  if (rp.stage !== "damaged" && rp.stage !== "repairing") fail_("ไม่สามารถตัดจำหน่ายใบนี้ได้");
  var at = moment_(input);
  var person = textIn_(input, "person", MAX_PERSON, "ชื่อผู้ทำรายการ");
  var note = textIn_(input, "note", MAX_TEXT, "หมายเหตุ");
  var fromStage = rp.stage;
  update_("repairs", rp, { stage: "scrapped", finished_at: at });
  var r = move_("scrap", rp.type_id, rp.size, rp.qty, fromStage, "scrapped",
    { person: person, note: rp.ticket_no + " " + note, repair_id: rp.id, moved_at: at });
  r.message = "ตัดจำหน่าย " + rp.qty + " ตัว แล้ว";
  return r;
}

function actionDeptSave_(input) {
  var name = phpTrim_(safeStr_(input.name));
  if (name === "") fail_("กรุณาระบุชื่อฝ่าย");
  if (mbLen_(name) > 100) fail_("ชื่อฝ่ายยาวเกิน 100 ตัวอักษร");
  var icon = phpTruthy_(input.icon) ? mbSubstr_(safeStr_(input.icon), 0, 40) : "fa-building";
  var color = phpTruthy_(input.color) ? mbSubstr_(safeStr_(input.color), 0, 10) : "#E2231A";
  var depts = table_("departments");
  // UNIQUE(name) with a case-insensitive collation -> ON DUPLICATE KEY UPDATE
  var existing = null;
  depts.rows.forEach(function (d) { if (!existing && eqCi_(d.name, name)) existing = d; });
  if (existing) {
    update_("departments", existing, { active: 1, icon: icon, color: color });
  } else {
    insert_("departments", { name: name, icon: icon, color: color, active: 1 });
  }
  audit_("setting", "dept_save", "เพิ่ม/แก้ไขฝ่าย: " + name);
  return { message: "เพิ่ม " + name + " แล้ว" };
}

function actionDeptDelete_(input) {
  var id = phpInt_(input.id);
  var dept = null;
  table_("departments").rows.forEach(function (d) { if (!dept && d.id === id) dept = d; });
  if (!dept || !phpTruthy_(dept.name)) fail_("ไม่พบฝ่าย");
  update_("departments", dept, { active: 0 });
  audit_("setting", "dept_delete", "ลบฝ่าย: " + dept.name);
  return { message: "ลบฝ่ายแล้ว" };
}

// Wipes all data rows (header row and cell formats are kept) of the
// transactional sheets. Runs under the script lock (runWrite_).
function actionResetData_() {
  var removed = {};
  RESET_TABLES.forEach(function (name) {
    var sheet = db_().getSheetByName(name);
    if (!sheet) throw new Error('ไม่พบชีต "' + name + '" กรุณารัน setupSystem() อีกครั้ง');
    var n = Math.max(0, sheet.getLastRow() - 1);
    if (n > 0) {
      var width = Math.max(sheet.getLastColumn(), SCHEMA[name].length);
      sheet.getRange(2, 1, n, width).clearContent(); // one call; keeps number formats
    }
    removed[name] = n;
    // Drop anything cached/buffered for this sheet so ids and doc numbers
    // are recomputed from the now-empty sheet (next id 1, next doc -0001).
    delete REQ_.tables[name];
    delete REQ_.appends[name];
  });
  REQ_.updates = REQ_.updates.filter(function (u) { return RESET_TABLES.indexOf(u.name) === -1; });

  audit_("setting", "รีเซ็ตข้อมูล",
    "รีเซ็ตข้อมูล: ล้างชีต movements (" + removed.movements + " แถว), repairs (" + removed.repairs +
    " แถว), audit_logs (" + removed.audit_logs + " แถว) · คงไว้: pallet_types, departments");
  return {
    removed: removed,
    message: "รีเซ็ตข้อมูลแล้ว — ลบประวัติเคลื่อนไหว " + removed.movements + " แถว, งานซ่อม " +
      removed.repairs + " แถว, บันทึกประวัติ " + removed.audit_logs + " แถว"
  };
}

/* ===================== business helpers ===================== */

function stockMap_() {
  var map = {};
  var add = function (type, size, st, q) {
    if (!map[type]) map[type] = {};
    if (!map[type][size]) map[type][size] = {};
    map[type][size][st] = (map[type][size][st] || 0) + q;
  };
  table_("movements").rows.forEach(function (m) {
    if (m.to_status != null) add(m.type_id, m.size, m.to_status, m.qty);
    if (m.from_status != null) add(m.type_id, m.size, m.from_status, -m.qty);
  });
  return map;
}

function qtyOf_(type, size, status) {
  var map = stockMap_();
  return (map[type] && map[type][size] && map[type][size][status]) || 0;
}

function deptOutstanding_() {
  var groups = {};
  var order = [];
  table_("movements").rows.forEach(function (m) {
    if (m.department == null || m.department === "") return;
    var key = m.department + "\u0000" + m.type_id + "\u0000" + m.size;
    if (!groups[key]) {
      groups[key] = { department: m.department, type_id: m.type_id, size: m.size, qty: 0 };
      order.push(key);
    }
    groups[key].qty += m.to_status === "issued" ? m.qty : (m.from_status === "issued" ? -m.qty : 0);
  });
  var rows = order.map(function (k) { return groups[k]; }).filter(function (r) { return r.qty !== 0; });
  return sortBy_(rows, function (a, b) { return a.department.localeCompare(b.department, "th"); });
}

function deptQty_(dept, type, size) {
  var list = deptOutstanding_();
  for (var i = 0; i < list.length; i++) {
    if (list[i].department === dept && list[i].type_id === type && list[i].size === size) return list[i].qty;
  }
  return 0;
}

function getType_(id) {
  var rows = table_("pallet_types").rows;
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  fail_("ไม่พบประเภทพาเลท");
}

function typeMap_() {
  var map = {};
  table_("pallet_types").rows.forEach(function (t) { map[t.id] = t; });
  return map;
}

function getRepair_(id) {
  var rows = table_("repairs").rows;
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  fail_("ไม่พบใบแจ้งซ่อม");
}

function baseInput_(input) {
  var type = phpInt_(input.type_id);
  var t = getType_(type);
  var size = phpTrim_(safeStr_(input.size));
  if (String(t.sizes).split(",").indexOf(size) === -1) fail_("กรุณาเลือกขนาดพาเลท");
  var qty = phpInt_(input.qty);
  if (qty <= 0) fail_("จำนวนต้องมากกว่า 0");
  return { type: type, size: size, qty: qty, t: t };
}

// date/time from the client (default: now), validated like DateTime::createFromFormat('Y-m-d H:i')
function moment_(input) {
  var now = nowParts_();
  var d = input.date == null ? now.date : safeStr_(input.date);
  var t = input.time == null ? now.time : safeStr_(input.time);
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{2})$/.exec(d + " " + t.slice(0, 5));
  if (!m) fail_("วันที่/เวลาไม่ถูกต้อง");
  var dt = new Date(0);
  dt.setUTCFullYear(Number(m[1]), Number(m[2]) - 1, Number(m[3])); // overflow (e.g. 02-30) rolls over like PHP
  dt.setUTCHours(Number(m[4]), Number(m[5]), 0, 0);
  return formatUtcParts_(dt).datetime.slice(0, 16) + ":00";
}

function nextNo_(prefix, table, col) {
  var tag = prefix + "-" + nowParts_().ymd + "-";
  var count = table_(table).rows.filter(function (r) {
    return String(r[col] || "").toUpperCase().indexOf(tag) === 0;
  }).length;
  return tag + ("0000" + (count + 1)).slice(-Math.max(4, String(count + 1).length));
}

function move_(action, type, size, qty, from, to, extra) {
  var doc = nextNo_(PREFIX[action], "movements", "doc_no");
  var movedAt = extra.moved_at || nowParts_().datetime;
  var row = insert_("movements", {
    doc_no: doc, action: action, type_id: type, size: size, qty: qty,
    from_status: from, to_status: to,
    department: extra.department == null ? null : extra.department,
    person: extra.person == null ? "" : extra.person,
    note: extra.note == null ? "" : extra.note,
    repair_id: extra.repair_id == null ? null : extra.repair_id,
    moved_at: movedAt,
    created_at: nowParts_().datetime
  });
  var t = getType_(type);
  var detail = ACT_NAME[action] + " " + t.code + " (" + t.name + ") ขนาด " + size + " ม. จำนวน " + qty + " ตัว" +
    " [" + (from ? STATUS_NAME[from] : "ภายนอก") + " → " + STATUS_NAME[to] + "]" +
    (!phpEmpty_(extra.department) ? " ฝ่าย: " + extra.department : "") +
    " · เวลาทำรายการ " + movedAt.slice(8, 10) + "/" + movedAt.slice(5, 7) + "/" + movedAt.slice(0, 4) + " " + movedAt.slice(11, 16) +
    (!phpEmpty_(extra.note) ? " · " + extra.note : "");
  var cat = ["damage", "repair_start", "repair_done", "scrap"].indexOf(action) !== -1 ? "repair" : "pallet";
  audit_(cat, action, detail, doc, phpTrim_(extra.person == null ? "" : extra.person));
  return { doc_no: doc, id: row.id };
}

function createTicket_(type, size, qty, source, dept, input, at) {
  var row = insert_("repairs", {
    ticket_no: nextNo_("RPR", "repairs", "ticket_no"),
    type_id: type, size: size, qty: qty, stage: "damaged", source: source,
    department: dept, cause: phpTrim_(safeStr_(input.cause)), reported_by: phpTrim_(safeStr_(input.person)),
    repairer: "", reported_at: at, started_at: null, finished_at: null, note: phpTrim_(safeStr_(input.note))
  });
  return row.id;
}

function audit_(category, action, detail, ref, actor) {
  actor = actor == null ? "" : actor;
  insert_("audit_logs", {
    category: category, action: action, ref: ref || "",
    detail: mbSubstr_(detail, 0, 500),
    actor: actor !== "" ? actor : REQ_.actor,
    ip: "web", // GAS cannot see the client IP
    created_at: nowParts_().datetime
  });
}

function textIn_(input, key, max, label) {
  var value = phpTrim_(safeStr_(input[key]));
  if (mbLen_(value) > max) fail_(label + "ยาวเกิน " + max + " ตัวอักษร");
  return value;
}

/* ===================== sheet storage ===================== */

var REQ_ = null; // per-request state: spreadsheet, loaded tables, pending writes

function resetRequest_(actor) {
  REQ_ = { ss: null, tables: {}, appends: {}, updates: [], actor: actor || "" };
}

function db_() {
  if (!REQ_.ss) {
    var id = PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY);
    if (!id) throw new Error("ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ใน Apps Script ก่อน");
    REQ_.ss = SpreadsheetApp.openById(id);
  }
  return REQ_.ss;
}

// Reads the whole sheet once per request and converts rows to typed objects.
function table_(name) {
  if (REQ_.tables[name]) return REQ_.tables[name];
  var sheet = db_().getSheetByName(name);
  if (!sheet) throw new Error('ไม่พบชีต "' + name + '" กรุณารัน setupSystem() อีกครั้ง');
  var values = sheet.getDataRange().getValues();
  var header = (values[0] || []).map(function (h) { return String(h).trim(); });
  var cols = SCHEMA[name];
  var idx = cols.map(function (c) { return header.indexOf(c[0]); });
  for (var c = 0; c < cols.length; c++) {
    if (idx[c] === -1) throw new Error('หัวตารางของชีต "' + name + '" ไม่ครบ (ไม่พบคอลัมน์ ' + cols[c][0] + ") กรุณารัน setupSystem()");
  }
  var rows = [];
  var maxId = 0;
  for (var r = 1; r < values.length; r++) {
    var raw = values[r];
    var empty = true;
    for (var k = 0; k < raw.length; k++) if (raw[k] !== "" && raw[k] !== null) { empty = false; break; }
    if (empty) continue;
    var obj = {};
    cols.forEach(function (col, i) { obj[col[0]] = fromCell_(col[1], raw[idx[i]]); });
    Object.defineProperty(obj, "_row", { value: r + 1, writable: true, enumerable: false });
    if (obj.id > maxId) maxId = obj.id;
    rows.push(obj);
  }
  REQ_.tables[name] = { name: name, sheet: sheet, header: header, idx: idx, rows: rows, maxId: maxId };
  return REQ_.tables[name];
}

function insert_(name, obj) {
  var t = table_(name);
  var row = {};
  SCHEMA[name].forEach(function (col) { row[col[0]] = obj[col[0]] === undefined ? null : obj[col[0]]; });
  t.maxId += 1;
  row.id = t.maxId;
  Object.defineProperty(row, "_row", { value: 0, writable: true, enumerable: false });
  t.rows.push(row);
  (REQ_.appends[name] = REQ_.appends[name] || []).push(row);
  return row;
}

function update_(name, row, fields) {
  Object.keys(fields).forEach(function (k) { row[k] = fields[k]; });
  if (row._row > 0) {
    for (var i = 0; i < REQ_.updates.length; i++) if (REQ_.updates[i].row === row) return;
    REQ_.updates.push({ name: name, row: row });
  } // rows inserted in this request are written with their latest values on flush
}

// Writes buffered rows: one setNumberFormats + one setValues per table.
function flush_() {
  Object.keys(REQ_.appends).forEach(function (name) {
    var list = REQ_.appends[name];
    if (!list.length) return;
    var t = REQ_.tables[name];
    var width = t.header.length;
    var start = Math.max(t.sheet.getLastRow(), 1) + 1;
    var needed = start + list.length - 1;
    var maxRows = t.sheet.getMaxRows();
    if (needed > maxRows) t.sheet.insertRowsAfter(maxRows, needed - maxRows);
    var formats = [];
    var values = list.map(function (row) {
      formats.push(rowFormats_(name, t));
      row._row = start + formats.length - 1;
      return rowValues_(name, t, row, width);
    });
    var range = t.sheet.getRange(start, 1, list.length, width);
    range.setNumberFormats(formats);
    range.setValues(values);
  });
  REQ_.updates.forEach(function (u) {
    var t = REQ_.tables[u.name];
    var range = t.sheet.getRange(u.row._row, 1, 1, t.header.length);
    var current = range.getValues()[0];
    range.setValues([rowValues_(u.name, t, u.row, t.header.length, current)]);
  });
  REQ_.appends = {};
  REQ_.updates = [];
}

function rowValues_(name, t, row, width, base) {
  var out = base ? base.slice() : [];
  while (out.length < width) out.push("");
  SCHEMA[name].forEach(function (col, i) { out[t.idx[i]] = toCell_(col[1], row[col[0]]); });
  return out;
}

function rowFormats_(name, t) {
  var out = [];
  for (var i = 0; i < t.header.length; i++) out.push("@");
  SCHEMA[name].forEach(function (col, i) { out[t.idx[i]] = col[1].indexOf("int") === 0 ? "0" : "@"; });
  return out;
}

function toCell_(type, v) {
  if (v === null || v === undefined) return "";
  if (type.indexOf("int") === 0) return Number(v);
  var s = String(v);
  // Never let user text become a formula (cells are also plain-text formatted).
  if (s.charAt(0) === "=") s = "'" + s;
  return s;
}

function fromCell_(type, v) {
  var nullable = type.charAt(type.length - 1) === "?";
  var base = nullable ? type.slice(0, -1) : type;
  if (v === "" || v === null || v === undefined) return nullable ? null : (base === "int" ? 0 : "");
  if (base === "int") {
    var n = Number(v);
    return isFinite(n) ? Math.trunc(n) : 0;
  }
  if (Object.prototype.toString.call(v) === "[object Date]") {
    // A cell Sheets converted to a Date (spreadsheet time zone is Asia/Bangkok).
    return formatUtcParts_(new Date(v.getTime() + TZ_OFFSET_MS)).datetime;
  }
  var s = String(v);
  if (s.charAt(0) === "'" && s.charAt(1) === "=") s = s.slice(1);
  return s;
}

/* ===================== setup ===================== */

/**
 * One-time setup. Run from the Apps Script editor and approve permissions.
 * If PALLET_SPREADSHEET_ID is already set, the existing spreadsheet is reused
 * (missing sheets/headers/seed data are added) instead of creating another.
 */
function setupSystem() {
  var properties = PropertiesService.getScriptProperties();
  var id = properties.getProperty(SPREADSHEET_ID_PROPERTY);
  var ss;
  var created = false;
  if (id) {
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (error) {
      throw new Error("เปิดสเปรดชีตเดิม (" + id + ") ไม่ได้: " + errorMessage_(error) +
        " — ถ้าต้องการสร้างใหม่ ให้ลบ Script Property " + SPREADSHEET_ID_PROPERTY + " แล้วรันอีกครั้ง");
    }
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    ss.setSpreadsheetTimeZone("Asia/Bangkok");
    created = true;
    properties.setProperty(SPREADSHEET_ID_PROPERTY, ss.getId());
  }

  var defaultSheets = created ? ss.getSheets() : [];
  TABLE_ORDER.forEach(function (name, i) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      if (created && i === 0 && defaultSheets.length) {
        sheet = defaultSheets[0];
        sheet.setName(name);
      } else {
        sheet = ss.insertSheet(name);
      }
    }
    var cols = SCHEMA[name];
    var headers = cols.map(function (c) { return c[0]; });
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    // Plain-text format for text/date columns so Sheets never converts values
    // (e.g. "2026-10-05 08:00:00" into a Date or "=..." into a formula).
    var maxRows = sheet.getMaxRows();
    var formats = [];
    var rowFormat = cols.map(function (c) { return c[1].indexOf("int") === 0 ? "0" : "@"; });
    for (var r = 0; r < maxRows; r++) formats.push(rowFormat);
    sheet.getRange(1, 1, maxRows, headers.length).setNumberFormats(formats);
  });

  REQ_ = { ss: ss, tables: {}, appends: {}, updates: [], actor: "" };
  if (!table_("pallet_types").rows.length) {
    SEED_TYPES.forEach(function (t) {
      insert_("pallet_types", { tkey: t[0], code: t[1], name: t[2], short: t[3], description: t[4], color: t[5], sizes: t[6], sort: t[7] });
    });
  }
  if (!table_("departments").rows.length) {
    SEED_DEPARTMENTS.forEach(function (d) {
      insert_("departments", { name: d[0], icon: d[1], color: d[2], active: 1 });
    });
  }
  flush_();
  REQ_ = null;

  var result = {
    created: created,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    passwordConfigured: Boolean(properties.getProperty(ACTION_PASSWORD_PROPERTY))
  };
  console.log(JSON.stringify(result));
  return result;
}

/* ===================== small utilities ===================== */

function fail_(message) {
  throw new Error(message);
}

function errorMessage_(error) {
  return (error && error.message) ? String(error.message) : String(error);
}

function errorResponse_(error) {
  var body = { ok: false, error: errorMessage_(error) };
  if (error && error.passwordError) body.passwordError = true;
  return jsonResponse_(body);
}

function jsonResponse_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function plain_(obj) {
  var out = {};
  Object.keys(obj).forEach(function (k) { out[k] = obj[k]; });
  return out;
}

function sortBy_(arr, cmp) { // stable sort
  return arr.map(function (v, i) { return [v, i]; })
    .sort(function (a, b) { return cmp(a[0], b[0]) || a[1] - b[1]; })
    .map(function (p) { return p[0]; });
}

function byMovedDesc_(a, b) {
  if (a.moved_at !== b.moved_at) return a.moved_at < b.moved_at ? 1 : -1;
  return b.id - a.id;
}

function str_(v) {
  if (v === null || v === undefined || v === false) return "";
  if (v === true) return "1";
  return String(v);
}

// Like str_ but rejects arrays/objects (PHP trim() on an array throws).
function safeStr_(v) {
  if (v !== null && typeof v === "object") fail_("ข้อมูลที่ส่งมาไม่ถูกต้อง");
  return str_(v);
}

function phpTrim_(s) {
  return String(s).replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

function phpInt_(v) {
  if (v === null || v === undefined || v === false) return 0;
  if (v === true) return 1;
  if (typeof v === "number") return isFinite(v) ? Math.trunc(v) : 0;
  if (typeof v === "object") return Array.isArray(v) && v.length ? 1 : 0;
  var m = /^[ \t\n\r\v\f]*([+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?)/.exec(String(v));
  if (!m) return 0;
  var n = Number(m[1]);
  return isFinite(n) ? Math.trunc(n) : 0;
}

function phpTruthy_(v) {
  if (v === null || v === undefined || v === false || v === 0 || v === "" || v === "0") return false;
  if (Array.isArray(v)) return v.length > 0;
  return true;
}

function phpEmpty_(v) {
  return !phpTruthy_(v);
}

function mbLen_(s) {
  return Array.from(String(s)).length;
}

function mbSubstr_(s, start, len) {
  return Array.from(String(s)).slice(start, start + len).join("");
}

function eqCi_(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return false;
  return String(a).toLowerCase() === String(b).toLowerCase();
}

// MySQL "col LIKE '%q%'" (case-insensitive, % and _ wildcards, \ escape).
function likeMatcher_(q) {
  var re = "";
  var chars = Array.from(q);
  for (var i = 0; i < chars.length; i++) {
    var ch = chars[i];
    if (ch === "\\" && i + 1 < chars.length) { re += escapeRe_(chars[++i]); continue; }
    if (ch === "%") re += "[\\s\\S]*";
    else if (ch === "_") re += "[\\s\\S]";
    else re += escapeRe_(ch);
  }
  var rx = new RegExp(re, "i");
  return function (v) { return v !== null && v !== undefined && rx.test(String(v)); };
}

function escapeRe_(ch) {
  return ch.replace(/[.*+?^${}()|[\]\\\/]/g, "\\$&");
}

// PHP fputcsv() with default delimiter/enclosure/escape, "\n" line ending.
function csvLine_(fields) {
  return fields.map(function (f) {
    var s = f === null || f === undefined ? "" : (f === true ? "1" : f === false ? "" : String(f));
    if (!/[,"\\\n\r\t ]/.test(s)) return s;
    var out = '"';
    var escaped = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (ch === "\\") escaped = true;
      else if (!escaped && ch === '"') out += '"';
      else escaped = false;
      out += ch;
    }
    return out + '"';
  }).join(",") + "\n";
}

function pad2_(n) {
  return (n < 10 ? "0" : "") + n;
}

function formatUtcParts_(d) {
  var y = String(d.getUTCFullYear());
  while (y.length < 4) y = "0" + y;
  var date = y + "-" + pad2_(d.getUTCMonth() + 1) + "-" + pad2_(d.getUTCDate());
  var time = pad2_(d.getUTCHours()) + ":" + pad2_(d.getUTCMinutes());
  return {
    date: date,
    time: time,
    datetime: date + " " + time + ":" + pad2_(d.getUTCSeconds()),
    ymd: y.slice(-2) + pad2_(d.getUTCMonth() + 1) + pad2_(d.getUTCDate())
  };
}

// Current wall-clock time in Asia/Bangkok.
function nowParts_() {
  return formatUtcParts_(new Date(Date.now() + TZ_OFFSET_MS));
}

function parseDt_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(s));
  if (!m) return NaN;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
}

function shiftDate_(date, days) {
  return formatUtcParts_(new Date(parseDt_(date) + days * 86400000)).date;
}

function shiftDateTime_(datetime, days) {
  return formatUtcParts_(new Date(parseDt_(datetime) + days * 86400000)).datetime;
}

function stamp_() {
  var p = nowParts_();
  return p.date.replace(/-/g, "") + "_" + p.datetime.slice(11).replace(/:/g, "");
}
