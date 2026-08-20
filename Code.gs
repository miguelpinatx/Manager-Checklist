/**********************************************************************
 *  MANAGER DAILY LOG — Google Sheet backend
 *  ------------------------------------------------------------------
 *  Paste this into Extensions > Apps Script on a new Google Sheet.
 *  Then set the office passcode below, deploy as a Web App, and paste
 *  the /exec URL into the manager app's Settings.
 *
 *  Full step-by-step instructions are in SETUP.md.
 **********************************************************************/

/* ====== SETTINGS — change these ====== */
var OFFICE_PASSCODE = "changeme";      // what main office types to open the dashboard
var STORE_CODE      = "changeme2";     // what each manager types once on their phone
var RETAIN_DAYS     = 15;              // keep this many days on the live tabs
var LOOKBACK_DAYS   = 3;               // how far back managers can read on their phones

/* Where the nightly digest goes. Leave a line blank and that department's
   items fall through to the owner instead. Leave OWNER_EMAIL blank and it
   goes to whichever Google account owns this script.                     */
var OWNER_EMAIL  = "";
var ROUTE_EMAIL  = {
  "Higher Management": "",
  "HR":                "",
  "Commissary":        "",
  "Maintenance":       ""
};
var DIGEST_HOUR  = 22;                 // 22 = 10pm, after the last store closes

/* One code for every store by default. To give each store its own code,
   fill in the names exactly as they appear in the manager app and they
   take over from STORE_CODE above.                                       */
var STORE_CODES = {
  // "Central":  "",
  // "Alameda":  "",
  // "Horizon":  "",
  // "Urban":    "",
  // "Doniphan": "",
  // "Piedras":  "",
  // "Montana":  ""
};
/* ====================================== */

var SHIFTS_SHEET   = "Shifts";
var BULLETIN_SHEET = "Vista Updates";
var TASKS_SHEET    = "Priority Tasks";
var TASKLOG_SHEET  = "Task log";
var ENTRIES_SHEET  = "Entries";
var MANAGERS_SHEET = "Managers";
var ARCHIVE_SUFFIX = " (archive)";

var SHIFT_COLS = ["Received","Date","Store","Shift","Manager","Day type",
                  "Departments","Logged","All good","Needs attention","Not covered",
                  "Issues","Notes","Walk confirmed","Handoff note","Full report"];
var ENTRY_COLS = ["Entry ID","Received","Date","Store","Shift","Manager","Time block",
                  "Department","Dept status","Type","Route to","Detail",
                  "Handled","Handled at"];

/* The "Managers" tab is the roster managers pick their name from, so the
   office gets one spelling instead of four.
   Store  — blank for every store, or one store's name
   Active — blank or TRUE to show them; FALSE retires them                */
var MANAGER_COLS = ["Store","Name","Active"];

/* The "Vista Updates" tab is where main office types announcements for managers.
   Store   — leave blank for every store, or type one store's name
   Starts / Ends — leave blank for "show it now, until I turn it off"
   Active  — leave blank or type TRUE to show it; FALSE hides it          */
var BULLETIN_COLS = ["Posted","Store","Headline","Message","Starts","Ends","Active"];

/* The "Priority Tasks" tab is the must-do list office sets for managers.
   Store  — blank for every store, or one store's name
   Shift  — blank for every shift, or Opening / 2nd Shift / Closing
   Active — blank or TRUE to show it; FALSE retires it                   */
var TASKS_COLS   = ["Posted","Store","Shift","Task","Active"];
var TASKLOG_COLS = ["Received","Date","Store","Shift","Manager","Task","Done","Marked at"];

/* ------------------------------------------------------------------ */
/* Receiving a submitted shift                                         */
/* ------------------------------------------------------------------ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    if (!p || !p.date || !p.shiftId) throw new Error("Missing date or shift");
    if (!codeOk(p.code, p.store)) return json({ ok: false, error: "bad-code" });

    var now = new Date();
    var c   = p.counts || {};
    var wrap = p.wrap || {};

    // A resend of the same store+date+shift replaces the earlier one.
    removeExisting(sheet(SHIFTS_SHEET, SHIFT_COLS), p);

    sheet(SHIFTS_SHEET, SHIFT_COLS).appendRow([
      now, p.date, p.store || "", p.shift || "", p.manager || "", p.dayType || "",
      c.total || 0, c.reported || 0, c.good || 0, c.attn || 0, c.skip || 0,
      c.issues || 0, c.entries || 0,
      wrap.done ? "Yes" : "No", wrap.note || "", p.report || ""
    ]);

    /* Anything already marked handled keeps that mark when a shift is resent. */
    var wasHandled = handledMap(p);

    var es = sheet(ENTRIES_SHEET, ENTRY_COLS);
    removeExisting(es, p);
    (p.entries || []).forEach(function (en, i) {
      var id = entryId(p, i, en);
      es.appendRow([
        id, now, p.date, p.store || "", p.shift || "", p.manager || "",
        en.block || "", en.department || "", statusWord(en.status),
        en.type || "", en.route || "", en.text || "",
        wasHandled[id] ? "Yes" : "", wasHandled[id] || ""
      ]);
    });

    var tasks = p.tasks || [];
    if (tasks.length) {
      var ts = sheet(TASKLOG_SHEET, TASKLOG_COLS);
      removeExisting(ts, p);
      tasks.forEach(function (t) {
        ts.appendRow([now, p.date, p.store || "", p.shift || "", p.manager || "",
                      t.task || "", t.done ? "Done" : "Not done", t.at || ""]);
      });
    }

    return json({ ok: true, saved: (p.entries || []).length, tasks: tasks.length });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ------------------------------------------------------------------ */
/* Serving the office dashboard + its data                             */
/* ------------------------------------------------------------------ */
function doGet(e) {
  var p = (e && e.parameter) || {};

  // the roster managers pick their name from
  if (p.action === "managers") {
    if (!codeOk(p.code, p.store)) return json({ ok: false, error: "bad-code" });
    return json({ ok: true, managers: readManagers(String(p.store || "")) });
  }

  // the priority tasks office wants ticked off
  if (p.action === "tasks") {
    if (!codeOk(p.code, p.store)) return json({ ok: false, error: "bad-code" });
    return json({ ok: true, tasks: readTasks(String(p.store || "")) });
  }

  // the bulletin managers see at the top of their shift
  if (p.action === "bulletin") {
    if (!codeOk(p.code, p.store)) return json({ ok: false, error: "bad-code" });
    return json({ ok: true, bulletins: readBulletins(String(p.store || "")) });
  }

  // managers reading the last few days for their own store, from their phone
  if (p.action === "recent") {
    if (!codeOk(p.code, p.store)) return json({ ok: false, error: "bad-code" });
    var days  = Math.min(Number(p.days) || LOOKBACK_DAYS, RETAIN_DAYS);
    var store = String(p.store || "");
    var cut = new Date(); cut.setHours(0,0,0,0); cut.setDate(cut.getDate() - (days - 1));
    var keep = function (r) {
      if (store && String(r.Store) !== store) return false;
      var d = asDate(r.Date);
      return d && d >= cut;
    };
    return json({
      ok: true, days: days, store: store,
      shifts:  readAll(SHIFTS_SHEET).filter(keep).map(function (r) {
        return { Date: r.Date, Shift: r.Shift, Manager: r.Manager, Issues: r.Issues, Notes: r.Notes };
      }),
      entries: readAll(ENTRIES_SHEET).filter(keep).map(function (r) {
        return { Date: r.Date, Shift: r.Shift, Manager: r.Manager, Department: r.Department,
                 Type: r.Type, "Route to": r["Route to"], Detail: r.Detail };
      })
    });
  }

  if (p.action === "handle") {
    return json(markHandled(p.pass, String(p.ids || "").split(",").filter(String),
                            String(p.on || "1") !== "0"));
  }

  if (p.action === "data") {
    if (String(p.pass || "") !== OFFICE_PASSCODE) return json({ ok: false, error: "Wrong passcode" });
    var all = String(p.all || "") === "1";      // reach back into the archive tabs too
    var sh = readAll(SHIFTS_SHEET), en = readAll(ENTRIES_SHEET);
    if (all) {
      sh = readAll(SHIFTS_SHEET  + ARCHIVE_SUFFIX).concat(sh);
      en = readAll(ENTRIES_SHEET + ARCHIVE_SUFFIX).concat(en);
    }
    return json({ ok: true, days: all ? 0 : RETAIN_DAYS, archived: all, shifts: sh, entries: en });
  }

  return HtmlService.createHtmlOutputFromFile("Office")
    .setTitle("Manager Log — Office")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/* ------------------------------------------------------------------ */
/* Retention — run daily from a time trigger (see setupDailyPurge)      */
/* ------------------------------------------------------------------ */
function purgeOld() {
  [SHIFTS_SHEET, ENTRIES_SHEET].forEach(function (name) {
    var sh = sheet(name, name === SHIFTS_SHEET ? SHIFT_COLS : ENTRY_COLS);
    var rows = sh.getDataRange().getValues();
    if (rows.length < 2) return;

    var cutoff = new Date();
    cutoff.setHours(0, 0, 0, 0);
    cutoff.setDate(cutoff.getDate() - RETAIN_DAYS);

    var dcol = 0; rows[0].forEach(function (h, i) { if (String(h).trim() === "Date") dcol = i; });
    var keep = [], move = [];
    for (var i = 1; i < rows.length; i++) {
      var d = asDate(rows[i][dcol]);
      (d && d < cutoff ? move : keep).push(rows[i]);
    }
    if (!move.length) return;

    var arch = sheet(name + ARCHIVE_SUFFIX, name === SHIFTS_SHEET ? SHIFT_COLS : ENTRY_COLS);
    arch.getRange(arch.getLastRow() + 1, 1, move.length, move[0].length).setValues(move);

    sh.getRange(2, 1, rows.length - 1, rows[0].length).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
  });
}

/** Run this ONCE from the editor to schedule the nightly cleanup. */
function setupDailyPurge() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "purgeOld") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("purgeOld").timeBased().everyDays(1).atHour(3).create();
}

/* ------------------------------------------------------------------ */
/* Closing the loop — one digest a night, and a Handled mark            */
/* ------------------------------------------------------------------ */

/** Stable id so a resent shift keeps whatever the office already handled. */
function entryId(p, i, en) {
  var s = [p.store || "", p.date, p.shiftId || p.shift || "", i,
           en.department || "", en.type || "", en.text || ""].join("|");
  var h = 0;
  for (var k = 0; k < s.length; k++) { h = ((h << 5) - h + s.charCodeAt(k)) | 0; }
  return "e" + Math.abs(h).toString(36);
}

function handledMap(p) {
  var out = {};
  readAll(ENTRIES_SHEET).forEach(function (r) {
    if (!sameDay(r.Date, p.date)) return;
    if (String(r.Store) !== String(p.store || "")) return;
    if (String(r.Shift) !== String(p.shift || "")) return;
    if (r.Handled) out[String(r["Entry ID"])] = r["Handled at"] || "Yes";
  });
  return out;
}

/** Office ticks something off. Gated by the office passcode, not the store code. */
function markHandled(pass, ids, on) {
  if (String(pass || "") !== OFFICE_PASSCODE) return { ok: false, error: "Wrong passcode" };
  var sh = sheet(ENTRIES_SHEET, ENTRY_COLS);
  if (sh.getLastRow() < 2) return { ok: true, changed: 0 };

  var c = colIndex(sh);
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var want = {}; (ids || []).forEach(function (id) { want[String(id)] = 1; });
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
  var n = 0;

  rows.forEach(function (r, i) {
    if (!want[String(r[c["Entry ID"]])]) return;
    sh.getRange(i + 2, c["Handled"] + 1).setValue(on ? "Yes" : "");
    sh.getRange(i + 2, c["Handled at"] + 1).setValue(on ? stamp : "");
    n++;
  });
  return { ok: true, changed: n };
}

function digestTo(route) {
  var a = String((ROUTE_EMAIL || {})[route] || "").trim();
  if (a) return a;
  return String(OWNER_EMAIL || "").trim() || Session.getEffectiveUser().getEmail();
}

/** Runs nightly. One email per inbox, covering every store, still-open items only. */
function sendDigest() {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var rows = readAll(ENTRIES_SHEET).filter(function (r) {
    return sameDay(r.Date, today) && !r.Handled && String(r["Route to"] || "").trim();
  });
  var shifts = readAll(SHIFTS_SHEET).filter(function (r) { return sameDay(r.Date, today); });

  /* group first by inbox, then by route, then by store */
  var byInbox = {};
  rows.forEach(function (r) {
    var route = String(r["Route to"]).trim();
    var to = digestTo(route);
    byInbox[to] = byInbox[to] || {};
    byInbox[to][route] = byInbox[to][route] || {};
    var st = String(r.Store || "no store");
    (byInbox[to][route][st] = byInbox[to][route][st] || []).push(r);
  });

  var missing = shiftsMissing(shifts);
  var owner = String(OWNER_EMAIL || "").trim() || Session.getEffectiveUser().getEmail();
  if (missing.length && !byInbox[owner]) byInbox[owner] = {};

  Object.keys(byInbox).forEach(function (to) {
    var L = [], count = 0;
    Object.keys(byInbox[to]).sort().forEach(function (route) {
      L.push(route.toUpperCase());
      L.push(new Array(route.length + 1).join("="));
      Object.keys(byInbox[to][route]).sort().forEach(function (store) {
        L.push("");
        L.push("  " + store);
        byInbox[to][route][store].forEach(function (r) {
          count++;
          L.push("    [" + r.Type + "] " + r.Department + " — " + r.Detail);
          L.push("        " + r.Shift + ", " + (r.Manager || "no name"));
        });
      });
      L.push("");
      L.push("");
    });

    if (to === owner && missing.length) {
      L.push("SHIFTS THAT NEVER CAME IN");
      L.push("=========================");
      missing.forEach(function (m) { L.push("  " + m); });
      L.push("");
    }

    if (!L.length) return;
    L.push("Mark things handled on the office dashboard so they drop off tomorrow's email.");

    MailApp.sendEmail(to,
      "Manager log — " + today + (count ? " — " + count + " open item" + (count === 1 ? "" : "s") : ""),
      L.join("\n"));
  });
}

/** Which store/shift combinations never arrived today. */
function shiftsMissing(shifts) {
  var expect = ["Opening Manager", "2nd Shift Manager", "Closing Manager"];
  var stores = {};
  readAll(SHIFTS_SHEET).forEach(function (r) { if (r.Store) stores[r.Store] = 1; });
  var seen = {};
  shifts.forEach(function (r) { seen[r.Store + "||" + r.Shift] = 1; });
  var out = [];
  Object.keys(stores).sort().forEach(function (st) {
    expect.forEach(function (sh) {
      if (!seen[st + "||" + sh]) out.push(st + " — " + sh);
    });
  });
  return out;
}

/** Run this ONCE from the editor to schedule the nightly digest. */
function setupDigest() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "sendDigest") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("sendDigest").timeBased().everyDays(1).atHour(DIGEST_HOUR).create();
}

/* ------------------------------------------------------------------ */
/* The roster managers pick their name from                             */
/* ------------------------------------------------------------------ */
function readManagers(store) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(MANAGERS_SHEET);
  if (!sh) { sheet(MANAGERS_SHEET, MANAGER_COLS); return []; }
  if (sh.getLastRow() < 2) return [];

  var rows = sh.getDataRange().getValues();
  var head = rows.shift(), col = {};
  head.forEach(function (h, i) { col[String(h).trim()] = i; });

  var out = [];
  rows.forEach(function (r) {
    var name = String(r[col["Name"]] || "").trim();
    if (!name) return;
    var active = String(r[col["Active"]] === undefined ? "" : r[col["Active"]]).trim().toLowerCase();
    if (active === "false" || active === "no" || active === "0") return;
    var forStore = String(r[col["Store"]] || "").trim();
    if (forStore && store && forStore !== store) return;
    if (out.indexOf(name) < 0) out.push(name);
  });
  return out.slice(0, 40);
}

function setupManagerSheet() { sheet(MANAGERS_SHEET, MANAGER_COLS); }

/** Which code this store expects. Falls back to the one shared code. */
function codeOk(code, store) {
  var want = String((STORE_CODES || {})[String(store || "").trim()] || "").trim();
  if (!want) {
    /* store names in the app may be bare ("Alameda") or full ("Vista Market — Alameda") */
    Object.keys(STORE_CODES || {}).forEach(function (k) {
      if (!want && k && String(store || "").indexOf(k) >= 0) want = String(STORE_CODES[k] || "").trim();
    });
  }
  if (!want) want = STORE_CODE;
  return String(code || "") === want;
}

/* ------------------------------------------------------------------ */
/* helpers                                                             */
/* ------------------------------------------------------------------ */
function sheet(name, cols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(cols);
    sh.getRange(1, 1, 1, cols.length).setFontWeight("bold").setBackground("#e8f2ec");
    sh.setFrozenRows(1);
  }
  return sh;
}

function colIndex(sh) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], m = {};
  head.forEach(function (h, i) { m[String(h).trim()] = i; });
  return m;
}

function removeExisting(sh, p) {
  var rows = sh.getDataRange().getValues();
  if (rows.length < 2) return;
  var c = {}; rows[0].forEach(function (h, i) { c[String(h).trim()] = i; });
  for (var i = rows.length - 1; i >= 1; i--) {
    if (sameDay(rows[i][c["Date"]], p.date) &&
        String(rows[i][c["Store"]]) === String(p.store || "") &&
        String(rows[i][c["Shift"]]) === String(p.shift || "")) {
      sh.deleteRow(i + 1);
    }
  }
}

function sameDay(cell, iso) {
  if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd") === iso;
  return String(cell) === String(iso);
}

function asDate(v) {
  if (v instanceof Date) return v;
  var s = String(v || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  var p = s.split("-");
  return new Date(+p[0], +p[1] - 1, +p[2]);
}

function statusWord(s) {
  return { good: "All good", attn: "Needs attention", skip: "Not covered" }[s] || "";
}

function readAll(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getDataRange().getValues();
  var head = rows.shift();
  var tz = Session.getScriptTimeZone();
  return rows.map(function (r) {
    var o = {};
    head.forEach(function (h, i) {
      var v = r[i];
      o[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, "yyyy-MM-dd HH:mm") : v;
    });
    return o;
  });
}

function taskId(store, shift, task) {
  var s = String(store || "") + "|" + String(shift || "") + "|" + String(task || "");
  var h = 0;
  for (var i = 0; i < s.length; i++) { h = ((h << 5) - h + s.charCodeAt(i)) | 0; }
  return "t" + Math.abs(h).toString(36);
}

function readTasks(store) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASKS_SHEET);
  if (!sh) { sheet(TASKS_SHEET, TASKS_COLS); return []; }
  if (sh.getLastRow() < 2) return [];

  var rows = sh.getDataRange().getValues();
  var head = rows.shift();
  var col  = {};
  head.forEach(function (h, i) { col[String(h).trim()] = i; });

  var out = [];
  rows.forEach(function (r) {
    var task = String(r[col["Task"]] || "").trim();
    if (!task) return;

    var active = String(r[col["Active"]] === undefined ? "" : r[col["Active"]]).trim().toLowerCase();
    if (active === "false" || active === "no" || active === "0") return;

    var forStore = String(r[col["Store"]] || "").trim();
    if (forStore && store && forStore !== store) return;

    var forShift = String(r[col["Shift"]] || "").trim();
    out.push({ id: taskId(forStore, forShift, task), task: task, shift: forShift, store: forStore });
  });

  return out.slice(0, 12);
}

/** Run once from the editor to create the Priority Tasks tab before your first entry. */
function setupTaskSheet() { sheet(TASKS_SHEET, TASKS_COLS); }

function readBulletins(store) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BULLETIN_SHEET);
  if (!sh) { sheet(BULLETIN_SHEET, BULLETIN_COLS); return []; }   // create it so office can find it
  if (sh.getLastRow() < 2) return [];

  var rows = sh.getDataRange().getValues();
  var head = rows.shift();
  var col  = {};
  head.forEach(function (h, i) { col[String(h).trim()] = i; });

  var today = new Date(); today.setHours(0, 0, 0, 0);
  var out = [];

  rows.forEach(function (r) {
    var headline = String(r[col["Headline"]] || "").trim();
    var message  = String(r[col["Message"]]  || "").trim();
    if (!headline && !message) return;

    var active = String(r[col["Active"]] === undefined ? "" : r[col["Active"]]).trim().toLowerCase();
    if (active === "false" || active === "no" || active === "0") return;

    var forStore = String(r[col["Store"]] || "").trim();
    if (forStore && store && forStore !== store) return;

    var starts = asDate(r[col["Starts"]]), ends = asDate(r[col["Ends"]]);
    if (starts && starts > today) return;
    if (ends   && ends   < today) return;

    var posted = r[col["Posted"]];
    out.push({
      headline: headline,
      message:  message,
      store:    forStore,
      posted:   posted instanceof Date
                  ? Utilities.formatDate(posted, Session.getScriptTimeZone(), "yyyy-MM-dd")
                  : String(posted || "")
    });
  });

  return out.reverse().slice(0, 3);   // newest rows first, at most three
}

/** Run once from the editor if you want the Bulletin tab created before your first post. */
function setupBulletinSheet() {
  sheet(BULLETIN_SHEET, BULLETIN_COLS);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BULLETIN_SHEET)
    .getRange("A2:G2").setValues([[new Date(), "", "Example — delete this row",
      "Leave Store blank to reach every store. Set Active to FALSE to take a bulletin down.", "", "", "FALSE"]]);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Used by the dashboard when it is served from this same Apps Script project. */
function getData(pass) {
  if (String(pass || "") !== OFFICE_PASSCODE) return { ok: false, error: "Wrong passcode" };
  return { ok: true, days: RETAIN_DAYS, shifts: readAll(SHIFTS_SHEET), entries: readAll(ENTRIES_SHEET) };
}
