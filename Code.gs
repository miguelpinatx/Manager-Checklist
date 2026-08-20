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
/* ====================================== */

var SHIFTS_SHEET   = "Shifts";
var BULLETIN_SHEET = "Vista Updates";
var ENTRIES_SHEET = "Entries";
var ARCHIVE_SUFFIX = " (archive)";

var SHIFT_COLS = ["Received","Date","Store","Shift","Manager","Day type",
                  "Departments","Logged","All good","Needs attention","Not covered",
                  "Issues","Notes","Full report"];
var ENTRY_COLS = ["Received","Date","Store","Shift","Manager","Time block",
                  "Department","Dept status","Type","Route to","Detail"];

/* The "Vista Updates" tab is where main office types announcements for managers.
   Store   — leave blank for every store, or type one store's name
   Starts / Ends — leave blank for "show it now, until I turn it off"
   Active  — leave blank or type TRUE to show it; FALSE hides it          */
var BULLETIN_COLS = ["Posted","Store","Headline","Message","Starts","Ends","Active"];

/* ------------------------------------------------------------------ */
/* Receiving a submitted shift                                         */
/* ------------------------------------------------------------------ */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    if (!p || !p.date || !p.shiftId) throw new Error("Missing date or shift");
    if (String(p.code || "") !== STORE_CODE) return json({ ok: false, error: "bad-code" });

    var now = new Date();
    var c   = p.counts || {};

    // A resend of the same store+date+shift replaces the earlier one.
    removeExisting(sheet(SHIFTS_SHEET,  SHIFT_COLS), p, [2,3,4]);
    removeExisting(sheet(ENTRIES_SHEET, ENTRY_COLS), p, [2,3,4]);

    sheet(SHIFTS_SHEET, SHIFT_COLS).appendRow([
      now, p.date, p.store || "", p.shift || "", p.manager || "", p.dayType || "",
      c.total || 0, c.reported || 0, c.good || 0, c.attn || 0, c.skip || 0,
      c.issues || 0, c.entries || 0, p.report || ""
    ]);

    var es = sheet(ENTRIES_SHEET, ENTRY_COLS);
    (p.entries || []).forEach(function (en) {
      es.appendRow([
        now, p.date, p.store || "", p.shift || "", p.manager || "",
        en.block || "", en.department || "", statusWord(en.status),
        en.type || "", en.route || "", en.text || ""
      ]);
    });

    return json({ ok: true, saved: (p.entries || []).length });
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

  // the bulletin managers see at the top of their shift
  if (p.action === "bulletin") {
    if (String(p.code || "") !== STORE_CODE) return json({ ok: false, error: "bad-code" });
    return json({ ok: true, bulletins: readBulletins(String(p.store || "")) });
  }

  // managers reading the last few days for their own store, from their phone
  if (p.action === "recent") {
    if (String(p.code || "") !== STORE_CODE) return json({ ok: false, error: "bad-code" });
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

  if (p.action === "data") {
    if (String(p.pass || "") !== OFFICE_PASSCODE) return json({ ok: false, error: "Wrong passcode" });
    return json({ ok: true, days: RETAIN_DAYS, shifts: readAll(SHIFTS_SHEET), entries: readAll(ENTRIES_SHEET) });
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

    var keep = [], move = [];
    for (var i = 1; i < rows.length; i++) {
      var d = asDate(rows[i][1]);
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

function removeExisting(sh, p, cols) {
  var rows = sh.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (sameDay(rows[i][1], p.date) &&
        String(rows[i][2]) === String(p.store || "") &&
        String(rows[i][3]) === String(p.shift || "")) {
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
