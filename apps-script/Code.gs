/**
 * ============================================================
 *  Opex Hub backend (Google Apps Script)
 *  ------------------------------------
 *  Bound to a Google Sheet. Receives submissions from the hub
 *  (login/user management) and the tool forms (5S Audit, OEE
 *  Entry), saves them into the sheet, and emails the admin.
 *
 *  SETUP (one time):
 *    1. Create a Google Sheet (or use the one this is bound to).
 *    2. Extensions -> Apps Script -> paste this file -> Save.
 *    3. Run the function  setup  once (from the toolbar).
 *    4. Deploy -> New deployment -> Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *    5. Copy the ".../exec" URL and paste it into:
 *         - Index.html  ->  var BACKEND_URL = "..."
 *         - 5s-audit-form.html  ->  const SHEET_ENDPOINT = "..."
 *         - oee-entry.html      ->  const SHEET_ENDPOINT = "..."
 * ============================================================
 */

var ADMIN_EMAIL = "alimdmortuza28@gmail.com";   // <-- all notifications go here
var SPREADSHEET_ID = "1-mPjdHWQlRZhhMHTiz9Gk-EIA1i-bCXa8Tn8WhgD73Q";   // your Google Sheet ID

var SHEET_USERS    = "Users";
var SHEET_REQUESTS = "Requests";
var SHEET_5S       = "5S Form";
var SHEET_OEE      = "OEE Entry";

function getSS() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  var ss = getSS();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function setup() {
  getSheet(SHEET_USERS).getRange(1, 1, 1, 6)
    .setValues([["email", "name", "role", "section", "hash", "perms"]]);
  getSheet(SHEET_REQUESTS).getRange(1, 1, 1, 5)
    .setValues([["type", "name", "email", "section", "when"]]);
  seedAdmin();
  // 5S Form / OEE Entry sheets get their headers automatically on first submit.
}

// Ensure the admin account exists in the sheet (needed for cross-device login).
function seedAdmin() {
  var sh = getSheet(SHEET_USERS);
  var email = "alimdmortuza28@gmail.com";

  // Ensure the header row exists.
  var lastCol = sh.getLastColumn();
  var hasHeader = lastCol >= 1 && sh.getRange(1, 1).getValue() === "email";
  if (!hasHeader) {
    sh.getRange(1, 1, 1, 6).setValues([["email", "name", "role", "section", "hash", "perms"]]);
  }

  var lastRow = sh.getLastRow();
  var exists = false;
  if (lastRow >= 2) {
    var emails = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase() === email) { exists = true; break; }
    }
  }
  if (!exists) {
    sh.appendRow([
      email, "Alim Mortuza", "admin", "",
      "e86f78a8a3caf0b60d8e74e5942aa6d86dc150cd3c03338aef25b7d2d7e3acc7",
      "[\"bulletin\",\"fives\",\"oee\",\"plan\",\"dashboard\",\"forms\"]"
    ]);
  }
}

/* ------------------------- HTTP entry points ------------------------- */

function doGet(e) {
  seedAdmin();
  var p = (e && e.parameter) ? e.parameter : {};
  var cb = p.callback || null;
  var result = { ok: true };

  if (p.action === "list") {
    result = { ok: true, users: readUsers(), requests: readRequests() };
  } else if (p.action === "5sStats") {
    var stats = fiveSStats(p.floor || "");
    result = {
      ok: true,
      previousAudit: stats.previousAudit,
      monthAverage: stats.monthAverage,
      monthCount: stats.monthCount
    };
  }

  var out = JSON.stringify(result);
  if (cb) out = cb + "(" + out + ")";   // JSONP
  return ContentService.createTextOutput(out)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Previous audit + this month's average for a floor (used by the 5S form).
function fiveSStats(floor) {
  var sh = getSheet(SHEET_5S);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { previousAudit: null, monthAverage: null, monthCount: 0 };

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  var idxFloor = headers.indexOf("Floor / Area");
  var idxDate = headers.indexOf("Audit Date");
  var idxScore = headers.indexOf("Score %");

  var rows = data.filter(function (r) {
    return String(r[idxFloor] || "") === floor;
  });

  var previousAudit = null;
  if (rows.length && idxScore >= 0 && idxDate >= 0) {
    var prev = rows[rows.length - 1];
    previousAudit = { scorePercent: prev[idxScore], date: prev[idxDate] };
  }

  var now = new Date();
  var ym = now.getFullYear() + "-" + ("0" + (now.getMonth() + 1)).slice(-2);
  var monthRows = rows.filter(function (r) {
    return String(r[idxDate] || "").slice(0, 7) === ym;
  });
  var monthAverage = null, monthCount = 0;
  if (monthRows.length) {
    var sum = 0;
    monthRows.forEach(function (r) { sum += Number(r[idxScore] || 0); });
    monthAverage = Math.round(sum / monthRows.length);
    monthCount = monthRows.length;
  }

  return { previousAudit: previousAudit, monthAverage: monthAverage, monthCount: monthCount };
}

function doPost(e) {
  seedAdmin();
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: "invalid json" });
  }

  // Tool-form submissions (carry a "sheet" field)
  if (data.sheet === "5S Form" || data.sheet === "OEE Entry") {
    return json(handleForm(data));
  }

  // Hub user-management calls (carry an "action" field)
  if (data.action) {
    return json(handleAction(data));
  }

  return json({ ok: false, error: "no action" });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------- Form handling ------------------------- */

function handleForm(data) {
  var sheetName = data.sheet;
  appendRowByHeader(sheetName, data);

  var subject = "[Opex Hub] New " + sheetName + " submission";
  var body = formatBody(data);
  try {
    MailApp.sendEmail(ADMIN_EMAIL, subject, body);
  } catch (err) {
    Logger.log("Email failed: " + err);
  }
  return { ok: true };
}

function formatBody(obj) {
  var lines = [];
  Object.keys(obj).forEach(function (k) {
    lines.push(k + ": " + obj[k]);
  });
  return lines.join("\n");
}

// Append an object as a row, creating/expanding headers as needed.
function appendRowByHeader(sheetName, obj) {
  var sh = getSheet(sheetName);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();

  var headers = [];
  if (lastRow >= 1 && lastCol >= 1) {
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
      .filter(function (h) { return h !== ""; });
  }

  var keys = Object.keys(obj);
  var newKeys = keys.filter(function (k) { return headers.indexOf(k) === -1; });

  if (newKeys.length) {
    if (headers.length === 0) {
      sh.getRange(1, 1, 1, newKeys.length).setValues([newKeys]);
    } else {
      sh.getRange(1, headers.length + 1, 1, newKeys.length).setValues([newKeys]);
    }
    headers = headers.concat(newKeys);
  }

  var row = headers.map(function (h) { return obj.hasOwnProperty(h) ? obj[h] : ""; });
  sh.appendRow(row);
}

/* ------------------------- User management ------------------------- */

function handleAction(data) {
  var action = data.action;

  if (action === "list") {
    return { ok: true, users: readUsers(), requests: readRequests() };
  }
  if (action === "upsert") {
    upsertUser(data);
    return { ok: true };
  }
  if (action === "request") {
    getSheet(SHEET_REQUESTS).appendRow([
      data.type || "", data.name || "", data.email || "",
      data.section || "", new Date().toString()
    ]);
    return { ok: true };
  }
  if (action === "removeRequest") {
    removeRowsByEmail(getSheet(SHEET_REQUESTS), 2, data.email);
    return { ok: true };
  }
  if (action === "resetpass") {
    updateCellByEmail(getSheet(SHEET_USERS), 4, data.hash, data.email);
    return { ok: true };
  }
  if (action === "deluser") {
    removeRowsByEmail(getSheet(SHEET_USERS), 0, data.email);
    return { ok: true };
  }
  return { ok: false, error: "unknown action" };
}

function readUsers() {
  var sh = getSheet(SHEET_USERS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var vals = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (!r[0]) return;
    var perms = [];
    try { perms = JSON.parse(r[5] || "[]"); } catch (e) { perms = []; }
    out.push({ id: r[0], name: r[1], role: r[2], section: r[3], hash: r[4], perms: perms });
  });
  return out;
}

function readRequests() {
  var sh = getSheet(SHEET_REQUESTS);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var vals = sh.getRange(2, 1, lastRow - 1, 5).getValues();
  var out = [];
  vals.forEach(function (r) {
    if (!r[0]) return;
    out.push({ type: r[0], name: r[1], email: r[2], section: r[3], when: r[4] });
  });
  return out;
}

function upsertUser(data) {
  var sh = getSheet(SHEET_USERS);
  var email = String(data.email || "").toLowerCase();
  var lastRow = sh.getLastRow();
  var foundRow = -1;

  if (lastRow >= 2) {
    var emails = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase() === email) { foundRow = i + 2; break; }
    }
  }

  var name = data.name || email;
  var role = data.role || "viewer";
  var section = data.section || "";
  var perms = JSON.stringify(data.perms || []);
  var hash = data.hash || "";

  if (foundRow === -1) {
    sh.appendRow([email, name, role, section, hash, perms]);
  } else {
    if (!hash) hash = sh.getRange(foundRow, 5).getValue(); // keep existing hash
    sh.getRange(foundRow, 1, 1, 6).setValues([[email, name, role, section, hash, perms]]);
  }
}

function removeRowsByEmail(sh, colIndex, email) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var target = String(email || "").toLowerCase();
  var rows = [];
  for (var r = 2; r <= lastRow; r++) {
    if (String(sh.getRange(r, colIndex + 1).getValue()).toLowerCase() === target) rows.push(r);
  }
  for (var i = rows.length - 1; i >= 0; i--) {
    sh.deleteRow(rows[i]);
  }
}

function updateCellByEmail(sh, colIndex, value, email) {
  var lastRow = sh.getLastRow();
  var target = String(email || "").toLowerCase();
  for (var r = 2; r <= lastRow; r++) {
    if (String(sh.getRange(r, 1).getValue()).toLowerCase() === target) {
      sh.getRange(r, colIndex + 1).setValue(value);
    }
  }
}
