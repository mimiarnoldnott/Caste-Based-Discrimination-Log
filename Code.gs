/**
 * Don't Forget Dalits — Google Sheets backend.
 *
 * SETUP:
 * 1. Create a new Google Sheet (or open an existing one to use).
 * 2. Extensions > Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. Set TEAM_SECRET below to a passphrase only your team knows — this
 *    is what separates "anyone can read" from "only the team can edit."
 *    Pick something you can share with teammates over a private channel
 *    (Signal, WhatsApp, etc.), not something guessable.
 * 5. Click Deploy > New deployment.
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" — this is intentional here: reading
 *      (doGet) is meant to be public, and writing (doPost) is protected
 *      separately by TEAM_SECRET below, not by this access setting.
 * 6. Click Deploy, authorize the script when prompted, then copy the
 *    "Web app URL" it gives you — it ends in /exec.
 * 7. Paste that URL into SHEETS_API_URL near the top of the app's HTML
 *    file (index.html), then host that file anywhere.
 *
 * The sheet named "Cases" is created automatically on first use, with
 * headers matching the app's fields. Don't rename or reorder columns
 * once you have real data in them, or add/edit calls will misalign.
 */

// Only requests carrying this exact string are allowed to add, edit, or
// delete cases. Reading the list never requires it. Change this any time
// by editing this line and redeploying (Deploy > Manage deployments >
// edit the existing deployment > Deploy, so the URL stays the same).
var TEAM_SECRET = "change-this-to-your-own-passphrase";

var SHEET_NAME = "Cases";

var HEADERS = [
  "id", "title", "dateOfIncident", "dateOfIncidentBS", "timeOfIncident",
  "state", "district", "locationDetail", "category", "status", "summary",
  "victimName", "victimGender", "victimDob", "victimDobBS", "victimCaste",
  "victimDod", "victimDodBS", "victimAgeAtDeath", "victimOrigin", "victimMoreInfo",
  "perpetrators", "perpMoreInfo",
  "legalCaseName", "legalCourt", "legalVerdict", "legalMoreInfo",
  "incidentMoreInfo",
  "sources", "sourcesMoreInfo",
  "tags",
  "documentedBy", "createdAt", "updatedAt"
];

var JSON_FIELDS = ["perpetrators", "sources", "tags"];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToCase_(row) {
  var obj = {};
  HEADERS.forEach(function (h, i) {
    var v = row[i];
    if (JSON_FIELDS.indexOf(h) !== -1) {
      try { obj[h] = v ? JSON.parse(v) : []; } catch (e) { obj[h] = []; }
    } else {
      obj[h] = (v === undefined || v === null) ? "" : String(v);
    }
  });
  return obj;
}

function caseToRow_(c) {
  return HEADERS.map(function (h) {
    var v = c[h];
    if (JSON_FIELDS.indexOf(h) !== -1) {
      return JSON.stringify(v || []);
    }
    return (v === undefined || v === null) ? "" : v;
  });
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET — list every case. */
function doGet(e) {
  try {
    var sheet = getSheet_();
    var lastRow = sheet.getLastRow();
    var cases = [];
    if (lastRow > 1) {
      var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
      cases = data.map(rowToCase_).filter(function (c) { return c.id; });
    }
    return jsonResponse_({ ok: true, cases: cases });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

/**
 * POST — add, update, or delete a case.
 * Body is JSON: {action: "add", case: {...}}
 *            or {action: "update", case: {...with id...}}
 *            or {action: "delete", id: "..."}
 *
 * Sent as text/plain from the browser to avoid a CORS preflight
 * request, which this endpoint doesn't handle — parse manually
 * regardless of the declared content type.
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;

    if (body.secret !== TEAM_SECRET) {
      return jsonResponse_({ ok: false, error: "unauthorized" });
    }

    var sheet = getSheet_();

    if (action === "add") {
      var c = body.case || {};
      c.id = "case-" + new Date().getTime() + "-" + Math.random().toString(36).slice(2, 8);
      var now = new Date().toISOString();
      c.createdAt = now;
      c.updatedAt = now;
      sheet.appendRow(caseToRow_(c));
      return jsonResponse_({ ok: true, id: c.id });
    }

    if (action === "update") {
      var c2 = body.case || {};
      var rowIndex = findRowById_(sheet, c2.id);
      if (rowIndex === -1) return jsonResponse_({ ok: false, error: "not_found" });
      var existing = rowToCase_(sheet.getRange(rowIndex, 1, 1, HEADERS.length).getValues()[0]);
      c2.createdAt = existing.createdAt || new Date().toISOString();
      c2.updatedAt = new Date().toISOString();
      sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([caseToRow_(c2)]);
      return jsonResponse_({ ok: true });
    }

    if (action === "delete") {
      var rowIndex2 = findRowById_(sheet, body.id);
      if (rowIndex2 === -1) return jsonResponse_({ ok: false, error: "not_found" });
      sheet.deleteRow(rowIndex2);
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: "unknown_action" });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function findRowById_(sheet, id) {
  var idCol = HEADERS.indexOf("id") + 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2;
  }
  return -1;
}
