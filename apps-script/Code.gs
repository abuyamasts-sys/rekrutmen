/**
 * Airtis Assessment -> Google Sheets collector
 *
 * Deploy:
 * - Extensions -> Apps Script (or script.google.com)
 * - Paste this file as Code.gs
 * - Set SPREADSHEET_ID below
 * - Deploy -> New deployment -> Web app
 *   - Execute as: Me
 *   - Who has access: Anyone (or Anyone with the link)
 * - Copy the Web app URL into `sheet-config.js` as `window.AIRTIS_SHEETS_ENDPOINT`
 */

const SPREADSHEET_ID = "1Y6ataIODMYO-GEDs33XKfrzV85ougtfiu0cplup5aBg";
const SHEET_NAME = "assessments";

// Optional shared secret: set a value here and in `sheet-config.js`
const SHARED_SECRET = "";

function doPost(e) {
  try {
    const bodyText = e && e.postData && e.postData.contents ? e.postData.contents : "";
    const body = bodyText ? JSON.parse(bodyText) : {};

    if (SHARED_SECRET && String(body.secret || "") !== String(SHARED_SECRET)) {
      return json({ ok: false, error: "unauthorized" }, 401);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    // Header (create once)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "received_at",
        "kind",
        "token",
        "name",
        "phone",
        "position",
        "payload_json"
      ]);
    }

    const kind = String(body.kind || "");
    const payload = body.payload || {};
    const token = (payload.session && payload.session.token) ? String(payload.session.token) : (payload.seed ? String(payload.seed) : "");
    const name = payload.session && payload.session.candidate && payload.session.candidate.name ? String(payload.session.candidate.name) : (payload.participant ? String(payload.participant) : "");
    const phone = payload.session && payload.session.candidate && payload.session.candidate.phone ? String(payload.session.candidate.phone) : "";
    const position = payload.session && payload.session.candidate && payload.session.candidate.position ? String(payload.session.candidate.position) : "";

    sheet.appendRow([
      new Date(),
      kind,
      token,
      name,
      phone,
      position,
      JSON.stringify(body)
    ]);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) }, 500);
  }
}

function json(obj, status) {
  const out = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  // Status code is not configurable in simple ContentService responses; included in body for debugging
  obj = obj || {};
  obj._status = status || 200;
  return out;
}
