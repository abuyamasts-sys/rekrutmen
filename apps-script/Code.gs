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
const HRD_SHEET_NAME = "HRD_TOKENS";

// Optional shared secret: set a value here and in `sheet-config.js`
const SHARED_SECRET = "";

// JSONP GET for cross-origin browser use (GitHub Pages)
// Example:
//   ?kind=hrd_token&position=SALES&name=...&phone=...&cb=callbackFn
function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    const kind = String(p.kind || "");
    const cb = String(p.cb || "");
    const secret = String(p.secret || "");

    if (SHARED_SECRET && secret !== String(SHARED_SECRET)) {
      return jsonp_(cb, { ok: false, error: "unauthorized" });
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (kind === "hrd_token") {
      const out = createHrdToken_(ss, {
        position: p.position,
        name: p.name,
        phone: p.phone
      });
      return jsonp_(cb, { ok: true, ...out });
    }

    if (kind === "validate_token") {
      const token = (p.token ? String(p.token) : "").trim();
      const out = validateHrdToken_(ss, token);
      return jsonp_(cb, { ok: true, ...out });
    }

    return jsonp_(cb, { ok: false, error: "unsupported_kind" });
  } catch (err) {
    return jsonp_("", { ok: false, error: String(err && err.message ? err.message : err) });
  }
}

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

    // HRD token generator (auto increment)
    if (kind === "hrd_token") {
      const out = createHrdToken_(ss, payload);
      return json({ ok: true, ...out }, 200);
    }

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

function createHrdToken_(ss, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = ss.getSheetByName(HRD_SHEET_NAME) || ss.insertSheet(HRD_SHEET_NAME);

    // Header (create once)
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "created_at",
        "date",
        "position",
        "seq",
        "token",
        "name",
        "phone"
      ]);
    }

    const tz = Session.getScriptTimeZone() || "Asia/Jakarta";
    const today = Utilities.formatDate(new Date(), tz, "yyyyMMdd");

    const position = (payload.position ? String(payload.position) : "").trim().toUpperCase() || "GENERAL";
    const name = (payload.name ? String(payload.name) : "").trim();
    const phone = (payload.phone ? String(payload.phone) : "").trim();

    const props = PropertiesService.getScriptProperties();
    const counterKey = "SEQ_" + today + "_" + position;
    const seq = Number(props.getProperty(counterKey) || "0") + 1;
    props.setProperty(counterKey, String(seq));

    const token = "ATS-" + today + "-" + position + "-" + ("000" + seq).slice(-3);

    sheet.appendRow([
      new Date(),
      today,
      position,
      seq,
      token,
      name,
      phone
    ]);

    return { token, seq, date: today, position };
  } finally {
    lock.releaseLock();
  }
}

function validateHrdToken_(ss, token) {
  const normalized = (token ? String(token) : "").trim();
  if (!normalized) return { valid: false, reason: "missing_token" };

  const sheet = ss.getSheetByName(HRD_SHEET_NAME);
  if (!sheet) return { valid: false, reason: "missing_sheet" };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { valid: false, reason: "empty_sheet" };

  // token column: E (5). Data starts at row 2 (row 1 = header)
  const tokenRange = sheet.getRange(2, 5, lastRow - 1, 1);
  const finder = tokenRange.createTextFinder(normalized).matchEntireCell(true);
  const cell = finder.findNext();
  if (!cell) return { valid: false, reason: "not_found" };

  const row = cell.getRow();
  const values = sheet.getRange(row, 1, 1, 7).getValues()[0] || [];
  return {
    valid: true,
    token: String(values[4] || normalized),
    date: values[1] ? String(values[1]) : "",
    position: values[2] ? String(values[2]) : "",
    seq: values[3] != null ? Number(values[3]) : null,
    name: values[5] ? String(values[5]) : "",
    phone: values[6] ? String(values[6]) : ""
  };
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

function jsonp_(cb, obj) {
  const callback = (cb || "").replace(/[^\w.$]/g, "");
  const safeCb = callback || "callback";
  const text = safeCb + "(" + JSON.stringify(obj || {}) + ");";
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
