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

// Readable report sheets (optional; will be auto-created)
const REPORT_CBT_SHEET = "REPORT_CBT";
const REPORT_PAULI_SHEET = "REPORT_PAULI";
const REPORT_DISC_SHEET = "REPORT_DISC";

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

    // Best-effort: write to readable report sheet(s)
    try {
      writeReportRow_(ss, kind, token, name, phone, position, body);
    } catch (err) {
      // Ignore reporting failures so raw collection still works
    }

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

function writeReportRow_(ss, kind, token, name, phone, position, body) {
  const k = String(kind || "").trim();
  if (!k) return;

  if (k === "cbt") return writeCbtReportRow_(ss, token, name, phone, position, body);
  if (k === "pauli") return writePauliReportRow_(ss, token, name, phone, position, body);
  if (k === "disc") return writeDiscReportRow_(ss, token, name, phone, position, body);
}

function ensureHeader_(sheet, headers) {
  if (sheet.getLastRow() > 0) return;
  sheet.appendRow(headers);
}

function writeCbtReportRow_(ss, token, name, phone, position, body) {
  const sheet = ss.getSheetByName(REPORT_CBT_SHEET) || ss.insertSheet(REPORT_CBT_SHEET);
  ensureHeader_(sheet, [
    "received_at",
    "token",
    "name",
    "phone",
    "position",
    "finished_at",
    "answered",
    "total",
    "correct",
    "iq",
    "category",
    "recommendation"
  ]);

  const payload = body && body.payload ? body.payload : {};
  const meta = payload.meta || {};
  const session = payload.session || {};
  const candidate = session.candidate || {};
  const summary = payload.summary || {};

  sheet.appendRow([
    new Date(),
    token || (session.token || ""),
    name || (candidate.name || ""),
    phone || (candidate.phone || ""),
    position || (candidate.position || ""),
    meta.finishedAt || "",
    summary.answered != null ? summary.answered : "",
    summary.total != null ? summary.total : "",
    summary.correct != null ? summary.correct : "",
    summary.iq != null ? summary.iq : "",
    summary.category || "",
    summary.recommendation || ""
  ]);
}

function writePauliReportRow_(ss, token, name, phone, position, body) {
  const sheet = ss.getSheetByName(REPORT_PAULI_SHEET) || ss.insertSheet(REPORT_PAULI_SHEET);
  ensureHeader_(sheet, [
    "received_at",
    "token",
    "name",
    "position",
    "started_at",
    "finished_at",
    "reason",
    "practice",
    "total_seconds",
    "column_seconds",
    "correct",
    "wrong",
    "total_attempts",
    "last_column"
  ]);

  const payload = body && body.payload ? body.payload : {};
  const meta = payload.meta || {};
  const scoring = payload.scoring || {};
  const attempts = Array.isArray(payload.attempts) ? payload.attempts : [];
  const practice = attempts.length ? (attempts[0] && attempts[0].practice ? 1 : 0) : "";

  sheet.appendRow([
    new Date(),
    token || (payload.seed || ""),
    name || (payload.participant || ""),
    position || "",
    meta.startedAt || "",
    meta.finishedAt || "",
    meta.reason || "",
    practice,
    meta.totalSeconds != null ? meta.totalSeconds : "",
    meta.columnSeconds != null ? meta.columnSeconds : "",
    scoring.correct != null ? scoring.correct : "",
    scoring.wrong != null ? scoring.wrong : "",
    scoring.totalAttempts != null ? scoring.totalAttempts : "",
    scoring.lastColumn != null ? scoring.lastColumn : ""
  ]);
}

function writeDiscReportRow_(ss, token, name, phone, position, body) {
  const sheet = ss.getSheetByName(REPORT_DISC_SHEET) || ss.insertSheet(REPORT_DISC_SHEET);
  ensureHeader_(sheet, [
    "received_at",
    "token",
    "name",
    "phone",
    "position",
    "started_at",
    "finished_at",
    "dominant",
    "secondary",
    "D",
    "I",
    "S",
    "C"
  ]);

  const payload = body && body.payload ? body.payload : {};
  const meta = payload.meta || {};
  const session = payload.session || {};
  const candidate = session.candidate || {};
  const scores = payload.scores || {};

  sheet.appendRow([
    new Date(),
    token || (session.token || ""),
    name || (candidate.name || ""),
    phone || (candidate.phone || ""),
    position || (candidate.position || ""),
    meta.startedAt || "",
    meta.finishedAt || "",
    payload.dominant || "",
    payload.secondary || "",
    scores.D != null ? scores.D : "",
    scores.I != null ? scores.I : "",
    scores.S != null ? scores.S : "",
    scores.C != null ? scores.C : ""
  ]);
}

// Run manually in Apps Script editor to rebuild REPORT_* sheets from raw `assessments` sheet.
function rebuildReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Missing sheet: " + SHEET_NAME);

  // Clear report sheets
  [REPORT_CBT_SHEET, REPORT_PAULI_SHEET, REPORT_DISC_SHEET].forEach((name) => {
    const s = ss.getSheetByName(name);
    if (s) s.clear();
  });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Columns: A received_at, B kind, C token, D name, E phone, F position, G payload_json
  const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  values.forEach((row) => {
    const kind = row[1];
    const token = row[2];
    const name = row[3];
    const phone = row[4];
    const position = row[5];
    const payloadJson = row[6];
    if (!payloadJson) return;
    let body;
    try { body = JSON.parse(String(payloadJson)); } catch { body = null; }
    if (!body) return;
    try { writeReportRow_(ss, kind, token, name, phone, position, body); } catch {}
  });
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
