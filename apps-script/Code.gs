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
const REPORT_ALL_SHEET = "REPORT_ALL";
const HRD_SUMMARY_SHEET = "HRD_SUMMARY";

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
      // Force as text to preserve leading zero(s) in Google Sheets
      phone ? ("'" + phone) : "",
      position,
      JSON.stringify(body)
    ]);

    // Best-effort: write to readable report sheet(s)
    try {
      writeReportRow_(ss, kind, token, name, phone, position, body);
    } catch (err) {
      // Ignore reporting failures so raw collection still works
    }

    // Best-effort: upsert into combined report sheet
    try {
      upsertCombinedReport_(ss, kind, token, name, phone, position, body);
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
      // Force as text to preserve leading zero(s) in Google Sheets
      phone ? ("'" + phone) : ""
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

function combinedHeaders_() {
  return [
    "updated_at",
    "token",
    "name",
    "phone",
    "position",
    // CBT
    "cbt_finished_at",
    "cbt_answered",
    "cbt_total",
    "cbt_correct",
    "cbt_iq",
    "cbt_category",
    "cbt_recommendation",
    // PAULI (non-practice)
    "pauli_started_at",
    "pauli_finished_at",
    "pauli_reason",
    "pauli_total_seconds",
    "pauli_column_seconds",
    "pauli_correct",
    "pauli_wrong",
    "pauli_total_attempts",
    "pauli_last_column",
    // DISC
    "disc_started_at",
    "disc_finished_at",
    "disc_dominant",
    "disc_secondary",
    "disc_D",
    "disc_I",
    "disc_S",
    "disc_C"
  ];
}

function getOrCreateCombinedSheet_(ss) {
  const sheet = ss.getSheetByName(REPORT_ALL_SHEET) || ss.insertSheet(REPORT_ALL_SHEET);
  ensureHeader_(sheet, combinedHeaders_());
  return sheet;
}

function findRowByToken_(sheet, token) {
  const t = (token || "").toString().trim();
  if (!t) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const tokenRange = sheet.getRange(2, 2, lastRow - 1, 1); // column B
  const finder = tokenRange.createTextFinder(t).matchEntireCell(true);
  const cell = finder.findNext();
  return cell ? cell.getRow() : null;
}

function toStr_(v) {
  return v == null ? "" : String(v);
}

function toNumOrBlank_(v) {
  if (v == null || v === "") return "";
  const n = Number(v);
  return Number.isFinite(n) ? n : "";
}

function detectPauliPractice_(payload) {
  const attempts = Array.isArray(payload?.attempts) ? payload.attempts : [];
  if (!attempts.length) return null;
  return attempts[0] && attempts[0].practice ? 1 : 0;
}

function upsertCombinedReport_(ss, kind, token, name, phone, position, body) {
  const k = String(kind || "").trim();
  const t = (token || "").toString().trim();
  if (!k || !t) return;

  const sheet = getOrCreateCombinedSheet_(ss);
  let row = findRowByToken_(sheet, t);
  if (!row) row = sheet.getLastRow() + 1;

  // Read existing row to preserve fields from other kinds
  const existing = row <= sheet.getLastRow()
    ? (sheet.getRange(row, 1, 1, combinedHeaders_().length).getValues()[0] || [])
    : new Array(combinedHeaders_().length).fill("");

  // Base identity fields (prefer latest non-empty)
  const next = existing.slice();
  next[0] = new Date(); // updated_at
  next[1] = t; // token
  if (toStr_(name).trim()) next[2] = toStr_(name).trim();
  if (toStr_(phone).trim()) next[3] = toStr_(phone).trim();
  if (toStr_(position).trim()) next[4] = toStr_(position).trim();

  const payload = body && body.payload ? body.payload : {};

  if (k === "cbt") {
    const meta = payload.meta || {};
    const session = payload.session || {};
    const candidate = session.candidate || {};
    const summary = payload.summary || {};

    if (!next[2] && toStr_(candidate.name).trim()) next[2] = toStr_(candidate.name).trim();
    if (!next[3] && toStr_(candidate.phone).trim()) next[3] = toStr_(candidate.phone).trim();
    if (!next[4] && toStr_(candidate.position).trim()) next[4] = toStr_(candidate.position).trim();

    next[5] = toStr_(meta.finishedAt);
    next[6] = toNumOrBlank_(summary.answered);
    next[7] = toNumOrBlank_(summary.total);
    next[8] = toNumOrBlank_(summary.correct);
    next[9] = toNumOrBlank_(summary.iq);
    next[10] = toStr_(summary.category);
    next[11] = toStr_(summary.recommendation);
  } else if (k === "pauli") {
    // Only store non-practice in combined report
    const isPractice = detectPauliPractice_(payload);
    if (isPractice === 1) {
      // ignore practice rows (keep combined report focused)
    } else {
      const meta = payload.meta || {};
      const scoring = payload.scoring || {};
      if (toStr_(payload.participant).trim() && !next[2]) next[2] = toStr_(payload.participant).trim();

      next[12] = toStr_(meta.startedAt);
      next[13] = toStr_(meta.finishedAt);
      next[14] = toStr_(meta.reason);
      next[15] = toNumOrBlank_(meta.totalSeconds);
      next[16] = toNumOrBlank_(meta.columnSeconds);
      next[17] = toNumOrBlank_(scoring.correct);
      next[18] = toNumOrBlank_(scoring.wrong);
      next[19] = toNumOrBlank_(scoring.totalAttempts);
      next[20] = toNumOrBlank_(scoring.lastColumn);
    }
  } else if (k === "disc") {
    const meta = payload.meta || {};
    const session = payload.session || {};
    const candidate = session.candidate || {};
    const scores = payload.scores || {};

    if (!next[2] && toStr_(candidate.name).trim()) next[2] = toStr_(candidate.name).trim();
    if (!next[3] && toStr_(candidate.phone).trim()) next[3] = toStr_(candidate.phone).trim();
    if (!next[4] && toStr_(candidate.position).trim()) next[4] = toStr_(candidate.position).trim();

    next[21] = toStr_(meta.startedAt);
    next[22] = toStr_(meta.finishedAt);
    next[23] = toStr_(payload.dominant);
    next[24] = toStr_(payload.secondary);
    next[25] = toNumOrBlank_(scores.D);
    next[26] = toNumOrBlank_(scores.I);
    next[27] = toNumOrBlank_(scores.S);
    next[28] = toNumOrBlank_(scores.C);
  }

  sheet.getRange(row, 1, 1, next.length).setValues([next]);
}

// Run manually in Apps Script editor to rebuild REPORT_* sheets from raw `assessments` sheet.
function rebuildReports() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Missing sheet: " + SHEET_NAME);

  // Clear report sheets
  [REPORT_CBT_SHEET, REPORT_PAULI_SHEET, REPORT_DISC_SHEET, REPORT_ALL_SHEET].forEach((name) => {
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
    try { upsertCombinedReport_(ss, kind, token, name, phone, position, body); } catch {}
  });
}

// Build HRD-friendly summary from REPORT_ALL (1 token = 1 row)
function syncHrdSummary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const source = ss.getSheetByName(REPORT_ALL_SHEET);
  if (!source) throw new Error("Missing sheet: " + REPORT_ALL_SHEET);

  const lastRow = source.getLastRow();
  const lastCol = source.getLastColumn();
  if (lastRow < 1 || lastCol < 1) throw new Error("REPORT_ALL is empty");

  const all = source.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = (all[0] || []).map((h) => String(h || "").trim());
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[h] = i; });

  function getCell(row, headerName) {
    const i = idx[headerName];
    return i == null ? "" : row[i];
  }

  const outHeader = [
    "Tanggal Tes",
    "Token",
    "Nama Kandidat",
    "No. HP",
    "Posisi Dilamar",
    "IQ",
    "Kategori IQ",
    "Ringkasan Pauli",
    "Ringkasan DISC",
    "Kesimpulan HR",
    "Rekomendasi Akhir"
  ];

  const rows = [];
  const phoneRichTexts = [];

  for (let r = 1; r < all.length; r++) {
    const row = all[r] || [];
    const token = String(getCell(row, "token") || "").trim();
    if (!token) continue;

    const updatedAt = getCell(row, "updated_at");
    const name = String(getCell(row, "name") || "").trim();
    const phone = String(getCell(row, "phone") || "").trim();
    const position = String(getCell(row, "position") || "").trim();
    const iqRaw = getCell(row, "cbt_iq");
    const iq = iqRaw === "" || iqRaw == null ? null : Number(iqRaw);
    const category = String(getCell(row, "cbt_category") || "").trim();

    const pauliCorrect = getCell(row, "pauli_correct");
    const pauliWrong = getCell(row, "pauli_wrong");
    const pauliTotal = getCell(row, "pauli_total_attempts");
    const pauliSummary = buildPauliSummary_(pauliCorrect, pauliWrong, pauliTotal);

    const discDom = String(getCell(row, "disc_dominant") || "").trim();
    const discSec = String(getCell(row, "disc_secondary") || "").trim();
    const discSummary = buildDiscSummary_(discDom, discSec);

    const cbtRec = String(getCell(row, "cbt_recommendation") || "").trim();
    const kesimpulan = buildKesimpulanHr_(iq, discDom, pauliWrong, position);
    const rekomAkhir = buildRekomendasiAkhir_(iq, cbtRec);

    const dateObj = normalizeToDate_(updatedAt);
    rows.push([
      dateObj || "",
      token,
      name,
      phone || "",
      position,
      iq == null || !Number.isFinite(iq) ? "" : iq,
      category,
      pauliSummary,
      discSummary,
      kesimpulan,
      rekomAkhir
    ]);

    phoneRichTexts.push(phoneToWhatsAppRichText_(phone));
  }

  const sheet = ss.getSheetByName(HRD_SUMMARY_SHEET) || ss.insertSheet(HRD_SUMMARY_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 1, outHeader.length).setValues([outHeader]);
  if (rows.length) sheet.getRange(2, 1, rows.length, outHeader.length).setValues(rows);

  // Phone as WhatsApp hyperlink (RichText; avoids locale formula separator issues)
  if (rows.length) {
    const rich = phoneRichTexts.map((rt) => [rt]);
    sheet.getRange(2, 4, rich.length, 1).setRichTextValues(rich);
  }

  // Formatting
  sheet.setFrozenRows(1);
  const headerRange = sheet.getRange(1, 1, 1, outHeader.length);
  headerRange
    .setFontWeight("bold")
    .setBackground("#0B2A4A")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getDataRange().setVerticalAlignment("middle");

  // Date format dd/mm/yyyy
  if (rows.length) sheet.getRange(2, 1, rows.length, 1).setNumberFormat("dd/MM/yyyy");

  // Wrap for Kesimpulan + Rekomendasi
  sheet.getRange(1, 10, Math.max(rows.length + 1, 1), 2).setWrap(true);

  // Comfortable alignment
  sheet.getRange(2, 6, Math.max(rows.length, 1), 2).setHorizontalAlignment("center"); // IQ + Category
  sheet.getRange(2, 1, Math.max(rows.length, 1), outHeader.length).setHorizontalAlignment("left");
  headerRange.setHorizontalAlignment("center");

  // Auto resize
  try { sheet.autoResizeColumns(1, outHeader.length); } catch {}
}

function normalizeToDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === "number" && Number.isFinite(v)) return new Date(v);
  if (!v) return null;
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function buildPauliSummary_(correct, wrong, total) {
  const c = correct === "" || correct == null ? null : Number(correct);
  const w = wrong === "" || wrong == null ? null : Number(wrong);
  const t = total === "" || total == null ? null : Number(total);
  if (!Number.isFinite(c) && !Number.isFinite(w) && !Number.isFinite(t)) return "";
  return "Benar " + (Number.isFinite(c) ? c : "-") + " | Salah " + (Number.isFinite(w) ? w : "-") + " | Total " + (Number.isFinite(t) ? t : "-");
}

function buildDiscSummary_(dominant, secondary) {
  const d = (dominant || "").toString().trim();
  const s = (secondary || "").toString().trim();
  if (!d && !s) return "";
  if (d && s) return d + " - " + s;
  return d || s;
}

function buildKesimpulanHr_(iq, discDominant, pauliWrong, position) {
  const pos = (position || "").toString().trim() || "posisi yang dilamar";
  const disc = (discDominant || "").toString().trim().toUpperCase();
  const iqNum = Number.isFinite(Number(iq)) ? Number(iq) : null;
  const wrongNum = pauliWrong === "" || pauliWrong == null ? null : Number(pauliWrong);

  let base;
  if (iqNum != null && iqNum >= 100) {
    if (pos.toLowerCase() === "sales" && (disc === "I" || disc === "S")) {
      base = "Cukup sesuai untuk posisi Sales dan berpotensi baik pada aspek komunikasi/relasi.";
    } else {
      base = "Cukup sesuai untuk " + pos + ".";
    }
  } else if (iqNum != null && iqNum < 100) {
    base = "Perlu pendampingan dan training untuk " + pos + ".";
  } else {
    base = "Perlu evaluasi lebih lanjut untuk " + pos + ".";
  }

  if (wrongNum != null && Number.isFinite(wrongNum) && wrongNum <= 2) {
    // keep one sentence only; add short clause if fits
    if (base.endsWith(".")) base = base.slice(0, -1);
    base += " dengan ketelitian cukup baik.";
  }

  // Ensure single concise sentence
  return base.replace(/\s+/g, " ").trim();
}

function buildRekomendasiAkhir_(iq, cbtRecommendation) {
  const iqNum = Number.isFinite(Number(iq)) ? Number(iq) : null;
  const rec = (cbtRecommendation || "").toString().toLowerCase();

  // Use recommendation text as a hint when present
  if (rec.includes("layak diprioritaskan")) return "Lanjut Interview";
  if (rec.includes("cukup layak")) return iqNum != null && iqNum < 90 ? "Perlu Review" : "Dipertimbangkan";
  if (rec.includes("perlu pertimbangan")) return "Perlu Review";

  if (iqNum == null) return "Perlu Review";
  if (iqNum >= 100) return "Lanjut Interview";
  if (iqNum >= 90) return "Dipertimbangkan";
  return "Perlu Review";
}

function normalizePhoneForWa_(phone) {
  const raw = (phone || "").toString().trim();
  if (!raw) return { display: "", wa: "" };

  // Prefer local display format (08...), even if the source lost leading zero
  const rawDigitsOnly = raw.replace(/[^\d]/g, "");
  let display = raw;
  if (rawDigitsOnly && rawDigitsOnly === rawDigitsOnly && rawDigitsOnly === raw) {
    if (rawDigitsOnly.startsWith("62") && rawDigitsOnly.length >= 9) display = "0" + rawDigitsOnly.slice(2);
    else if (rawDigitsOnly.startsWith("8") && rawDigitsOnly.length >= 8) display = "0" + rawDigitsOnly;
  }

  // build wa number digits only, with country code 62
  let digits = raw.replace(/[^\d+]/g, "");
  if (digits.startsWith("+")) digits = digits.slice(1);

  if (digits.startsWith("08")) digits = "62" + digits.slice(1);
  else if (digits.startsWith("8")) digits = "62" + digits; // fallback
  else if (digits.startsWith("62")) digits = digits;
  else if (digits.startsWith("0")) digits = "62" + digits.slice(1);

  // final cleanup
  digits = digits.replace(/[^\d]/g, "");
  return { display, wa: digits };
}

function phoneToWhatsAppFormula_(phone) {
  const p = normalizePhoneForWa_(phone);
  if (!p.display) return "";
  if (!p.wa) return p.display;
  return '=HYPERLINK("https://wa.me/' + p.wa + '","' + p.display + '")';
}

function phoneToWhatsAppRichText_(phone) {
  const p = normalizePhoneForWa_(phone);
  const builder = SpreadsheetApp.newRichTextValue();
  builder.setText(p.display || "");
  if (p.display && p.wa) builder.setLinkUrl("https://wa.me/" + p.wa);
  return builder.build();
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
