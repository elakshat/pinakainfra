// ═══════════════════════════════════════════════════════
// PINAKA INFRA — Collection System v5
// Google Apps Script Backend
// ═══════════════════════════════════════════════════════
//
// SHEET STRUCTURE (auto-created on first run):
//
// 1. "Entries"    — All collection records (DO NOT edit manually)
// 2. "Settings"   — Global config, editable by Satyendra Babu
//    Columns: Key | Value | Notes
//    Keys:
//      admin_pin           → Admin PIN (default 9999)
//      accountant_pin      → Satyendra Babu PIN (default 1414)
//      md_whatsapp         → MD number with country code e.g. 919876543210
//      amount_residential  → Fixed amount Rs (default 50)
//      amount_commercial   → Fixed amount Rs (default 100)
//
// 3. "Collectors" — Collector list, editable by Satyendra Babu
//    Columns: ID | Name | PIN | Daily Target (Rs) | Color | Notes
//
// HOW TO ADD/CHANGE A COLLECTOR:
//   Open the "Collectors" sheet → add a new row with ID, Name, 3-digit PIN, Target, Color
//   Changes take effect next time the app is opened.
//
// HOW TO CHANGE PINS:
//   Settings sheet → find the row → change Value column
//
// ═══════════════════════════════════════════════════════

const ENTRIES_SHEET    = "Entries";
const SETTINGS_SHEET   = "Settings";
const COLLECTORS_SHEET = "Collectors";
const MASTER_DATA_URL = "https://script.google.com/macros/s/AKfycbwRRSc5J2X3SYjJ_9pFDq7eoVwTICDICtyc8DUC_H-DIsgJHjnzmehbKrk7CCv8vZ0G/exec";

// ── IMPORTANT: Paste your Google Sheet ID here ──────────
// Open your Sheet → look at the URL:
// https://docs.google.com/spreadsheets/d/YOUR_ID_HERE/edit
const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"; // ← Replace this
// ────────────────────────────────────────────────────────

const HEADERS = [
  "Receipt No","Date","Time",
  "Collector ID","Collector Name",
  "Category","Zone",
  "Payer Name","Address","Amount","Mode","Mobile","Notes",
  "Synced At"
];

function fetchMasterConfig_(app) {
  if (!MASTER_DATA_URL) return null;
  const cacheKey = "master_config_" + app;
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    const url = MASTER_DATA_URL + "?action=config&app=" + encodeURIComponent(app);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error("HTTP " + res.getResponseCode());
    const text = res.getContentText();
    const data = JSON.parse(text);
    if (!data || data.ok !== true) throw new Error(data && data.error ? data.error : "Master returned not ok");
    cache.put(cacheKey, text, 120);
    return data;
  } catch (e) {
    Logger.log("Master data fetch failed for " + app + ": " + e.message);
    return null;
  }
}

function masterUserChargeCollectors_() {
  const cfg = fetchMasterConfig_("usercharge");
  const rows = cfg && Array.isArray(cfg.employees) ? cfg.employees : [];
  const colors = ["#1B5E20", "#1565C0", "#4A148C", "#6D4C41", "#00695C"];
  const collectors = rows
    .filter(u => u && u.id && u.name && u.pin)
    .map((u, i) => ({
      id: String(u.id || "").trim(),
      name: String(u.name || "").trim(),
      pin: String(u.pin || "").trim(),
      target: 3000,
      color: colors[i % colors.length]
    }));
  return collectors.length ? collectors : null;
}

// ─── SPREADSHEET HELPER ─────────────────────────────────
// Works in BOTH contexts:
//   Web app (doGet/doPost) → getActiveSpreadsheet() returns null, openById required
//   Editor / initializeSheets → falls back to getActiveSpreadsheet if ID not set yet
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "YOUR_SPREADSHEET_ID_HERE") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      "SPREADSHEET_ID not set. Paste your Sheet ID at the top of the script " +
      "(find it in the Sheet URL between /d/ and /edit)."
    );
  }
  return ss;
}

function ensureEntriesSheetStructure(sh) {
  if (!sh) return; // null guard — prevents crash if sheet lookup fails
  const lastCol = sh.getLastColumn();
  if (lastCol === 13) {
    const lastRow = sh.getLastRow();
    sh.insertColumnAfter(8);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight("bold")
      .setBackground("#1B5E20")
      .setFontColor("#FFFFFF");
    if (lastRow > 1) {
      const blanks = Array.from({ length: lastRow - 1 }, () => [""]);
      sh.getRange(2, 9, blanks.length, 1).setValues(blanks);
    }
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 100);
    sh.setColumnWidth(3, 90);
    sh.setColumnWidth(7, 180);
    sh.setColumnWidth(8, 160);
    sh.setColumnWidth(9, 180);
  }
}

// ─── SHEET SETUP ────────────────────────────────────────

function setupEntriesSheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(ENTRIES_SHEET);
  if (!sh) sh = ss.insertSheet(ENTRIES_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight("bold")
      .setBackground("#1B5E20")
      .setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180); // Receipt No
    sh.setColumnWidth(2, 100); // Date
    sh.setColumnWidth(3, 90);  // Time
    sh.setColumnWidth(7, 180); // Zone
    sh.setColumnWidth(8, 160); // Payer Name
    sh.setColumnWidth(9, 180); // Address
  }
  ensureEntriesSheetStructure(sh);
  return sh;
}

function setupSettingsSheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    const defaults = [
      ["Key",                 "Value",  "Notes — do not edit the Key column"],
      ["admin_pin",           "9999",   "Admin login PIN (4 digits)"],
      ["accountant_pin",      "1414",   "Satyendra Babu PIN (4 digits) — view and export only"],
      ["md_whatsapp",         "",       "MD WhatsApp number with country code e.g. 919876543210"],
      ["amount_residential",  "50",     "Fixed collection amount for Residential (Rs)"],
      ["amount_commercial",   "100",    "Fixed collection amount for Commercial (Rs)"],
    ];
    sh.getRange(1, 1, defaults.length, 3).setValues(defaults);
    sh.getRange(1, 1, 1, 3)
      .setFontWeight("bold")
      .setBackground("#1B5E20")
      .setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 180);
    sh.setColumnWidth(2, 180);
    sh.setColumnWidth(3, 320);
    sh.getRange(2, 1, defaults.length - 1, 1).setFontWeight("bold");
  }
  return sh;
}

function setupCollectorsSheet() {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(COLLECTORS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(COLLECTORS_SHEET);
    const defaults = [
      ["ID",      "Name",        "PIN", "Daily Target (Rs)", "Color",    "Notes"],
      ["uc-collector", "User Charge Collection", "3333", "3000", "#1B5E20",  "Default user charge login"],
      ["shubham", "Shubham",     "678", "3000",              "#1B5E20",  ""],
      ["vishal",  "Vishal",      "789", "3000",              "#1565C0",  ""],
      ["x",       "Collector 3", "890", "3000",              "#4A148C",  "Pending — update Name and ID when decided"],
    ];
    sh.getRange(1, 1, defaults.length, 6).setValues(defaults);
    sh.getRange(1, 1, 1, 6)
      .setFontWeight("bold")
      .setBackground("#1B5E20")
      .setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 100);
    sh.setColumnWidth(2, 150);
    sh.setColumnWidth(3, 80);
    sh.setColumnWidth(4, 150);
    sh.setColumnWidth(5, 100);
    sh.setColumnWidth(6, 280);
  }
  return sh;
}

// Run this once from the Apps Script editor after first deployment
function initializeSheets() {
  setupEntriesSheet();
  setupSettingsSheet();
  setupCollectorsSheet();
  Logger.log("All sheets initialized.");
}

// ─── CONFIG READ ─────────────────────────────────────────

function readConfig() {
  const config = {
    admin_pin:          "9999",
    accountant_pin:     "1414",
    md_whatsapp:        "",
    amount_residential: 50,
    amount_commercial:  100,
    collectors:         []
  };

  // Read Settings sheet
  try {
    const sh = setupSettingsSheet();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        const key = String(rows[i][0] || "").trim().toLowerCase();
        const val = String(rows[i][1] || "").trim();
        if (!key) continue;
        if      (key === "admin_pin")          config.admin_pin          = val || "9999";
        else if (key === "accountant_pin")     config.accountant_pin     = val || "1414";
        else if (key === "md_whatsapp")        config.md_whatsapp        = val;
        else if (key === "amount_residential") config.amount_residential = Number(val) || 50;
        else if (key === "amount_commercial")  config.amount_commercial  = Number(val) || 100;
      }
    }
  } catch(e) { Logger.log("Settings read error: " + e); }

  // Read Collectors sheet
  try {
    const sh = setupCollectorsSheet();
    const lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      const rows = sh.getRange(2, 1, lastRow - 1, 5).getValues();
      for (var i = 0; i < rows.length; i++) {
        const id     = String(rows[i][0] || "").trim();
        const name   = String(rows[i][1] || "").trim();
        const pin    = String(rows[i][2] || "").trim();
        const target = Number(rows[i][3]) || 3000;
        const color  = String(rows[i][4] || "#1B5E20").trim();
        if (id && pin) {
          config.collectors.push({ id, name: name || id, pin, target, color });
        }
      }
    }
  } catch(e) { Logger.log("Collectors read error: " + e); }

  const masterCollectors = masterUserChargeCollectors_();
  if (masterCollectors) config.collectors = masterCollectors;

  return config;
}

// ─── RECEIPT NUMBER ──────────────────────────────────────

// Returns next receipt number for a collector on a given date
// Format: PI-YYMMDD-PFX-NNN  e.g. PI-260509-SHU-001
function getNextRno(cid, date) {
  const sh = setupEntriesSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return buildRno(cid, date, 1);

  // Count existing entries for this collector on this date
  const rows = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // rno, date, time, cid
  var count = 0;
  for (var i = 0; i < rows.length; i++) {
    const rowCid = String(rows[i][3] || "").trim();
    var rowDate  = rows[i][1];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, "Asia/Kolkata", "yyyy-MM-dd");
    } else {
      rowDate = String(rowDate || "").replace(/^'/, "").trim();
    }
    if (rowCid === cid && rowDate === date) count++;
  }
  return buildRno(cid, date, count + 1);
}

function buildRno(cid, date, n) {
  const dateStr = date ? date.replace(/-/g, "").slice(2) : "000000"; // YYMMDD
  var px;
  if      (cid === "x")       px = "COL3";
  else if (cid === "shubham") px = "SHU";
  else if (cid === "vishal")  px = "VIS";
  else                         px = cid.slice(0, 3).toUpperCase();
  return "PI-" + dateStr + "-" + px + "-" + String(n).padStart(3, "0");
}

// ─── HTTP HANDLERS ───────────────────────────────────────

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sh   = setupEntriesSheet();
    const now  = new Date();

    if (data.entry) {
      appendEntry(sh, data.entry, now);
      return jsonOut({ ok: true, count: 1 });
    }
    if (data.entries && Array.isArray(data.entries)) {
      data.entries.forEach(function(en){ appendEntry(sh, en, now); });
      return jsonOut({ ok: true, count: data.entries.length });
    }
    return jsonOut({ ok: false, error: "No entry field" });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const action = (e.parameter && e.parameter.action) || "ping";

  // ── PING ──
  if (action === "ping") {
    return jsonOut({ ok: true, msg: "Pinaka Collection API v5", version: 5 });
  }

  // ── CONFIG ──
  if (action === "config") {
    return jsonOut({ ok: true, config: readConfig() });
  }

  // ── NEXT RECEIPT NUMBER ──
  if (action === "nextRno") {
    const cid  = (e.parameter && e.parameter.cid)  || "";
    const date = (e.parameter && e.parameter.date) || "";
    if (!cid || !date) return jsonOut({ ok: false, error: "Missing cid or date" });
    const rno = getNextRno(cid, date);
    return jsonOut({ ok: true, rno: rno });
  }

  // ── LIST ENTRIES (month-filtered) ──
  if (action === "list") {
    const cid   = (e.parameter && e.parameter.cid)   || "";
    const month = (e.parameter && e.parameter.month) || ""; // e.g. "2026-05"
    const sh    = setupEntriesSheet();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return jsonOut({ ok: true, entries: [] });

    const rows    = sh.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    const tz      = "Asia/Kolkata";
    const entries = [];

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];

      // Normalize date
      var dateOut = r[1];
      if (dateOut instanceof Date) {
        dateOut = Utilities.formatDate(dateOut, tz, "yyyy-MM-dd");
      } else {
        dateOut = String(dateOut || "").replace(/^'/, "").trim();
      }

      // Filter by collector
      if (cid && String(r[3]) !== cid) continue;

      // Filter by month prefix (server-side, prevents full-sheet load issues)
      if (month && !dateOut.startsWith(month)) continue;

      // Normalize time
      var timeOut = r[2];
      if (timeOut instanceof Date) {
        timeOut = Utilities.formatDate(timeOut, tz, "hh:mm a");
      } else {
        timeOut = String(timeOut || "").replace(/^'/, "").trim();
      }

      entries.push({
        rno:      r[0],
        date:     dateOut,
        time:     timeOut,
        cid:      r[3],
        cname:    r[4],
        cat:      r[5],
        zone:     r[6],
        payer:    r[7],
        address:  r[8],
        amount:   r[9],
        mode:     r[10],
        mobile:   r[11],
        notes:    r[12],
        syncedAt: r[13]
      });
    }
    return jsonOut({ ok: true, entries: entries });
  }

  return jsonOut({ ok: false, error: "Unknown action" });
}

// ─── APPEND ENTRY ────────────────────────────────────────

function appendEntry(sh, en, now) {
  // Prevent duplicate receipt numbers (safe re-sync)
  if (en.rno) {
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const rnos = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < rnos.length; i++) {
        if (String(rnos[i][0]) === String(en.rno)) return; // Already exists
      }
    }
  }

  // Prefix with apostrophe so Sheets stores as plain text (no date auto-conversion)
  const dateText = en.date ? ("'" + en.date) : "";
  const timeText = en.time ? ("'" + en.time) : "";

  sh.appendRow([
    en.rno     || "",
    dateText,
    timeText,
    en.cid     || "",
    en.cname   || "",
    en.cat     || "",
    en.zone    || "",
    en.payer   || "",
    en.address || "",
    Number(en.amount) || 0,
    en.mode    || "",
    en.mobile  || "",
    en.notes   || "",
    Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")
  ]);
}

// ─── HELPER ─────────────────────────────────────────────

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── ONE-TIME UTILITIES ──────────────────────────────────

// Run once from editor if date/time columns show as Date objects instead of text
function fixExistingDates() {
  const sh = setupEntriesSheet();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log("No rows to fix."); return; }
  const tz    = "Asia/Kolkata";
  const range = sh.getRange(2, 2, lastRow - 1, 2); // B and C columns
  const vals  = range.getValues();
  const fixed = [];
  for (var i = 0; i < vals.length; i++) {
    var d = vals[i][0];
    var t = vals[i][1];
    if (d instanceof Date) d = Utilities.formatDate(d, tz, "yyyy-MM-dd");
    if (t instanceof Date) t = Utilities.formatDate(t, tz, "hh:mm a");
    fixed.push([String(d || ""), String(t || "")]);
  }
  range.setNumberFormat("@");
  range.setValues(fixed);
  Logger.log("Fixed " + fixed.length + " rows.");
}
