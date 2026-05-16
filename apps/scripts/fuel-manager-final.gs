// ================================================================
// PINAKA INFRA — Diesel Fuel Manager · Google Apps Script v5
// ================================================================
// Handles:
//   GET  ?action=config      → vehicles + users + last_readings + settings
//   GET  ?action=list        → all fuel log entries
//   GET  ?action=test        → diagnostic
//   POST {action:"saveEntry",    entry:{...}}
//   POST {action:"saveVehicle",  vehicle:{...}, editNo:"old_no"|null}
//   POST {action:"deleteVehicle",no:"UP71CT0440"}
//   POST {action:"saveUser",     user:{...}, editId:"old_id"|null}
//   POST {action:"deleteUser",   id:"gajendra"}
//   POST {action:"saveSettings", settings:{admin_whatsapp:"...", fuel_rate:92.5}}
// ================================================================
// WHAT CHANGED IN v5 (vs v4):
//   Fuel Log sheet  → 3 new columns: Cost (₹) | Efficiency | Efficiency Flag
//   Vehicles sheet  → 1 new column:  Target Efficiency
//   Settings        → fuel_rate key added (auto-handled, no structural change)
//   getVehicles()   → returns targetEfficiency per vehicle
//   saveVehicle()   → saves targetEfficiency
//   getEntries()    → returns cost, efficiency, efficiencyFlag
//   saveEntry()     → writes cost, efficiency, efficiencyFlag to sheet
//   setupSampleSheets() → updated headers + adds fuel_rate in Settings
// ================================================================
// DEPLOY SETTINGS (important!):
//   Execute as: Me
//   Who has access: Anyone  ← MUST be "Anyone", not "logged-in users"
// ================================================================

const MASTER_DATA_URL = 'https://script.google.com/macros/s/AKfycbwRRSc5J2X3SYjJ_9pFDq7eoVwTICDICtyc8DUC_H-DIsgJHjnzmehbKrk7CCv8vZ0G/exec';
const FUEL_BACKEND_SHEET = 'Diesel_Report_Source';
const FUEL_BACKEND_HEADERS = [
  'Report Date',
  'Shift',
  'Slip NO.',
  'Vehicle No.',
  'Vehicle Type',
  'Last fulling Date',
  'Last Filling Diesel (Ltr)',
  'Last Filling  (Hrs)',
  'Last Filling  (km)',
  'Today Filling  (Hrs)',
  'Today Filling  (km)',
  'Total Runing (Hrs)',
  'Total Runing (Km)',
  'Total Trip',
  'Total Diesel Filled (In Ltr)',
  'Total Petrol Filled (In Ltr)',
  'Route/Ward',
  'Driver Name',
  'Father Name',
  'Remark'
];
const FUEL_HEADER_ALIASES = {
  'Report Date': ['Date', 'Report, Date'],
  'Slip NO.': ['Slip No', 'Slip No.', 'SL NO.'],
  'Vehicle No.': ['Vehicle No'],
  'Driver Name': ['Driver'],
  'Last Filling  (Hrs)': ['Prev Reading'],
  'Last Filling  (km)': ['Prev Reading'],
  'Today Filling  (Hrs)': ['Curr Reading'],
  'Today Filling  (km)': ['Curr Reading'],
  'Total Runing (Hrs)': ['Running'],
  'Total Runing (Km)': ['Running'],
  'Total Diesel Filled (In Ltr)': ['Litres']
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fuel Backend')
    .addItem('Cleanup Diesel Report Source Headers', 'cleanupDieselReportSourceHeaders')
    .addToUi();
}

function getFuelBackendSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FUEL_BACKEND_SHEET);
  if (!sh) sh = ss.insertSheet(FUEL_BACKEND_SHEET);
  ensureFuelBackendHeaders_(sh);
  return sh;
}

function ensureFuelBackendHeaders_(sh) {
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, FUEL_BACKEND_HEADERS.length).setValues([FUEL_BACKEND_HEADERS]);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(h => String(h || "").trim());
    let nextCol = current.length + 1;
    FUEL_BACKEND_HEADERS.forEach(header => {
      if (!current.includes(header)) {
        sh.getRange(1, nextCol).setValue(header);
        current.push(header);
        nextCol++;
      }
    });
  }
  sh.getRange(1, 1, 1, sh.getLastColumn())
    .setFontWeight("bold")
    .setBackground("#1B5E20")
    .setFontColor("#FFFFFF");
  sh.setFrozenRows(1);
}

function fuelHeaderMap_(sh) {
  ensureFuelBackendHeaders_(sh);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const map = {};
  headers.forEach((header, i) => {
    if (header) map[header] = i;
  });
  Object.keys(FUEL_HEADER_ALIASES).forEach(canonical => {
    if (map[canonical] === undefined) return;
    FUEL_HEADER_ALIASES[canonical].forEach(alias => {
      if (map[alias] === undefined) map[alias] = map[canonical];
    });
  });
  return { headers, map };
}

function fuelCell_(row, map, header) {
  const idx = map[header];
  return idx === undefined ? "" : row[idx];
}

function cleanupDieselReportSourceHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FUEL_BACKEND_SHEET);
  if (!sh) throw new Error(FUEL_BACKEND_SHEET + " sheet not found");
  const allowed = new Set(FUEL_BACKEND_HEADERS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  for (let i = headers.length - 1; i >= 0; i--) {
    if (!allowed.has(headers[i])) sh.deleteColumn(i + 1);
  }
  ensureFuelBackendHeaders_(sh);
  SpreadsheetApp.getUi().alert("Diesel_Report_Source headers cleaned. Only approved fuel report columns remain.");
}

function isPetrolFuel_(e) {
  const text = [
    e && e.fuelType,
    e && e.vehicleType,
    e && e.remark
  ].join(" ").toUpperCase();
  return /\b(PETROL|BIKE|MOTORCYCLE|SCOOTER)\b/.test(text);
}

function previousFuelInfo_(sh, hm, vehicleNo) {
  const target = String(vehicleNo || "").trim().toUpperCase();
  if (!target || sh.getLastRow() < 2) return {};
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    const r = rows[i];
    const vno = String(fuelCell_(r, hm.map, "Vehicle No.") || "").trim().toUpperCase();
    if (vno !== target) continue;
    return {
      date: fuelCell_(r, hm.map, "Report Date"),
      dieselLitres: parseFloat(fuelCell_(r, hm.map, "Total Diesel Filled (In Ltr)")) || 0,
      petrolLitres: parseFloat(fuelCell_(r, hm.map, "Total Petrol Filled (In Ltr)")) || 0
    };
  }
  return {};
}

function appendFuelEntry_(sh, e, entryDate) {
  const hm = fuelHeaderMap_(sh);
  const row = new Array(hm.headers.length).fill("");
  const set = (header, value) => {
    const idx = hm.map[header];
    if (idx !== undefined) row[idx] = value;
  };
  const metric = String(e.metric || "KM").toUpperCase() === "HRS" ? "HRS" : "KM";
  const litres = parseFloat(e.litres) || 0;
  const petrol = isPetrolFuel_(e);
  const previous = previousFuelInfo_(sh, hm, e.vehicleNo);
  set("Slip NO.", String(e.slipNo || ""));
  set("Report Date", entryDate);
  set("Shift", String(e.shift || ""));
  set("Vehicle No.", String(e.vehicleNo || ""));
  set("Vehicle Type", String(e.vehicleType || ""));
  set("Driver Name", String(e.driver || ""));
  set("Last fulling Date", previous.date || "");
  set("Last Filling Diesel (Ltr)", petrol ? "" : (previous.dieselLitres || ""));
  set(metric === "HRS" ? "Last Filling  (Hrs)" : "Last Filling  (km)", parseFloat(e.prevReading) || 0);
  set(metric === "HRS" ? "Today Filling  (Hrs)" : "Today Filling  (km)", parseFloat(e.curReading) || 0);
  set(metric === "HRS" ? "Total Runing (Hrs)" : "Total Runing (Km)", parseFloat(e.running) || 0);
  set("Total Trip", 0);
  set("Total Diesel Filled (In Ltr)", petrol ? 0 : litres);
  set("Total Petrol Filled (In Ltr)", petrol ? litres : 0);
  set("Remark", String(e.remark || ""));
  sh.appendRow(row);
}

function fetchMasterConfig_(app) {
  if (!MASTER_DATA_URL) return null;
  const cacheKey = 'master_config_' + app;
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
    const url = MASTER_DATA_URL + '?action=config&app=' + encodeURIComponent(app);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
    const text = res.getContentText();
    const data = JSON.parse(text);
    if (!data || data.ok !== true) throw new Error(data && data.error ? data.error : 'Master returned not ok');
    cache.put(cacheKey, text, 120);
    return data;
  } catch (e) {
    Logger.log('Master data fetch failed for ' + app + ': ' + e.message);
    return null;
  }
}

function masterFuelVehicles_() {
  const cfg = fetchMasterConfig_('fuel');
  const rows = cfg && Array.isArray(cfg.vehicles) ? cfg.vehicles : [];
  return rows
    .filter(v => v && v.no)
    .map(v => ({
      no: String(v.no || '').trim().toUpperCase(),
      type: String(v.type || '').trim().toUpperCase(),
      driver: String(v.driver || '').trim(),
      metric: String(v.fuelMetric || 'KM').trim().toUpperCase() === 'HRS' ? 'HRS' : 'KM',
      fuelType: String(v.fuelType || v.fuel || '').trim().toUpperCase(),
      targetEfficiency: parseFloat(v.targetEfficiency) || 0
    }));
}

function masterFuelUsers_() {
  const cfg = fetchMasterConfig_('fuel');
  const rows = cfg && Array.isArray(cfg.employees) ? cfg.employees : [];
  const users = rows
    .filter(u => u && u.id && u.name && u.pin)
    .map(u => ({
      id: String(u.id || '').trim(),
      name: String(u.name || '').trim(),
      pin: String(u.pin || '').trim(),
      admin: false,
      whatsapp: String(u.whatsapp || u.mobile || '').trim()
    }));
  return users.length ? users : null;
}

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || "config").toLowerCase();
  let result;
  try {
    if      (action === "config")       result = getConfig();
    else if (action === "list")         result = getEntries();
    else if (action === "dieselsource") result = getDieselReportSource(e.parameter || {});
    else if (action === "test")         result = testConnection();
    else                                result = { ok: false, msg: "Unknown action: " + action };
  } catch (err) {
    Logger.log("doGet error: " + err.message);
    result = { ok: false, msg: err.message };
  }
  return jsonOut(result);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ ok: false, msg: "No POST body" });
    }
    const data = JSON.parse(e.postData.contents);
    const action = String(data.action || "saveEntry");

    if      (action === "saveEntry")     { saveEntry(data.entry);                        return jsonOut({ ok: true }); }
    else if (action === "saveVehicle")   { saveVehicle(data.vehicle, data.editNo||null); return jsonOut({ ok: true }); }
    else if (action === "deleteVehicle") { deleteVehicle(data.no);                       return jsonOut({ ok: true }); }
    else if (action === "saveUser")      { saveUser(data.user, data.editId||null);        return jsonOut({ ok: true }); }
    else if (action === "deleteUser")    { deleteUser(data.id);                           return jsonOut({ ok: true }); }
    else if (action === "saveSettings")  { saveSettings(data.settings);                  return jsonOut({ ok: true }); }
    else                                 { return jsonOut({ ok: false, msg: "Unknown action: " + action }); }
  } catch (err) {
    Logger.log("doPost error: " + err.message);
    return jsonOut({ ok: false, msg: err.message });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// DIAGNOSTIC
// ════════════════════════════════════════════════════════════════
function getDieselReportSource(params) {
  const sh = getFuelBackendSheet_();

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return {
      ok: true,
      sheet: FUEL_BACKEND_SHEET,
      rowCount: 0,
      columnCount: 0,
      headers: [],
      rows: []
    };
  }

  const maxRows = Math.max(1, Math.min(parseInt(params.limit || "2000", 10) || 2000, 10000));
  const readRows = Math.min(lastRow, maxRows + 1);
  const values = sh.getRange(1, 1, readRows, lastCol).getDisplayValues();
  const headers = values[0].map(h => String(h || "").trim());
  const rows = values.slice(1)
    .filter(row => row.some(cell => String(cell || "").trim() !== ""))
    .map(row => {
      const obj = {};
      headers.forEach((header, i) => {
        obj[header || ("Column " + (i + 1))] = row[i];
      });
      return obj;
    });

  return {
    ok: true,
    sheet: FUEL_BACKEND_SHEET,
    rowCount: Math.max(lastRow - 1, 0),
    returnedRows: rows.length,
    columnCount: lastCol,
    headers,
    rows
  };
}

function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = ss.getSheets().map(s => s.getName());
  const get = (n) => ss.getSheetByName(n);
  const fuelSh = get(FUEL_BACKEND_SHEET);
  const fuelCols = fuelSh ? fuelSh.getLastColumn() : 0;
  return {
    ok: true,
    spreadsheet: ss.getName(),
    allSheets: names,
    sheetStatus: {
      "Vehicles":  get("Vehicles")  ? (get("Vehicles").getLastRow()-1)  + " data rows" : "MISSING",
      "Users":     get("Users")     ? (get("Users").getLastRow()-1)     + " data rows" : "MISSING",
      "Settings":  get("Settings")  ? (get("Settings").getLastRow()-1)  + " data rows" : "MISSING",
      [FUEL_BACKEND_SHEET]: fuelSh ? (fuelSh.getLastRow()-1) + " entries · " + fuelCols + " cols (need 19)" : "Not created yet",
      "Fuel Log":  get("Fuel Log") ? (get("Fuel Log").getLastRow()-1) + " legacy entries" : "Not used"
    },
    vehicles: getVehicles(),
    users:    getUsers().map(u => ({ ...u, pin: "****" }))
  };
}

// ════════════════════════════════════════════════════════════════
// ONE-CLICK SETUP — Run once from Apps Script editor
// ════════════════════════════════════════════════════════════════
function setupSampleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const GREEN = "#1B5E20", WHITE = "#FFFFFF";

  // ── VEHICLES (7 columns now — added Target Efficiency) ──
  let vSh = ss.getSheetByName("Vehicles") || ss.insertSheet("Vehicles");
  vSh.clearContents();
  vSh.getRange(1,1,1,7).setValues([["Sl No","Vehicle No","Type","Driver","Metric","Active","Target Efficiency"]])
     .setFontWeight("bold").setBackground(GREEN).setFontColor(WHITE);
  const vData = [
    [1,  "UP71CT0440",  "MAGIC",    "Ashish",       "KM",  "Y", 0],
    [2,  "UP71CT0449",  "MAGIC",    "Gaurav",        "KM",  "Y", 0],
    [3,  "17332",       "MAGIC",    "Shameem",       "KM",  "Y", 0],
    [4,  "NEW",         "COMPACTOR","Susheel",       "HRS", "Y", 0],
    [5,  "RAISH",       "DP BIN",   "Niyaz",         "KM",  "Y", 0],
    [6,  "UP71CT3458",  "MAGIC",    "Ankit",         "KM",  "Y", 0],
    [7,  "UP71BT6195",  "DUMPER",   "Sanjay",        "KM",  "Y", 0],
    [8,  "UP71G0378",   "COMPACTOR","Ravi Kumar",    "HRS", "Y", 0],
    [9,  "UP71CT3457",  "MAGIC",    "Akash",         "KM",  "Y", 0],
    [10, "UP71CT0447",  "MAGIC",    "Arun Shukla",   "KM",  "Y", 0],
    [11, "CAR",         "CAR",      "Staff",         "KM",  "Y", 0],
    [12, "UP71CT3459",  "MAGIC",    "Vanshi Lal",    "KM",  "Y", 0],
    [13, "UP71AT8186",  "JCB",      "Kuldeep",       "HRS", "Y", 0],
    [14, "UP70PT4525",  "COMPACTOR","Hardeep",       "HRS", "Y", 0],
    [15, "UP71CT2545",  "DUMPER",   "Lal Bahadur",   "KM",  "Y", 0],
    [16, "UP71BT6196",  "DUMPER",   "Vimlesh",       "KM",  "Y", 0],
    [17, "R1",          "TRACTOR",  "Ram Vishal",    "HRS", "Y", 0],
    [18, "P1",          "TRACTOR",  "Pradeep",       "HRS", "Y", 0],
    [19, "UP71CT2544",  "DUMPER",   "Ravi",          "KM",  "Y", 0]
  ];
  vSh.getRange(2,1,vData.length,7).setValues(vData);
  vSh.setFrozenRows(1);
  vSh.autoResizeColumns(1,7);

  // ── USERS (unchanged) ──
  let uSh = ss.getSheetByName("Users") || ss.insertSheet("Users");
  uSh.clearContents();
  uSh.getRange(1,1,1,5).setValues([["ID","Name","PIN","Admin (Y/N)","WhatsApp No"]])
     .setFontWeight("bold").setBackground(GREEN).setFontColor(WHITE);
  uSh.getRange(2,1,2,5).setValues([
    ["gajendra","Gajendra Singh","1234","N",""],
    ["admin",   "Admin",         "9999","Y",""]
  ]);
  uSh.setFrozenRows(1);
  uSh.autoResizeColumns(1,5);

  // ── SETTINGS (added fuel_rate) ──
  let sSh = ss.getSheetByName("Settings") || ss.insertSheet("Settings");
  sSh.clearContents();
  sSh.getRange(1,1,1,2).setValues([["Key","Value"]])
     .setFontWeight("bold").setBackground(GREEN).setFontColor(WHITE);
  sSh.getRange(2,1,2,2).setValues([
    ["admin_whatsapp", ""],
    ["fuel_rate",      "0"]
  ]);
  sSh.setFrozenRows(1);
  sSh.autoResizeColumns(1,2);

  SpreadsheetApp.getUi().alert(
    "✅ Setup complete! (v5)\n\n" +
    "Sheets created:\n" +
    "• Vehicles (19 vehicles) — 7 columns incl. Target Efficiency\n" +
    "• Users (2 users)\n" +
    "• Settings (admin_whatsapp + fuel_rate)\n\n" +
    "Fuel Log will be created automatically on first fuel entry.\n" +
    "It will have 19 columns including Cost, Efficiency, Efficiency Flag.\n\n" +
    "NEXT: Deploy → Manage Deployments → Edit → New Version → Deploy\n" +
    "Make sure 'Who has access' = Anyone"
  );
}

// ════════════════════════════════════════════════════════════════
// MIGRATION HELPER — Run once if upgrading from v4 to v5
// Adds the 3 new columns to an existing Fuel Log sheet
// and adds Target Efficiency column to existing Vehicles sheet
// ════════════════════════════════════════════════════════════════
function migrateV4toV5() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const GREEN = "#1B5E20", WHITE = "#FFFFFF";
  let messages = [];

  // ── Fuel Log: add 3 new columns if sheet has exactly 16 ──
  const fuelSh = ss.getSheetByName("Fuel Log");
  if (fuelSh) {
    const lastCol = fuelSh.getLastColumn();
    if (lastCol === 16) {
      // Add headers for new columns 17, 18, 19
      fuelSh.getRange(1, 17, 1, 3).setValues([["Cost (Rs)","Efficiency","Efficiency Flag"]])
             .setFontWeight("bold").setBackground(GREEN).setFontColor(WHITE);
      // Fill existing data rows with defaults
      const lastRow = fuelSh.getLastRow();
      if (lastRow > 1) {
        const defaultData = [];
        for (let i = 2; i <= lastRow; i++) defaultData.push([0, 0, "OK"]);
        fuelSh.getRange(2, 17, defaultData.length, 3).setValues(defaultData);
      }
      fuelSh.autoResizeColumns(17,3);
      messages.push("✅ Fuel Log: added Cost, Efficiency, Efficiency Flag columns.");
    } else if (lastCol >= 19) {
      messages.push("ℹ Fuel Log already has " + lastCol + " columns — no migration needed.");
    } else {
      messages.push("⚠ Fuel Log has " + lastCol + " columns — unexpected, please check manually.");
    }
  } else {
    messages.push("ℹ Fuel Log sheet not found — it will be created on first entry.");
  }

  // ── Vehicles: add Target Efficiency column if sheet has exactly 6 ──
  const vSh = ss.getSheetByName("Vehicles");
  if (vSh) {
    const lastCol = vSh.getLastColumn();
    if (lastCol === 6) {
      vSh.getRange(1, 7).setValue("Target Efficiency")
         .setFontWeight("bold").setBackground(GREEN).setFontColor(WHITE);
      const lastRow = vSh.getLastRow();
      if (lastRow > 1) {
        const zeros = [];
        for (let i = 2; i <= lastRow; i++) zeros.push([0]);
        vSh.getRange(2, 7, zeros.length, 1).setValues(zeros);
      }
      vSh.autoResizeColumns(7,1);
      messages.push("✅ Vehicles: added Target Efficiency column.");
    } else if (lastCol >= 7) {
      messages.push("ℹ Vehicles already has " + lastCol + " columns — no migration needed.");
    } else {
      messages.push("⚠ Vehicles has " + lastCol + " columns — unexpected, please check manually.");
    }
  } else {
    messages.push("⚠ Vehicles sheet not found.");
  }

  // ── Settings: add fuel_rate if missing ──
  const sSh = ss.getSheetByName("Settings");
  if (sSh && sSh.getLastRow() > 1) {
    const rows = sSh.getRange(2, 1, sSh.getLastRow()-1, 1).getValues();
    const hasRate = rows.some(r => String(r[0]||"").trim() === "fuel_rate");
    if (!hasRate) {
      sSh.appendRow(["fuel_rate", "0"]);
      messages.push("✅ Settings: added fuel_rate key.");
    } else {
      messages.push("ℹ Settings: fuel_rate already exists.");
    }
  } else {
    messages.push("⚠ Settings sheet missing — run setupSampleSheets().");
  }

  SpreadsheetApp.getUi().alert("Migration Result:\n\n" + messages.join("\n"));
}

// ════════════════════════════════════════════════════════════════
// GET CONFIG
// ════════════════════════════════════════════════════════════════
function getConfig() {
  return {
    ok:            true,
    vehicles:      getVehicles(),
    users:         getUsers(),
    last_readings: getLastReadings(),
    settings:      getSettings()
  };
}

// ════════════════════════════════════════════════════════════════
// VEHICLES — READ  (v5: reads 7 cols, returns targetEfficiency)
// ════════════════════════════════════════════════════════════════
function getVehicles() {
  try {
    const masterVehicles = masterFuelVehicles_();
    if (masterVehicles.length) return masterVehicles;

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicles");
    if (!sh || sh.getLastRow() < 2) return [];
    // Read up to 7 columns — column 7 is Target Efficiency (may not exist on old sheets)
    const numCols = Math.min(sh.getLastColumn(), 7);
    const rows = sh.getRange(2, 1, sh.getLastRow()-1, numCols).getValues();
    return rows
      .filter(r => String(r[1]||"").trim() && String(r[5]||"Y").trim().toUpperCase() !== "N")
      .map(r => ({
        no:               String(r[1]||"").trim(),
        type:             String(r[2]||"").trim().toUpperCase(),
        driver:           String(r[3]||"").trim(),
        metric:           String(r[4]||"KM").trim().toUpperCase() === "HRS" ? "HRS" : "KM",
        targetEfficiency: parseFloat(r[6]) || 0   // ← NEW in v5
      }));
  } catch(e) { Logger.log("getVehicles: "+e.message); return []; }
}

// ════════════════════════════════════════════════════════════════
// VEHICLES — WRITE  (v5: saves targetEfficiency as column 7)
// ════════════════════════════════════════════════════════════════
function saveVehicle(v, editNo) {
  if (!v || !v.no) throw new Error("Vehicle number required");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Vehicles");
  if (!sh) {
    sh = ss.insertSheet("Vehicles");
    sh.appendRow(["Sl No","Vehicle No","Type","Driver","Metric","Active","Target Efficiency"]);
    sh.getRange(1,1,1,7).setFontWeight("bold").setBackground("#1B5E20").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }
  // Ensure column 7 header exists (migration safety)
  if (sh.getLastColumn() < 7) {
    sh.getRange(1, 7).setValue("Target Efficiency")
      .setFontWeight("bold").setBackground("#1B5E20").setFontColor("#FFFFFF");
  }

  const numCols = Math.min(sh.getLastColumn(), 7);
  const rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow()-1, numCols).getValues();
  const metric = String(v.metric||"KM").trim().toUpperCase() === "HRS" ? "HRS" : "KM";
  const targetEff = parseFloat(v.targetEfficiency) || 0;

  if (editNo) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][1]||"").trim() === String(editNo).trim()) {
        // Update columns 2–7 (Vehicle No through Target Efficiency)
        sh.getRange(i+2, 2, 1, 6).setValues([[
          String(v.no).trim().toUpperCase(),
          String(v.type||"").trim().toUpperCase(),
          String(v.driver||"").trim(),
          metric,
          String(rows[i][5]||"Y"),   // preserve Active flag
          targetEff
        ]]);
        return;
      }
    }
  }

  // Check for duplicate
  const exists = rows.some(r => String(r[1]||"").trim() === String(v.no).trim().toUpperCase());
  if (exists) throw new Error("Vehicle " + v.no + " already exists");

  // Append new row (7 columns)
  sh.appendRow([
    rows.length + 1,
    String(v.no).trim().toUpperCase(),
    String(v.type||"").trim().toUpperCase(),
    String(v.driver||"").trim(),
    metric,
    "Y",
    targetEff
  ]);
  sh.autoResizeColumns(1,7);
}

function deleteVehicle(no) {
  if (!no) throw new Error("Vehicle number required");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicles");
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  for (let i = rows.length-1; i >= 0; i--) {
    if (String(rows[i][1]||"").trim() === String(no).trim()) {
      sh.deleteRow(i+2);
      return;
    }
  }
}

// ════════════════════════════════════════════════════════════════
// USERS — READ (unchanged)
// ════════════════════════════════════════════════════════════════
function getUsers() {
  try {
    const masterUsers = masterFuelUsers_();
    if (masterUsers) return masterUsers;

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    if (!sh || sh.getLastRow() < 2) return [];
    const rows = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
    return rows
      .filter(r => String(r[0]||"").trim() && String(r[2]||"").trim())
      .map(r => ({
        id:       String(r[0]||"").trim(),
        name:     String(r[1]||"").trim(),
        pin:      String(r[2]||"").trim(),
        admin:    String(r[3]||"N").trim().toUpperCase() === "Y",
        whatsapp: String(r[4]||"").trim()
      }));
  } catch(e) { Logger.log("getUsers: "+e.message); return []; }
}

// ════════════════════════════════════════════════════════════════
// USERS — WRITE (unchanged)
// ════════════════════════════════════════════════════════════════
function saveUser(u, editId) {
  if (!u || !u.id) throw new Error("User ID required");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Users");
  if (!sh) {
    sh = ss.insertSheet("Users");
    sh.appendRow(["ID","Name","PIN","Admin (Y/N)","WhatsApp No"]);
    sh.getRange(1,1,1,5).setFontWeight("bold").setBackground("#1B5E20").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }
  const rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
  if (editId) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]||"").trim() === String(editId).trim()) {
        sh.getRange(i+2, 1, 1, 5).setValues([[
          String(u.id).trim(),
          String(u.name||"").trim(),
          String(u.pin||"").trim(),
          u.admin ? "Y" : "N",
          String(u.whatsapp||"").trim()
        ]]);
        return;
      }
    }
  }
  const exists = rows.some(r => String(r[0]||"").trim() === String(u.id).trim());
  if (exists) throw new Error("User ID '" + u.id + "' already exists");
  sh.appendRow([
    String(u.id).trim(),
    String(u.name||"").trim(),
    String(u.pin||"").trim(),
    u.admin ? "Y" : "N",
    String(u.whatsapp||"").trim()
  ]);
  sh.autoResizeColumns(1,5);
}

function deleteUser(id) {
  if (!id) throw new Error("User ID required");
  if (id === "admin") throw new Error("Cannot delete main admin user");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
  if (!sh || sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
  for (let i = rows.length-1; i >= 0; i--) {
    if (String(rows[i][0]||"").trim() === String(id).trim()) {
      sh.deleteRow(i+2);
      return;
    }
  }
}

// ════════════════════════════════════════════════════════════════
// SETTINGS (unchanged — key-value store handles fuel_rate automatically)
// ════════════════════════════════════════════════════════════════
function getSettings() {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
    if (!sh || sh.getLastRow() < 2) return {};
    const rows = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
    const obj = {};
    rows.forEach(r => {
      const k = String(r[0]||"").trim();
      if (k) obj[k] = String(r[1]||"").trim();
    });
    // Parse fuel_rate as number before returning
    if (obj.fuel_rate !== undefined) obj.fuel_rate = parseFloat(obj.fuel_rate) || 0;
    return obj;
  } catch(e) { Logger.log("getSettings: "+e.message); return {}; }
}

function saveSettings(settings) {
  if (!settings) throw new Error("No settings provided");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Settings");
  if (!sh) {
    sh = ss.insertSheet("Settings");
    sh.appendRow(["Key","Value"]);
    sh.getRange(1,1,1,2).setFontWeight("bold").setBackground("#1B5E20").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }
  const rows = sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  for (const [key, val] of Object.entries(settings)) {
    let found = false;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]||"").trim() === key) {
        sh.getRange(i+2, 2).setValue(String(val||""));
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([key, String(val||"")]);
  }
}

// ════════════════════════════════════════════════════════════════
// FUEL LOG — READ
// ════════════════════════════════════════════════════════════════
function getLastReadings() {
  try {
    const sh = getFuelBackendSheet_();
    if (!sh || sh.getLastRow() < 2) return {};
    // Columns 1-16 are unchanged — vehicleNo=col5(r[4]), curReading=col9(r[8])
    const hm = fuelHeaderMap_(sh);
    const rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    const map = {};
    rows.forEach((r, idx) => {
      const vno = String(fuelCell_(r, hm.map, "Vehicle No.")||"").trim().toUpperCase();
      const hrsReading = parseFloat(fuelCell_(r, hm.map, "Today Filling  (Hrs)")) || 0;
      const kmReading = parseFloat(fuelCell_(r, hm.map, "Today Filling  (km)")) || 0;
      const rdg = hrsReading > 0 ? hrsReading : kmReading;
      if (!vno || rdg <= 0) return;
      let dateStr = "";
      try {
        const raw = fuelCell_(r, hm.map, "Report Date");
        dateStr = raw instanceof Date && !isNaN(raw)
          ? Utilities.formatDate(raw, "Asia/Kolkata", "yyyy-MM-dd")
          : String(raw||"").trim().slice(0,10);
      } catch(_) {}
      const timeStr = String(fuelCell_(r, hm.map, "Time")||"").trim();
      const ex = map[vno];
      if (!ex || dateStr > ex.date || (dateStr === ex.date && (timeStr > ex.time || idx > ex.index))) {
        map[vno] = { reading: rdg, date: dateStr, time: timeStr, index: idx };
      }
    });
    Object.keys(map).forEach(k => delete map[k].index);
    return map;
  } catch(e) { Logger.log("getLastReadings: "+e.message); return {}; }
}

function getEntries() {
  try {
    const sh = getFuelBackendSheet_();
    if (!sh || sh.getLastRow() < 2) return { ok: true, entries: [] };

    const hm = fuelHeaderMap_(sh);
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const out = [];

    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      const slipNo = String(fuelCell_(r, hm.map, "Slip NO.") || "").trim();
      if (!slipNo) continue;

      let dateStr = "";
      try {
        const raw = fuelCell_(r, hm.map, "Report Date");
        dateStr = raw instanceof Date && !isNaN(raw)
          ? Utilities.formatDate(raw, "Asia/Kolkata", "yyyy-MM-dd")
          : String(raw || "").trim().slice(0, 10);
      } catch (_) {}

      const hrsReading = parseFloat(fuelCell_(r, hm.map, "Today Filling  (Hrs)")) || 0;
      const dieselLitres = parseFloat(fuelCell_(r, hm.map, "Total Diesel Filled (In Ltr)")) || 0;
      const petrolLitres = parseFloat(fuelCell_(r, hm.map, "Total Petrol Filled (In Ltr)")) || 0;
      out.push({
        slipNo,
        date: dateStr,
        time: "",
        shift: String(fuelCell_(r, hm.map, "Shift") || ""),
        vehicleNo: String(fuelCell_(r, hm.map, "Vehicle No.") || ""),
        vehicleType: String(fuelCell_(r, hm.map, "Vehicle Type") || ""),
        driver: String(fuelCell_(r, hm.map, "Driver Name") || ""),
        prevReading: parseFloat(fuelCell_(r, hm.map, "Last Filling  (Hrs)")) || parseFloat(fuelCell_(r, hm.map, "Last Filling  (km)")) || 0,
        curReading: hrsReading || parseFloat(fuelCell_(r, hm.map, "Today Filling  (km)")) || 0,
        running: parseFloat(fuelCell_(r, hm.map, "Total Runing (Hrs)")) || parseFloat(fuelCell_(r, hm.map, "Total Runing (Km)")) || 0,
        litres: dieselLitres || petrolLitres,
        fuelType: petrolLitres > 0 ? "PETROL" : "DIESEL",
        metric: hrsReading > 0 ? "HRS" : "KM",
        remark: String(fuelCell_(r, hm.map, "Remark") || ""),
        userName: "",
        userId: "",
        cost: 0,
        efficiency: 0,
        efficiencyFlag: "OK"
      });
    }
    return { ok: true, entries: out };
  } catch (e) {
    Logger.log("getEntries: " + e.message);
    return { ok: false, msg: e.message, entries: [] };
  }
}


// ════════════════════════════════════════════════════════════════
// FUEL LOG — WRITE  (v5: 19 columns now)
// ════════════════════════════════════════════════════════════════
// Fuel Log column map (v5):
//  1  Slip No
//  2  Date
//  3  Time
//  4  Shift
//  5  Vehicle No
//  6  Vehicle Type
//  7  Driver
//  8  Prev Reading
//  9  Curr Reading
// 10  Running
// 11  Litres
// 12  Metric
// 13  Remark
// 14  User Name
// 15  User ID
// 16  Logged At      ← internal timestamp, not synced back to app
// 17  Cost (Rs)      ← NEW in v5
// 18  Efficiency     ← NEW in v5
// 19  Efficiency Flag ← NEW in v5
// ════════════════════════════════════════════════════════════════
function saveEntry(e) {
  if (!e) throw new Error("No entry data");
  const sh = getFuelBackendSheet_();

  if (!sh) {
    sh = ss.insertSheet("Fuel Log");
    sh.appendRow([
      "Slip No","Date","Time","Shift","Vehicle No","Vehicle Type",
      "Driver","Prev Reading","Curr Reading","Running","Litres",
      "Metric","Remark","User Name","User ID","Logged At",
      "Cost (Rs)","Efficiency","Efficiency Flag"   // ← v5 additions
    ]);
    sh.getRange(1,1,1,19).setFontWeight("bold").setBackground("#1B5E20").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1,19);
  }

  // Safety: add new headers if sheet was created by v4 (16 columns)
  if (sh.getLastColumn() === 16) {
    sh.getRange(1, 17, 1, 3).setValues([["Cost (Rs)","Efficiency","Efficiency Flag"]])
      .setFontWeight("bold").setBackground("#1B5E20").setFontColor("#FFFFFF");
    sh.autoResizeColumns(17,3);
  }

  let entryDate;
  try {
    entryDate = e.date ? new Date(e.date + "T00:00:00+05:30") : new Date();
    if (isNaN(entryDate)) entryDate = new Date();
  } catch(_) { entryDate = new Date(); }

  appendFuelEntry_(sh, e, entryDate);
  return;

  sh.appendRow([
    String(e.slipNo      || ""),
    entryDate,
    String(e.time        || ""),
    String(e.shift       || ""),
    String(e.vehicleNo   || ""),
    String(e.vehicleType || ""),
    String(e.driver      || ""),
    parseFloat(e.prevReading)    || 0,
    parseFloat(e.curReading)     || 0,
    parseFloat(e.running)        || 0,
    parseFloat(e.litres)         || 0,
    String(e.metric      || "KM"),
    String(e.remark      || ""),
    String(e.userName    || ""),
    String(e.userId      || ""),
    new Date(),                               // Logged At (server timestamp)
    parseFloat(e.cost)           || 0,        // ← NEW in v5
    parseFloat(e.efficiency)     || 0,        // ← NEW in v5
    String(e.efficiencyFlag      || "OK")     // ← NEW in v5
  ]);
}
