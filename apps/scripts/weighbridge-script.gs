// ═══════════════════════════════════════════════════════════════════
// PINAKA INFRA — WEIGHBRIDGE GOOGLE APPS SCRIPT
// Paste this entire file into Google Apps Script (script.google.com)
// File → New → Script → Replace all content → Save → Deploy
// ═══════════════════════════════════════════════════════════════════

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ─── SHEET NAMES ───
const SHEET_USERS    = 'Users';
const SHEET_VEHICLES = 'Vehicles';
const SHEET_SETTINGS = 'Settings';
const MASTER_DATA_URL = 'https://script.google.com/macros/s/AKfycbwRRSc5J2X3SYjJ_9pFDq7eoVwTICDICtyc8DUC_H-DIsgJHjnzmehbKrk7CCv8vZ0G/exec';

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

function masterWeighbridgeVehicles_() {
  const cfg = fetchMasterConfig_('weighbridge');
  const rows = cfg && Array.isArray(cfg.vehicles) ? cfg.vehicles : [];
  return rows
    .filter(v => v && v.no)
    .map(v => ({
      num: String(v.no || '').trim().toUpperCase(),
      type: String(v.type || '').trim(),
      driver: String(v.driver || '').trim()
    }));
}

function masterWeighbridgeUsers_() {
  const cfg = fetchMasterConfig_('weighbridge');
  const rows = cfg && Array.isArray(cfg.employees) ? cfg.employees : [];
  const users = rows
    .filter(u => u && u.id && u.name && u.pin)
    .map(u => ({
      id: String(u.id || '').trim(),
      name: String(u.name || '').trim(),
      pin: String(u.pin || '').trim(),
      role: String(u.role || 'operator').trim().toLowerCase() || 'operator'
    }));
  return users.length ? users : null;
}

// ─── CORS HEADERS ───
function setCORS(output) {
  if (!output) output = ContentService.createTextOutput('{"error":"null output"}');
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ═══════════════════════════════════════════════════════════════════
// GET HANDLER
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let result = {};

  try {
    if (action === 'getRST') {
      result = handleGetRST(e.parameter || {});
    } else if (action === 'getVehicles') {
      result = handleGetVehicles();
    } else if (action === 'getUsers') {
      result = handleGetUsers();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  return setCORS(ContentService.createTextOutput(JSON.stringify(result)));
}

// ═══════════════════════════════════════════════════════════════════
// POST HANDLER
// ═══════════════════════════════════════════════════════════════════
function doPost(e) {
  let result = {};
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('No POST body received');
    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';

    if (action === 'addTrips') {
      result = handleAddTripsLocked(body.trips || []);
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  return setCORS(ContentService.createTextOutput(JSON.stringify(result)));
}

// ═══════════════════════════════════════════════════════════════════
// GET RST — Returns next RST number + today's trips
// ═══════════════════════════════════════════════════════════════════
function handleGetRST(params) {
  const settings = getSettings();
  const sheets = getTripSheets();

  // Find today's trips (header is row 0)
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const range = (params && params.range) || 'today';
  const trips = [];
  let maxRST = (parseInt(settings.rst_start) || 1) - 1;

  for (const sheet of sheets) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue; // empty row
      const rowDate = row[1] ? Utilities.formatDate(new Date(row[1]), 'Asia/Kolkata', 'yyyy-MM-dd') : '';
      const rstVal = parseInt(row[4]) || 0;
      if (rstVal > maxRST) maxRST = rstVal;

      if (shouldIncludeTripDate(rowDate, today, range)) {
        trips.push({
          id: row[0].toString(),
          rst: row[4],
          date: rowDate,
          tin: row[2] ? Utilities.formatDate(new Date(row[2]), 'Asia/Kolkata', 'HH:mm') : '',
          tout: row[3] ? Utilities.formatDate(new Date(row[3]), 'Asia/Kolkata', 'HH:mm') : '',
          veh: row[5],
          vtype: row[6],
          driver: row[7],
          gross: row[8],
          tare: row[9],
          net: row[10],
          by: row[12],
          synced: true
        });
      }
    }
  }

  return { rst: maxRST + 1, trips: trips };
}

function shouldIncludeTripDate(rowDate, today, range) {
  if (!rowDate) return false;
  if (range === 'all') return true;
  if (range === 'month') return rowDate.slice(0, 7) === today.slice(0, 7);
  if (range === 'week') {
    const start = new Date();
    start.setDate(start.getDate() - 7);
    const startKey = Utilities.formatDate(start, 'Asia/Kolkata', 'yyyy-MM-dd');
    return rowDate >= startKey;
  }
  return rowDate === today;
}

// ═══════════════════════════════════════════════════════════════════
// GET VEHICLES — From Vehicles sheet
// ═══════════════════════════════════════════════════════════════════
function handleGetVehicles() {
  const masterVehicles = masterWeighbridgeVehicles_();
  if (masterVehicles.length) {
    const users = handleGetUsers().users;
    return { vehicles: masterVehicles, users };
  }

  const sh = getOrCreate(SHEET_VEHICLES);
  const data = sh.getDataRange().getValues();
  const vehicles = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[3] && row[3].toString().toUpperCase() === 'NO') continue;
    vehicles.push({ num: row[0], type: row[1], driver: row[2] });
  }

  const users = handleGetUsers().users;
  return { vehicles, users };
}

// ═══════════════════════════════════════════════════════════════════
// GET USERS — From Users sheet
// ═══════════════════════════════════════════════════════════════════
function handleGetUsers() {
  const masterUsers = masterWeighbridgeUsers_();
  if (masterUsers) return { users: masterUsers };

  const sh = getOrCreate(SHEET_USERS);
  const data = sh.getDataRange().getValues();
  const users = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    users.push({ id: row[0], name: row[1], pin: row[2].toString(), role: row[3] });
  }

  return { users };
}

// ═══════════════════════════════════════════════════════════════════
// ADD TRIPS — Save to monthly sheet
// ═══════════════════════════════════════════════════════════════════
function handleAddTrips(trips) {
  if (!trips.length) return { ok: true, nextRST: null };

  const monthSheet = getOrCreateMonthSheet();
  const existing = monthSheet.getDataRange().getValues();
  const existingIds = new Set(existing.slice(1).map(r => r[0].toString()));

  let maxRST = 0;
  const existingRSTs = existing.slice(1).map(r => parseInt(r[4]) || 0);
  if (existingRSTs.length) maxRST = Math.max(...existingRSTs);

  for (const trip of trips) {
    if (existingIds.has(trip.id.toString())) continue;

    const dateVal = trip.date ? new Date(trip.date + 'T00:00:00') : new Date();
    const tinVal  = trip.tin  ? parseTime(trip.date, trip.tin)  : '';
    const toutVal = trip.tout ? parseTime(trip.date, trip.tout) : '';

    monthSheet.appendRow([
      trip.id,           // Col A — Internal ID
      dateVal,           // Col B — Date
      tinVal,            // Col C — Time IN
      toutVal,           // Col D — Time OUT
      trip.rst,          // Col E — RST No
      trip.veh,          // Col F — Vehicle No
      trip.vtype,        // Col G — Vehicle Type
      trip.driver,       // Col H — Driver
      trip.gross,        // Col I — Gross KG
      trip.tare,         // Col J — Tare KG
      trip.net,          // Col K — Net KG
      (trip.net/1000).toFixed(3), // Col L — Net MT
      trip.by,           // Col M — Entered By
      new Date()         // Col N — Timestamp
    ]);

    if (parseInt(trip.rst) > maxRST) maxRST = parseInt(trip.rst);
  }

  // Format the sheet
  formatMonthSheet(monthSheet);

  return { ok: true, nextRST: maxRST + 1 };
}

// ═══════════════════════════════════════════════════════════════════
// SHEET HELPERS
// ═══════════════════════════════════════════════════════════════════

function handleAddTripsLocked(trips) {
  if (!trips.length) return { ok: true, nextRST: null, savedTrips: [] };

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const settings = getSettings();
    const monthSheet = getOrCreateMonthSheet();
    const existing = monthSheet.getDataRange().getValues();
    const existingIds = new Set(existing.slice(1).map(r => r[0].toString()));
    const savedTrips = [];

    let maxRST = getHighestRST(settings);

    for (const trip of trips) {
      const tripId = trip.id.toString();
      if (existingIds.has(tripId)) continue;

      const assignedRST = ++maxRST;
      const dateVal = trip.date ? new Date(trip.date + 'T00:00:00') : new Date();
      const tinVal  = trip.tin  ? parseTime(trip.date, trip.tin)  : '';
      const toutVal = trip.tout ? parseTime(trip.date, trip.tout) : '';

      monthSheet.appendRow([
        trip.id,
        dateVal,
        tinVal,
        toutVal,
        assignedRST,
        trip.veh,
        trip.vtype,
        trip.driver,
        trip.gross,
        trip.tare,
        trip.net,
        (trip.net/1000).toFixed(3),
        trip.by,
        new Date()
      ]);

      savedTrips.push({ id: tripId, rst: assignedRST });
      existingIds.add(tripId);
    }

    formatMonthSheet(monthSheet);

    return { ok: true, nextRST: maxRST + 1, savedTrips: savedTrips };
  } finally {
    lock.releaseLock();
  }
}

function getTripSheets() {
  const reserved = new Set([SHEET_USERS, SHEET_VEHICLES, SHEET_SETTINGS]);
  const sheets = SS.getSheets().filter(sh => {
    if (reserved.has(sh.getName())) return false;
    if (sh.getLastRow() < 1 || sh.getLastColumn() < 5) return false;
    const headers = sh.getRange(1, 1, 1, 5).getValues()[0];
    return headers[0] === 'ID' && headers[1] === 'Date' && headers[4] === 'RST No';
  });
  return sheets.length ? sheets : [getOrCreateMonthSheet()];
}

function getHighestRST(settings) {
  let maxRST = (parseInt(settings.rst_start) || 1) - 1;
  for (const sheet of getTripSheets()) {
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const rst = parseInt(values[i][4]) || 0;
      if (rst > maxRST) maxRST = rst;
    }
  }
  return maxRST;
}

function getOrCreateMonthSheet() {
  const monthName = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'MMM-yyyy');
  return getOrCreate(monthName, true);
}

function getOrCreate(name, isData) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    if (name === SHEET_USERS) setupUsersSheet(sh);
    else if (name === SHEET_VEHICLES) setupVehiclesSheet(sh);
    else if (name === SHEET_SETTINGS) setupSettingsSheet(sh);
    else if (isData) setupDataSheet(sh);
  }
  return sh;
}

function setupDataSheet(sh) {
  const headers = [
    'ID','Date','Time IN','Time OUT','RST No',
    'Vehicle No','Vehicle Type','Driver',
    'Gross KG','Tare KG','Net KG','Net MT',
    'Entered By','Timestamp'
  ];
  sh.appendRow(headers);

  // Style header
  const hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(11);
  sh.setFrozenRows(1);

  // Column widths
  sh.setColumnWidth(1, 120);  // ID
  sh.setColumnWidth(2, 100);  // Date
  sh.setColumnWidth(3, 90);   // Time IN
  sh.setColumnWidth(4, 90);   // Time OUT
  sh.setColumnWidth(5, 80);   // RST
  sh.setColumnWidth(6, 160);  // Vehicle No
  sh.setColumnWidth(7, 110);  // Type
  sh.setColumnWidth(8, 130);  // Driver
  sh.setColumnWidth(9, 90);   // Gross
  sh.setColumnWidth(10, 90);  // Tare
  sh.setColumnWidth(11, 90);  // Net KG
  sh.setColumnWidth(12, 80);  // Net MT
  sh.setColumnWidth(13, 100); // By
  sh.setColumnWidth(14, 140); // Timestamp
}

function setupUsersSheet(sh) {
  sh.appendRow(['ID', 'Name', 'PIN', 'Role']);
  const hdr = sh.getRange(1,1,1,4);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  // Default users
  sh.appendRow(['admin',      'Kartikey',   '0000', 'admin']);
  sh.appendRow(['weighbridge','Abhishek',   '1234', 'operator']);
  sh.appendRow(['accountant', 'Satyendra',  '5678', 'accountant']);
  sh.appendRow(['pm',         'Ritesh',     '9999', 'viewer']);
  // Instructions
  sh.getRange('F1').setValue('← Change PINs here. Roles: admin, operator, accountant, viewer');
  sh.setColumnWidths(1, 6, 140);
}

function setupVehiclesSheet(sh) {
  sh.appendRow(['Vehicle No', 'Vehicle Type', 'Driver', 'Active']);
  const hdr = sh.getRange(1,1,1,4);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  // Pre-load all vehicles
  const vehs = [
    ['UP71G0376','Compactor','Rajesh','YES'],
    ['UP71G0377','Compactor','Rohit Kumar','YES'],
    ['UP71G0378','Compactor','Ravi Kumar','YES'],
    ['UP70PT4525','Compactor','Hardeep Singh','YES'],
    ['MAIPCAPBDJ6K98790','Compactor','Susheel','YES'],
    ['ECE07620JR0015917','JCB','Dilip Kumar','YES'],
    ['UP71AT8186','JCB','Kuldeep','YES'],
    ['MBNBEBBBEKNG01659','Bull','Raj Bahadur','YES'],
    ['UP71AB4405','Bull','Shyam Bahadur','YES'],
    ['T052603265DL','Tractor Trolly','Pradip Kumar','YES'],
    ['T052602039DL','Tractor Trolly','Ram Vishal','YES'],
    ['UP71CT2559','DP Bin Small','Anand Shukla','YES'],
    ['UP71CT2551','DP Bin Small','Rajat Kumar','YES'],
    ['MAT804305N8D07858','DP Bin Double','Pankaj','YES'],
    ['UP71G0371','DP Bin Single','Sunil','YES'],
    ['UP71G0411','Dumper','Guddu Pal','YES'],
    ['UP71CT2546','Dumper','Riyaz Ali','YES'],
    ['UP71G0425','Dumper','Arun Kumar','YES'],
    ['UP71G0424','Dumper','Lal Bahadur','YES'],
    ['UP71CT2544','Dumper','Ravi Kumar','YES'],
    ['UP71CT2545','Dumper','Niyaz','YES'],
    ['UP71BT6195','Dumper','Sanjay Kumar','YES'],
    ['UP71BT6196','Dumper','Vimlesh Kumar','YES'],
    ['UP71CT3457','Magic','Akash Kumar','YES'],
    ['UP71CT3458','Magic','Ankit','YES'],
    ['UP71CT3459','Magic','Vanshi Lal','YES'],
    ['UP71CT3460','Magic','Kailash Kumar','YES'],
    ['UP71CT3461','Magic','Santosh','YES'],
    ['UP71CT0439','Magic','Ashwani','YES'],
    ['UP71CT0440','Magic','Ashish','YES'],
    ['UP71CT0445','Magic','Dileep Pal','YES'],
    ['UP71CT0447','Magic','Arun Kumar','YES'],
    ['UP71CT0448','Magic','Harishchandra','YES'],
    ['UP71CT0449','Magic','Gaurav','YES'],
    ['UP71CT0450','Magic','Amit Kumar','YES'],
    ['UP71G0379','Magic','Imran Ahmad','YES'],
    ['UP71G0384','Magic','Sumit','YES'],
    ['MA1FN2XURS6G11445','Magic','Ram Kishor','YES'],
    ['MA1FN2XURS6C10162','Magic','Talha Rahmaan','YES'],
    ['MA1FN2XURS6C10254','Magic','Shaeem Ahmad','YES'],
    ['MA1FN2XURS6G11530','Magic','Gulab Singh','YES'],
    ['MA1FN2XURS6G11466','Magic','Dilshad Islam','YES'],
    ['KBCAZ23ALS3638814','Poclain','—','YES'],
  ];
  for(const v of vehs) sh.appendRow(v);
  sh.setColumnWidths(1, 4, 180);
  sh.getRange('F1').setValue('← To add/remove vehicles or change drivers, edit here. Set Active=NO to hide.');
}

function setupSettingsSheet(sh) {
  sh.appendRow(['Key', 'Value', 'Description']);
  const hdr = sh.getRange(1,1,1,3);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.appendRow(['rst_start', '4652', 'Starting RST number — app will continue from highest found']);
  sh.appendRow(['project',   'Fatehpur MSW', 'Project name']);
  sh.appendRow(['md_mobile', '', 'MD WhatsApp number for daily reports (with country code, no +)']);
  sh.setColumnWidths(1, 3, 200);
}

function formatMonthSheet(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;
  // Date format Col B
  sh.getRange(2, 2, last-1, 1).setNumberFormat('dd-mmm-yyyy');
  // Time format Col C, D
  sh.getRange(2, 3, last-1, 2).setNumberFormat('hh:mm AM/PM');
  // Number format Gross/Tare/Net
  sh.getRange(2, 9, last-1, 3).setNumberFormat('#,##0');
  // MT format
  sh.getRange(2, 12, last-1, 1).setNumberFormat('0.000');
  // Alternate row colors
  for (let i = 2; i <= last; i++) {
    if (i % 2 === 0) sh.getRange(i, 1, 1, 14).setBackground('#F0F7F0');
    else sh.getRange(i, 1, 1, 14).setBackground('#FFFFFF');
  }
}

function getSettings() {
  const sh = getOrCreate(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function parseTime(dateStr, timeStr) {
  if (!timeStr || !dateStr) return '';
  try {
    const [h, m] = timeStr.split(':');
    const d = new Date(dateStr + 'T00:00:00');
    d.setHours(parseInt(h), parseInt(m), 0, 0);
    return d;
  } catch(e) { return ''; }
}

// ═══════════════════════════════════════════════════════════════════
// SETUP FUNCTION — Run this ONCE manually after pasting the script
// Tools → Run → runSetup
// ═══════════════════════════════════════════════════════════════════
function runSetup() {
  getOrCreate(SHEET_USERS);
  getOrCreate(SHEET_VEHICLES);
  getOrCreate(SHEET_SETTINGS);
  getOrCreateMonthSheet();
  SpreadsheetApp.getUi().alert(
    '✅ Weighbridge Setup Complete!\n\n' +
    'Sheets created:\n' +
    '• Users — manage logins & PINs\n' +
    '• Vehicles — your 43 vehicles pre-loaded\n' +
    '• Settings — RST start, project name\n' +
    '• ' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'MMM-yyyy') + ' — this month\'s data\n\n' +
    'Next: Deploy as Web App\n' +
    'Deploy → New Deployment → Web App\n' +
    'Execute as: Me | Who can access: Anyone\n' +
    'Copy the URL into weighbridge-app.html → SCRIPT_URL'
  );
}

// ═══════════════════════════════════════════════════════════════════
// MONTHLY SUMMARY — Run manually or schedule
// ═══════════════════════════════════════════════════════════════════
function getMonthlySummary() {
  const sh = getOrCreateMonthSheet();
  const data = sh.getDataRange().getValues();
  const summary = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[5]) continue;
    const veh = row[5];
    const type = row[6];
    const net = parseFloat(row[10]) || 0;
    if (!summary[veh]) summary[veh] = { type, trips: 0, netKG: 0 };
    summary[veh].trips++;
    summary[veh].netKG += net;
  }

  Logger.log('=== MONTHLY VEHICLE SUMMARY ===');
  Object.entries(summary)
    .sort((a,b) => b[1].netKG - a[1].netKG)
    .forEach(([veh, s]) => {
      Logger.log(`${veh} (${s.type}): ${s.trips} trips | ${(s.netKG/1000).toFixed(3)} MT`);
    });
}
