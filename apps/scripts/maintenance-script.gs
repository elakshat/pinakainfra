// ═══════════════════════════════════════════════════════════════════
// PINAKA INFRA — MAINTENANCE APP — GOOGLE APPS SCRIPT
// Paste this entire file into Google Apps Script
// Go to: script.google.com → New Project → Replace all → Save → Deploy
// ═══════════════════════════════════════════════════════════════════

// If this script is not bound to a spreadsheet, paste the spreadsheet ID here.
// Leave blank when the script is opened from Extensions > Apps Script inside the Sheet.
const SPREADSHEET_ID = '';

function getSS() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No active spreadsheet. Open the Sheet > Extensions > Apps Script, or set SPREADSHEET_ID in maintenance-script.gs.');
  }
  return ss;
}

// ── SHEET NAMES ─────────────────────────────────────────────────────
const SHEET_USERS    = 'Users';
const SHEET_ASSETS   = 'Assets';
const SHEET_SETTINGS = 'Settings';
const SHEET_VENDORS  = 'Vendors';
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

function masterMaintenanceNames_() {
  const cfg = fetchMasterConfig_('maintenance');
  const rows = cfg && Array.isArray(cfg.employees) ? cfg.employees : [];
  return rows
    .filter(e => e && e.name)
    .map(e => String(e.name || '').trim())
    .filter(Boolean);
}

function masterMaintenanceVehicles_() {
  const cfg = fetchMasterConfig_('maintenance');
  const rows = cfg && Array.isArray(cfg.vehicles) ? cfg.vehicles : [];
  return rows
    .filter(v => v && v.no)
    .map(v => ({
      num: String(v.no || '').trim().toUpperCase(),
      type: String(v.type || '').trim(),
      driver: String(v.driver || '').trim()
    }));
}

// ═══════════════════════════════════════════════════════════════════
// GET HANDLER
// FIX 1: Removed setCORS() wrapper — setHeader() does not exist on
//         GAS TextOutput objects. CORS is handled automatically by
//         Google when deployed as "Anyone" access.
// FIX 2: Guard against undefined `e` — happens when doGet is run
//         directly from the Apps Script editor (no HTTP request).
//         Always test via the deployed Web App URL, not Run button.
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  e = e || {};
  const params = e.parameter || {};
  const action = params.action || '';
  let result = {};

  try {
    if (action === 'getConfig')       result = handleGetConfig();
    else if (action === 'getTickets') result = handleGetTickets(params.month || 'all');
    else if (action === 'getVendors') result = handleGetVendors();
    else if (action === 'debug')      result = handleDebug();
    else result = { error: 'Unknown action. Use ?action=getConfig, ?action=getTickets, ?action=getVendors or ?action=debug' };
  } catch(err) {
    result = { error: err.message || err.toString() };
  }

  const json = JSON.stringify(result);
  const callback = /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(params.callback || '')
    ? params.callback
    : '';
  const output = ContentService.createTextOutput(
    callback ? callback + '(' + json + ');' : json
  );
  return output.setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
// POST HANDLER
// ═══════════════════════════════════════════════════════════════════
function doPost(e) {
  let result = {};
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';

    if (action === 'addTickets') result = handleAddTickets(body.tickets || []);
    else result = { error: 'Unknown action' };
  } catch(err) {
    result = { error: err.message || err.toString() };
  }

  const output = ContentService.createTextOutput(JSON.stringify(result));
  return output.setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
// GET CONFIG — Returns authorised names, assets, next ticket ID
// ═══════════════════════════════════════════════════════════════════
function handleGetConfig() {
  const masterNames = masterMaintenanceNames_();
  const masterVehicles = masterMaintenanceVehicles_();

  // Authorised names from Users sheet
  const usersSh = getOrCreate(SHEET_USERS);
  const usersData = usersSh.getDataRange().getValues();
  const names = [];
  for (let i = 1; i < usersData.length; i++) {
    if (usersData[i][0] && usersData[i][1] === 'YES') {
      names.push(usersData[i][0].toString());
    }
  }

  // Assets (vehicles) from Assets sheet
  const assetsSh = getOrCreate(SHEET_ASSETS);
  const assetsData = assetsSh.getDataRange().getValues();
  const vehicles = [];
  for (let i = 1; i < assetsData.length; i++) {
    if (!assetsData[i][0]) continue;
    vehicles.push({
      num:    assetsData[i][0].toString(),
      type:   assetsData[i][1].toString(),
      driver: assetsData[i][2].toString()
    });
  }
  const finalNames = masterNames.length ? masterNames : names;
  const finalVehicles = masterVehicles.length ? masterVehicles : vehicles;

  // Next ticket ID from all month sheets
  const ss = getSS();
  const monthRe = /^[A-Z][a-z]{2}-\d{4}$/;
  const sheets = ss.getSheets().filter(sh => monthRe.test(sh.getName()));
  let maxID = getSettings().ticket_start || 1;
  sheets.forEach(sh => {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const idStr = data[i][0] ? data[i][0].toString() : '';
      const match = idStr.match(/MNT-(\d+)/);
      if (match) {
        const n = parseInt(match[1]);
        if (n >= maxID) maxID = n + 1;
      }
    }
  });

  return { names: finalNames, vehicles: finalVehicles, nextId: maxID };
}

// ═══════════════════════════════════════════════════════════════════
// GET TICKETS — Returns tickets for a given month (for history view)
// ═══════════════════════════════════════════════════════════════════
function cellText(row, idx) {
  return row[idx] == null ? '' : row[idx].toString();
}

function handleGetTickets(monthStr) {
  const ss = getSS();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const monthRe = /^[A-Z][a-z]{2}-\d{4}$/;
  const sheets = (monthStr && monthStr !== 'all') ? [ss.getSheetByName(monthStr)].filter(Boolean) : ss.getSheets().filter(sh => monthRe.test(sh.getName()));
  if (!sheets.length) return { tickets: [] };
  const tickets = [];
  sheets.forEach(sh => {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const dateVal = row[1];
      const dateStr = dateVal instanceof Date ? Utilities.formatDate(dateVal, 'Asia/Kolkata', 'yyyy-MM-dd') : cellText(row, 1);
      if (dateStr !== todayStr) continue;
      const vendorText = cellText(row, 16);
      const vendorMatch = vendorText.match(/^(.*) \((VND-\d+)\)$/);
      tickets.push({
        id: cellText(row, 0), date: dateStr, time: cellText(row, 2), by: cellText(row, 3), auth: cellText(row, 4),
        cat: cellText(row, 5) || 'general', asset: cellText(row, 6), ctype: cellText(row, 7), desc: cellText(row, 8),
        amount: parseFloat(row[10]) || 0, status: (cellText(row, 11) || 'PENDING').toUpperCase(),
        amtPaid: row[12] ? parseFloat(row[12]) : null, paidDate: cellText(row, 13),
        gst: cellText(row, 15).toUpperCase() === 'YES',
        vendorName: vendorMatch ? vendorMatch[1] : '', vendorCode: vendorMatch ? vendorMatch[2] : ''
      });
    }
  });
  return { tickets };
}

function handleAddTickets(tickets) {
  if (!tickets.length) return { ok: true };

  const monthSheet = getOrCreateMonthSheet();
  ensureDataSheetColumns(monthSheet);
  const vendorsSheet = getOrCreateVendorsSheet();
  const existing   = monthSheet.getDataRange().getValues();
  const existingIds = new Set(existing.slice(1).map(r => r[0].toString()));

  let maxID = 0;
  const existingNums = existing.slice(1)
    .map(r => { const m = r[0].toString().match(/MNT-(\d+)/); return m ? parseInt(m[1]) : 0; })
    .filter(n => n > 0);
  if (existingNums.length) maxID = Math.max(...existingNums);

  for (const t of tickets) {
    if (existingIds.has(t.id)) continue;

    const dateVal = t.date ? new Date(t.date + 'T00:00:00') : new Date();
    const vendor = resolveVendorForTicket(vendorsSheet, t);

    monthSheet.appendRow([
      t.id,                          // Col A — Ticket ID
      dateVal,                       // Col B — Date
      t.time || '',                  // Col C — Time
      t.by || '',                    // Col D — Raised By
      t.auth ? 'YES' : 'FLAGGED',   // Col E — Authorised
      t.cat || '',                   // Col F — Category
      t.asset || '',                 // Col G — Asset
      t.ctype || '',                 // Col H — Complaint Type
      t.desc || '',                  // Col I — Description
      t.work || '',                  // Col J — Work Required
      t.amount || 0,                 // Col K — Amount Requested
      'PENDING',                     // Col L — Status (Satyendra changes to PAID in sheet)
      '',                            // Col M — Amount Paid (Satyendra fills)
      '',                            // Col N - Date of Payment (Satyendra fills)
      new Date(),                    // Col O - Timestamp
      isGstTicket(t) ? 'YES' : 'NO',    // Col P - GST Bill
      vendor ? vendor.name + ' (' + vendor.code + ')' : '' // Col Q - Vendor
    ]);

    const idNum = parseInt(t.id.replace('MNT-', '')) || 0;
    if (idNum > maxID) maxID = idNum;
  }

  formatMonthSheet(monthSheet);

  return { ok: true, nextId: maxID + 1 };
}

// ═══════════════════════════════════════════════════════════════════
// SHEET HELPERS
// ═══════════════════════════════════════════════════════════════════

function getOrCreateMonthSheet() {
  return getOrCreate(getCurrentMonthName(), true);
}

function getCurrentMonthName() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'MMM-yyyy');
}

function getOrCreate(name, isData) {
  const ss = getSS();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (name === SHEET_USERS)         setupUsersSheet(sh);
    else if (name === SHEET_ASSETS)   setupAssetsSheet(sh);
    else if (name === SHEET_SETTINGS) setupSettingsSheet(sh);
    else if (name === SHEET_VENDORS)  setupVendorsSheet(sh);
    else if (isData)                  setupDataSheet(sh);
  }
  return sh;
}

function setupDataSheet(sh) {
  const headers = [
    'Ticket ID', 'Date', 'Time', 'Raised By', 'Authorised',
    'Category', 'Asset', 'Complaint Type', 'Description',
    'Work Required', 'Amount Requested', 'Status',
    'Amount Paid', 'Date of Payment', 'Timestamp',
    'GST Bill', 'Vendor'
  ];
  sh.appendRow(headers);

  // Style header row
  const hdr = sh.getRange(1, 1, 1, headers.length);
  hdr.setBackground('#1B5E20')
     .setFontColor('#FFFFFF')
     .setFontWeight('bold')
     .setFontSize(11);
  sh.setFrozenRows(1);

  // Column widths
  sh.setColumnWidth(1, 100);   // Ticket ID
  sh.setColumnWidth(2, 100);   // Date
  sh.setColumnWidth(3, 80);    // Time
  sh.setColumnWidth(4, 150);   // Raised By
  sh.setColumnWidth(5, 90);    // Authorised
  sh.setColumnWidth(6, 100);   // Category
  sh.setColumnWidth(7, 200);   // Asset
  sh.setColumnWidth(8, 180);   // Complaint Type
  sh.setColumnWidth(9, 250);   // Description
  sh.setColumnWidth(10, 200);  // Work Required
  sh.setColumnWidth(11, 130);  // Amount Requested
  sh.setColumnWidth(12, 100);  // Status
  sh.setColumnWidth(13, 120);  // Amount Paid
  sh.setColumnWidth(14, 130);  // Date of Payment
  sh.setColumnWidth(15, 160);  // Timestamp
  sh.setColumnWidth(16, 90);   // GST Bill
  sh.setColumnWidth(17, 220);  // Vendor

  // Add notes on Status and Amount Paid columns
  sh.getRange(1, 12).setNote('Satyendra: Change PENDING to PAID after payment is made.');
  sh.getRange(1, 13).setNote('Satyendra: Enter the actual amount paid here.');
  sh.getRange(1, 14).setNote('Satyendra: Enter the date payment was made.');
}

function setupUsersSheet(sh) {
  sh.appendRow(['Name', 'Authorised (YES/NO)', 'Designation', 'Notes']);
  const hdr = sh.getRange(1, 1, 1, 4);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);

  // Pre-load all authorised names
  const names = [
    ['Ritesh Kapoor',           'YES', 'Project Head'],
    ['Abhinav Singh',           'YES', 'Project Manager'],
    ['Aman Pratap Singh',       'YES', 'Supervisor Parking'],
    ['Sandeep Tiwari',          'YES', 'Authorised'],
    ['Gajendra',                'YES', 'Authorised'],
    ['Jitendra Kumar',          'YES', 'Mechanic'],
    ['Mohd. Shahil',            'YES', 'Mechanic'],
    ['Sanjeev Kumar',           'YES', 'Supervisor'],
    ['Mohd. Adil Fareed',       'YES', 'Asst. Supervisor'],
    ['Shubham Singh',           'YES', 'Supervisor'],
    ['Sachin Singh',            'YES', 'Supervisor'],
    ['Saiyad Khalid Ali',       'YES', 'Supervisor'],
    ['Chhote Lal',              'YES', 'Supervisor'],
    ['Arvind Kumar Maurya',     'YES', 'Supervisor MRF'],
    ['Abhishek Kumar',          'YES', 'Supervisor MRF'],
    ['Utpal Pratap Singh',      'YES', 'Supervisor MRF'],
    ['Vishal Singh',            'YES', 'Supervisor MRF'],
    ['Yuvraj Singh',            'YES', 'Supervisor Parking'],
    ['Mahendra Kumar',          'YES', 'Supervisor Parking'],
    ['Mayank Dwivedi',          'YES', 'Supervisor C&T'],
    ['Raghvendra Pratap Singh', 'YES', 'Supervisor C&T'],
    ['Prabhat Kumar',           'YES', 'Supervisor C&T'],
    ['Arvind Pal',              'YES', 'Supervisor C&T'],
    ['Shiv Om Mishra',          'YES', 'Asst. Supervisor'],
    ['Priyanshu',               'YES', 'Supervisor C&T'],
  ];
  names.forEach(row => sh.appendRow(row));
  sh.setColumnWidths(1, 4, 180);
  sh.getRange('F1').setValue('To add/remove: change Authorised column to YES or NO. Do NOT delete rows.');
}

function setupAssetsSheet(sh) {
  sh.appendRow(['Vehicle No', 'Vehicle Type', 'Driver', 'Active (YES/NO)']);
  const hdr = sh.getRange(1, 1, 1, 4);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);

  const vehs = [
    ['UP71G0376',           'Compactor',      'Rajesh',           'YES'],
    ['UP71G0377',           'Compactor',      'Rohit Kumar',      'YES'],
    ['UP71G0378',           'Compactor',      'Ravi Kumar',       'YES'],
    ['UP70PT4525',          'Compactor',      'Hardeep Singh',    'YES'],
    ['MAIPCAPBDJ6K98790',   'Compactor',      'Susheel',          'YES'],
    ['ECE07620JR0015917',   'JCB',            'Dilip Kumar',      'YES'],
    ['UP71AT8186',          'JCB',            'Kuldeep',          'YES'],
    ['MBNBEBBBEKNG01659',   'Bull',           'Raj Bahadur',      'YES'],
    ['UP71AB4405',          'Bull',           'Shyam Bahadur',    'YES'],
    ['T052603265DL',        'Tractor Trolly', 'Pradip Kumar',     'YES'],
    ['T052602039DL',        'Tractor Trolly', 'Ram Vishal',       'YES'],
    ['UP71CT2559',          'DP Bin Small',   'Anand Shukla',     'YES'],
    ['UP71CT2551',          'DP Bin Small',   'Rajat Kumar',      'YES'],
    ['MAT804305N8D07858',   'DP Bin Double',  'Pankaj',           'YES'],
    ['UP71G0371',           'DP Bin Single',  'Sunil',            'YES'],
    ['UP71G0411',           'Dumper',         'Guddu Pal',        'YES'],
    ['UP71CT2546',          'Dumper',         'Riyaz Ali',        'YES'],
    ['UP71G0425',           'Dumper',         'Arun Kumar',       'YES'],
    ['UP71G0424',           'Dumper',         'Lal Bahadur',      'YES'],
    ['UP71CT2544',          'Dumper',         'Ravi Kumar',       'YES'],
    ['UP71CT2545',          'Dumper',         'Niyaz',            'YES'],
    ['UP71BT6195',          'Dumper',         'Sanjay Kumar',     'YES'],
    ['UP71BT6196',          'Dumper',         'Vimlesh Kumar',    'YES'],
    ['UP71CT3457',          'Magic',          'Akash Kumar',      'YES'],
    ['UP71CT3458',          'Magic',          'Ankit',            'YES'],
    ['UP71CT3459',          'Magic',          'Vanshi Lal',       'YES'],
    ['UP71CT3460',          'Magic',          'Kailash Kumar',    'YES'],
    ['UP71CT3461',          'Magic',          'Santosh',          'YES'],
    ['UP71CT0439',          'Magic',          'Ashwani',          'YES'],
    ['UP71CT0440',          'Magic',          'Ashish',           'YES'],
    ['UP71CT0445',          'Magic',          'Dileep Pal',       'YES'],
    ['UP71CT0447',          'Magic',          'Arun Kumar',       'YES'],
    ['UP71CT0448',          'Magic',          'Harishchandra',    'YES'],
    ['UP71CT0449',          'Magic',          'Gaurav',           'YES'],
    ['UP71CT0450',          'Magic',          'Amit Kumar',       'YES'],
    ['UP71G0379',           'Magic',          'Imran Ahmad',      'YES'],
    ['UP71G0384',           'Magic',          'Sumit',            'YES'],
    ['MA1FN2XURS6G11445',   'Magic',          'Ram Kishor',       'YES'],
    ['MA1FN2XURS6C10162',   'Magic',          'Talha Rahmaan',    'YES'],
    ['MA1FN2XURS6C10254',   'Magic',          'Shaeem Ahmad',     'YES'],
    ['MA1FN2XURS6G11530',   'Magic',          'Gulab Singh',      'YES'],
    ['MA1FN2XURS6G11466',   'Magic',          'Dilshad Islam',    'YES'],
    ['KBCAZ23ALS3638814',   'Poclain',        '—',                'YES'],
  ];
  vehs.forEach(row => sh.appendRow(row));
  sh.setColumnWidths(1, 4, 200);
  sh.getRange('F1').setValue('To update driver name: edit column C. Set Active=NO to hide from app.');
}

function setupSettingsSheet(sh) {
  sh.appendRow(['Key', 'Value', 'Description']);
  const hdr = sh.getRange(1, 1, 1, 3);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.appendRow(['ticket_start', 1,              'Starting ticket number']);
  sh.appendRow(['wa_group',     '',             'WhatsApp group number (with country code, no +)']);
  sh.appendRow(['project',      'Fatehpur MSW', 'Project name']);
  sh.setColumnWidths(1, 3, 220);
}

function formatMonthSheet(sh) {
  const last = sh.getLastRow();
  if (last < 2) return;

  // Date format Col B
  sh.getRange(2, 2, last - 1, 1).setNumberFormat('dd-mmm-yyyy');
  // Amount format Col K and M
  sh.getRange(2, 11, last - 1, 1).setNumberFormat('₹#,##0');
  sh.getRange(2, 13, last - 1, 1).setNumberFormat('₹#,##0');
  // Date of payment Col N
  sh.getRange(2, 14, last - 1, 1).setNumberFormat('dd-mmm-yyyy');

  // Colour rows by status
  for (let i = 2; i <= last; i++) {
    const statusCell = sh.getRange(i, 12).getValue().toString().toUpperCase();
    let bg = '#FFFFFF';
    if (statusCell === 'PENDING') bg = '#FFF8E1';
    else if (statusCell === 'PAID') bg = '#E8F5E9';
    sh.getRange(i, 1, 1, Math.max(17, sh.getLastColumn())).setBackground(bg);
  }
}

function ensureDataSheetColumns(sh) {
  const required = ['GST Bill', 'Vendor'];
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => v.toString());
  required.forEach(h => {
    if (!headers.includes(h)) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(h);
      headers.push(h);
    }
  });
}

function setupVendorsSheet(sh) {
  sh.appendRow(['Vendor Code', 'Vendor Name', 'Registered On', 'First Ticket ID']);
  const hdr = sh.getRange(1, 1, 1, 4);
  hdr.setBackground('#1B5E20').setFontColor('#FFFFFF').setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 4, 180);
}

function getOrCreateVendorsSheet() { return getOrCreate(SHEET_VENDORS); }

function handleDebug() {
  const ss = getSS();
  return {
    ok: true,
    version: 'maintenance-vendor-gst-2026-05-11',
    spreadsheet: ss.getName(),
    vendorsSheetExists: !!ss.getSheetByName(SHEET_VENDORS),
    currentMonth: getCurrentMonthName()
  };
}

function handleGetVendors() {
  const sh = getOrCreateVendorsSheet();
  const data = sh.getDataRange().getValues();
  const vendors = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || !data[i][1]) continue;
    vendors.push({ code: data[i][0].toString(), name: data[i][1].toString() });
  }
  return { vendors };
}

function nextVendorCode(sh) {
  const data = sh.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const m = data[i][0] ? data[i][0].toString().match(/VND-(\d+)/) : null;
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return 'VND-' + String(max + 1).padStart(4, '0');
}

function isGstTicket(t) {
  return t && (t.gst === true || String(t.gst || '').toUpperCase() === 'YES' || String(t.gst || '').toLowerCase() === 'true');
}

function resolveVendorForTicket(sh, t) {
  if (!isGstTicket(t) || !t.vendorName) return null;
  const name = t.vendorName.toString().trim();
  if (!name) return null;
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const existing = data[i][1] ? data[i][1].toString().trim() : '';
    if (existing && existing.toLowerCase() === name.toLowerCase()) return { code: data[i][0].toString(), name: existing };
  }
  const code = nextVendorCode(sh);
  sh.appendRow([code, name, new Date(), t.id || '']);
  sh.getRange(2, 3, Math.max(1, sh.getLastRow() - 1), 1).setNumberFormat('dd-mmm-yyyy');
  return { code, name };
}

function getSettings() {
  const sh = getOrCreate(SHEET_SETTINGS);
  const data = sh.getDataRange().getValues();
  const s = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) s[data[i][0]] = data[i][1];
  }
  return s;
}

// ═══════════════════════════════════════════════════════════════════
// MONTHLY SUMMARY — Run manually to get a quick report
// ═══════════════════════════════════════════════════════════════════
function getMonthlySummary() {
  const sh = getOrCreateMonthSheet();
  const data = sh.getDataRange().getValues();

  let totalTickets = 0, totalRequested = 0, totalPaid = 0;
  let pending = 0, paid = 0;
  const byCategory = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    totalTickets++;
    const amt     = parseFloat(row[10]) || 0;
    const amtPaid = parseFloat(row[12]) || 0;
    const status  = row[11].toString().toUpperCase();
    const cat     = row[5].toString();

    totalRequested += amt;
    if (status === 'PAID') { paid++; totalPaid += amtPaid; }
    else pending++;

    byCategory[cat] = (byCategory[cat] || 0) + amt;
  }

  Logger.log('=== PINAKA INFRA MAINTENANCE SUMMARY ===');
  Logger.log('Month: ' + getCurrentMonthName());
  Logger.log('Total Tickets: ' + totalTickets);
  Logger.log('Pending: ' + pending + ' | Paid: ' + paid);
  Logger.log('Total Requested: Rs.' + totalRequested.toLocaleString('en-IN'));
  Logger.log('Total Paid: Rs.' + totalPaid.toLocaleString('en-IN'));
  Logger.log('--- By Category ---');
  Object.entries(byCategory).forEach(([cat, amt]) => {
    Logger.log(cat + ': Rs.' + amt.toLocaleString('en-IN'));
  });
}

// ═══════════════════════════════════════════════════════════════════
// SETUP FUNCTION — Run this ONCE after pasting the script
// Tools → Run → runSetup
// ═══════════════════════════════════════════════════════════════════
function runSetup() {
  getOrCreate(SHEET_USERS);
  getOrCreate(SHEET_ASSETS);
  getOrCreate(SHEET_SETTINGS);
  getOrCreateVendorsSheet();
  getOrCreateMonthSheet();

  SpreadsheetApp.getUi().alert(
    'PINAKA INFRA — Maintenance App Setup Complete\n\n' +
    'Sheets created:\n' +
    '- Users (23 authorised names pre-loaded)\n' +
    '- Assets (all 43 vehicles pre-loaded)\n' +
    '- Settings\n' +
    '- ' + getCurrentMonthName() + ' (this month data sheet)\n\n' +
    'NEXT STEPS:\n' +
    '1. Deploy → New Deployment → Web App\n' +
    '2. Execute as: Me\n' +
    '3. Who can access: Anyone\n' +
    '4. Deploy → Copy the Web App URL\n' +
    '5. Open maintenance-app.html in a text editor\n' +
    '6. Find: YOUR_GOOGLE_APPS_SCRIPT_URL_HERE\n' +
    '7. Replace with your copied URL → Save\n\n' +
    'SATYENDRA BABU INSTRUCTIONS:\n' +
    'When a ticket is paid — open the ' + getCurrentMonthName() + ' sheet\n' +
    'Find the ticket row → Fill columns:\n' +
    '- Column L (Status): Type PAID\n' +
    '- Column M (Amount Paid): Enter actual amount paid\n' +
    '- Column N (Date of Payment): Enter payment date\n' +
    'App will automatically show PAID status.'
  );
}
