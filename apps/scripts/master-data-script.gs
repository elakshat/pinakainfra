// Pinaka Infra - Master Data Sheet
// DEV-SAFE: create a new Google Sheet, paste this script, run setupMasterSheet().
// Do not paste into live app sheets until the unified app is ready.

const MASTER_VERSION = 'master-data-dev-2026-05-13';
const MASTER_SPREADSHEET_ID = ''; // Leave blank when bound to the master Google Sheet.

const SHEET_EMPLOYEES = 'Employees';
const SHEET_VEHICLES = 'Vehicles';
const SHEET_SETTINGS = 'Settings';
const SHEET_AUDIT = 'Audit Log';
const SHEET_USER_MANAGER = 'User PIN Manager';
const SHEET_PORTAL_USERS = 'Portal Users';
const SHEET_PORTAL_APPS = 'Portal Apps';

const EMPLOYEE_HEADERS = [
  'Employee ID',
  'Full Name',
  'Mobile',
  'Father/Husband Name',
  'DOB',
  'DOJ',
  'Designation',
  'Department',
  'Role',
  'Site',
  'PIN',
  'WhatsApp',
  'Active',
  'Fuel Access',
  'Weighbridge Access',
  'MRF Access',
  'Maintenance Access',
  'Joining Access',
  'Monthly Salary',
  'CTC',
  'Bank Name',
  'Account Holder Name',
  'Account Number',
  'IFSC Code',
  'Notes',
  'Created At',
  'Updated At',
  'User Charge Access'
];

const VEHICLE_HEADERS = [
  'Vehicle No',
  'Vehicle Type',
  'Driver Name',
  'Driver Mobile',
  'Ward No',
  'Ward Name',
  'Working Area',
  'Active',
  'Fuel Enabled',
  'Weighbridge Enabled',
  'MRF Enabled',
  'Maintenance Enabled',
  'Fuel Metric',
  'Target Efficiency',
  'Notes',
  'Created At',
  'Updated At',
  'Fuel Type'
];

const SETTINGS_HEADERS = ['Key', 'Value', 'Description'];
const AUDIT_HEADERS = ['Timestamp', 'Action', 'Entity', 'Key', 'User', 'Notes'];
const USER_MANAGER_HEADERS = [
  'Employee ID',
  'Full Name',
  'PIN',
  'App Access',
  'Site',
  'Role',
  'Active',
  'Status'
];

const PORTAL_USER_HEADERS = [
  'Login ID',
  'User Name',
  'PIN',
  'Active',
  'Admin',
  'Fuel',
  'Weighbridge',
  'User Charge',
  'MRF',
  'Staff Joining',
  'Maintenance',
  'Site',
  'Role',
  'Status'
];

const PORTAL_APP_HEADERS = [
  'App Key',
  'App Name',
  'PIN Required',
  'Status'
];

function getMasterSS_() {
  if (MASTER_SPREADSHEET_ID) return SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the master Google Sheet > Extensions > Apps Script, or set MASTER_SPREADSHEET_ID.');
  return ss;
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Master Data')
    .addItem('Open Portal Apps', 'setupPortalAppsSheet')
    .addItem('Open Portal Users - PIN & Access', 'setupPortalUsersSheet')
    .addItem('Set Default App PINs', 'setDefaultAppPins')
    .addItem('Sync Portal Users', 'syncPortalUsers')
    .addItem('Setup / Repair Master Tabs', 'setupMasterSheet')
    .addToUi();
}

function yes_(value) {
  if (value === true) return 'YES';
  const s = String(value || '').trim().toUpperCase();
  return ['Y', 'YES', 'TRUE', '1', 'ACTIVE'].includes(s) ? 'YES' : 'NO';
}

function bool_(value) {
  return yes_(value) === 'YES';
}

function getOrCreate_(name, headers) {
  const ss = getMasterSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  ensureHeaders_(sh, headers);
  return sh;
}

function ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach((header, i) => {
      if (current[i] !== header) sh.getRange(1, i + 1).setValue(header);
    });
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#1B5E20')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
}

function rowObjects_(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(String);
  return data.slice(1).filter(r => r.some(v => v !== '')).map((row, idx) => {
    const obj = {_row: idx + 2};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function findRowByKey_(sh, keyHeader, keyValue) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return -1;
  const headers = data[0].map(String);
  const idx = headers.indexOf(keyHeader);
  if (idx < 0) throw new Error('Missing key header: ' + keyHeader);
  const key = String(keyValue || '').trim().toUpperCase();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idx] || '').trim().toUpperCase() === key) return r + 1;
  }
  return -1;
}

function writeRecord_(sh, headers, keyHeader, record) {
  const rowNo = findRowByKey_(sh, keyHeader, record[keyHeader]);
  const existing = rowNo > 0 ? sh.getRange(rowNo, 1, 1, headers.length).getValues()[0] : null;
  const existingObj = {};
  if (existing) headers.forEach((h, i) => existingObj[h] = existing[i]);
  const merged = Object.assign({}, existingObj, record);
  if (!merged['Created At']) merged['Created At'] = now_();
  merged['Updated At'] = now_();
  const values = headers.map(h => merged[h] === undefined ? '' : merged[h]);
  if (rowNo > 0) sh.getRange(rowNo, 1, 1, headers.length).setValues([values]);
  else sh.appendRow(values);
  return {updated: rowNo > 0, row: rowNo > 0 ? rowNo : sh.getLastRow()};
}

function getEmployeeRecordById_(id) {
  const sh = getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  const key = String(id || '').trim().toUpperCase();
  return rowObjects_(sh).find(r => String(r['Employee ID'] || '').trim().toUpperCase() === key) || null;
}

function applyAccessList_(record, accessText) {
  const access = String(accessText || '').trim().toLowerCase();
  if (!access) return;
  const parts = access.split(/[,\s]+/).map(s => s.trim()).filter(Boolean);
  const accessFields = [
    'Fuel Access',
    'Weighbridge Access',
    'MRF Access',
    'User Charge Access',
    'Maintenance Access',
    'Joining Access'
  ];
  const keys = {
    fuel: 'Fuel Access',
    weighbridge: 'Weighbridge Access',
    weigh: 'Weighbridge Access',
    wb: 'Weighbridge Access',
    mrf: 'MRF Access',
    usercharge: 'User Charge Access',
    uc: 'User Charge Access',
    collection: 'User Charge Access',
    maintenance: 'Maintenance Access',
    maint: 'Maintenance Access',
    joining: 'Joining Access',
    staff: 'Joining Access'
  };
  accessFields.forEach(field => record[field] = 'NO');
  parts.forEach(part => {
    if (part === 'admin' || part === 'all') {
      accessFields.forEach(field => record[field] = 'YES');
      return;
    }
    if (keys[part]) record[keys[part]] = 'YES';
  });
}

function hasDataRowsInFirstColumn_(sh) {
  if (sh.getLastRow() < 2) return false;
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1)
    .getValues()
    .some(row => String(row[0] || '').trim() !== '');
}

function setupPortalUsersSheet() {
  const ss = getMasterSS_();
  const sh = getOrCreate_(SHEET_PORTAL_USERS, PORTAL_USER_HEADERS);
  sh.getRange(1, 1, 1, PORTAL_USER_HEADERS.length)
    .setBackground('#1B5E20')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, PORTAL_USER_HEADERS.length, 130);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(14, 240);
  sh.getRange('P1').setValue('Instructions');
  sh.getRange('P2').setValue('One row per portal user. Tick the app boxes this user can open.');
  sh.getRange('P3').setValue('Admin checked gives all app access. Active unchecked disables login.');
  sh.getRange('P4').setValue('After editing, run Master Data > Sync Portal Users.');
  if (!hasDataRowsInFirstColumn_(sh)) {
    sh.getRange(2, 1, 5, PORTAL_USER_HEADERS.length).setValues([
      ['admin', 'Admin', '4321', true, true, true, true, true, true, true, true, 'fatehpur', 'Admin', 'All apps'],
      ['fuel-manager', 'Fuel Manager', '1111', true, false, true, false, false, false, false, false, 'fatehpur', 'Fuel Manager', 'Fuel only'],
      ['weighbridge', 'Weighbridge Incharge', '2222', true, false, false, true, false, false, false, false, 'fatehpur', 'Weighbridge Incharge', 'Weighbridge only'],
      ['mrf-supervisor', 'MRF Supervisor', '4444', true, false, false, false, false, true, false, false, 'fatehpur', 'MRF Supervisor', 'MRF only'],
      ['uc-collector', 'User Charge Collector', '3333', true, false, false, false, true, false, false, false, 'fatehpur', 'Collector', 'User Charge only']
    ]);
  }
  applyPortalUserCheckboxes_(sh);
  ss.setActiveSheet(sh);
  return {ok:true, sheet:SHEET_PORTAL_USERS};
}

function defaultPortalPinRows_() {
  return [
    ['fuel-manager', 'Fuel Manager', '1111', true, false, true, false, false, false, false, false, 'fatehpur', 'Fuel Manager', 'Fuel only'],
    ['weighbridge', 'Weighbridge Incharge', '2222', true, false, false, true, false, false, false, false, 'fatehpur', 'Weighbridge Incharge', 'Weighbridge only'],
    ['uc-collector', 'User Charge Collector', '3333', true, false, false, false, true, false, false, false, 'fatehpur', 'Collector', 'User Charge only'],
    ['mrf-supervisor', 'MRF Supervisor', '4444', true, false, false, false, false, true, false, false, 'fatehpur', 'MRF Supervisor', 'MRF only']
  ];
}

function setDefaultAppPins() {
  const ss = getMasterSS_();
  const sh = getOrCreate_(SHEET_PORTAL_USERS, PORTAL_USER_HEADERS);
  setupPortalUsersSheet();
  defaultPortalPinRows_().forEach(row => {
    const record = {};
    PORTAL_USER_HEADERS.forEach((header, idx) => record[header] = row[idx]);
    writeRecord_(sh, PORTAL_USER_HEADERS, 'Login ID', record);
  });
  applyPortalUserCheckboxes_(sh);
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    'Default app PINs are ready in Portal Users:\n\n' +
    'Fuel Manager: 1111\n' +
    'Weighbridge: 2222\n' +
    'User Charge Collection: 3333\n' +
    'MRF Supervisor: 4444\n\n' +
    'To change them later, edit the PIN cells in Portal Users. Keep PINs unique.'
  );
  return {ok:true, updated:4};
}

function applyPortalUserCheckboxes_(sh) {
  const firstCheckboxCol = PORTAL_USER_HEADERS.indexOf('Active') + 1;
  const checkboxCols = PORTAL_USER_HEADERS.indexOf('Maintenance') - PORTAL_USER_HEADERS.indexOf('Active') + 1;
  const rowsToFormat = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, firstCheckboxCol, rowsToFormat, checkboxCols).insertCheckboxes();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const range = sh.getRange(2, firstCheckboxCol, lastRow - 1, checkboxCols);
  const values = range.getValues().map(row => row.map((value, idx) => {
    if (idx === 0 && value === '') return true;
    return bool_(value);
  }));
  range.setValues(values);
}

function setupPortalAppsSheet() {
  const ss = getMasterSS_();
  const sh = getOrCreate_(SHEET_PORTAL_APPS, PORTAL_APP_HEADERS);
  sh.getRange(1, 1, 1, PORTAL_APP_HEADERS.length)
    .setBackground('#1B5E20')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, PORTAL_APP_HEADERS.length, 150);
  sh.setColumnWidth(2, 220);
  sh.getRange('F1').setValue('Instructions');
  sh.getRange('F2').setValue('PIN Required checked = app asks for PIN.');
  sh.getRange('F3').setValue('PIN Required unchecked = app opens directly from portal.');
  if (!hasDataRowsInFirstColumn_(sh)) {
    sh.getRange(2, 1, 6, PORTAL_APP_HEADERS.length).setValues([
      ['fuel', 'Fuel Manager', true, 'Protected'],
      ['weighbridge', 'Weighbridge', true, 'Protected'],
      ['userCharge', 'User Charge Collection', true, 'Protected'],
      ['mrf', 'MRF Supervisor', true, 'Protected'],
      ['joining', 'Staff Joining', false, 'Open without PIN'],
      ['maintenance', 'Maintenance', false, 'Open without PIN']
    ]);
  }
  const rowsToFormat = Math.max(sh.getMaxRows() - 1, 1);
  sh.getRange(2, 3, rowsToFormat, 1).insertCheckboxes();
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const range = sh.getRange(2, 3, lastRow - 1, 1);
    range.setValues(range.getValues().map(row => [bool_(row[0])]));
  }
  ss.setActiveSheet(sh);
  return {ok:true, sheet:SHEET_PORTAL_APPS};
}

function portalUserRecordFromRow_(row, idx) {
  const admin = bool_(row[idx.Admin]);
  return {
    'Employee ID': String(row[idx['Login ID']] || '').trim(),
    'Full Name': String(row[idx['User Name']] || '').trim(),
    PIN: String(row[idx.PIN] || '').trim(),
    Active: yes_(row[idx.Active] === '' ? true : row[idx.Active]),
    Site: String(row[idx.Site] || '').trim(),
    Role: String(row[idx.Role] || '').trim(),
    Designation: String(row[idx.Role] || '').trim(),
    'Fuel Access': yes_(admin || row[idx.Fuel]),
    'Weighbridge Access': yes_(admin || row[idx.Weighbridge]),
    'User Charge Access': yes_(admin || row[idx['User Charge']]),
    'MRF Access': yes_(admin || row[idx.MRF]),
    'Joining Access': yes_(admin || row[idx['Staff Joining']]),
    'Maintenance Access': yes_(admin || row[idx.Maintenance])
  };
}

function portalRecordFromEmployeeRecord_(record) {
  const fuel = bool_(record['Fuel Access']);
  const weighbridge = bool_(record['Weighbridge Access']);
  const userCharge = bool_(record['User Charge Access']);
  const mrf = bool_(record['MRF Access']);
  const joining = bool_(record['Joining Access']);
  const maintenance = bool_(record['Maintenance Access']);
  const admin = fuel && weighbridge && userCharge && mrf && joining && maintenance;
  return {
    'Login ID': String(record['Employee ID'] || '').trim(),
    'User Name': String(record['Full Name'] || '').trim(),
    PIN: String(record.PIN || '').trim(),
    Active: bool_(record.Active === undefined ? true : record.Active),
    Admin: admin,
    Fuel: admin || fuel,
    Weighbridge: admin || weighbridge,
    'User Charge': admin || userCharge,
    MRF: admin || mrf,
    'Staff Joining': admin || joining,
    Maintenance: admin || maintenance,
    Site: String(record.Site || '').trim(),
    Role: String(record.Role || record.Designation || '').trim(),
    Status: 'Updated from PIN manager at ' + now_()
  };
}

function employeeRecordFromPortalRecord_(record) {
  const admin = bool_(record.Admin);
  return {
    'Employee ID': String(record['Login ID'] || '').trim(),
    'Full Name': String(record['User Name'] || '').trim(),
    PIN: String(record.PIN || '').trim(),
    Active: yes_(record.Active === undefined ? true : record.Active),
    Site: String(record.Site || '').trim(),
    Role: String(record.Role || '').trim(),
    Designation: String(record.Role || '').trim(),
    'Fuel Access': yes_(admin || record.Fuel),
    'Weighbridge Access': yes_(admin || record.Weighbridge),
    'User Charge Access': yes_(admin || record['User Charge']),
    'MRF Access': yes_(admin || record.MRF),
    'Joining Access': yes_(admin || record['Staff Joining']),
    'Maintenance Access': yes_(admin || record.Maintenance)
  };
}

function hasDuplicatePortalPin_(pin, loginId) {
  pin = String(pin || '').trim();
  if (!pin) return false;
  const ss = getMasterSS_();
  const sh = ss.getSheetByName(SHEET_PORTAL_USERS);
  if (!sh || sh.getLastRow() < 2) return false;
  const id = String(loginId || '').trim().toUpperCase();
  return rowObjects_(sh).some(row => {
    const rowId = String(row['Login ID'] || '').trim().toUpperCase();
    const rowPin = String(row.PIN || '').trim();
    const active = bool_(row.Active === '' ? true : row.Active);
    return active && rowPin === pin && rowId !== id;
  });
}

function writePortalUserAndEmployee_(portalRecord, actor, auditAction) {
  const portalSh = getOrCreate_(SHEET_PORTAL_USERS, PORTAL_USER_HEADERS);
  const portalResult = writeRecord_(portalSh, PORTAL_USER_HEADERS, 'Login ID', portalRecord);
  applyPortalUserCheckboxes_(portalSh);
  const empRecord = employeeRecordFromPortalRecord_(portalRecord);
  const empSh = getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  const empResult = writeRecord_(empSh, EMPLOYEE_HEADERS, 'Employee ID', empRecord);
  logAudit_(auditAction || 'portal user update', 'employee', empRecord['Employee ID'], actor || '', empRecord['Full Name']);
  return {portal: portalResult, employee: empResult};
}

function syncPortalUsers() {
  const sh = getOrCreate_(SHEET_PORTAL_USERS, PORTAL_USER_HEADERS);
  if (!hasDataRowsInFirstColumn_(sh)) {
    setupPortalUsersSheet();
    return {ok:true, synced:0, message:'Portal Users sheet created. Add rows and run sync again.'};
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = {};
  PORTAL_USER_HEADERS.forEach(h => idx[h] = headers.indexOf(h));
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), PORTAL_USER_HEADERS.length)).getValues();
  const pinCounts = {};
  rows.forEach(row => {
    const pin = String(row[idx.PIN] || '').trim();
    const active = yes_(row[idx.Active] === '' ? true : row[idx.Active]) === 'YES';
    if (pin && active) pinCounts[pin] = (pinCounts[pin] || 0) + 1;
  });
  const empSh = getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  let synced = 0;
  rows.forEach((row, i) => {
    const id = String(row[idx['Login ID']] || '').trim();
    const name = String(row[idx['User Name']] || '').trim();
    const statusCell = sh.getRange(i + 2, idx.Status + 1);
    if (!id && !name) return;
    if (!id || !name) {
      statusCell.setValue('Missing Login ID or User Name');
      return;
    }
    const pin = String(row[idx.PIN] || '').trim();
    if (pin && pinCounts[pin] > 1) {
      statusCell.setValue('Duplicate PIN. Give this user a unique PIN before syncing.');
      return;
    }
    const record = portalUserRecordFromRow_(row, idx);
    const result = writeRecord_(empSh, EMPLOYEE_HEADERS, 'Employee ID', record);
    statusCell.setValue((result.updated ? 'Updated' : 'Added') + ' employee row ' + result.row + ' at ' + now_());
    logAudit_(result.updated ? 'portal user update' : 'portal user create', 'employee', id, Session.getActiveUser().getEmail(), name);
    synced++;
  });
  SpreadsheetApp.getUi().alert('Portal Users sync complete.\nRows synced: ' + synced);
  return {ok:true, synced:synced};
}

function setupUserPinManagerSheet() {
  return setupPortalUsersSheet();
}

function syncUserPinManager() {
  const sh = getOrCreate_(SHEET_USER_MANAGER, USER_MANAGER_HEADERS);
  if (sh.getLastRow() < 2) {
    setupUserPinManagerSheet();
    return {ok:true, synced:0, message:'User PIN Manager sheet created. Add rows and run sync again.'};
  }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = {};
  USER_MANAGER_HEADERS.forEach(h => idx[h] = headers.indexOf(h));
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), USER_MANAGER_HEADERS.length)).getValues();
  let synced = 0;
  rows.forEach((row, i) => {
    const id = String(row[idx['Employee ID']] || '').trim();
    const name = String(row[idx['Full Name']] || '').trim();
    const pin = String(row[idx.PIN] || '').trim();
    if (!id && !name && !pin) return;
    const statusCell = sh.getRange(i + 2, idx.Status + 1);
    if (!id || !name) {
      statusCell.setValue('Missing Employee ID or Full Name');
      return;
    }
    if (pin && hasDuplicatePortalPin_(pin, id)) {
      statusCell.setValue('Duplicate PIN. Give this user a unique PIN before syncing.');
      return;
    }
    const record = {
      'Employee ID': id,
      'Full Name': name,
      Active: yes_(row[idx.Active] === '' ? true : row[idx.Active])
    };
    if (pin) record.PIN = pin;
    const site = String(row[idx.Site] || '').trim();
    const role = String(row[idx.Role] || '').trim();
    if (site) record.Site = site;
    if (role) {
      record.Role = role;
      record.Designation = role;
    }
    applyAccessList_(record, row[idx['App Access']]);
    const portalRecord = portalRecordFromEmployeeRecord_(record);
    const result = writePortalUserAndEmployee_(portalRecord, Session.getActiveUser().getEmail(), 'user pin manager update');
    statusCell.setValue((result.portal.updated ? 'Updated' : 'Added') + ' portal row ' + result.portal.row + ' at ' + now_());
    synced++;
  });
  SpreadsheetApp.getUi().alert('User PIN Manager sync complete.\nRows synced: ' + synced + '\nPortal Users also updated.');
  return {ok:true, synced:synced};
}

function promptAddOrUpdateUserPin() {
  const ui = SpreadsheetApp.getUi();
  const idRes = ui.prompt('Add / Update User PIN', 'Employee ID / Login ID (example: fuel-manager)', ui.ButtonSet.OK_CANCEL);
  if (idRes.getSelectedButton() !== ui.Button.OK) return;
  const id = String(idRes.getResponseText() || '').trim();
  if (!id) {
    ui.alert('Employee ID is required.');
    return;
  }

  setupPortalUsersSheet();
  const portalSh = getOrCreate_(SHEET_PORTAL_USERS, PORTAL_USER_HEADERS);
  const existingPortal = rowObjects_(portalSh).find(row => String(row['Login ID'] || '').trim().toUpperCase() === id.toUpperCase()) || null;
  const existing = existingPortal || getEmployeeRecordById_(id);
  const namePrompt = existing
    ? 'Full name (leave blank to keep: ' + (existing['User Name'] || existing['Full Name']) + ')'
    : 'Full name for this new user';
  const nameRes = ui.prompt('User Name', namePrompt, ui.ButtonSet.OK_CANCEL);
  if (nameRes.getSelectedButton() !== ui.Button.OK) return;
  const name = String(nameRes.getResponseText() || '').trim() || (existing ? String(existing['User Name'] || existing['Full Name'] || '').trim() : '');
  if (!name) {
    ui.alert('Full name is required for a new user.');
    return;
  }

  const pinRes = ui.prompt('PIN', 'Enter login PIN. Leave blank to keep existing PIN.', ui.ButtonSet.OK_CANCEL);
  if (pinRes.getSelectedButton() !== ui.Button.OK) return;
  const pin = String(pinRes.getResponseText() || '').trim();
  const finalPin = pin || (existingPortal ? String(existingPortal.PIN || '').trim() : (existing ? String(existing.PIN || '').trim() : ''));
  if (finalPin && hasDuplicatePortalPin_(finalPin, id)) {
    ui.alert('This PIN is already used by another active portal user. Please choose a unique PIN.');
    return;
  }

  const accessRes = ui.prompt(
    'App Access',
    'Enter access names separated by comma: fuel, weighbridge, mrf, usercharge, maintenance, joining. Leave blank to keep existing access.',
    ui.ButtonSet.OK_CANCEL
  );
  if (accessRes.getSelectedButton() !== ui.Button.OK) return;
  const accessText = String(accessRes.getResponseText() || '').trim();

  const siteRes = ui.prompt('Site', 'Site for this user, needed for MRF users. Leave blank to keep existing.', ui.ButtonSet.OK_CANCEL);
  if (siteRes.getSelectedButton() !== ui.Button.OK) return;
  const site = String(siteRes.getResponseText() || '').trim();

  const roleRes = ui.prompt('Role', 'Role/designation text. Leave blank to keep existing.', ui.ButtonSet.OK_CANCEL);
  if (roleRes.getSelectedButton() !== ui.Button.OK) return;
  const role = String(roleRes.getResponseText() || '').trim();

  const portalRecord = existingPortal
    ? Object.assign({}, existingPortal)
    : (existing ? portalRecordFromEmployeeRecord_(existing) : {
      Active: true,
      Admin: false,
      Fuel: false,
      Weighbridge: false,
      'User Charge': false,
      MRF: false,
      'Staff Joining': false,
      Maintenance: false
    });
  portalRecord['Login ID'] = id;
  portalRecord['User Name'] = name;
  portalRecord.PIN = finalPin;
  portalRecord.Active = true;
  if (site) portalRecord.Site = site;
  if (role) portalRecord.Role = role;
  if (accessText) {
    const accessRecord = {};
    applyAccessList_(accessRecord, accessText);
    const mapped = portalRecordFromEmployeeRecord_(Object.assign({
      'Employee ID': id,
      'Full Name': name,
      PIN: finalPin,
      Active: 'YES',
      Site: portalRecord.Site,
      Role: portalRecord.Role
    }, accessRecord));
    portalRecord.Admin = mapped.Admin;
    portalRecord.Fuel = mapped.Fuel;
    portalRecord.Weighbridge = mapped.Weighbridge;
    portalRecord['User Charge'] = mapped['User Charge'];
    portalRecord.MRF = mapped.MRF;
    portalRecord['Staff Joining'] = mapped['Staff Joining'];
    portalRecord.Maintenance = mapped.Maintenance;
  }

  const result = writePortalUserAndEmployee_(portalRecord, Session.getActiveUser().getEmail(), existingPortal ? 'update portal pin/access' : 'create portal user');
  ui.alert((result.portal.updated ? 'Updated' : 'Added') + ' portal user: ' + name + '\nPortal Users row: ' + result.portal.row);
}

function logAudit_(action, entity, key, user, notes) {
  const sh = getOrCreate_(SHEET_AUDIT, AUDIT_HEADERS);
  sh.appendRow([now_(), action, entity, key || '', user || '', notes || '']);
}

function seedEmployees_() {
  const sh = getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['admin', 'Admin', '', '', '', '', 'Admin', 'Operations', 'Admin', 'fatehpur', '4321', '', 'YES', 'YES', 'YES', 'YES', 'YES', 'YES', '', '', '', '', '', '', 'Default admin login'],
    ['fuel-manager', 'Fuel Manager', '', '', '', '', 'Fuel Incharge', 'Operations', 'Fuel Manager', 'fatehpur', '1111', '', 'YES', 'YES', 'NO', 'NO', 'NO', 'NO', '', '', '', '', '', '', 'Default fuel login'],
    ['weighbridge', 'Weighbridge Incharge', '', '', '', '', 'Weighbridge Operator', 'Operations', 'Weighbridge Incharge', 'fatehpur', '2222', '', 'YES', 'NO', 'YES', 'NO', 'NO', 'NO', '', '', '', '', '', '', 'Default weighbridge login'],
    ['uc-collector', 'User Charge Collector', '', '', '', '', 'Collector', 'Operations', 'User Charge Collection', 'fatehpur', '3333', '', 'YES', 'NO', 'NO', 'YES', 'NO', 'NO', '', '', '', '', '', '', 'Default user charge login'],
    ['mrf-supervisor', 'MRF Supervisor', '', '', '', '', 'MRF Supervisor', 'MRF', 'Supervisor', 'fatehpur', '4444', '', 'YES', 'NO', 'NO', 'NO', 'YES', 'NO', '', '', '', '', '', '', 'Default MRF app login'],
    ['vishal', 'Vishal Singh', '', '', '', '', 'MRF Supervisor', 'MRF', 'Supervisor', 'malaka', '789', '', 'YES', 'NO', 'NO', 'YES', 'YES', 'NO', '', '', '', '', '', '', 'MRF default'],
    ['arvind', 'Arvind Kumar Maurya', '', '', '', '', 'MRF Supervisor', 'MRF', 'Supervisor', 'raddiya', '456', '', 'YES', 'NO', 'NO', 'YES', 'YES', 'NO', '', '', '', '', '', '', 'MRF default'],
    ['utpal', 'Utpal Pratap Singh', '', '', '', '', 'MRF Supervisor', 'MRF', 'Supervisor', 'ajgawan', '123', '', 'YES', 'NO', 'NO', 'YES', 'YES', 'NO', '', '', '', '', '', '', 'MRF default'],
    ['abhishek', 'Abhishek Kumar', '', '', '', '', 'MRF Supervisor', 'MRF', 'Supervisor', 'mithnapur', '321', '', 'YES', 'NO', 'NO', 'YES', 'YES', 'NO', '', '', '', '', '', '', 'MRF default']
  ];
  rows.forEach(r => sh.appendRow(r.concat([now_(), now_(), 'NO'])));
}

function seedVehicles_() {
  const sh = getOrCreate_(SHEET_VEHICLES, VEHICLE_HEADERS);
  if (sh.getLastRow() > 1) return;
  const rows = [
    ['UP71CT3457', 'MAGIC', 'Akash Kumar', '', 'W13', 'Raddiya', 'C&T', 'YES', 'YES', 'YES', 'YES', 'YES', 'KMPL', '', 'Default MRF/C&T vehicle'],
    ['UP71CT3458', 'MAGIC', 'Ankit', '', 'W19', 'Murain Tola', 'C&T', 'YES', 'YES', 'YES', 'YES', 'YES', 'KMPL', '', 'Default MRF/C&T vehicle'],
    ['UP71CT3459', 'MAGIC', 'Vanshi Lal', '', 'W16', 'Asti', 'C&T', 'YES', 'YES', 'YES', 'YES', 'YES', 'KMPL', '', 'Default MRF/C&T vehicle'],
    ['UP71CT3460', 'MAGIC', 'Kailash Kumar', '', 'W14', 'Abu Nagar', 'C&T', 'YES', 'YES', 'YES', 'YES', 'YES', 'KMPL', '', 'Default MRF/C&T vehicle'],
    ['UP71CT3461', 'MAGIC', 'Santosh', '', 'W27', 'Collector Ganj', 'C&T', 'YES', 'YES', 'YES', 'YES', 'YES', 'KMPL', '', 'Default MRF/C&T vehicle'],
    ['UP71G0376', 'COMPACTOR', 'Rajesh', '', '', '', 'C&T', 'YES', 'YES', 'YES', 'NO', 'YES', 'HPL', '', 'Default operations vehicle'],
    ['UP71G0377', 'COMPACTOR', 'Rohit Kumar', '', '', '', 'C&T', 'YES', 'YES', 'YES', 'NO', 'YES', 'HPL', '', 'Default operations vehicle'],
    ['ECE07620JR0015917', 'JCB', 'Dilip Kumar', '', '', '', 'C&T', 'YES', 'YES', 'NO', 'NO', 'YES', 'HPL', '', 'Default machinery'],
    ['KBCAZ23ALS3638814', 'POCLAIN', '', '', '', '', 'MRF', 'YES', 'YES', 'NO', 'YES', 'YES', 'HPL', '', 'Default MRF machinery']
  ];
  rows.forEach(r => sh.appendRow(r.concat([now_(), now_()])));
}

function seedSettings_() {
  const sh = getOrCreate_(SHEET_SETTINGS, SETTINGS_HEADERS);
  if (sh.getLastRow() > 1) return;
  sh.appendRow(['version', MASTER_VERSION, 'Master data script version']);
  sh.appendRow(['project', 'Fatehpur MSW', 'Project name']);
  sh.appendRow(['mode', 'development', 'Use development until unified app is fully tested']);
}

function setupMasterSheet() {
  getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  getOrCreate_(SHEET_VEHICLES, VEHICLE_HEADERS);
  getOrCreate_(SHEET_SETTINGS, SETTINGS_HEADERS);
  getOrCreate_(SHEET_AUDIT, AUDIT_HEADERS);
  setupPortalAppsSheet();
  setupPortalUsersSheet();
  seedEmployees_();
  seedVehicles_();
  seedSettings_();
  logAudit_('setup', 'master', '', Session.getActiveUser().getEmail(), 'Master sheet initialized');
  return {
    ok: true,
    version: MASTER_VERSION,
    spreadsheet: getMasterSS_().getName(),
    sheets: getMasterSS_().getSheets().map(s => s.getName())
  };
}

function employeeFromRow_(r) {
  return {
    id: String(r['Employee ID'] || ''),
    name: String(r['Full Name'] || ''),
    mobile: String(r.Mobile || ''),
    fatherName: String(r['Father/Husband Name'] || ''),
    dob: String(r.DOB || ''),
    doj: String(r.DOJ || ''),
    designation: String(r.Designation || ''),
    department: String(r.Department || ''),
    role: String(r.Role || ''),
    site: String(r.Site || ''),
    pin: String(r.PIN || ''),
    whatsapp: String(r.WhatsApp || ''),
    active: bool_(r.Active),
    access: {
      fuel: bool_(r['Fuel Access']),
      weighbridge: bool_(r['Weighbridge Access']),
      mrf: bool_(r['MRF Access']),
      usercharge: bool_(r['User Charge Access']),
      uc: bool_(r['User Charge Access']),
      maintenance: bool_(r['Maintenance Access']),
      joining: bool_(r['Joining Access'])
    }
  };
}

function vehicleFromRow_(r) {
  return {
    no: String(r['Vehicle No'] || ''),
    type: String(r['Vehicle Type'] || ''),
    driver: String(r['Driver Name'] || ''),
    driverMobile: String(r['Driver Mobile'] || ''),
    wardNo: String(r['Ward No'] || ''),
    wardName: String(r['Ward Name'] || ''),
    workingArea: String(r['Working Area'] || ''),
    active: bool_(r.Active),
    enabled: {
      fuel: bool_(r['Fuel Enabled']),
      weighbridge: bool_(r['Weighbridge Enabled']),
      mrf: bool_(r['MRF Enabled']),
      maintenance: bool_(r['Maintenance Enabled'])
    },
    fuelMetric: String(r['Fuel Metric'] || ''),
    fuelType: String(r['Fuel Type'] || ''),
    targetEfficiency: r['Target Efficiency'] || ''
  };
}

function getEmployees_(app) {
  const sh = getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  let list = rowObjects_(sh).map(employeeFromRow_).filter(e => e.id && e.name && e.active);
  if (app) list = list.filter(e => e.access[String(app).toLowerCase()] === true);
  return list;
}

function portalUserFromSheetRow_(r) {
  const admin = bool_(r.Admin);
  return {
    id: String(r['Login ID'] || ''),
    name: String(r['User Name'] || ''),
    pin: String(r.PIN || ''),
    active: bool_(r.Active === '' ? true : r.Active),
    site: String(r.Site || ''),
    role: String(r.Role || ''),
    access: {
      fuel: admin || bool_(r.Fuel),
      weighbridge: admin || bool_(r.Weighbridge),
      usercharge: admin || bool_(r['User Charge']),
      uc: admin || bool_(r['User Charge']),
      mrf: admin || bool_(r.MRF),
      joining: admin || bool_(r['Staff Joining']),
      maintenance: admin || bool_(r.Maintenance)
    }
  };
}

function getPortalUsers_() {
  const ss = getMasterSS_();
  const sh = ss.getSheetByName(SHEET_PORTAL_USERS);
  if (!sh || sh.getLastRow() < 2) return getEmployees_('').filter(e => e.pin);
  const users = rowObjects_(sh)
    .map(portalUserFromSheetRow_)
    .filter(u => u.id && u.name && u.active && u.pin);
  const pinCounts = {};
  users.forEach(u => pinCounts[u.pin] = (pinCounts[u.pin] || 0) + 1);
  return users.filter(u => pinCounts[u.pin] === 1);
}

function getPortalApps_() {
  const ss = getMasterSS_();
  const sh = ss.getSheetByName(SHEET_PORTAL_APPS);
  if (!sh || sh.getLastRow() < 2) {
    return [
      {key:'fuel', name:'Fuel Manager', pinRequired:true},
      {key:'weighbridge', name:'Weighbridge', pinRequired:true},
      {key:'userCharge', name:'User Charge Collection', pinRequired:true},
      {key:'mrf', name:'MRF Supervisor', pinRequired:true},
      {key:'joining', name:'Staff Joining', pinRequired:false},
      {key:'maintenance', name:'Maintenance', pinRequired:false}
    ];
  }
  return rowObjects_(sh)
    .map(r => ({
      key: String(r['App Key'] || '').trim(),
      name: String(r['App Name'] || '').trim(),
      pinRequired: bool_(r['PIN Required'])
    }))
    .filter(app => app.key);
}

function getVehicles_(app) {
  const sh = getOrCreate_(SHEET_VEHICLES, VEHICLE_HEADERS);
  let list = rowObjects_(sh).map(vehicleFromRow_).filter(v => v.no && v.active);
  if (app) list = list.filter(v => v.enabled[String(app).toLowerCase()] === true);
  return list;
}

function upsertEmployee_(data) {
  const record = {};
  record['Employee ID'] = String(data.id || data.employeeId || '').trim();
  if (!record['Employee ID']) throw new Error('Employee ID is required');
  record['Full Name'] = String(data.name || data.fullName || '').trim();
  record.Mobile = String(data.mobile || '');
  record['Father/Husband Name'] = String(data.fatherName || data.fatherHusbandName || '');
  record.DOB = String(data.dob || '');
  record.DOJ = String(data.doj || '');
  record.Designation = String(data.designation || '');
  record.Department = String(data.department || '');
  record.Role = String(data.role || '');
  record.Site = String(data.site || '');
  record.PIN = String(data.pin || '');
  record.WhatsApp = String(data.whatsapp || '');
  record.Active = yes_(data.active === undefined ? true : data.active);
  record['Fuel Access'] = yes_(data.fuelAccess || (data.access && data.access.fuel));
  record['Weighbridge Access'] = yes_(data.weighbridgeAccess || (data.access && data.access.weighbridge));
  record['MRF Access'] = yes_(data.mrfAccess || (data.access && data.access.mrf));
  record['User Charge Access'] = yes_(data.userChargeAccess || data.ucAccess || (data.access && (data.access.usercharge || data.access.uc)));
  record['Maintenance Access'] = yes_(data.maintenanceAccess || (data.access && data.access.maintenance));
  record['Joining Access'] = yes_(data.joiningAccess || (data.access && data.access.joining));
  record['Monthly Salary'] = data.monthlySalary || '';
  record.CTC = data.ctc || '';
  record['Bank Name'] = String(data.bankName || '');
  record['Account Holder Name'] = String(data.accountHolderName || '');
  record['Account Number'] = String(data.accountNumber || '');
  record['IFSC Code'] = String(data.ifscCode || '');
  record.Notes = String(data.notes || '');
  const sh = getOrCreate_(SHEET_EMPLOYEES, EMPLOYEE_HEADERS);
  const result = writeRecord_(sh, EMPLOYEE_HEADERS, 'Employee ID', record);
  logAudit_(result.updated ? 'update' : 'create', 'employee', record['Employee ID'], '', record['Full Name']);
  return result;
}

function upsertVehicle_(data) {
  const record = {};
  record['Vehicle No'] = String(data.no || data.vehicleNo || '').trim().toUpperCase();
  if (!record['Vehicle No']) throw new Error('Vehicle No is required');
  record['Vehicle Type'] = String(data.type || data.vehicleType || '');
  record['Driver Name'] = String(data.driver || data.driverName || '');
  record['Driver Mobile'] = String(data.driverMobile || '');
  record['Ward No'] = String(data.wardNo || '');
  record['Ward Name'] = String(data.wardName || '');
  record['Working Area'] = String(data.workingArea || data.area || '');
  record.Active = yes_(data.active === undefined ? true : data.active);
  record['Fuel Enabled'] = yes_(data.fuelEnabled || (data.enabled && data.enabled.fuel));
  record['Weighbridge Enabled'] = yes_(data.weighbridgeEnabled || (data.enabled && data.enabled.weighbridge));
  record['MRF Enabled'] = yes_(data.mrfEnabled || (data.enabled && data.enabled.mrf));
  record['Maintenance Enabled'] = yes_(data.maintenanceEnabled || (data.enabled && data.enabled.maintenance));
  record['Fuel Metric'] = String(data.fuelMetric || data.metric || '');
  record['Target Efficiency'] = data.targetEfficiency || '';
  record['Fuel Type'] = String(data.fuelType || data.fuel || '').trim().toUpperCase();
  record.Notes = String(data.notes || '');
  const sh = getOrCreate_(SHEET_VEHICLES, VEHICLE_HEADERS);
  const result = writeRecord_(sh, VEHICLE_HEADERS, 'Vehicle No', record);
  logAudit_(result.updated ? 'update' : 'create', 'vehicle', record['Vehicle No'], '', record['Vehicle Type']);
  return result;
}

function doGet(e) {
  e = e || {};
  const p = e.parameter || {};
  const action = String(p.action || 'debug');
  try {
    if (action === 'debug') {
      const ss = getMasterSS_();
      return respond_({ok:true, version:MASTER_VERSION, spreadsheet:ss.getName(), sheets:ss.getSheets().map(s => s.getName())});
    }
    if (action === 'setup') return respond_(setupMasterSheet());
    if (action === 'portalUsers') return respond_({ok:true, users:getPortalUsers_()});
    if (action === 'portalApps') return respond_({ok:true, apps:getPortalApps_()});
    if (action === 'portalConfig') return respond_({ok:true, users:getPortalUsers_(), apps:getPortalApps_()});
    if (action === 'employees') return respond_({ok:true, employees:getEmployees_(p.app || '')});
    if (action === 'vehicles') return respond_({ok:true, vehicles:getVehicles_(p.app || '')});
    if (action === 'config') return respond_({ok:true, employees:getEmployees_(p.app || ''), vehicles:getVehicles_(p.app || '')});
    return respond_({ok:false, error:'Unknown action'});
  } catch (err) {
    return respond_({ok:false, error:err.message || String(err)});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = String(body.action || '');
    if (action === 'setup') return respond_(setupMasterSheet());
    if (action === 'upsertEmployee') return respond_(Object.assign({ok:true}, upsertEmployee_(body.employee || body.data || body)));
    if (action === 'upsertVehicle') return respond_(Object.assign({ok:true}, upsertVehicle_(body.vehicle || body.data || body)));
    if (action === 'bulkUpsertEmployees') {
      const results = (body.employees || []).map(upsertEmployee_);
      return respond_({ok:true, results});
    }
    if (action === 'bulkUpsertVehicles') {
      const results = (body.vehicles || []).map(upsertVehicle_);
      return respond_({ok:true, results});
    }
    return respond_({ok:false, error:'Unknown action'});
  } catch (err) {
    return respond_({ok:false, error:err.message || String(err)});
  }
}
