// MSW Complaints Google Sheets backend
// Deploy as: Apps Script > Deploy > Web app > Execute as Me > Anyone.
// Put the deployed /exec URL into COMPLAINTS_SCRIPT_URL on the Node host.

const COMPLAINTS_SHEET_NAME = 'Complaints';
const SETTINGS_SHEET_NAME = 'Settings';
const COMPLAINT_HEADERS = [
  'Complaint ID',
  'Created At',
  'Updated At',
  'Name',
  'Mobile',
  'Ward',
  'Type',
  'Description',
  'Status',
  'Assigned Officer',
  'Admin Remark',
  'Priority',
  'Created By UID',
  'Created By Phone'
];

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || 'ping');
  try {
    if (action === 'ping') return jsonOut_({ ok: true, msg: 'MSW Complaints API', version: 1 });
    if (action === 'setup') return jsonOut_(setupComplaintsSheet_());
    if (action === 'list') return jsonOut_({ ok: true, complaints: listComplaints_(params) });
    return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '');
    if (action === 'createComplaint') return jsonOut_({ ok: true, complaint: createComplaint_(body.complaint || body) });
    if (action === 'updateComplaint') return jsonOut_({ ok: true, complaint: updateComplaint_(body.complaintId, body.patch || {}) });
    if (action === 'setup') return jsonOut_(setupComplaintsSheet_());
    return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: err.message || String(err) });
  }
}

function setupComplaintsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the target Google Sheet, or bind it to the Sheet.');
  const sh = ensureComplaintsSheet_(ss);
  ensureSettingsSheet_(ss);
  return { ok: true, spreadsheet: ss.getName(), sheet: sh.getName(), headers: COMPLAINT_HEADERS.length };
}

function ensureComplaintsSheet_(ss) {
  let sh = ss.getSheetByName(COMPLAINTS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(COMPLAINTS_SHEET_NAME);
  const range = sh.getRange(1, 1, 1, COMPLAINT_HEADERS.length);
  const existing = range.getValues()[0];
  const needsHeader = COMPLAINT_HEADERS.some((h, i) => existing[i] !== h);
  if (needsHeader) {
    range.setValues([COMPLAINT_HEADERS]);
    range.setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureSettingsSheet_(ss) {
  let sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET_NAME);
    sh.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setFontWeight('bold');
    sh.getRange(2, 1, 3, 2).setValues([
      ['admin_mobile', '7052000032'],
      ['wards', '20'],
      ['officers', 'Ramesh Sharma|Ward 7; Suresh Verma|Ward 3; Priya Singh|Ward 1']
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function createComplaint_(input) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureComplaintsSheet_(ss);
  const now = new Date().toISOString();
  const record = normalizeComplaint_(Object.assign({}, input, {
    complaintId: input.complaintId || nextComplaintId_(),
    status: input.status || 'Open',
    assignedOfficer: input.assignedOfficer || '',
    adminRemark: input.adminRemark || '',
    priority: input.priority || 'Medium',
    createdAt: input.createdAt || now,
    updatedAt: now
  }));
  sh.appendRow(recordToRow_(record));
  return record;
}

function updateComplaint_(complaintId, patch) {
  if (!complaintId) throw new Error('Missing complaintId');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureComplaintsSheet_(ss);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(complaintId)) {
      const existing = rowToRecord_(rows[i]);
      const updated = normalizeComplaint_(Object.assign({}, existing, patch, {
        complaintId: existing.complaintId,
        updatedAt: new Date().toISOString()
      }));
      sh.getRange(i + 1, 1, 1, COMPLAINT_HEADERS.length).setValues([recordToRow_(updated)]);
      return updated;
    }
  }
  throw new Error('Complaint not found: ' + complaintId);
}

function listComplaints_(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureComplaintsSheet_(ss);
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COMPLAINT_HEADERS.length).getValues();
  return rows
    .filter(row => String(row[0] || '').trim())
    .map(rowToRecord_)
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));
}

function normalizeComplaint_(record) {
  return {
    complaintId: String(record.complaintId || nextComplaintId_()).trim(),
    name: String(record.name || '').trim(),
    mobile: String(record.mobile || '').replace(/\D/g, '').slice(0, 10),
    ward: String(record.ward || '').trim(),
    type: String(record.type || '').trim(),
    description: String(record.description || '').trim(),
    status: String(record.status || 'Open').trim(),
    assignedOfficer: String(record.assignedOfficer || '').trim(),
    adminRemark: String(record.adminRemark || '').trim(),
    priority: String(record.priority || 'Medium').trim(),
    createdAt: toIso_(record.createdAt || new Date()),
    updatedAt: toIso_(record.updatedAt || new Date()),
    createdByUid: String(record.createdByUid || '').trim(),
    createdByPhone: String(record.createdByPhone || '').trim()
  };
}

function recordToRow_(r) {
  return [
    r.complaintId,
    r.createdAt,
    r.updatedAt,
    r.name,
    r.mobile,
    r.ward,
    r.type,
    r.description,
    r.status,
    r.assignedOfficer,
    r.adminRemark,
    r.priority,
    r.createdByUid,
    r.createdByPhone
  ];
}

function rowToRecord_(row) {
  return normalizeComplaint_({
    complaintId: row[0],
    createdAt: row[1],
    updatedAt: row[2],
    name: row[3],
    mobile: row[4],
    ward: row[5],
    type: row[6],
    description: row[7],
    status: row[8],
    assignedOfficer: row[9],
    adminRemark: row[10],
    priority: row[11],
    createdByUid: row[12],
    createdByPhone: row[13]
  });
}

function nextComplaintId_() {
  return 'CMP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(100 + Math.random() * 900);
}

function toIso_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return value.toISOString();
  const text = String(value || '').trim();
  if (!text) return new Date().toISOString();
  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? text : parsed.toISOString();
}

function jsonOut_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
