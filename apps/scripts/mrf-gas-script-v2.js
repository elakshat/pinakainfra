// ══════════════════════════════════════════════════════════════════════
// PINAKA INFRA — MRF SUPERVISOR APP  |  Google Apps Script v2
// ══════════════════════════════════════════════════════════════════════

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

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

function masterMrfVehicles_() {
  const cfg = fetchMasterConfig_('mrf');
  const rows = cfg && Array.isArray(cfg.vehicles) ? cfg.vehicles : [];
  return rows
    .filter(v => v && v.no)
    .map(v => ({
      no: String(v.no || '').trim().toUpperCase(),
      type: String(v.type || '').trim().toUpperCase(),
      wardNo: String(v.wardNo || '').trim(),
      wardName: String(v.wardName || v.workingArea || '').trim(),
      driver: String(v.driver || '').trim(),
      active: 'Y'
    }));
}

function masterMrfUsers_() {
  const cfg = fetchMasterConfig_('mrf');
  const rows = cfg && Array.isArray(cfg.employees) ? cfg.employees : [];
  const users = rows
    .filter(u => u && u.id && u.name && u.pin && u.site)
    .map(u => ({
      id: String(u.id || '').trim(),
      name: String(u.name || '').trim(),
      pin: String(u.pin || '').trim(),
      admin: false,
      site: String(u.site || '').trim()
    }));
  return users.length ? users : null;
}

function getSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); if (headers) sh.appendRow(headers); }
  return sh;
}

function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const keys = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    keys.forEach((k, i) => { obj[k] = normalizeCell(row[i], k); });
    return obj;
  });
}

function normalizeCell(v, key) {
  if (v instanceof Date) {
    const k = String(key || '');
    if (/^(Date)$/.test(k)) return Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd');
    if (/Time|Arrival|Departure|StartTime|EndTime/i.test(k)) return Utilities.formatDate(v, 'Asia/Kolkata', 'HH:mm');
    return Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
  }
  return v !== undefined && v !== null ? String(v) : '';
}

function mapEntry(sheetName, e) {
  if (sheetName === 'Trip Log') return {
    slipNo:e.SlipNo, date:e.Date, time:e.Time, vehicleNo:e.VehicleNo, vehicleType:e.VehicleType,
    wardNo:e.WardNo, wardName:e.WardName, driver:e.Driver, weight:Number(e.Weight)||0,
    arrival:e.Arrival, departure:e.Departure, supervisor:e.Supervisor, site:e.Site||'', synced:true
  };
  if (sheetName === 'Compost Log') return {date:e.Date, wetCollected:Number(e.WetCollected)||0, wetProcessed:Number(e.WetProcessed)||0, method:e.Method, compostGenerated:Number(e.CompostGenerated)||0, utilization:e.Utilization, quantity:e.Quantity, supervisor:e.Supervisor, site:e.Site||'', synced:true};
  if (sheetName === 'Dry Waste Log') return {date:e.Date, received:Number(e.Received)||0, processed:Number(e.Processed)||0, cardboard:Number(e.Cardboard)||0, paper:Number(e.Paper)||0, plastic:Number(e.Plastic)||0, cloth:Number(e.Cloth)||0, glass:Number(e.Glass)||0, packaging:Number(e.Packaging)||0, rubber:Number(e.Rubber)||0, metal:Number(e.Metal)||0, qtyUtilised:e.QtyUtilised, supervisor:e.Supervisor, site:e.Site||'', synced:true};
  if (sheetName === 'Trommel Wet Log') return {date:e.Date, startTime:e.StartTime, endTime:e.EndTime, shutdownHours:Number(e.ShutdownHours)||0, runningHours:Number(e.RunningHours)||0, compostScreened:Number(e.CompostScreened)||0, remarks:e.Remarks, supervisor:e.Supervisor, synced:true};
  if (sheetName === 'Trommel Dry Log') return {date:e.Date, startTime:e.StartTime, endTime:e.EndTime, shutdownHours:Number(e.ShutdownHours)||0, runningHours:Number(e.RunningHours)||0, dryInput:Number(e.DryInput)||0, plantRejected:Number(e.PlantRejected)||0, bioSoil:Number(e.BioSoil)||0, insectsRejection:Number(e.InsectsRejection)||0, remarks:e.Remarks, supervisor:e.Supervisor, synced:true};
  if (sheetName === 'RDF Dispatch Log') return {slipNo:e.SlipNo, date:e.Date, vehicleNo:e.VehicleNo, driverName:e.DriverName, grossWeight:Number(e.GrossWeight)||0, tareWeight:Number(e.TareWeight)||0, netWeight:Number(e.NetWeight)||0, disposalMethod:e.DisposalMethod, receiverName:e.ReceiverName, supervisor:e.Supervisor, synced:true};
  if (sheetName === 'Poclain Log') return {date:e.Date, stationFrom:e.StationFrom, stationTo:e.StationTo, readingStart:Number(e.ReadingStart)||0, readingEnd:Number(e.ReadingEnd)||0, totalHours:Number(e.TotalHours)||0, purpose:e.Purpose, diesel:e.Diesel, mobilOil:e.MobilOil, repairs:e.Repairs, replacement:e.Replacement, driverName:e.DriverName, remarks:e.Remarks, supervisor:e.Supervisor, synced:true};
  if (sheetName === 'CD Waste Log') return {date:e.Date, totalGenerated:Number(e.TotalGenerated)||0, totalCollected:Number(e.TotalCollected)||0, bricks:Number(e.Bricks)||0, concrete:Number(e.Concrete)||0, solid:Number(e.Solid)||0, wood:Number(e.Wood)||0, mortar:Number(e.Mortar)||0, supervisor:e.Supervisor, synced:true};
  return e;
}

// ── SETUP ──────────────────────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getSheet(ss, 'Vehicles', ['No','Type','WardNo','WardName','Driver','Active']);
  getSheet(ss, 'Users', ['ID','Name','PIN','Admin','Site']);
  getSheet(ss, 'Settings', ['Key','Value']);
  getSheet(ss, 'Trip Log', ['SlipNo','Date','Time','VehicleNo','VehicleType','WardNo','WardName','Driver','Weight','Arrival','Departure','Supervisor','Site','LoggedAt']);
  getSheet(ss, 'Compost Log', ['Date','WetCollected','WetProcessed','Method','CompostGenerated','Utilization','Quantity','Supervisor','Site','LoggedAt']);
  getSheet(ss, 'Dry Waste Log', ['Date','Received','Processed','Cardboard','Paper','Plastic','Cloth','Glass','Packaging','Rubber','Metal','QtyUtilised','Supervisor','Site','LoggedAt']);
  getSheet(ss, 'Trommel Wet Log', ['Date','StartTime','EndTime','ShutdownHours','RunningHours','CompostScreened','Remarks','Supervisor','LoggedAt']);
  getSheet(ss, 'Trommel Dry Log', ['Date','StartTime','EndTime','ShutdownHours','RunningHours','DryInput','PlantRejected','BioSoil','InsectsRejection','Remarks','Supervisor','LoggedAt']);
  getSheet(ss, 'RDF Dispatch Log', ['SlipNo','Date','VehicleNo','DriverName','GrossWeight','TareWeight','NetWeight','DisposalMethod','ReceiverName','Supervisor','LoggedAt']);
  getSheet(ss, 'Poclain Log', ['Date','StationFrom','StationTo','ReadingStart','ReadingEnd','TotalHours','Purpose','Diesel','MobilOil','Repairs','Replacement','DriverName','Remarks','Supervisor','LoggedAt']);
  getSheet(ss, 'CD Waste Log', ['Date','TotalGenerated','TotalCollected','Bricks','Concrete','Solid','Wood','Mortar','Supervisor','LoggedAt']);
  
  // Default vehicles
  const vs = ss.getSheetByName('Vehicles');
  if (vs.getLastRow() < 2) {
    const vdata = [
      ['UP71CT3457','MAGIC','W13','Raddiya','Akash Kumar','Y'],
      ['UP71CT3458','MAGIC','W19','Murain Tola','Ankit','Y'],
      ['UP71CT3459','MAGIC','W16','Asti','Vanshi Lal','Y'],
      ['UP71CT3460','MAGIC','W14','Abu Nagar','Kailash Kumar','Y'],
      ['UP71CT3461','MAGIC','W27','Collector Ganj','Santosh','Y'],
      ['UP71CT0439','MAGIC','W15','Maswani','Ashwani','Y'],
      ['UP71CT0440','MAGIC','W2','Awas Vikas','Ashish','Y'],
      ['UP71CT0445','MAGIC','W22','Isainpurwa','Dileep Pal','Y'],
      ['UP71CT0447','MAGIC','W29','Krishna Vihari Nagar','Arun Kumar','Y'],
      ['UP71CT0448','MAGIC','W34','Ahmad Ganj','Harishchandra','Y'],
      ['UP71CT0449','MAGIC','W28','Mahajari','Gaurav','Y'],
      ['UP71CT0450','MAGIC','W24','Amar Jai','Amit Kumar','Y'],
      ['UP71G0379','MAGIC','W21','Shadipur','Imran Ahmad','Y'],
      ['UP71G0384','MAGIC','--','2nd Shift','Sumit','Y'],
    ];
    vdata.forEach(r => vs.appendRow(r));
  }
  
  // Default users
  const us = ss.getSheetByName('Users');
  if (us.getLastRow() < 2) {
    us.appendRow(['vishal','Vishal Singh','789','N','malaka']);
    us.appendRow(['arvind','Arvind Kumar Maurya','456','N','raddiya']);
    us.appendRow(['utpal','Utpal Pratap Singh','123','N','ajgawan']);
    us.appendRow(['abhishek','Abhishek Kumar','321','N','mithnapur']);
    us.appendRow(['admin','Admin','9999','Y','malaka']);
  }
  
  // Default settings
  const st = ss.getSheetByName('Settings');
  if (st.getLastRow() < 2) {
    st.appendRow(['admin_whatsapp','']);
    st.appendRow(['editing_allowed','true']);
  }
  
  return 'Setup complete';
}

// ── GET ────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = (e.parameter && e.parameter.action) || 'test';
    const site = (e.parameter && e.parameter.site) || '';
    
    if (action === 'config') return respond(getConfig(ss));
    if (action === 'trips') return respond(getEntries(ss, 'Trip Log', site));
    if (action === 'compost') return respond(getEntries(ss, 'Compost Log', site));
    if (action === 'drywaste') return respond(getEntries(ss, 'Dry Waste Log', site));
    if (action === 'trommelwet') return respond(getEntries(ss, 'Trommel Wet Log', ''));
    if (action === 'trommeldry') return respond(getEntries(ss, 'Trommel Dry Log', ''));
    if (action === 'rdf') return respond(getEntries(ss, 'RDF Dispatch Log', ''));
    if (action === 'poclain') return respond(getEntries(ss, 'Poclain Log', ''));
    if (action === 'cd') return respond(getEntries(ss, 'CD Waste Log', ''));
    if (action === 'test') return respond(testConn(ss));
    return respond({ok:false, error:'Unknown action'});
  } catch(err) {
    return respond({ok:false, error:err.message});
  }
}

function getConfig(ss) {
  const vsh = ss.getSheetByName('Vehicles');
  const ush = ss.getSheetByName('Users');
  const ssh = ss.getSheetByName('Settings');
  let vehicles = vsh ? sheetToObjects(vsh).filter(v=>v.Active==='Y').map(v=>({no:v.No,type:v.Type,wardNo:v.WardNo,wardName:v.WardName,driver:v.Driver,active:v.Active})) : [];
  let users = ush ? sheetToObjects(ush).map(u=>({id:u.ID,name:u.Name,pin:u.PIN,admin:u.Admin==='Y',site:u.Site})) : [];
  const masterVehicles = masterMrfVehicles_();
  const masterUsers = masterMrfUsers_();
  if (masterVehicles.length) vehicles = masterVehicles;
  if (masterUsers) users = masterUsers;
  const settingsArr = ssh ? sheetToObjects(ssh) : [];
  const settings = {};
  settingsArr.forEach(r => { settings[r.Key] = r.Value; });
  return {ok:true, vehicles, users, settings};
}

function getEntries(ss, sheetName, site) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return {ok:true, entries:[]};
  let entries = sheetToObjects(sh);
  if (site) entries = entries.filter(e => !e.Site || e.Site === site);
  entries = entries.map(e => mapEntry(sheetName, e));
  return {ok:true, entries};
}

function testConn(ss) {
  const sheets = ss.getSheets().map(s=>s.getName());
  return {ok:true, sheets, spreadsheet: ss.getName()};
}

// ── POST ───────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const now = new Date().toISOString();
    
    if (action === 'saveTrip') return respond(saveTrip(ss, data.entry, now));
    if (action === 'saveCompost') return respond(saveCompost(ss, data.entry, now));
    if (action === 'saveDryWaste') return respond(saveDryWaste(ss, data.entry, now));
    if (action === 'saveTrommelWet') return respond(saveTrommelWet(ss, data.entry, now));
    if (action === 'saveTrommelDry') return respond(saveTrommelDry(ss, data.entry, now));
    if (action === 'saveRDF') return respond(saveRDF(ss, data.entry, now));
    if (action === 'savePoclain') return respond(savePoclain(ss, data.entry, now));
    if (action === 'saveCD') return respond(saveCD(ss, data.entry, now));
    if (action === 'saveVehicle') return respond(saveVehicle(ss, data.vehicle, data.editNo));
    if (action === 'deleteVehicle') return respond(deleteVehicle(ss, data.no));
    if (action === 'saveUser') return respond(saveUser(ss, data.user, data.editId));
    if (action === 'deleteUser') return respond(deleteUser(ss, data.id));
    if (action === 'saveSettings') return respond(saveSettings(ss, data.settings));
    if (action === 'editEntry') return respond(editEntry(ss, data));
    if (action === 'setup') { setupSheets(); return respond({ok:true, msg:'Setup complete'}); }
    return respond({ok:false, error:'Unknown action'});
  } catch(err) {
    return respond({ok:false, error:err.message});
  }
}

function saveTrip(ss, e, now) {
  const sh = getSheet(ss, 'Trip Log');
  // Check for edit
  if (e._editSlip) {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(e._editSlip)) {
        sh.getRange(i+1, 1, 1, 14).setValues([[
          e.slipNo||e._editSlip, e.date, e.time, e.vehicleNo, e.vehicleType,
          e.wardNo, e.wardName, e.driver, e.weight, e.arrival, e.departure,
          e.supervisor, e.site||'', now
        ]]);
        return {ok:true};
      }
    }
  }
  // Count today's slips
  const todayData = sh.getDataRange().getValues();
  const today = e.date;
  const todayCount = todayData.slice(1).filter(r=>normalizeCell(r[1], 'Date')===today).length;
  const slipNo = 'MRF-'+today.slice(2).replace(/-/g,'')+'-'+String(todayCount+1).padStart(3,'0');
  sh.appendRow([slipNo, e.date, e.time, e.vehicleNo, e.vehicleType,
    e.wardNo, e.wardName, e.driver, e.weight, e.arrival, e.departure,
    e.supervisor, e.site||'', now]);
  return {ok:true, slipNo};
}

function saveCompost(ss, e, now) {
  const sh = getSheet(ss, 'Compost Log');
  // Remove existing entry for same date+site
  deleteRowByDateSite(sh, e.date, e.site||'', 0, 8);
  sh.appendRow([e.date, e.wetCollected, e.wetProcessed, e.method,
    e.compostGenerated, e.utilization, e.quantity||'', e.supervisor, e.site||'', now]);
  return {ok:true};
}

function saveDryWaste(ss, e, now) {
  const sh = getSheet(ss, 'Dry Waste Log');
  deleteRowByDateSite(sh, e.date, e.site||'', 0, 12);
  sh.appendRow([e.date, e.received, e.processed, e.cardboard, e.paper,
    e.plastic, e.cloth, e.glass, e.packaging, e.rubber, e.metal,
    e.qtyUtilised, e.supervisor, e.site||'', now]);
  return {ok:true};
}

function saveTrommelWet(ss, e, now) {
  const sh = getSheet(ss, 'Trommel Wet Log');
  deleteRowByDate(sh, e.date, 0);
  sh.appendRow([e.date, e.startTime, e.endTime, e.shutdownHours,
    e.runningHours, e.compostScreened, e.remarks||'', e.supervisor, now]);
  return {ok:true};
}

function saveTrommelDry(ss, e, now) {
  const sh = getSheet(ss, 'Trommel Dry Log');
  deleteRowByDate(sh, e.date, 0);
  sh.appendRow([e.date, e.startTime, e.endTime, e.shutdownHours,
    e.runningHours, e.dryInput, e.plantRejected, e.bioSoil,
    e.insectsRejection, e.remarks||'', e.supervisor, now]);
  return {ok:true};
}

function saveRDF(ss, e, now) {
  const sh = getSheet(ss, 'RDF Dispatch Log');
  if (e.slipNo) {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(e.slipNo)) {
        sh.getRange(i+1, 1, 1, 11).setValues([[e.slipNo, e.date, e.vehicleNo, e.driverName, e.grossWeight, e.tareWeight, e.netWeight, e.disposalMethod||'', e.receiverName||'', e.supervisor, now]]);
        return {ok:true, slipNo:e.slipNo};
      }
    }
  }
  const rows = sh.getLastRow() - 1;
  const slipNo = e.slipNo || 'RDF-'+e.date.slice(2).replace(/-/g,'')+'-'+String(rows+1).padStart(3,'0');
  sh.appendRow([slipNo, e.date, e.vehicleNo, e.driverName,
    e.grossWeight, e.tareWeight, e.netWeight,
    e.disposalMethod||'', e.receiverName||'', e.supervisor, now]);
  return {ok:true, slipNo};
}

function savePoclain(ss, e, now) {
  const sh = getSheet(ss, 'Poclain Log');
  sh.appendRow([e.date, e.stationFrom||'', e.stationTo||'',
    e.readingStart, e.readingEnd, e.totalHours,
    e.purpose||'Waste Sifting', e.diesel||'',
    e.mobilOil||'', e.repairs||'', e.replacement||'',
    e.driverName||'', e.remarks||'', e.supervisor, now]);
  return {ok:true};
}

function saveCD(ss, e, now) {
  const sh = getSheet(ss, 'CD Waste Log');
  deleteRowByDate(sh, e.date, 0);
  sh.appendRow([e.date, e.totalGenerated, e.totalCollected,
    e.bricks||0, e.concrete||0, e.solid||0, e.wood||0, e.mortar||0,
    e.supervisor, now]);
  return {ok:true};
}

function deleteRowByDate(sh, date, dateCol) {
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][dateCol]) === String(date)) sh.deleteRow(i+1);
  }
}

function deleteRowByDateSite(sh, date, site, dateCol, siteCol) {
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][dateCol]) === String(date) && String(data[i][siteCol]) === String(site)) {
      sh.deleteRow(i+1);
    }
  }
}

function editEntry(ss, data) {
  const {sheet, keyField, keyValue, fields} = data;
  const sh = ss.getSheetByName(sheet);
  if (!sh) return {ok:false, error:'Sheet not found'};
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const keyIdx = headers.indexOf(keyField);
  if (keyIdx < 0) return {ok:false, error:'Key field not found'};
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][keyIdx]) === String(keyValue)) {
      Object.entries(fields).forEach(([k, v]) => {
        const ci = headers.indexOf(k);
        if (ci >= 0) sh.getRange(i+1, ci+1).setValue(v);
      });
      return {ok:true};
    }
  }
  return {ok:false, error:'Entry not found'};
}

function saveVehicle(ss, v, editNo) {
  const sh = getSheet(ss, 'Vehicles');
  if (editNo) {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(editNo)) {
        sh.getRange(i+1,1,1,6).setValues([[v.no,v.type,v.wardNo,v.wardName,v.driver,v.active]]);
        return {ok:true};
      }
    }
  }
  sh.appendRow([v.no,v.type,v.wardNo,v.wardName,v.driver,v.active||'Y']);
  return {ok:true};
}

function deleteVehicle(ss, no) {
  const sh = ss.getSheetByName('Vehicles'); if (!sh) return {ok:false};
  const data = sh.getDataRange().getValues();
  for (let i = data.length-1; i>=1; i--) {
    if (String(data[i][0])===String(no)) { sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:'Not found'};
}

function saveUser(ss, u, editId) {
  const sh = getSheet(ss, 'Users');
  if (editId) {
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(editId)) {
        sh.getRange(i+1,1,1,5).setValues([[u.id,u.name,u.pin,u.admin?'Y':'N',u.site||'']]);
        return {ok:true};
      }
    }
  }
  sh.appendRow([u.id,u.name,u.pin,u.admin?'Y':'N',u.site||'']);
  return {ok:true};
}

function deleteUser(ss, id) {
  const sh = ss.getSheetByName('Users'); if (!sh) return {ok:false};
  const data = sh.getDataRange().getValues();
  for (let i = data.length-1; i>=1; i--) {
    if (String(data[i][0])===String(id)) { sh.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:'Not found'};
}

function saveSettings(ss, settings) {
  const sh = getSheet(ss, 'Settings');
  const data = sh.getDataRange().getValues();
  Object.entries(settings).forEach(([k,v]) => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(k)) {
        sh.getRange(i+1,2).setValue(v); found = true; break;
      }
    }
    if (!found) sh.appendRow([k,v]);
  });
  return {ok:true};
}
