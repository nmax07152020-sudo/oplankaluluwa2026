/**
 * OPLAN KALULUWA 2026
 * ANTIPOLO CITY HEALTH OFFICE
 *
 * Google Sheets = Database
 * Apps Script = Backend / API / Authentication
 * GitHub + Vercel = Frontend
 *
 * V12 FEATURES
 * - Public dashboard (no login)
 * - TEAM_LEADER login, locked to own deployment unit
 * - ADMIN login, full management
 * - Admin CRUD for personnel/duty assignment slots
 * - Admin user account management
 * - Reports per unit
 * - Separate PATIENTS tab
 * - Dashboard totals
 * - Session token auth via CacheService
 *
 * IMPORTANT:
 * Run upgradeOplanDatabaseV12() once after replacing your Code.gs.
 * It preserves REPORTS and PATIENTS records and does not wipe DUTY_INPUT
 * if it already contains data.
 */

const CFG = {
  TITLE1: 'ANTIPOLO CITY HEALTH OFFICE',
  TITLE2: 'SCHEDULE OF STANDBY MEDIC',
  TITLE3: 'OPLAN KALULUWA YEAR 2026',
  COVERAGE: 'OCTOBER 31, 2026 - NOVEMBER 2, 2026',
  TIMEZONE: Session.getScriptTimeZone() || 'Asia/Manila',

  MEDICS_PER_SHIFT: 3,

  SHEETS: {
    settings:'SETTINGS',
    cemeteries:'CEMETERIES',
    shifts:'SHIFTS',
    input:'DUTY_INPUT',
    personnel:'PERSONNEL',
    duty:'DUTY_SCHEDULE',
    users:'USERS',
    reports:'REPORTS',
    patients:'PATIENTS',
    summary:'UNIT_REPORT_SUMMARY'
  },

  SESSION_TTL_SECONDS: 21600
};

const UNITS = [{"group": 1, "name": "Himlayang Katoliko / Sta Elena Memorial Park / Antipolo Public Cemetery 1", "covered": ["Himlayang Katoliko", "Sta Elena Memorial Park", "Antipolo Public Cemetery 1 (Besides Sta. Elena Memorial Park)"]}, {"group": 2, "name": "Our Lady of Peace Crematorium and Columbarium 1 / Antipolo Public Cemetery 2 / Antipolo Public Cemetery 3", "covered": ["Our Lady of Peace Crematorium and Columbarium 1", "Antipolo Public Cemetery 2", "Antipolo Public Cemetery 3 (Infront of Antipolo Public Cemetery 2)"]}, {"group": 3, "name": "Our Lady of Peace Memorial Park 2 / Our Lady of Peace Memorial Park 3", "covered": ["Our Lady of Peace Memorial Park 2", "Our Lady of Peace Memorial Park 3 (Infront of OLPMP 2)"]}, {"group": 4, "name": "Transfiguration of Christ Parish Church", "covered": ["Transfiguration of Christ Parish Church"]}, {"group": 5, "name": "Pantay Public Cemetery", "covered": ["Pantay Public Cemetery"]}, {"group": 6, "name": "Boso-Boso Public Cemetery", "covered": ["Boso-Boso Public Cemetery"]}, {"group": 7, "name": "Heaven’s Gate 2 Memorial Park", "covered": ["Heaven’s Gate 2 Memorial Park"]}, {"group": 8, "name": "Loyola Gardens of Antipolo", "covered": ["Loyola Gardens of Antipolo"]}, {"group": 9, "name": "Providence Memorial Park", "covered": ["Providence Memorial Park"]}, {"group": 10, "name": "Sto. Rosario Memorial Park", "covered": ["Sto. Rosario Memorial Park"]}, {"group": 11, "name": "Heaven’s Gate 1 Memorial Park", "covered": ["Heaven’s Gate 1 Memorial Park"]}, {"group": 12, "name": "Gethsemane Memorial Park", "covered": ["Gethsemane Memorial Park"]}, {"group": 13, "name": "Haven of Angels Memorial and Crematorium", "covered": ["Haven of Angels Memorial and Crematorium"]}, {"group": 14, "name": "Valley of Sympathy Memorial Park", "covered": ["Valley of Sympathy Memorial Park"]}];
const SHIFT_CONFIG = [{"id": "S1", "date": "October 31, 2026", "period": "2:00 PM - 10:00 PM", "coverage": "Oct. 31"}, {"id": "S2", "date": "Oct. 31 - Nov. 1, 2026", "period": "10:00 PM - 6:00 AM", "coverage": "Oct. 31 - Nov. 1"}, {"id": "S3", "date": "November 1, 2026", "period": "6:00 AM - 2:00 PM", "coverage": "Nov. 1"}, {"id": "S4", "date": "November 1, 2026", "period": "2:00 PM - 10:00 PM", "coverage": "Nov. 1"}, {"id": "S5", "date": "Nov. 1 - 2, 2026", "period": "10:00 PM - 6:00 AM", "coverage": "Nov. 1 - 2"}];
const REPORT_TYPES = ["Vital Signs", "Medicine Request", "First Aid", "Medical Assistance", "Emergency Response", "Lost Person", "Traffic Assistance", "Security Incident", "Other"];

// Fixed Team Leader login passwords, one password per deployment unit.
// Team Leader login is: select unit + password (no username).
const DEFAULT_UNIT_PASSWORDS = {
  '1':'chounit1',
  '2':'chounit2',
  '3':'chounit3',
  '4':'chounit4',
  '5':'chounit5',
  '6':'chounit6',
  '7':'chounit7',
  '8':'chounit8',
  '9':'chounit9',
  '10':'chounit10',
  '11':'chounit11',
  '12':'chounit12',
  '13':'chounit13',
  '14':'chounit14'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('OPLAN KALULUWA')
    .addItem('Upgrade / Setup V12', 'upgradeOplanDatabaseV12')
    .addItem('Sync Duty Input', 'syncDutyInput')
    .addItem('Rebuild Report Summary', 'buildUnitReportSummary')
    .addSeparator()
    .addItem('Show Default Admin Login', 'showDefaultAdmin')
    .addToUi();
}

/**
 * V12 setup/migration.
 * Does not clear REPORTS, PATIENTS, or existing DUTY_INPUT data.
 */
function upgradeOplanDatabaseV12() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, CFG.SHEETS.settings, [
    ['Key','Value'],
    ['SYSTEM_TITLE_1',CFG.TITLE1],
    ['SYSTEM_TITLE_2',CFG.TITLE2],
    ['SYSTEM_TITLE_3',CFG.TITLE3],
    ['COVERAGE',CFG.COVERAGE],
    ['TIMEZONE',CFG.TIMEZONE],
    ['TOTAL_DEPLOYMENT_UNITS',UNITS.length],
    ['MEDICS_PER_SHIFT',CFG.MEDICS_PER_SHIFT],
    ['SHIFT_BLOCKS',SHIFT_CONFIG.length],
    ['TOTAL_DUTY_SLOTS',UNITS.length * SHIFT_CONFIG.length * CFG.MEDICS_PER_SHIFT]
  ]);

  ensureCemeteries_(ss);
  ensureShifts_(ss);
  ensureDutyInput_(ss);

  ensureSheet_(ss, CFG.SHEETS.personnel, [
    ['Personnel ID','Full Name','Role','Cemetery / Post','Team Leader','Status','Assignment Count','Contact Number']
  ]);

  ensureSheet_(ss, CFG.SHEETS.duty, [
    ['Duty ID','Cemetery / Post','Group','Team Leader','Date / Coverage','Shift ID','Shift / Time','Personnel ID','Personnel Name','Role','Duty Status','Last Sync']
  ]);

  ensureSheet_(ss, CFG.SHEETS.users, [
    ['User ID','Username','Password Hash','Role','Cemetery / Post','Personnel ID','Full Name','Status','Created At','Updated At']
  ]);

  ensureReports_(ss);
  ensureSheet_(ss, CFG.SHEETS.patients, [
    ['Patient ID','Created At','Full Name','Age','Gender','Address','Contact Number','Linked Report ID','Cemetery / Post','Recorded By','Notes']
  ]);

  ensureSheet_(ss, CFG.SHEETS.summary, [
    ['Cemetery / Post','Team Leader','Total Reports','Total Persons','Vital Signs','Medicine Request','First Aid','Medical Assistance','Emergency Response','Other']
  ]);

  // Add a default admin account only when no active ADMIN exists.
  ensureDefaultAdmin_(ss);
  ensureDefaultTeamLeaders_();

  syncDutyInput();
  buildUnitReportSummary();
  formatSheets_();

  SpreadsheetApp.getUi().alert(
    'Oplan Kaluluwa V12 is ready.\\n\\n' +
    'Default Admin login:\\n' +
    'Username: admin\\n' +
    'Password: admin123\\n\\n' +
    'Change the Admin password from the Admin > User Access page.'
  );
}

function showDefaultAdmin() {
  SpreadsheetApp.getUi().alert(
    'Default Admin login\\n\\nUsername: admin\\nPassword: admin123\\n\\n' +
    'Change this password after first login.'
  );
}

function ensureSheet_(ss, name, headerRows) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }

  const header = headerRows[0];
  const current = sh.getRange(1,1,1,Math.max(header.length, sh.getLastColumn() || header.length)).getValues()[0];
  const same = header.every((h,i) => String(current[i] || '').trim() === h);

  if (!same) {
    // For configuration / empty sheets, write the new header.
    // Existing data rows are left intact when possible.
    sh.getRange(1,1,1,header.length).setValues([header]);
  }

  sh.setFrozenRows(1);
  return sh;
}

function ensureCemeteries_(ss) {
  const sh = ensureSheet_(ss, CFG.SHEETS.cemeteries, [
    ['Cemetery ID','Group','Cemetery / Post','Covered Locations','Team Leader','Status']
  ]);

  if (sh.getLastRow() < 2) {
    const rows = UNITS.map((u,i) => [
      'CEM' + String(i+1).padStart(3,'0'),
      u.group,
      u.name,
      (u.covered || []).join(' | '),
      '',
      'ACTIVE'
    ]);
    sh.getRange(2,1,rows.length,6).setValues(rows);
  }
}

function ensureShifts_(ss) {
  const sh = ensureSheet_(ss, CFG.SHEETS.shifts, [
    ['Shift ID','Shift','Date / Coverage','Start','End','Required Medics Per Location']
  ]);

  if (sh.getLastRow() < 2) {
    const rows = SHIFT_CONFIG.map(x => {
      const parts = x.period.split(' - ');
      return [x.id,x.period,x.date,parts[0],parts[1],3];
    });
    sh.getRange(2,1,rows.length,6).setValues(rows);
  }
}

function ensureDutyInput_(ss) {
  const sh = ensureSheet_(ss, CFG.SHEETS.input, [
    ['Group','Cemetery / Post','Team Leader','Shift ID','Shift / Time','Date / Coverage','Staff 1','Staff 2','Staff 3']
  ]);

  // Do not wipe existing assignments.
  if (sh.getLastRow() < 2) {
    const rows = [];
    UNITS.forEach(u => {
      SHIFT_CONFIG.forEach(s => {
        rows.push([u.group,u.name,'',s.id,s.period,s.date,'','','']);
      });
    });
    sh.getRange(2,1,rows.length,9).setValues(rows);
  }
}

function ensureReports_(ss) {
  ensureSheet_(ss, CFG.SHEETS.reports, [
    ['Report ID','Timestamp','Report Date','Report Time','Cemetery / Post','Team Leader','Staff / Reporter','Report Type','Patient / Person Count','Patient ID','Status','Details','Action Taken']
  ]);
}

function ensureDefaultAdmin_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.users);
  const rows = getObjects_(sh);
  const existingAdmin = rows.find(r =>
    String(r['Role']).toUpperCase() === 'ADMIN' &&
    String(r['Status']).toUpperCase() === 'ACTIVE'
  );
  if (existingAdmin) return;

  const now = new Date();
  sh.appendRow([
    'USR-ADMIN',
    'admin',
    hashPassword_('admin123'),
    'ADMIN',
    '',
    '',
    'System Administrator',
    'ACTIVE',
    now,
    now
  ]);
}

function setupContentService_() {
  return ContentService;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalize_(value) {
  return String(value == null ? '' : value).trim();
}

function hashPassword_(password) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );
  return raw.map(b => {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
}

function createSession_(user) {
  const token = randomToken_();
  CacheService.getScriptCache().put(
    'OPLAN_SESSION_' + token,
    JSON.stringify({
      userId:user['User ID'],
      username:user['Username'],
      role:String(user['Role']).toUpperCase(),
      unit:user['Cemetery / Post'] || '',
      personnelId:user['Personnel ID'] || '',
      fullName:user['Full Name'] || ''
    }),
    CFG.SESSION_TTL_SECONDS
  );
  return token;
}

function getSession_(token) {
  const t = normalize_(token);
  if (!t) return null;
  const raw = CacheService.getScriptCache().get('OPLAN_SESSION_' + t);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function requireSession_(token) {
  const session = getSession_(token);
  if (!session) throw new Error('Session expired. Please login again.');
  return session;
}

function requireRole_(token, roles) {
  const session = requireSession_(token);
  const allowed = roles.map(x=>String(x).toUpperCase());
  if (allowed.indexOf(String(session.role).toUpperCase()) === -1) {
    throw new Error('Access denied.');
  }
  return session;
}

function userPayload_(user) {
  return {
    userId:user['User ID'],
    username:user['Username'] || '',
    role:String(user['Role']).toUpperCase(),
    unit:user['Cemetery / Post'] || '',
    personnelId:user['Personnel ID'] || '',
    fullName:user['Full Name'] || ''
  };
}

function authenticateAdmin_(username, password) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.users);
  const users = getObjects_(sh);
  const u = normalize_(username).toLowerCase();
  const p = hashPassword_(password);

  const user = users.find(row =>
    String(row['Role']).toUpperCase() === 'ADMIN' &&
    normalize_(row['Username']).toLowerCase() === u &&
    normalize_(row['Password Hash']) === p &&
    normalize_(row['Status']).toUpperCase() === 'ACTIVE'
  );

  if (!user) throw new Error('Invalid Admin username or password.');

  return {
    ok:true,
    token:createSession_(user),
    user:userPayload_(user)
  };
}

function authenticateTeamLeader_(unit, password) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.users);
  const users = getObjects_(sh);
  const selectedUnit = normalize_(unit);
  const p = hashPassword_(password);

  if (!selectedUnit) throw new Error('Please select your deployment unit.');

  const user = users.find(row =>
    String(row['Role']).toUpperCase() === 'TEAM_LEADER' &&
    normalize_(row['Cemetery / Post']) === selectedUnit &&
    normalize_(row['Password Hash']) === p &&
    normalize_(row['Status']).toUpperCase() === 'ACTIVE'
  );

  if (!user) throw new Error('Invalid unit password.');

  return {
    ok:true,
    token:createSession_(user),
    user:userPayload_(user)
  };
}

function doGet(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const action = normalize_(p.action);

    if (action === 'publicDashboard' || !action) {
      return json_(getPublicDashboard_());
    }

    if (action === 'unitsForLogin') {
      return json_({
        ok:true,
        units:UNITS.map(u=>({group:u.group,name:u.name}))
      });
    }

    if (action === 'unitData') {
      return json_(getUnitData_(p.token, p.unit));
    }

    if (action === 'adminData') {
      return json_(getAdminData_(p.token));
    }

    if (action === 'me') {
      return json_({ok:true,user:requireSession_(p.token)});
    }

    throw new Error('Unknown GET action.');
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = normalize_(body.action);
    const data = body.data || {};

    if (action === 'loginAdmin') {
      return json_(authenticateAdmin_(data.username, data.password));
    }

    if (action === 'loginTeamLeader') {
      return json_(authenticateTeamLeader_(data.unit, data.password));
    }

    // Backward compatibility for the previous login action.
    if (action === 'login') {
      const role = normalize_(data.role || 'TEAM_LEADER').toUpperCase();
      if (role === 'ADMIN') {
        return json_(authenticateAdmin_(data.username, data.password));
      }
      return json_(authenticateTeamLeader_(data.unit, data.password));
    }

    if (action === 'logout') {
      const token = normalize_(body.token);
      if (token) CacheService.getScriptCache().remove('OPLAN_SESSION_' + token);
      return json_({ok:true});
    }

    if (action === 'submitReport') {
      const session = requireRole_(body.token, ['ADMIN','TEAM_LEADER']);
      const clean = Object.assign({}, data);

      // Team Leader cannot choose another unit.
      if (session.role === 'TEAM_LEADER') {
        clean.cemetery = session.unit;
        clean.teamLeader = session.fullName;
        clean.reporter = session.fullName;
      }

      const result = saveReport_(clean);
      buildUnitReportSummary();
      return json_(result);
    }

    if (action === 'adminSaveAssignment') {
      requireRole_(body.token, ['ADMIN']);
      return json_(adminSaveAssignment_(data));
    }

    if (action === 'adminDeleteAssignment') {
      requireRole_(body.token, ['ADMIN']);
      return json_(adminDeleteAssignment_(data));
    }

    if (action === 'adminSaveUser') {
      requireRole_(body.token, ['ADMIN']);
      return json_(adminSaveUser_(data));
    }

    if (action === 'adminDeleteUser') {
      requireRole_(body.token, ['ADMIN']);
      return json_(adminDeleteUser_(data));
    }

    if (action === 'teamLeaderCredentials') {
      return json_(getTeamLeaderCredentials_(body.token));
    }

    if (action === 'syncDuty') {
      requireRole_(body.token, ['ADMIN']);
      syncDutyInput();
      buildUnitReportSummary();
      return json_({ok:true});
    }

    throw new Error('Unknown POST action.');
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}

/* =========================
   PUBLIC DATA
   ========================= */

function getPublicDashboard_() {
  const summary = getObjects_(SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.summary));
  const reports = getObjects_(SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.reports));
  const duty = getObjects_(SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.duty));

  return {
    ok:true,
    title1:CFG.TITLE1,
    title2:CFG.TITLE2,
    title3:CFG.TITLE3,
    coverage:CFG.COVERAGE,
    totalUnits:UNITS.length,
    totalPersonnel:unique_(duty.map(x=>x['Personnel Name'])).length,
    totalReports:reports.length,
    totalPersons:reports.reduce((a,r)=>a+Number(r['Patient / Person Count'] || 0),0),
    units:summary.map(x=>({
      cemetery:x['Cemetery / Post'],
      totalReports:Number(x['Total Reports'] || 0),
      totalPersons:Number(x['Total Persons'] || 0)
    }))
  };
}

/* =========================
   UNIT / ADMIN DATA
   ========================= */

function getUnitData_(token, requestedUnit) {
  const session = requireRole_(token, ['ADMIN','TEAM_LEADER']);
  let unit = normalize_(requestedUnit);

  if (session.role === 'TEAM_LEADER') unit = session.unit;
  if (!unit) unit = session.unit || UNITS[0].name;

  const duty = getObjects_(SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.duty))
    .filter(x=>normalize_(x['Cemetery / Post']) === unit);

  const reports = getObjects_(SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.reports))
    .filter(x=>normalize_(x['Cemetery / Post']) === unit);

  const cemetery = getObjects_(SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.cemeteries))
    .find(x=>normalize_(x['Cemetery / Post']) === unit) || {};

  return {
    ok:true,
    session:session,
    unit:{
      name:unit,
      group:cemetery['Group'] || '',
      covered:normalize_(cemetery['Covered Locations']).split(' | ').filter(Boolean),
      teamLeader:cemetery['Team Leader'] || '',
      status:cemetery['Status'] || 'ACTIVE'
    },
    duty:duty,
    reports:reports.map(publicReportForAuthorized_),
    reportTotals:reportTotals_(reports)
  };
}

function getAdminData_(token) {
  requireRole_(token, ['ADMIN']);

  const ss = SpreadsheetApp.getActive();
  return {
    ok:true,
    users:getObjects_(ss.getSheetByName(CFG.SHEETS.users)).map(safeUserForAdmin_),
    assignments:getAssignmentRecords_(),
    units:getObjects_(ss.getSheetByName(CFG.SHEETS.cemeteries)),
    shifts:getObjects_(ss.getSheetByName(CFG.SHEETS.shifts)),
    dashboard:getPublicDashboard_(),
    reports:getObjects_(ss.getSheetByName(CFG.SHEETS.reports)).map(publicReportForAuthorized_)
  };
}

function reportTotals_(reports) {
  const out = {totalReports:reports.length,totalPersons:0};
  REPORT_TYPES.forEach(t=>out[t]=0);

  reports.forEach(r=>{
    const count=Number(r['Patient / Person Count'] || 0);
    out.totalPersons += count;
    const t=normalize_(r['Report Type']);
    if (Object.prototype.hasOwnProperty.call(out,t)) out[t]+=count;
    else out.Other+=count;
  });
  return out;
}

function publicReportForAuthorized_(r) {
  return {
    reportId:r['Report ID'],
    timestamp:r['Timestamp'],
    reportDate:r['Report Date'],
    reportTime:r['Report Time'],
    cemetery:r['Cemetery / Post'],
    teamLeader:r['Team Leader'],
    reporter:r['Staff / Reporter'],
    type:r['Report Type'],
    count:Number(r['Patient / Person Count'] || 0),
    patientId:r['Patient ID'] || '',
    status:r['Status'],
    details:r['Details'],
    actionTaken:r['Action Taken']
  };
}

/* =========================
   DUTY ASSIGNMENT MANAGEMENT
   ========================= */

function getAssignmentRecords_() {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.input);
  const values=sh.getDataRange().getValues();
  if (values.length<2) return [];

  const headers=values[0];
  const idx = {};
  headers.forEach((h,i)=>idx[h]=i);

  const out=[];
  for(let r=1;r<values.length;r++){
    const row=values[r];
    const site=normalize_(row[idx['Cemetery / Post']]);
    const shift=normalize_(row[idx['Shift ID']]);
    if(!site || !shift) continue;

    for(let slot=1;slot<=3;slot++){
      const col=idx['Staff ' + slot];
      out.push({
        row:r+1,
        group:row[idx['Group']],
        cemetery:site,
        teamLeader:normalize_(row[idx['Team Leader']]),
        shiftId:shift,
        shift:normalize_(row[idx['Shift / Time']]),
        date:normalize_(row[idx['Date / Coverage']]),
        slot:slot,
        personnelName:normalize_(row[col]),
        status:normalize_(row[col]) ? 'ASSIGNED' : 'VACANT'
      });
    }
  }
  return out;
}

function adminSaveAssignment_(data) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.input);
  const row=Number(data.row);
  const slot=Number(data.slot);

  if(row<2 || slot<1 || slot>3) throw new Error('Invalid assignment row/slot.');

  const name=normalize_(data.personnelName);
  const leader=normalize_(data.teamLeader);

  // Verify the row actually belongs to a configured unit/shift.
  const rowVals=sh.getRange(row,1,1,9).getValues()[0];
  if(!normalize_(rowVals[1]) || !normalize_(rowVals[3])) throw new Error('Invalid assignment row.');

  sh.getRange(row,7 + (slot-1)).setValue(name);

  if (leader !== '') sh.getRange(row,3).setValue(leader);

  syncDutyInput();
  buildUnitReportSummary();
  return {ok:true,message:'Duty assignment saved.'};
}

function adminDeleteAssignment_(data) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.input);
  const row=Number(data.row);
  const slot=Number(data.slot);

  if(row<2 || slot<1 || slot>3) throw new Error('Invalid assignment row/slot.');

  sh.getRange(row,7 + (slot-1)).clearContent();

  syncDutyInput();
  buildUnitReportSummary();
  return {ok:true,message:'Personnel assignment removed.'};
}

function syncDutyInput() {
  const ss=SpreadsheetApp.getActive();
  const input=ss.getSheetByName(CFG.SHEETS.input);
  if(!input || input.getLastRow()<2) return;

  const values=input.getRange(2,1,input.getLastRow()-1,9).getValues();
  const personnelMap=new Map();
  const dutyRows=[];
  const now=new Date();

  values.forEach(r=>{
    const group=normalize_(r[0]);
    const site=normalize_(r[1]);
    const leader=normalize_(r[2]);
    const shiftId=normalize_(r[3]);
    const shiftTime=normalize_(r[4]);
    const coverage=normalize_(r[5]);

    if(!site || !shiftId) return;

    for(let slot=0;slot<3;slot++){
      const name=normalize_(r[6+slot]);
      if(!name) continue;

      const key=name.toLowerCase() + '|' + site;
      if(!personnelMap.has(key)){
        personnelMap.set(key,{
          id:'PER-' + Utilities.getUuid().slice(0,8).toUpperCase(),
          name:name,
          role:'Standby Medic',
          site:site,
          leader:leader,
          status:'ACTIVE',
          count:0
        });
      }
      const p=personnelMap.get(key);
      p.count++;

      dutyRows.push([
        'DUTY-' + Utilities.getUuid().slice(0,8).toUpperCase(),
        site,
        group,
        leader,
        coverage,
        shiftId,
        shiftTime,
        p.id,
        p.name,
        p.role,
        'ON DUTY',
        now
      ]);
    }
  });

  const personnelRows=Array.from(personnelMap.values()).map(p=>[
    p.id,p.name,p.role,p.site,p.leader,p.status,p.count,''
  ]);

  replaceData_(ss.getSheetByName(CFG.SHEETS.personnel),personnelRows);
  replaceData_(ss.getSheetByName(CFG.SHEETS.duty),dutyRows);
  syncTeamLeadersToCemeteries_();
}

function syncTeamLeadersToCemeteries_() {
  const ss=SpreadsheetApp.getActive();
  const input=ss.getSheetByName(CFG.SHEETS.input);
  const cem=ss.getSheetByName(CFG.SHEETS.cemeteries);
  if(!input || !cem) return;

  const map={};
  if(input.getLastRow()>1){
    input.getRange(2,1,input.getLastRow()-1,9).getValues().forEach(r=>{
      const site=normalize_(r[1]);
      const leader=normalize_(r[2]);
      if(site && leader && !map[site]) map[site]=leader;
    });
  }

  if(cem.getLastRow()>1){
    const rows=cem.getRange(2,1,cem.getLastRow()-1,6).getValues();
    rows.forEach(r=>r[4]=map[normalize_(r[2])] || r[4] || '');
    cem.getRange(2,1,rows.length,6).setValues(rows);
  }
}

function replaceData_(sheet,rows) {
  const cols=sheet.getLastColumn();
  if(sheet.getLastRow()>1) sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(cols,1)).clearContent();
  if(!rows.length) return;
  if(sheet.getMaxColumns()<rows[0].length) sheet.insertColumnsAfter(sheet.getMaxColumns(),rows[0].length-sheet.getMaxColumns());
  sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}

/* =========================
   USER ACCESS MANAGEMENT
   ========================= */

function safeUserForAdmin_(u) {
  const role=String(u['Role']).toUpperCase();
  const unit=u['Cemetery / Post'] || '';
  const match=UNITS.find(x=>x.name===unit);
  return {
    userId:u['User ID'],
    username:u['Username'],
    role:role,
    unit:unit,
    personnelId:u['Personnel ID'] || '',
    fullName:u['Full Name'] || '',
    status:u['Status'] || 'ACTIVE',
    createdAt:u['Created At'] || '',
    updatedAt:u['Updated At'] || '',
    loginPassword:role==='TEAM_LEADER' && match ? DEFAULT_UNIT_PASSWORDS[String(match.group)] : ''
  };
}

function adminSaveUser_(data) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.users);

  const userId=normalize_(data.userId);
  const role=normalize_(data.role).toUpperCase();
  const username=role === 'TEAM_LEADER' ? '' : normalize_(data.username).toLowerCase();
  const unit=normalize_(data.unit);
  const personnelId=normalize_(data.personnelId);
  const fullName=normalize_(data.fullName);
  const status=normalize_(data.status || 'ACTIVE').toUpperCase();

  if(['ADMIN','TEAM_LEADER'].indexOf(role)===-1) throw new Error('Invalid role.');
  if(role==='ADMIN' && !username) throw new Error('Admin username is required.');
  if(role==='TEAM_LEADER' && !unit) throw new Error('Team Leader must have a locked cemetery / deployment unit.');

  let password=String(data.password || '').trim();
  // Team Leader passwords are tied to the selected unit.
  if(role==='TEAM_LEADER' && !password){
    const target=UNITS.find(x=>x.name===unit);
    if(!target) throw new Error('Invalid deployment unit.');
    password=DEFAULT_UNIT_PASSWORDS[String(target.group)] || '';
  }
  if(!userId && role==='ADMIN' && !password) throw new Error('Password is required when creating an Admin user.');

  const values=sh.getDataRange().getValues();
  const headers=values[0];
  const idx={}; headers.forEach((h,i)=>idx[h]=i);

  let rowIndex=-1;
  for(let i=1;i<values.length;i++){
    if(normalize_(values[i][idx['User ID']])===userId && userId){
      rowIndex=i+1;
      break;
    }
    if(role==='ADMIN' && normalize_(values[i][idx['Username']]).toLowerCase()===username && username){
      rowIndex=i+1;
    }
    if(role==='TEAM_LEADER' &&
       String(values[i][idx['Role']]).toUpperCase()==='TEAM_LEADER' &&
       normalize_(values[i][idx['Cemetery / Post']])===unit &&
       normalize_(values[i][idx['User ID']])!==userId){
      rowIndex=i+1;
    }
  }

  const now=new Date();
  const finalId=userId || ('USR-' + Utilities.getUuid().slice(0,8).toUpperCase());
  const passwordHash=password ? hashPassword_(password) : '';

  if(rowIndex===-1){
    sh.appendRow([
      finalId,username,passwordHash,role,unit,personnelId,fullName,status,now,now
    ]);
  } else {
    const old=sh.getRange(rowIndex,1,1,headers.length).getValues()[0];
    old[idx['User ID']]=finalId;
    old[idx['Username']]=username;
    if(passwordHash) old[idx['Password Hash']]=passwordHash;
    old[idx['Role']]=role;
    old[idx['Cemetery / Post']]=unit;
    old[idx['Personnel ID']]=personnelId;
    old[idx['Full Name']]=fullName;
    old[idx['Status']]=status;
    old[idx['Updated At']]=now;
    sh.getRange(rowIndex,1,1,headers.length).setValues([old]);
  }

  // Reflect Team Leader to CEMETERIES.
  if(role==='TEAM_LEADER' && unit){
    setCemeteryTeamLeader_(unit,fullName);
  }

  return {ok:true,message:'User account saved.'};
}

function setCemeteryTeamLeader_(unit,fullName) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.cemeteries);
  if(!sh || sh.getLastRow()<2) return;
  const rows=sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
  rows.forEach(r=>{
    if(normalize_(r[2])===unit) r[4]=fullName;
  });
  sh.getRange(2,1,rows.length,6).setValues(rows);
}


function getTeamLeaderCredentials_(token) {
  requireRole_(token, ['ADMIN']);
  return {
    ok:true,
    credentials:UNITS.map(u=>({
      unit:u.name,
      password:DEFAULT_UNIT_PASSWORDS[String(u.group)]
    }))
  };
}

function adminDeleteUser_(data) {
  const sh=SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.users);
  const userId=normalize_(data.userId);
  if(!userId) throw new Error('User ID is required.');

  const values=sh.getDataRange().getValues();
  const headers=values[0];
  const idCol=headers.indexOf('User ID');
  const roleCol=headers.indexOf('Role');

  for(let i=1;i<values.length;i++){
    if(normalize_(values[i][idCol])===userId){
      if(String(values[i][roleCol]).toUpperCase()==='ADMIN' && values.filter(r=>String(r[roleCol]).toUpperCase()==='ADMIN' && String(r[headers.indexOf('Status')]).toUpperCase()==='ACTIVE').length<=1){
        throw new Error('Cannot delete the last active Admin account.');
      }
      sh.deleteRow(i+1);
      return {ok:true,message:'User account deleted.'};
    }
  }
  throw new Error('User account not found.');
}

/* =========================
   REPORTS / PATIENTS
   ========================= */

function saveReport_(report) {
  const ss=SpreadsheetApp.getActive();
  const reports=ss.getSheetByName(CFG.SHEETS.reports);
  const patients=ss.getSheetByName(CFG.SHEETS.patients);

  const cemetery=normalize_(report.cemetery);
  if(!cemetery) throw new Error('Cemetery / Post is required.');

  const patient=report.patient || {};
  const fullName=normalize_(patient.fullName);
  const ageText=normalize_(patient.age);
  const age=ageText==='' ? '' : Number(ageText);
  const gender=normalize_(patient.gender);
  const address=normalize_(patient.address);
  const contact=normalize_(patient.contactNumber);

  const hasPatient=[fullName,ageText,gender,address,contact].some(Boolean);
  if(hasPatient && !fullName) throw new Error('Patient Full Name is required when patient information is provided.');

  const now=new Date();
  const reportId='REP-' + Utilities.getUuid().slice(0,8).toUpperCase();
  const patientId=hasPatient ? 'PAT-' + Utilities.getUuid().slice(0,8).toUpperCase() : '';

  if(hasPatient){
    patients.appendRow([
      patientId,now,fullName,age,gender,address,contact,
      reportId,cemetery,normalize_(report.reporter),''
    ]);
  }

  reports.appendRow([
    reportId,
    now,
    Utilities.formatDate(now,CFG.TIMEZONE,'yyyy-MM-dd'),
    Utilities.formatDate(now,CFG.TIMEZONE,'HH:mm:ss'),
    cemetery,
    normalize_(report.teamLeader),
    normalize_(report.reporter),
    normalize_(report.type || 'Other'),
    Number(report.count || 0),
    patientId,
    normalize_(report.status || 'Completed'),
    normalize_(report.details),
    normalize_(report.actionTaken)
  ]);

  return {ok:true,reportId:reportId,patientId:patientId};
}

function buildUnitReportSummary() {
  const ss=SpreadsheetApp.getActive();
  const summary=ss.getSheetByName(CFG.SHEETS.summary);
  const reports=ss.getSheetByName(CFG.SHEETS.reports);
  if(!summary || !reports) return;

  if(summary.getLastRow()>1) summary.getRange(2,1,summary.getLastRow()-1,10).clearContent();

  const reportRows=reports.getLastRow()>1 ? reports.getRange(2,1,reports.getLastRow()-1,13).getValues() : [];
  const leaderMap={};

  const cem=ss.getSheetByName(CFG.SHEETS.cemeteries);
  if(cem && cem.getLastRow()>1){
    cem.getRange(2,1,cem.getLastRow()-1,6).getValues().forEach(r=>leaderMap[normalize_(r[2])]=normalize_(r[4]));
  }

  const map={};
  UNITS.forEach(u=>{
    map[u.name]={site:u.name,leader:leaderMap[u.name]||'',totalReports:0,totalPersons:0};
    REPORT_TYPES.forEach(t=>map[u.name][t]=0);
  });

  reportRows.forEach(r=>{
    const site=normalize_(r[4]);
    if(!map[site]) return;
    const type=normalize_(r[7]);
    const count=Number(r[8]||0);
    map[site].totalReports++;
    map[site].totalPersons+=count;
    if(Object.prototype.hasOwnProperty.call(map[site],type)) map[site][type]+=count;
    else map[site].Other+=count;
  });

  const rows=UNITS.map(u=>{
    const x=map[u.name];
    return [
      x.site,x.leader,x.totalReports,x.totalPersons,
      x['Vital Signs'],x['Medicine Request'],x['First Aid'],
      x['Medical Assistance'],x['Emergency Response'],x['Other']
    ];
  });

  if(rows.length) summary.getRange(2,1,rows.length,10).setValues(rows);
}


function ensureDefaultTeamLeaders_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.SHEETS.users);
  const users = getObjects_(sh);
  const now = new Date();

  UNITS.forEach(u => {
    const existing = users.find(row =>
      String(row['Role']).toUpperCase()==='TEAM_LEADER' &&
      normalize_(row['Cemetery / Post'])===u.name
    );

    if (existing) return;

    const pass = DEFAULT_UNIT_PASSWORDS[String(u.group)];
    sh.appendRow([
      'USR-TL' + String(u.group).padStart(2,'0'),
      '',
      hashPassword_(pass),
      'TEAM_LEADER',
      u.name,
      '',
      'To be assigned',
      'ACTIVE',
      now,
      now
    ]);
  });
}

function getObjects_(sheet) {
  if(!sheet || sheet.getLastRow()<2) return [];
  const vals=sheet.getDataRange().getValues();
  const headers=vals[0];
  return vals.slice(1).filter(r=>r.some(v=>String(v).trim()!=='')).map(r=>{
    const o={};
    headers.forEach((h,i)=>o[h]=r[i]);
    return o;
  });
}

function unique_(values) {
  return [...new Set(values.map(normalize_).filter(Boolean))];
}

function formatSheets_() {
  const ss=SpreadsheetApp.getActive();
  Object.keys(CFG.SHEETS).forEach(k=>{
    const sh=ss.getSheetByName(CFG.SHEETS[k]);
    if(!sh) return;
    const cols=sh.getLastColumn();
    if(cols){
      sh.getRange(1,1,1,cols)
        .setFontWeight('bold')
        .setBackground('#07386e')
        .setFontColor('#ffffff');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1,cols);
    }
  });
}
