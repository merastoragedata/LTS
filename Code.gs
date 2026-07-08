/*************************************************************************
 * MSETCL — Load Trimming Scheme (LTS) Portal
 * Backend : Google Apps Script (Sheets = database, Drive = file store)
 * Frontend: Index.html (served via HtmlService, uses google.script.run)
 *
 * FIRST-TIME SETUP
 *   1. Create a new Apps Script project.
 *   2. Paste this file as "Code.gs" and add "Index.html".
 *   3. Run  setup()  once  -> creates the DB spreadsheet, Drive folder,
 *      the default admin login and seeds the Karjat 400 kV LTS records.
 *   4. Deploy > New deployment > Web app > Execute as "Me",
 *      Access "Anyone" (or "Anyone with Google account" for tighter use).
 *
 * DEFAULT ADMIN   username: admin    password: admin@123
 *   (change the password from the Admin page after first login)
 *************************************************************************/

var DB_NAME       = 'MSETCL_LTS_Portal_DB';
var ROOT_FOLDER   = 'MSETCL_LTS_Portal_Files';
var PW_SALT       = 'MSETCL-LTS-2026';        // app-level salt (change if you like)
var SESSION_HOURS = 8;

/* ------------------------------------------------------------------ *
 *  ENTRY POINT — serve the single-page app
 * ------------------------------------------------------------------ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MSETCL LTS Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================================================================== *
 *  LOW-LEVEL STORE HELPERS
 * ================================================================== */
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }

function getDb_() {
  var id = props_().getProperty('DB_ID');
  if (!id) throw new Error('Portal not initialised. Run setup() once.');
  return SpreadsheetApp.openById(id);
}
function getRootFolder_() {
  var id = props_().getProperty('ROOT_ID');
  return DriveApp.getFolderById(id);
}
function sheet_(name) {
  var ss = getDb_();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  return sh;
}

/* read a sheet into array-of-objects keyed by header row */
function readObjects_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var head = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var o = {};
    for (var c = 0; c < head.length; c++) o[head[c]] = values[r][c];
    o.__row = r + 1;
    out.push(o);
  }
  return out;
}
function headers_(name) { return sheet_(name).getRange(1, 1, 1, sheet_(name).getLastColumn()).getValues()[0]; }

function appendObject_(name, obj) {
  var sh = sheet_(name);
  var head = headers_(name);
  var row = head.map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
  sh.appendRow(row);
}
function updateRowById_(name, id, patch) {
  var sh = sheet_(name);
  var head = headers_(name);
  var idCol = head.indexOf('id');
  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(id)) {
      for (var k in patch) {
        var c = head.indexOf(k);
        if (c >= 0) sh.getRange(r + 1, c + 1).setValue(patch[k]);
      }
      return true;
    }
  }
  return false;
}
function deleteById_(name, id) {
  var sh = sheet_(name);
  var head = headers_(name);
  var idCol = head.indexOf('id');
  var values = sh.getDataRange().getValues();
  for (var r = values.length - 1; r >= 1; r--) {
    if (String(values[r][idCol]) === String(id)) sh.deleteRow(r + 1);
  }
  return true;
}
function uuid_()  { return Utilities.getUuid(); }
function now_()   { return new Date().toISOString(); }
function hash_(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw + PW_SALT);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* ================================================================== *
 *  SETUP / SCHEMA
 * ================================================================== */
function setup() {
  var p = props_();
  var ss = p.getProperty('DB_ID') ? SpreadsheetApp.openById(p.getProperty('DB_ID'))
                                  : SpreadsheetApp.create(DB_NAME);
  p.setProperty('DB_ID', ss.getId());

  var folder = p.getProperty('ROOT_ID') ? DriveApp.getFolderById(p.getProperty('ROOT_ID'))
                                        : DriveApp.createFolder(ROOT_FOLDER);
  p.setProperty('ROOT_ID', folder.getId());

  var schema = {
    Users:        ['id','username','passHash','role','substationCode','status','createdAt'],
    Substations:  ['id','code','name','division','status','createdBy','createdAt'],
    Shares:       ['id','substationCode','username','createdAt'],
    LTS:          ['id','substationCode','endType','srNo','division','substation','status','date',
                   'tfIct','hvLv','ctr','relayType','maxLoad','alarmPsm','stage1','stage2',
                   'feeders','loadRelief','remarks','healthiness','lastOpDate','visibleToViewer',
                   'createdAt','updatedAt'],
    CustomFields: ['id','ltsId','fieldName','fieldValue','createdAt'],
    Operations:   ['id','ltsId','substationCode','opTitle','interruptFrom','interruptTo',
                   'durationMin','blocks','createdBy','createdAt'],
    Uploads:      ['id','ltsId','substationCode','driveFileId','fileName','caption','mimeType','createdAt']
  };
  Object.keys(schema).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(schema[name]);
  });
  var def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  /* default admin */
  if (readObjects_('Users').filter(function (u) { return u.role === 'admin'; }).length === 0) {
    appendObject_('Users', { id: uuid_(), username: 'admin', passHash: hash_('admin@123'),
      role: 'admin', substationCode: '*', status: 'approved', createdAt: now_() });
  }

  seedKarjat_();
  return 'Setup complete. DB: ' + ss.getUrl();
}

function seedKarjat_() {
  var code = 'K400';
  if (readObjects_('Substations').filter(function (s) { return s.code === code; }).length) return;

  appendObject_('Substations', { id: uuid_(), code: code, name: '400kV Karjat',
    division: '400kV R.S. Dn Karjat', status: 'approved', createdBy: 'admin', createdAt: now_() });

  var base = { substationCode: code, endType: 'own', division: '400kV R.S. Dn Karjat',
    substation: '400kV Karjat', status: 'Approved', healthiness: 'Yes',
    lastOpDate: 'NA', visibleToViewer: 'yes' };

  var rows = [
    { srNo: 1, date: '09.04.2026', tfIct: '400/220/33kV 501 MVA ICT 1 & 2', hvLv: 'HV',
      ctr: '1500/1 A', relayType: 'Numerical NR PCS9611S', maxLoad: '270 MW / 405 A on each ICT',
      alarmPsm: '48% , td 1 sec', stage1: '53% , 795 A , 1.3 sec', stage2: '—',
      feeders: 'Stage 1: 33kV fdrs in (1) 220kV Jeur, Shirsuphal, Kurkumbh, Baramati & Bhose SS and (2) 132kV Karmala, Parewadi, Daund & Janai SS through PLCC Link',
      loadRelief: '196 MW / 196 MW', remarks: 'In case of tripping of any one ICT (N-1 criteria)' },
    { srNo: 2, date: '21.04.2025', tfIct: '400kV Girwali Ckt. 1 & 2', hvLv: '—',
      ctr: '3000/1 A', relayType: 'Distance Relay Main II — Numerical GE Micom P444',
      maxLoad: '624 MW / 920 A on each Ckt.', alarmPsm: 'NA',
      stage1: '534 mA , 1600 A , 1.3 sec', stage2: '—',
      feeders: 'Stage 1: 33kV fdrs in (1) 220kV Jeur, Shirsuphal, Kurkumbh, Baramati, Tembhurni, Malinagar & Bhose SS and (2) 132kV Karmala, Parewadi, Daund & Janai SS through PLCC Link',
      loadRelief: '324 MW / 324 MW', remarks: 'In case of tripping of any one ckt (N-1 contingency)' },
    { srNo: 3, date: '04.07.2026', tfIct: 'UV LTS (Under Voltage in HV B/U relay)', hvLv: '—',
      ctr: '400kV/110V = 3636.36', relayType: 'Numerical NR PCS9611S',
      maxLoad: 'Under Voltage upto 371.5 kV', alarmPsm: 'NA',
      stage1: '101.75 V , 370 kV , 5 sec', stage2: '100.375 V , 365 kV , 5 sec',
      feeders: 'Stage 1: 33kV fdrs in 220kV Jeur & Bhose SS, 132kV Karmala & Parewadi SS. Stage 2: 220kV Shirsuphal, Kurkumbh, Baramati SS, 132kV Daund & Janai SS through PLCC Link',
      loadRelief: 'Stage 1: 84 MW / 84 MW ; Stage 2: 112 MW / 112 MW',
      remarks: 'In case of under voltage in 400kV Grid' }
  ];
  rows.forEach(function (r) {
    var o = {}; for (var k in base) o[k] = base[k]; for (var k2 in r) o[k2] = r[k2];
    o.id = uuid_(); o.createdAt = now_(); o.updatedAt = now_();
    appendObject_('LTS', o);
  });
}

/* ================================================================== *
 *  SESSIONS / AUTH
 * ================================================================== */
function makeSession_(user) {
  var token = uuid_() + uuid_();
  var payload = JSON.stringify({ username: user.username, role: user.role,
    substationCode: user.substationCode });
  cache_().put('sess_' + token, payload, SESSION_HOURS * 3600);
  return token;
}
function requireSession_(token) {
  var raw = cache_().get('sess_' + token);
  if (!raw) throw new Error('Session expired. Please log in again.');
  var s = JSON.parse(raw);
  cache_().put('sess_' + token, raw, SESSION_HOURS * 3600); // sliding refresh
  return s;
}
function requireAdmin_(token) {
  var s = requireSession_(token);
  if (s.role !== 'admin') throw new Error('Admin permission required.');
  return s;
}

/* which substation codes may this session read? */
function accessibleCodes_(s) {
  if (s.role === 'admin') return readObjects_('Substations')
    .filter(function (x) { return x.status === 'approved'; }).map(function (x) { return x.code; });
  var codes = {};
  if (s.substationCode && s.substationCode !== '*') codes[s.substationCode] = true;
  readObjects_('Shares').forEach(function (sh) {
    if (sh.username === s.username) codes[sh.substationCode] = true;
  });
  return Object.keys(codes);
}
function canEdit_(s, code) { return s.role === 'admin' || (s.role === 'editor' && s.substationCode === code); }

/* ---------------- public auth endpoints ---------------- */
function apiLogin(username, password) {
  username = String(username || '').trim();
  var u = readObjects_('Users').filter(function (x) {
    return String(x.username).toLowerCase() === username.toLowerCase(); })[0];
  if (!u) throw new Error('User not found.');
  if (u.status !== 'approved') throw new Error('Account pending admin approval.');
  if (u.passHash !== hash_(password)) throw new Error('Incorrect password.');
  return { token: makeSession_(u), user: { username: u.username, role: u.role, substationCode: u.substationCode } };
}

function apiRegister(username, password, subName, subCode) {
  username = String(username || '').trim();
  subCode  = String(subCode || '').trim().toUpperCase();
  if (!username || !password || !subCode) throw new Error('All fields are required.');
  if (readObjects_('Users').filter(function (x) {
      return String(x.username).toLowerCase() === username.toLowerCase(); }).length)
    throw new Error('Username already taken.');

  var subs = readObjects_('Substations');
  var existing = subs.filter(function (s) { return s.code === subCode; })[0];
  if (!existing) {
    appendObject_('Substations', { id: uuid_(), code: subCode, name: subName || subCode,
      division: '', status: 'pending', createdBy: username, createdAt: now_() });
  }
  appendObject_('Users', { id: uuid_(), username: username, passHash: hash_(password),
    role: 'editor', substationCode: subCode, status: 'pending', createdAt: now_() });
  return { ok: true, message: 'Registration submitted. An admin must approve your account and substation.' };
}

/* ---------------- context ---------------- */
function apiGetContext(token) {
  var s = requireSession_(token);
  var codes = accessibleCodes_(s);
  var subs = readObjects_('Substations').filter(function (x) {
    return x.status === 'approved' && (s.role === 'admin' || codes.indexOf(x.code) >= 0);
  }).map(function (x) { return { code: x.code, name: x.name, division: x.division }; });
  return { user: { username: s.username, role: s.role, substationCode: s.substationCode }, substations: subs };
}

/* ================================================================== *
 *  ADMIN
 * ================================================================== */
function apiAdminData(token) {
  requireAdmin_(token);
  return {
    pendingUsers: readObjects_('Users').filter(function (u) { return u.status === 'pending'; })
      .map(function (u) { return { username: u.username, role: u.role, substationCode: u.substationCode, createdAt: u.createdAt }; }),
    pendingSubs: readObjects_('Substations').filter(function (x) { return x.status === 'pending'; })
      .map(function (x) { return { code: x.code, name: x.name, createdBy: x.createdBy }; }),
    users: readObjects_('Users').map(function (u) {
      return { username: u.username, role: u.role, substationCode: u.substationCode, status: u.status }; }),
    substations: readObjects_('Substations').map(function (x) {
      return { code: x.code, name: x.name, division: x.division, status: x.status }; })
  };
}
function apiApproveUser(token, username) {
  requireAdmin_(token);
  var u = readObjects_('Users').filter(function (x) { return x.username === username; })[0];
  if (!u) throw new Error('User not found.');
  updateRowById_('Users', u.id, { status: 'approved' });
  var sub = readObjects_('Substations').filter(function (x) { return x.code === u.substationCode; })[0];
  if (sub && sub.status === 'pending') updateRowById_('Substations', sub.id, { status: 'approved' });
  return apiAdminData(token);
}
function apiRejectUser(token, username) {
  requireAdmin_(token);
  var u = readObjects_('Users').filter(function (x) { return x.username === username; })[0];
  if (u) deleteById_('Users', u.id);
  return apiAdminData(token);
}
function apiApproveSubstation(token, code) {
  requireAdmin_(token);
  var s = readObjects_('Substations').filter(function (x) { return x.code === code; })[0];
  if (s) updateRowById_('Substations', s.id, { status: 'approved' });
  return apiAdminData(token);
}
function apiCreateSubstation(token, name, code, division, editorUser, editorPass) {
  requireAdmin_(token);
  code = String(code || '').trim().toUpperCase();
  if (!code || !name) throw new Error('Substation name and code are required.');
  if (readObjects_('Substations').filter(function (x) { return x.code === code; }).length)
    throw new Error('Substation code already exists.');
  appendObject_('Substations', { id: uuid_(), code: code, name: name, division: division || '',
    status: 'approved', createdBy: 'admin', createdAt: now_() });
  if (editorUser && editorPass) {
    if (readObjects_('Users').filter(function (x) {
        return String(x.username).toLowerCase() === String(editorUser).toLowerCase(); }).length)
      throw new Error('Editor username already taken.');
    appendObject_('Users', { id: uuid_(), username: editorUser, passHash: hash_(editorPass),
      role: 'editor', substationCode: code, status: 'approved', createdAt: now_() });
  }
  return apiAdminData(token);
}

/* editor/admin: create a viewer login for a substation */
function apiCreateViewer(token, code, username, password) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Only the substation editor or an admin can add a viewer.');
  if (readObjects_('Users').filter(function (x) {
      return String(x.username).toLowerCase() === String(username).toLowerCase(); }).length)
    throw new Error('Username already taken.');
  appendObject_('Users', { id: uuid_(), username: username, passHash: hash_(password),
    role: 'viewer', substationCode: code, status: 'approved', createdAt: now_() });
  return { ok: true };
}

/* editor/admin: share read access of a substation with another registered user */
function apiShareSubstation(token, code, username) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Only the substation editor or an admin can share.');
  if (!readObjects_('Users').filter(function (x) { return x.username === username; }).length)
    throw new Error('No such user.');
  appendObject_('Shares', { id: uuid_(), substationCode: code, username: username, createdAt: now_() });
  return { ok: true };
}

function apiChangePassword(token, oldPw, newPw) {
  var s = requireSession_(token);
  var u = readObjects_('Users').filter(function (x) { return x.username === s.username; })[0];
  if (!u || u.passHash !== hash_(oldPw)) throw new Error('Current password is incorrect.');
  updateRowById_('Users', u.id, { passHash: hash_(newPw) });
  return { ok: true };
}

/* ================================================================== *
 *  LTS CRUD
 * ================================================================== */
var LTS_FIELDS = ['endType','srNo','division','substation','status','date','tfIct','hvLv','ctr',
  'relayType','maxLoad','alarmPsm','stage1','stage2','feeders','loadRelief','remarks',
  'healthiness','lastOpDate','visibleToViewer'];

function apiGetLTSList(token, code) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access to this substation.');
  var editor = canEdit_(s, code);
  var list = readObjects_('LTS').filter(function (l) { return l.substationCode === code; });
  if (!editor) list = list.filter(function (l) { return String(l.visibleToViewer).toLowerCase() === 'yes'; });
  list.sort(function (a, b) { return (Number(a.srNo) || 0) - (Number(b.srNo) || 0); });
  return { canEdit: editor, own: list.filter(function (l) { return l.endType !== 'other'; }),
           other: list.filter(function (l) { return l.endType === 'other'; }) };
}

function apiGetLTS(token, code, id) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  var l = readObjects_('LTS').filter(function (x) { return x.id === id; })[0];
  if (!l) throw new Error('LTS not found.');
  var cf = readObjects_('CustomFields').filter(function (x) { return x.ltsId === id; })
    .map(function (x) { return { id: x.id, fieldName: x.fieldName, fieldValue: x.fieldValue }; });
  var up = readObjects_('Uploads').filter(function (x) { return x.ltsId === id; })
    .map(function (x) { return { id: x.id, driveFileId: x.driveFileId, fileName: x.fileName,
      caption: x.caption, mimeType: x.mimeType }; });
  return { lts: l, customFields: cf, uploads: up, canEdit: canEdit_(s, code) };
}

function apiSaveLTS(token, code, obj) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var clean = {};
  LTS_FIELDS.forEach(function (f) { clean[f] = obj[f] === undefined ? '' : obj[f]; });
  if (obj.id) {
    clean.updatedAt = now_();
    updateRowById_('LTS', obj.id, clean);
    return { id: obj.id };
  }
  clean.id = uuid_(); clean.substationCode = code;
  clean.createdAt = now_(); clean.updatedAt = now_();
  if (!clean.srNo) {
    var n = readObjects_('LTS').filter(function (l) { return l.substationCode === code; }).length;
    clean.srNo = n + 1;
  }
  appendObject_('LTS', clean);
  return { id: clean.id };
}

function apiDeleteLTS(token, code, id) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  readObjects_('CustomFields').filter(function (x) { return x.ltsId === id; })
    .forEach(function (x) { deleteById_('CustomFields', x.id); });
  readObjects_('Uploads').filter(function (x) { return x.ltsId === id; }).forEach(function (x) {
    try { DriveApp.getFileById(x.driveFileId).setTrashed(true); } catch (e) {}
    deleteById_('Uploads', x.id);
  });
  deleteById_('LTS', id);
  return { ok: true };
}

function apiToggleVisible(token, code, id, visible) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  updateRowById_('LTS', id, { visibleToViewer: visible ? 'yes' : 'no' });
  return { ok: true };
}

/* -------- custom fields -------- */
function apiAddCustomField(token, code, ltsId, name, value) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var id = uuid_();
  appendObject_('CustomFields', { id: id, ltsId: ltsId, fieldName: name, fieldValue: value, createdAt: now_() });
  return { id: id };
}
function apiDeleteCustomField(token, code, id) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  deleteById_('CustomFields', id);
  return { ok: true };
}

/* ================================================================== *
 *  UPLOADS  (Drive)
 * ================================================================== */
function subFolder_(code) {
  var root = getRootFolder_();
  var it = root.getFoldersByName(code);
  return it.hasNext() ? it.next() : root.createFolder(code);
}
function apiUploadFile(token, code, ltsId, base64, fileName, mimeType, caption) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var file = subFolder_(code).createFile(blob);
  var id = uuid_();
  appendObject_('Uploads', { id: id, ltsId: ltsId, substationCode: code, driveFileId: file.getId(),
    fileName: fileName, caption: caption || fileName, mimeType: mimeType, createdAt: now_() });
  return { id: id, driveFileId: file.getId() };
}
function apiGetFile(token, fileId) {
  var s = requireSession_(token);
  var meta = readObjects_('Uploads').filter(function (x) { return x.driveFileId === fileId; })[0];
  if (!meta) throw new Error('File not found.');
  if (accessibleCodes_(s).indexOf(meta.substationCode) < 0) throw new Error('No access.');
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  return { base64: Utilities.base64Encode(blob.getBytes()), mimeType: meta.mimeType, fileName: meta.fileName };
}
function apiDeleteUpload(token, code, id) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var u = readObjects_('Uploads').filter(function (x) { return x.id === id; })[0];
  if (u) { try { DriveApp.getFileById(u.driveFileId).setTrashed(true); } catch (e) {} deleteById_('Uploads', id); }
  return { ok: true };
}

/* ================================================================== *
 *  OPERATIONS
 * ================================================================== */
function apiSaveOperation(token, code, op) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  if (op.id) {
    updateRowById_('Operations', op.id, {
      opTitle: op.opTitle || '', interruptFrom: op.interruptFrom || '', interruptTo: op.interruptTo || '',
      durationMin: op.durationMin || '', blocks: JSON.stringify(op.blocks || []) });
    return { id: op.id };
  }
  var id = uuid_();
  appendObject_('Operations', { id: id, ltsId: op.ltsId, substationCode: code,
    opTitle: op.opTitle || '', interruptFrom: op.interruptFrom || '', interruptTo: op.interruptTo || '',
    durationMin: op.durationMin || '', blocks: JSON.stringify(op.blocks || []),
    createdBy: s.username, createdAt: now_() });
  return { id: id };
}
function apiGetOperations(token, code) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  var lts = {};
  readObjects_('LTS').filter(function (l) { return l.substationCode === code; })
    .forEach(function (l) { lts[l.id] = l.tfIct; });
  return readObjects_('Operations').filter(function (o) { return o.substationCode === code; })
    .map(function (o) {
      return { id: o.id, ltsId: o.ltsId, ltsName: lts[o.ltsId] || '(deleted)', opTitle: o.opTitle,
        interruptFrom: o.interruptFrom, interruptTo: o.interruptTo, durationMin: o.durationMin,
        blocks: safeJson_(o.blocks), createdBy: o.createdBy, createdAt: o.createdAt };
    }).sort(function (a, b) { return String(b.interruptFrom).localeCompare(String(a.interruptFrom)); });
}
function apiDeleteOperation(token, code, id) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  deleteById_('Operations', id);
  return { ok: true };
}
function safeJson_(str) { try { return JSON.parse(str || '[]'); } catch (e) { return []; } }

/* ================================================================== *
 *  EXPORTS — build a formatted temp Sheet, export as XLSX / PDF
 *  scope: 'list' (all LTS in substation) | 'single' (one LTS) | 'ops'
 * ================================================================== */
function apiExport(token, code, format, scope, id) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  var temp = SpreadsheetApp.create('LTS_export_' + Date.now());
  try {
    var sh = temp.getSheets()[0];
    var title, matrix;
    if (scope === 'single') {
      var d = apiGetLTS(token, code, id);
      title = 'LTS Detail — ' + d.lts.tfIct;
      matrix = ltsDetailMatrix_(d);
    } else if (scope === 'ops') {
      title = 'LTS Operations — ' + code;
      matrix = opsMatrix_(apiGetOperations(token, code));
    } else {
      title = 'Load Trimming Schemes — ' + code;
      matrix = ltsListMatrix_(readObjects_('LTS').filter(function (l) { return l.substationCode === code; })
        .sort(function (a, b) { return (Number(a.srNo) || 0) - (Number(b.srNo) || 0); }));
    }
    writeFormatted_(sh, title, matrix);
    SpreadsheetApp.flush();

    var blob;
    if (format === 'pdf') {
      blob = exportBlob_(temp.getId(), 'pdf');
      blob.setName(safeName_(title) + '.pdf');
    } else {
      blob = exportBlob_(temp.getId(), 'xlsx');
      blob.setName(safeName_(title) + '.xlsx');
    }
    return { base64: Utilities.base64Encode(blob.getBytes()),
             mimeType: blob.getContentType(), fileName: blob.getName() };
  } finally {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
  }
}

function safeName_(t) { return String(t).replace(/[^a-z0-9\- ]/gi, '').replace(/\s+/g, '_').slice(0, 60); }

/* header row is matrix[0]; subsequent rows are data */
function writeFormatted_(sh, title, matrix) {
  var cols = matrix[0].length;
  sh.getRange(1, 1, 1, cols).merge().setValue(title)
    .setFontSize(14).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#0b2a4a').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);

  var startRow = 3;
  sh.getRange(startRow, 1, matrix.length, cols).setValues(matrix)
    .setWrap(true).setVerticalAlignment('top').setBorder(true, true, true, true, true, true);

  sh.getRange(startRow, 1, 1, cols).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#12507e').setHorizontalAlignment('center');

  /* smart-ish column widths: sample content length, capped */
  for (var c = 1; c <= cols; c++) {
    var maxLen = 8;
    for (var r = 0; r < matrix.length; r++) {
      var v = String(matrix[r][c - 1] || '');
      v.split('\n').forEach(function (line) { if (line.length > maxLen) maxLen = line.length; });
    }
    var w = Math.min(360, Math.max(70, maxLen * 7 + 16));
    sh.setColumnWidth(c, w);
  }
  sh.getRange(startRow, 1, matrix.length, cols).setFontSize(9);
  sh.setFrozenRows(startRow);
}

function ltsListMatrix_(list) {
  var head = ['Sr','End','ICT / Line','HV/LV','CTR','Relay','Max Load','Alarm PSM',
    'Stage-1 Trip','Stage-2 Trip','Feeders in scheme','Load Relief (Obtd/Reqd)',
    'Remarks','Healthy','Last Op','Status','Date'];
  var rows = list.map(function (l) {
    return [l.srNo, l.endType === 'other' ? 'Other' : 'Own', l.tfIct, l.hvLv, l.ctr, l.relayType,
      l.maxLoad, l.alarmPsm, l.stage1, l.stage2, l.feeders, l.loadRelief, l.remarks,
      l.healthiness, l.lastOpDate, l.status, l.date]; });
  return [head].concat(rows);
}
function ltsDetailMatrix_(d) {
  var l = d.lts;
  var pairs = [['Field', 'Value'],
    ['Sr No', l.srNo], ['Scheme end', l.endType === 'other' ? 'Other substation end' : 'Own substation end'],
    ['Division', l.division], ['Substation', l.substation], ['Status', l.status], ['Date', l.date],
    ['ICT / Line', l.tfIct], ['HV/LV', l.hvLv], ['Connected CTR', l.ctr], ['Relay', l.relayType],
    ['Max load recorded', l.maxLoad], ['Alarm PSM setting', l.alarmPsm],
    ['Stage-1 (Trip)', l.stage1], ['Stage-2 (Trip)', l.stage2],
    ['Feeders in scheme', l.feeders], ['Load relief (obtd/reqd)', l.loadRelief],
    ['Remarks / purpose', l.remarks], ['Healthiness', l.healthiness], ['Last operating date', l.lastOpDate]];
  d.customFields.forEach(function (c) { pairs.push([c.fieldName, c.fieldValue]); });
  return pairs;
}
function opsMatrix_(ops) {
  var head = ['LTS', 'Operation', 'Interruption From', 'Interruption To', 'Duration (min)', 'Details', 'Logged by'];
  var rows = ops.map(function (o) {
    var text = (o.blocks || []).map(function (b) {
      return (b.label ? b.label + ': ' : '') + stripHtml_(b.html); }).join('\n\n');
    return [o.ltsName, o.opTitle, fmtDt_(o.interruptFrom), fmtDt_(o.interruptTo), o.durationMin, text, o.createdBy];
  });
  return [head].concat(rows);
}
function stripHtml_(h) { return String(h || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim(); }
function fmtDt_(iso) { if (!iso) return ''; var d = new Date(iso); return isNaN(d) ? iso :
  Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm'); }

function exportBlob_(ssId, format) {
  var params = format === 'pdf'
    ? 'format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=true&fzr=true&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4'
    : 'format=xlsx';
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' + params;
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  return resp.getBlob();
}
