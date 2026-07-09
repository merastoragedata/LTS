/*************************************************************************
 * MSETCL — Load Trimming Scheme (LTS) Portal
 * Backend : Google Apps Script (Sheets = database, Drive = file store)
 * Frontend: Index.html — hosted separately (GitHub Pages / any static host)
 *           and talks to this backend over fetch() -> doPost().
 *
 * FIRST-TIME SETUP
 *   1. Create a new Apps Script project and paste THIS file as "Code.gs".
 *      (You do NOT need to add Index.html here — it is hosted elsewhere.)
 *   2. Run  setup()  once -> creates the DB spreadsheet, the Drive folder,
 *      the default admin login and seeds the Karjat 400 kV LTS records.
 *   3. Deploy > New deployment > Web app >
 *        Execute as        : Me
 *        Who has access    : Anyone   (required so the hosted page can call it)
 *   4. Copy the Web-app URL (…/exec). Give it to me and I will hard-code it
 *      into Index.html; then upload Index.html to GitHub Pages / your host.
 *
 * DEFAULT ADMIN   username: admin    password: admin@123
 *   (change the password from the Account page after first login)
 *************************************************************************/

var DB_NAME       = 'MSETCL_LTS_Portal_DB';
var ROOT_FOLDER   = 'MSETCL_LTS_Portal_Files';
var PW_SALT       = 'MSETCL-LTS-2026';        // app-level salt (change if you like)
var SESSION_HOURS = 8;

/* ------------------------------------------------------------------ *
 *  ENTRY POINTS
 *
 *  The frontend (Index.html) is hosted separately (e.g. GitHub Pages),
 *  so it is a DIFFERENT ORIGIN from this script. It therefore talks to
 *  this backend over fetch() -> doPost(), NOT google.script.run.
 *
 *  CORS NOTE (important for the frontend):
 *    The browser must send the POST as a "simple request" so it does NOT
 *    trigger a CORS pre-flight (which Apps Script cannot answer). That
 *    means the fetch MUST use  Content-Type: text/plain  and no custom
 *    headers. An anonymous Apps Script web app then returns the response
 *    with  Access-Control-Allow-Origin: *  automatically. This is handled
 *    for you in Index.html — just deploy with access = "Anyone".
 *
 *  Request body (JSON string):  { "fn": "apiLogin", "args": [ ... ] }
 *  Response body (JSON string): { "ok": true,  "data": <result> }
 *                          or   { "ok": false, "error": "<message>" }
 * ------------------------------------------------------------------ */

/* only these functions may be invoked from the web — never setup()/helpers */
var API_WHITELIST = {
  apiLogin:1, apiRegister:1, apiGetContext:1,
  apiAdminData:1, apiApproveUser:1, apiRejectUser:1, apiApproveSubstation:1,
  apiCreateSubstation:1, apiCreateViewer:1, apiShareSubstation:1, apiChangePassword:1,
  apiListShares:1, apiRevokeShare:1,
  apiGetLTSList:1, apiGetLTS:1, apiSaveLTS:1, apiDeleteLTS:1, apiToggleVisible:1,
  apiGetVersions:1, apiSaveLTSVersion:1,
  apiGetDiagram:1, apiSaveDiagram:1, apiDeleteDiagram:1,
  apiAddCustomField:1, apiDeleteCustomField:1,
  apiUploadFile:1, apiGetFile:1, apiDeleteUpload:1,
  apiGetAnnotations:1, apiSaveAnnotation:1, apiDeleteAnnotation:1,
  apiSaveOperation:1, apiGetOperations:1, apiDeleteOperation:1,
  apiExport:1
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Empty request');
    var req  = JSON.parse(e.postData.contents);
    var fn   = req.fn;
    var args = req.args || [];
    if (!API_WHITELIST[fn]) throw new Error('Unknown action: ' + fn);
    var result = globalThis[fn].apply(null, args);
    return json_({ ok: true, data: result });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}

/* Visiting the /exec URL in a browser just shows a friendly status page.
   (Read-only ping calls may also GET  ?fn=apiGetContext&args=["<token>"] ) */
function doGet(e) {
  if (e && e.parameter && e.parameter.fn) {
    try {
      var fn = e.parameter.fn;
      if (!API_WHITELIST[fn]) throw new Error('Unknown action: ' + fn);
      var args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      return json_({ ok: true, data: globalThis[fn].apply(null, args) });
    } catch (err) {
      return json_({ ok: false, error: (err && err.message) ? err.message : String(err) });
    }
  }
  var ready = !!props_().getProperty('DB_ID');
  return HtmlService.createHtmlOutput(
    '<html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:system-ui,Segoe UI,Roboto,sans-serif;background:#0b1622;color:#dbe7f3;' +
    'display:flex;min-height:100vh;align-items:center;justify-content:center;margin:0}' +
    '.c{text-align:center;padding:28px 34px;border:1px solid #22364b;border-radius:14px;background:#101d2c}' +
    'b{color:#4bd0a0}</style></head><body><div class="c">' +
    '<h2>MSETCL LTS Portal — Backend</h2>' +
    '<p>Status: <b>' + (ready ? 'running &amp; initialised' : 'deployed — run setup() once') + '</b></p>' +
    '<p style="color:#7f96ad;font-size:13px">This is the API endpoint. Open the portal from your hosted web page.</p>' +
    '</div></body></html>'
  ).setTitle('MSETCL LTS Portal — Backend');
}

/* ================================================================== *
 *  LOW-LEVEL STORE HELPERS
 * ================================================================== */
function props_() { return PropertiesService.getScriptProperties(); }
function cache_() { return CacheService.getScriptCache(); }

/* ---- per-execution memo caches (an Apps Script run is short-lived, so
   these are safe and cut most of the latency: openById + getDataRange
   were previously repeated on every single read) ---- */
var _SS = null, _SH = {}, _ROWS = {};

function getDb_() {
  if (_SS) return _SS;
  var id = props_().getProperty('DB_ID');
  if (!id) throw new Error('Portal not initialised. Run setup() once.');
  _SS = SpreadsheetApp.openById(id);
  return _SS;
}
function getRootFolder_() {
  var id = props_().getProperty('ROOT_ID');
  return DriveApp.getFolderById(id);
}
function sheet_(name) {
  if (_SH[name]) return _SH[name];
  var sh = getDb_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name);
  _SH[name] = sh;
  return sh;
}
/* call after ANY write so later reads in the same execution see fresh data */
function invalidate_(name) {
  if (name) delete _ROWS[name]; else _ROWS = {};
}

/* read a sheet into array-of-objects keyed by header row (memoised) */
function readObjects_(name) {
  if (_ROWS[name]) return _ROWS[name];
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) { _ROWS[name] = []; return []; }
  var head = values[0];
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var o = {};
    for (var c = 0; c < head.length; c++) o[head[c]] = values[r][c];
    o.__row = r + 1;
    out.push(o);
  }
  _ROWS[name] = out;
  return out;
}
function headers_(name) {
  var sh = sheet_(name);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function appendObject_(name, obj) {
  var sh = sheet_(name);
  var head = headers_(name);
  var row = head.map(function (h) { return obj[h] === undefined || obj[h] === null ? '' : obj[h]; });
  sh.appendRow(row);
  invalidate_(name);
}
function updateRowById_(name, id, patch) {
  var sh = sheet_(name);
  var head = headers_(name);
  var idCol = head.indexOf('id');
  var values = sh.getDataRange().getValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(id)) {
      var row = values[r].slice();
      for (var k in patch) {
        var c = head.indexOf(k);
        if (c >= 0) row[c] = patch[k];
      }
      /* one batched write instead of one call per changed cell */
      sh.getRange(r + 1, 1, 1, head.length).setValues([row]);
      invalidate_(name);
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
  invalidate_(name);
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
    Shares:       ['id','substationCode','username','level','createdAt'],
    LTS:          ['id','substationCode','endType','srNo','division','substation','status','date',
                   'tfIct','hvLv','ctr','relayType','maxLoad','alarmPsm','stage1','stage2',
                   'feeders','loadRelief','remarks','healthiness','lastOpDate','visibleToViewer',
                   'createdAt','updatedAt'],
    /* every revision of an LTS setting is kept; the newest is the live one */
    LTSVersions:  ['id','ltsId','substationCode','versionNo','functionalFrom','functionalTill',
                   'data','customFields','approvalRemarks','approvalFileId','approvalFileName',
                   'createdBy','createdAt'],
    CustomFields: ['id','ltsId','fieldName','fieldValue','createdAt'],
    Operations:   ['id','ltsId','substationCode','opTitle','interruptFrom','interruptTo',
                   'durationMin','blocks','createdBy','createdAt'],
    /* versionId binds an attachment to the revision it was filed under */
    Uploads:      ['id','ltsId','versionId','substationCode','driveFileId','fileName','caption',
                   'mimeType','createdAt'],
    /* user-drawn single-line diagram, one per LTS (JSON scene graph) */
    Diagrams:     ['id','ltsId','substationCode','json','updatedBy','updatedAt'],
    /* highlight rectangles drawn on an uploaded image/PDF page */
    Annotations:  ['id','driveFileId','substationCode','page','x','y','w','h','color','note',
                   'createdBy','createdAt']
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
  var code = 'J966';
  /* Legacy installs used K400. Rename first so we never seed a duplicate. */
  if (readObjects_('Substations').filter(function (s) { return s.code === 'K400'; }).length) {
    renameSubstationCode('K400', code);
  }
  if (readObjects_('Substations').filter(function (s) { return s.code === code; }).length) return;

  appendObject_('Substations', { id: uuid_(), code: code, name: '400kV Karjat',
    division: '400kV R.S. Dn Karjat', status: 'approved', createdBy: 'admin', createdAt: now_() });

  var base = { substationCode: code, endType: 'own', division: '400kV R.S. Dn Karjat',
    substation: '400kV Karjat', status: 'Approved', healthiness: 'Yes',
    lastOpDate: '', visibleToViewer: 'yes' };

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
/* A substation is private to the editor who owns it. Others see it only via
   an explicit Shares row granted by that editor or by an admin. */
function sharesFor_(username) {
  return readObjects_('Shares').filter(function (sh) {
    return String(sh.username).toLowerCase() === String(username).toLowerCase();
  });
}
function accessibleCodes_(s) {
  if (s.role === 'admin') return readObjects_('Substations')
    .filter(function (x) { return x.status === 'approved'; }).map(function (x) { return x.code; });
  var codes = {};
  if (s.substationCode && s.substationCode !== '*') codes[s.substationCode] = true;
  sharesFor_(s.username).forEach(function (sh) { codes[sh.substationCode] = true; });
  return Object.keys(codes);
}
/* edit rights: admin, the owning editor, or an editor-level share */
function canEdit_(s, code) {
  if (s.role === 'admin') return true;
  if (s.role === 'editor' && s.substationCode === code) return true;
  if (s.role === 'viewer') return false;
  return sharesFor_(s.username).some(function (sh) {
    return sh.substationCode === code && String(sh.level).toLowerCase() === 'editor';
  });
}
/* only the OWNING editor or an admin may hand out access */
function canGrant_(s, code) {
  return s.role === 'admin' || (s.role === 'editor' && s.substationCode === code);
}

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
/* level: 'viewer' (read-only) or 'editor' (full edit on that substation) */
function apiShareSubstation(token, code, username, level) {
  var s = requireSession_(token);
  if (!canGrant_(s, code)) throw new Error('Only the owning editor or an admin can grant access.');
  level = String(level || 'viewer').toLowerCase();
  if (level !== 'viewer' && level !== 'editor') throw new Error('Level must be viewer or editor.');

  var u = readObjects_('Users').filter(function (x) {
    return String(x.username).toLowerCase() === String(username).toLowerCase(); })[0];
  if (!u) throw new Error('No such user.');
  if (u.username === s.username) throw new Error('You already own this substation.');
  if (level === 'editor' && u.role === 'viewer')
    throw new Error('"' + u.username + '" is a viewer account and cannot be granted editor access.');

  /* upsert — re-sharing updates the level instead of stacking rows */
  var existing = readObjects_('Shares').filter(function (sh) {
    return sh.substationCode === code &&
           String(sh.username).toLowerCase() === String(u.username).toLowerCase(); })[0];
  if (existing) updateRowById_('Shares', existing.id, { level: level });
  else appendObject_('Shares', { id: uuid_(), substationCode: code,
        username: u.username, level: level, createdAt: now_() });
  return { ok: true, username: u.username, level: level };
}

function apiListShares(token, code) {
  var s = requireSession_(token);
  if (!canGrant_(s, code)) throw new Error('Not permitted.');
  return readObjects_('Shares').filter(function (sh) { return sh.substationCode === code; })
    .map(function (sh) { return { id: sh.id, username: sh.username,
        level: String(sh.level || 'viewer').toLowerCase(), createdAt: sh.createdAt }; });
}

function apiRevokeShare(token, code, shareId) {
  var s = requireSession_(token);
  if (!canGrant_(s, code)) throw new Error('Not permitted.');
  var sh = readObjects_('Shares').filter(function (x) { return x.id === shareId; })[0];
  if (!sh || sh.substationCode !== code) throw new Error('Share not found.');
  deleteById_('Shares', shareId);
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

/* Last-operated is DERIVED from the Operations log, never typed by hand.
   Returns a map: ltsId -> most recent interruptFrom (ISO string). */
function lastOpMap_(code) {
  var map = {};
  readObjects_('Operations').forEach(function (o) {
    if (o.substationCode !== code) return;
    var t = o.interruptFrom;
    if (!t) return;
    if (!map[o.ltsId] || new Date(t) > new Date(map[o.ltsId])) map[o.ltsId] = t;
  });
  return map;
}

function apiGetLTSList(token, code) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access to this substation.');
  var editor = canEdit_(s, code);
  var list = readObjects_('LTS').filter(function (l) { return l.substationCode === code; });
  if (!editor) list = list.filter(function (l) { return String(l.visibleToViewer).toLowerCase() === 'yes'; });
  list.sort(function (a, b) { return (Number(a.srNo) || 0) - (Number(b.srNo) || 0); });
  var lom = lastOpMap_(code);
  list.forEach(function (l) { l.lastOpDate = lom[l.id] || ''; });
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
  l.lastOpDate = lastOpMap_(code)[id] || '';
  ensureBaselineVersion_(l);
  return { lts: l, customFields: cf, uploads: up, versions: shapeVersions_(id),
           canEdit: canEdit_(s, code) };
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

/* ================================================================== *
 *  LTS VERSIONING
 *
 *  Every revision of an LTS setting is preserved. The newest version is
 *  the live one; it is mirrored into the LTS row so all existing reads
 *  (list, detail, exports) keep working unchanged.
 *
 *  functionalFrom  — supplied by the editor ("functional from" date)
 *  functionalTill  — derived: the functionalFrom of the NEXT version,
 *                    or blank ("till date") for the current version.
 *  Attachments are bound to the version that was current when uploaded.
 * ================================================================== */

/* newest first */
function versionsFor_(ltsId) {
  return readObjects_('LTSVersions')
    .filter(function (v) { return v.ltsId === ltsId; })
    .sort(function (a, b) { return Number(b.versionNo) - Number(a.versionNo); });
}

function currentVersion_(ltsId) {
  var vs = versionsFor_(ltsId);
  return vs.length ? vs[0] : null;
}

/* shape a stored version row for the client, computing functionalTill */
function shapeVersions_(ltsId) {
  var vs = versionsFor_(ltsId);          /* newest first */
  var ups = readObjects_('Uploads').filter(function (u) { return u.ltsId === ltsId; });
  return vs.map(function (v, i) {
    var next = vs[i - 1];                /* the version that superseded this one */
    return {
      id: v.id,
      versionNo: Number(v.versionNo),
      isCurrent: i === 0,
      functionalFrom: v.functionalFrom || '',
      functionalTill: i === 0 ? '' : (next ? next.functionalFrom : ''),
      data: safeJson_(v.data, {}),
      customFields: safeJson_(v.customFields, []),
      approvalRemarks: v.approvalRemarks || '',
      approvalFileId: v.approvalFileId || '',
      approvalFileName: v.approvalFileName || '',
      createdBy: v.createdBy,
      createdAt: v.createdAt,
      uploads: ups.filter(function (u) { return u.versionId === v.id; })
        .map(function (u) { return { id: u.id, driveFileId: u.driveFileId, fileName: u.fileName,
              caption: u.caption, mimeType: u.mimeType }; })
    };
  });
}

function safeJson_(s, dflt) {
  if (s === undefined || s === null || s === '') return dflt;
  if (typeof s === 'object') return s;
  try { return JSON.parse(s); } catch (e) { return dflt; }
}

function apiGetVersions(token, code, ltsId) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  return shapeVersions_(ltsId);
}

/**
 * Save a NEW version of an LTS. Becomes the current version; the previous
 * one is closed off at this version's functionalFrom date.
 *
 * payload = { ltsId, functionalFrom, fields:{...}, customFields:[{fieldName,fieldValue}],
 *             approvalRemarks, approval:{base64,mimeType,fileName} | null }
 */
function apiSaveLTSVersion(token, code, payload) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');

  var ltsId = payload.ltsId;
  var lts = readObjects_('LTS').filter(function (x) { return x.id === ltsId; })[0];
  if (!lts) throw new Error('LTS not found.');
  if (lts.substationCode !== code) throw new Error('LTS does not belong to this substation.');

  var from = String(payload.functionalFrom || '').trim();
  if (!from) throw new Error('"Functional from" date is required.');

  var prev = currentVersion_(ltsId);
  if (prev && prev.functionalFrom && new Date(from) < new Date(prev.functionalFrom)) {
    throw new Error('"Functional from" cannot precede the current version (' +
      fmtD_(prev.functionalFrom) + ').');
  }
  if (!payload.approvalRemarks || !String(payload.approvalRemarks).trim()) {
    throw new Error('Approval remarks are required when a setting is revised.');
  }

  /* optional approval copy -> Drive */
  var appFileId = '', appFileName = '';
  if (payload.approval && payload.approval.base64) {
    var f = saveToDrive_(code, payload.approval);
    appFileId = f.getId();
    appFileName = payload.approval.fileName || f.getName();
  }

  var fields = {};
  LTS_FIELDS.forEach(function (k) {
    if (k === 'lastOpDate') return;                     /* derived, never versioned */
    fields[k] = payload.fields[k] === undefined ? '' : payload.fields[k];
  });

  var vid = uuid_();
  var vno = prev ? Number(prev.versionNo) + 1 : 1;
  appendObject_('LTSVersions', {
    id: vid, ltsId: ltsId, substationCode: code, versionNo: vno,
    functionalFrom: from, functionalTill: '',
    data: JSON.stringify(fields),
    customFields: JSON.stringify(payload.customFields || []),
    approvalRemarks: String(payload.approvalRemarks).trim(),
    approvalFileId: appFileId, approvalFileName: appFileName,
    createdBy: s.username, createdAt: now_()
  });

  /* close the previous version's validity window */
  if (prev) updateRowById_('LTSVersions', prev.id, { functionalTill: from });

  /* mirror the new version into the live LTS row */
  var mirror = {};
  Object.keys(fields).forEach(function (k) { mirror[k] = fields[k]; });
  mirror.updatedAt = now_();
  updateRowById_('LTS', ltsId, mirror);

  /* rebuild the live custom-field rows to match this version */
  readObjects_('CustomFields').filter(function (x) { return x.ltsId === ltsId; })
    .forEach(function (x) { deleteById_('CustomFields', x.id); });
  (payload.customFields || []).forEach(function (c) {
    if (!c.fieldName) return;
    appendObject_('CustomFields', { id: uuid_(), ltsId: ltsId,
      fieldName: c.fieldName, fieldValue: c.fieldValue || '', createdAt: now_() });
  });

  return { id: vid, versionNo: vno };
}

/* seed v1 for LTS rows that pre-date versioning, so history is never empty */
function ensureBaselineVersion_(lts) {
  if (versionsFor_(lts.id).length) return;
  var fields = {};
  LTS_FIELDS.forEach(function (k) { if (k !== 'lastOpDate') fields[k] = lts[k] === undefined ? '' : lts[k]; });
  var cfs = readObjects_('CustomFields').filter(function (x) { return x.ltsId === lts.id; })
    .map(function (x) { return { fieldName: x.fieldName, fieldValue: x.fieldValue }; });
  appendObject_('LTSVersions', {
    id: uuid_(), ltsId: lts.id, substationCode: lts.substationCode, versionNo: 1,
    functionalFrom: lts.date || lts.createdAt || now_(), functionalTill: '',
    data: JSON.stringify(fields), customFields: JSON.stringify(cfs),
    approvalRemarks: 'Original approved setting (baseline record).',
    approvalFileId: '', approvalFileName: '',
    createdBy: 'system', createdAt: lts.createdAt || now_()
  });
}

/** Run once after upgrading: creates v1 for every existing LTS. */
function migrateBaselineVersions() {
  var n = 0;
  readObjects_('LTS').forEach(function (l) {
    if (!versionsFor_(l.id).length) { ensureBaselineVersion_(l); n++; }
  });
  invalidate_();
  var msg = 'Baseline versions created: ' + n;
  Logger.log(msg);
  return msg;
}

/* ================================================================== *
 *  DIAGRAMS  (per-LTS, user-drawn scene graph)
 * ================================================================== */
function apiGetDiagram(token, code, ltsId) {
  var s = requireSession_(token);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  var d = readObjects_('Diagrams').filter(function (x) { return x.ltsId === ltsId; })[0];
  return d ? { json: safeJson_(d.json, null), updatedBy: d.updatedBy, updatedAt: d.updatedAt }
           : { json: null };
}

function apiSaveDiagram(token, code, ltsId, json) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var str = typeof json === 'string' ? json : JSON.stringify(json);
  if (str.length > 45000) throw new Error('Diagram is too large to store in one cell. Simplify it.');
  var d = readObjects_('Diagrams').filter(function (x) { return x.ltsId === ltsId; })[0];
  if (d) updateRowById_('Diagrams', d.id, { json: str, updatedBy: s.username, updatedAt: now_() });
  else appendObject_('Diagrams', { id: uuid_(), ltsId: ltsId, substationCode: code,
        json: str, updatedBy: s.username, updatedAt: now_() });
  return { ok: true };
}

function apiDeleteDiagram(token, code, ltsId) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var d = readObjects_('Diagrams').filter(function (x) { return x.ltsId === ltsId; })[0];
  if (d) deleteById_('Diagrams', d.id);
  return { ok: true };
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
  readObjects_('LTSVersions').filter(function (x) { return x.ltsId === id; }).forEach(function (x) {
    if (x.approvalFileId) { try { DriveApp.getFileById(x.approvalFileId).setTrashed(true); } catch (e) {} }
    deleteById_('LTSVersions', x.id);
  });
  readObjects_('Diagrams').filter(function (x) { return x.ltsId === id; })
    .forEach(function (x) { deleteById_('Diagrams', x.id); });
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
/* file = {base64, mimeType, fileName} */
function saveToDrive_(code, file) {
  var bytes = Utilities.base64Decode(file.base64);
  var blob = Utilities.newBlob(bytes, file.mimeType, file.fileName);
  return subFolder_(code).createFile(blob);
}
function apiUploadFile(token, code, ltsId, base64, fileName, mimeType, caption) {
  var s = requireSession_(token);
  if (!canEdit_(s, code)) throw new Error('Editor permission required.');
  var file = saveToDrive_(code, { base64: base64, mimeType: mimeType, fileName: fileName });
  /* attachments belong to whichever version is current at upload time */
  var lts = readObjects_('LTS').filter(function (x) { return x.id === ltsId; })[0];
  if (lts) ensureBaselineVersion_(lts);
  var cur = currentVersion_(ltsId);
  var id = uuid_();
  appendObject_('Uploads', { id: id, ltsId: ltsId, versionId: cur ? cur.id : '',
    substationCode: code, driveFileId: file.getId(),
    fileName: fileName, caption: caption || fileName, mimeType: mimeType, createdAt: now_() });
  return { id: id, driveFileId: file.getId() };
}
function apiGetFile(token, fileId) {
  var s = requireSession_(token);
  var code = null, name = '', mime = '';

  var meta = readObjects_('Uploads').filter(function (x) { return x.driveFileId === fileId; })[0];
  if (meta) { code = meta.substationCode; name = meta.fileName; mime = meta.mimeType; }
  else {
    /* approval copies live on the version row, not in Uploads */
    var v = readObjects_('LTSVersions').filter(function (x) { return x.approvalFileId === fileId; })[0];
    if (!v) throw new Error('File not found.');
    code = v.substationCode; name = v.approvalFileName || 'approval';
  }
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  var blob = DriveApp.getFileById(fileId).getBlob();
  return { base64: Utilities.base64Encode(blob.getBytes()),
           mimeType: mime || blob.getContentType(), fileName: name };
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
/* opts = { from:'YYYY-MM-DD', to:'YYYY-MM-DD', ltsId:'...' }  — all optional */
function filterOpsRange_(ops, opts) {
  opts = opts || {};
  if (opts.ltsId) ops = ops.filter(function (o) { return o.ltsId === opts.ltsId; });
  if (opts.from) {
    var f = new Date(opts.from);
    ops = ops.filter(function (o) { return new Date(o.interruptFrom) >= f; });
  }
  if (opts.to) {
    var t = new Date(opts.to + 'T23:59:59');
    ops = ops.filter(function (o) { return new Date(o.interruptFrom) <= t; });
  }
  return ops;
}
function rangeSuffix_(opts) {
  opts = opts || {};
  if (opts.from || opts.to) {
    return ' (' + (opts.from ? fmtD_(opts.from) : 'start') + ' to ' + (opts.to ? fmtD_(opts.to) : 'date') + ')';
  }
  return '';
}

function apiExport(token, code, format, scope, id, opts) {
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
      title = 'LTS Operations — ' + code + rangeSuffix_(opts);
      matrix = opsMatrix_(filterOpsRange_(apiGetOperations(token, code), opts));
    } else if (scope === 'ltsops') {
      /* every operation of ONE scheme, optionally date-bounded */
      var o2 = (opts || {}); o2.ltsId = id;
      var rows = filterOpsRange_(apiGetOperations(token, code), o2);
      var nm = rows.length ? rows[0].ltsName : '';
      if (!nm) {
        var lx = readObjects_('LTS').filter(function (x) { return x.id === id; })[0];
        nm = lx ? lx.tfIct : id;
      }
      title = 'Operation history — ' + nm + rangeSuffix_(opts);
      matrix = opsMatrix_(rows);
    } else if (scope === 'version') {
      /* id arrives as "<ltsId>|<versionId>" */
      var parts = String(id).split('|');
      var v = shapeVersions_(parts[0]).filter(function (x) { return x.id === parts[1]; })[0];
      if (!v) throw new Error('Version not found.');
      title = 'LTS Version ' + v.versionNo + ' — ' + (v.data.tfIct || '');
      matrix = versionMatrix_(v);
    } else if (scope === 'op1') {
      var one = apiGetOperations(token, code).filter(function (o) { return o.id === id; });
      if (!one.length) throw new Error('Operation not found.');
      title = 'LTS Operation — ' + one[0].ltsName;
      matrix = opsMatrix_(one);
    } else {
      title = 'Load Trimming Schemes — ' + code;
      var lom = lastOpMap_(code);
      matrix = ltsListMatrix_(readObjects_('LTS').filter(function (l) { return l.substationCode === code; })
        .sort(function (a, b) { return (Number(a.srNo) || 0) - (Number(b.srNo) || 0); })
        .map(function (l) { l.lastOpDate = lom[l.id] || ''; return l; }));
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

/* ---- Golden & White export theme ---- */
var GOLD       = '#B8912F';   /* header fill / rules      */
var GOLD_DEEP  = '#7A5E12';   /* title text               */
var GOLD_TINT  = '#FBF6E6';   /* alternate row band       */
var GOLD_LINE  = '#E2D3A2';   /* inner borders            */
var INK        = '#2B2410';   /* body text                */

/* header row is matrix[0]; subsequent rows are data */
function writeFormatted_(sh, title, matrix) {
  var cols = matrix[0].length;
  var rows = matrix.length;

  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, rows + 4, cols).setBackground('#FFFFFF');

  /* ---- title band: white with deep-gold text + gold rule ---- */
  sh.getRange(1, 1, 1, cols).merge().setValue(title)
    .setFontSize(15).setFontWeight('bold').setFontFamily('Georgia')
    .setFontColor(GOLD_DEEP).setBackground('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 38);

  /* subtitle / provenance line */
  sh.getRange(2, 1, 1, cols).merge()
    .setValue('MSETCL — Load Trimming Scheme Register    ·    Generated ' +
              Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm') + ' hrs')
    .setFontSize(8).setFontColor('#8A7A55').setBackground('#FFFFFF')
    .setHorizontalAlignment('center');
  sh.setRowHeight(2, 16);
  /* gold rule under the title block */
  sh.getRange(2, 1, 1, cols).setBorder(null, null, true, null, null, null, GOLD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var startRow = 4;
  var body = sh.getRange(startRow, 1, rows, cols);
  body.setValues(matrix).setWrap(true).setVerticalAlignment('top')
      .setFontFamily('Calibri').setFontSize(9).setFontColor(INK)
      .setBorder(true, true, true, true, true, true, GOLD_LINE, SpreadsheetApp.BorderStyle.SOLID);

  /* ---- header row: gold fill, white bold text ---- */
  sh.getRange(startRow, 1, 1, cols)
    .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(9)
    .setBackground(GOLD).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, GOLD, SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(startRow, 30);

  /* ---- zebra banding in gold tint ---- */
  for (var r = 1; r < rows; r++) {
    if (r % 2 === 0) sh.getRange(startRow + r, 1, 1, cols).setBackground(GOLD_TINT);
  }

  /* first column emphasised (Sr / field name) */
  if (rows > 1) {
    sh.getRange(startRow + 1, 1, rows - 1, 1).setFontWeight('bold').setFontColor(GOLD_DEEP);
  }

  /* closing gold rule */
  sh.getRange(startRow + rows - 1, 1, 1, cols)
    .setBorder(null, null, true, null, null, null, GOLD, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  /* smart-ish column widths: sample content length, capped */
  for (var c = 1; c <= cols; c++) {
    var maxLen = 8;
    for (var r2 = 0; r2 < rows; r2++) {
      var v = String(matrix[r2][c - 1] || '');
      v.split('\n').forEach(function (line) { if (line.length > maxLen) maxLen = line.length; });
    }
    sh.setColumnWidth(c, Math.min(360, Math.max(70, maxLen * 7 + 16)));
  }
  sh.setFrozenRows(startRow);
}

function ltsListMatrix_(list) {
  var head = ['Sr','End','ICT / Line','HV/LV','CTR','Relay','Max Load','Alarm PSM',
    'Stage-1 Trip','Stage-2 Trip','Feeders in scheme','Load Relief (Obtd/Reqd)',
    'Remarks','Healthy','Last Op','Status','Date'];
  var rows = list.map(function (l) {
    return [l.srNo, l.endType === 'other' ? 'Other' : 'Own', l.tfIct, l.hvLv, l.ctr, l.relayType,
      l.maxLoad, l.alarmPsm, l.stage1, l.stage2, l.feeders, l.loadRelief, l.remarks,
      l.healthiness, l.lastOpDate ? fmtD_(l.lastOpDate) : 'Never', l.status, l.date]; });
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
    ['Remarks / purpose', l.remarks], ['Healthiness', l.healthiness],
    ['Last operated (from Operations log)', l.lastOpDate ? fmtDt_(l.lastOpDate) : 'Never operated']];
  d.customFields.forEach(function (c) { pairs.push([c.fieldName, c.fieldValue]); });
  return pairs;
}
function versionMatrix_(v) {
  var d = v.data || {};
  var pairs = [['Field', 'Value'],
    ['Version', 'v' + v.versionNo + (v.isCurrent ? ' (current)' : '')],
    ['Functional from', v.functionalFrom ? fmtD_(v.functionalFrom) : ''],
    ['Functional till', v.isCurrent ? 'till date' : (v.functionalTill ? fmtD_(v.functionalTill) : '')],
    ['Sr No', d.srNo], ['Scheme end', d.endType === 'other' ? 'Other substation end' : 'Own substation end'],
    ['Division', d.division], ['Substation', d.substation], ['Status', d.status], ['Date', d.date],
    ['ICT / Line', d.tfIct], ['HV/LV', d.hvLv], ['Connected CTR', d.ctr], ['Relay', d.relayType],
    ['Max load recorded', d.maxLoad], ['Alarm PSM setting', d.alarmPsm],
    ['Stage-1 (Trip)', d.stage1], ['Stage-2 (Trip)', d.stage2],
    ['Feeders in scheme', d.feeders], ['Load relief (obtd/reqd)', d.loadRelief],
    ['Remarks / purpose', d.remarks], ['Healthiness', d.healthiness]];
  (v.customFields || []).forEach(function (c) { pairs.push([c.fieldName, c.fieldValue]); });
  pairs.push(['Approval remarks', v.approvalRemarks || '']);
  pairs.push(['Approval copy', v.approvalFileName || 'not attached']);
  pairs.push(['Revised by', v.createdBy]);
  pairs.push(['Revised on', fmtDt_(v.createdAt)]);
  return pairs;
}
function opsMatrix_(ops) {
  var head = ['LTS', 'Operation', 'Interruption From', 'Interruption To', 'Duration (hh:mm)', 'Details', 'Logged by'];
  var rows = ops.slice().sort(function (a, b) {
    return new Date(b.interruptFrom) - new Date(a.interruptFrom);   /* newest first */
  }).map(function (o) {
    var text = (o.blocks || []).map(function (b) {
      return (b.label ? b.label + ': ' : '') + stripHtml_(b.html); }).join('\n\n');
    return [o.ltsName, o.opTitle, fmtDt_(o.interruptFrom), fmtDt_(o.interruptTo),
            fmtDur_(o.durationMin), text, o.createdBy];
  });
  return [head].concat(rows);
}
function stripHtml_(h) { return String(h || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').trim(); }
/* 12-Jan-2026, 09:36 hrs */
function fmtDt_(iso) {
  if (!iso) return '';
  var d = new Date(iso);
  return isNaN(d) ? String(iso)
    : Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy, HH:mm') + ' hrs';
}
/* 12-Jan-2026 */
function fmtD_(iso) {
  if (!iso) return '';
  var d = new Date(iso);
  return isNaN(d) ? String(iso) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
}
/* minutes -> hh:mm (hours not capped at 24) */
function fmtDur_(min) {
  var m = Number(min);
  if (!isFinite(m) || m < 0) return '';
  m = Math.round(m);
  var h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
}

function exportBlob_(ssId, format) {
  var gid = SpreadsheetApp.openById(ssId).getSheets()[0].getSheetId();
  var params = format === 'pdf'
    ? 'format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&sheetnames=false' +
      '&printtitle=false&pagenumbers=true&fzr=true' +
      '&top_margin=0.40&bottom_margin=0.40&left_margin=0.40&right_margin=0.40&gid=' + gid
    : 'format=xlsx';
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?' + params;
  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var rc = resp.getResponseCode();
  if (rc !== 200) {
    throw new Error('Export failed (HTTP ' + rc + '). Re-authorise the script and confirm Drive access.');
  }
  var blob = resp.getBlob();
  /* Google sometimes returns an HTML error body with a 200 — detect it */
  var ct = String(blob.getContentType() || '');
  if (ct.indexOf('text/html') === 0) {
    throw new Error('Export returned an error page instead of a file. Re-run authorisation.');
  }
  blob.setContentType(format === 'pdf'
    ? 'application/pdf'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  return blob;
}

/* ================================================================== *
 *  ANNOTATIONS — highlight rectangles on an attachment, stored as
 *  fractions of page size so they survive any zoom level.
 * ================================================================== */
function fileSubCode_(fileId) {
  var u = readObjects_('Uploads').filter(function (x) { return x.driveFileId === fileId; })[0];
  if (u) return u.substationCode;
  var v = readObjects_('LTSVersions').filter(function (x) { return x.approvalFileId === fileId; })[0];
  if (v) return v.substationCode;
  throw new Error('File not found.');
}

function apiGetAnnotations(token, fileId) {
  var s = requireSession_(token);
  var code = fileSubCode_(fileId);
  if (accessibleCodes_(s).indexOf(code) < 0) throw new Error('No access.');
  return readObjects_('Annotations')
    .filter(function (a) { return a.driveFileId === fileId; })
    .map(function (a) {
      return { id: a.id, page: Number(a.page) || 1,
        x: Number(a.x), y: Number(a.y), w: Number(a.w), h: Number(a.h),
        color: a.color || '#ffd54a', note: a.note || '',
        createdBy: a.createdBy, createdAt: a.createdAt };
    });
}

/* a = {page,x,y,w,h,color,note} with x/y/w/h as 0..1 fractions */
function apiSaveAnnotation(token, fileId, a) {
  var s = requireSession_(token);
  var code = fileSubCode_(fileId);
  if (!canEdit_(s, code)) throw new Error('Editor permission required to annotate.');
  var id = uuid_();
  appendObject_('Annotations', { id: id, driveFileId: fileId, substationCode: code,
    page: a.page || 1, x: a.x, y: a.y, w: a.w, h: a.h,
    color: a.color || '#ffd54a', note: a.note || '',
    createdBy: s.username, createdAt: now_() });
  return { id: id };
}

function apiDeleteAnnotation(token, id) {
  var s = requireSession_(token);
  var a = readObjects_('Annotations').filter(function (x) { return x.id === id; })[0];
  if (!a) throw new Error('Highlight not found.');
  if (!canEdit_(s, a.substationCode)) throw new Error('Editor permission required.');
  deleteById_('Annotations', id);
  return { ok: true };
}

/* ==================================================================
 *  MAINTENANCE — run manually from the Apps Script editor
 * ================================================================== */

/**
 * One-shot: rename the Karjat substation code K400 -> J966 across every
 * sheet that stores a substation code. Safe to run twice (idempotent).
 * Run this ONCE from the editor, then reload the portal.
 */
function migrateKarjatCodeToJ966() {
  renameSubstationCode('K400', 'J966');
}

/**
 * Adds any columns introduced after your DB was first created.
 * Safe to run repeatedly. Run once after updating Code.gs.
 */
function migrateSchema() {
  var added = [];
  var ss = getDb_();

  /* new tables */
  var newTables = {
    LTSVersions: ['id','ltsId','substationCode','versionNo','functionalFrom','functionalTill',
                  'data','customFields','approvalRemarks','approvalFileId','approvalFileName',
                  'createdBy','createdAt'],
    Diagrams:    ['id','ltsId','substationCode','json','updatedBy','updatedAt'],
    Annotations: ['id','driveFileId','substationCode','page','x','y','w','h','color','note',
                  'createdBy','createdAt']
  };
  Object.keys(newTables).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.appendRow(newTables[name]); added.push(name + ' (table)'); }
  });

  if (addColumnIfMissing_('Shares', 'level', 'createdAt')) added.push('Shares.level');
  if (addColumnIfMissing_('Uploads', 'versionId', 'substationCode')) added.push('Uploads.versionId');

  /* existing shares were read-only before levels existed */
  var sh2 = sheet_('Shares');
  var head = headers_('Shares');
  var li = head.indexOf('level');
  if (li >= 0 && sh2.getLastRow() > 1) {
    var rng = sh2.getRange(2, li + 1, sh2.getLastRow() - 1, 1);
    var vals = rng.getValues();
    var changed = false;
    for (var i = 0; i < vals.length; i++) {
      if (!String(vals[i][0]).trim()) { vals[i][0] = 'viewer'; changed = true; }
    }
    if (changed) rng.setValues(vals);
  }
  invalidate_();

  /* baseline v1 for every pre-existing LTS, then bind orphan uploads to it */
  migrateBaselineVersions();
  var ups = readObjects_('Uploads').filter(function (u) { return !String(u.versionId).trim(); });
  ups.forEach(function (u) {
    var cur = currentVersion_(u.ltsId);
    if (cur) updateRowById_('Uploads', u.id, { versionId: cur.id });
  });
  if (ups.length) added.push(ups.length + ' upload(s) bound to v1');

  invalidate_();
  var msg = added.length ? 'Added: ' + added.join(', ') : 'Schema already current';
  Logger.log(msg);
  return msg;
}

/* insert `col` immediately before `beforeCol` (or append) if not present */
function addColumnIfMissing_(sheetName, col, beforeCol) {
  var sh = sheet_(sheetName);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (head.indexOf(col) >= 0) return false;
  var at = beforeCol ? head.indexOf(beforeCol) : -1;
  var pos = at >= 0 ? at + 1 : head.length + 1;
  sh.insertColumnBefore(pos);
  sh.getRange(1, pos).setValue(col);
  invalidate_(sheetName);
  return true;
}

function renameSubstationCode(oldCode, newCode) {
  var ss = getDb_();
  var touched = [];
  [['Substations', 'code'],
   ['Users', 'substationCode'],
   ['Shares', 'substationCode'],
   ['LTS', 'substationCode'],
   ['Operations', 'substationCode'],
   ['Uploads', 'substationCode'],
   ['CustomFields', 'substationCode']].forEach(function (pair) {
    var name = pair[0], col = pair[1];
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return;
    var ci = vals[0].indexOf(col);
    if (ci < 0) return;
    var n = 0;
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][ci]) === oldCode) { vals[r][ci] = newCode; n++; }
    }
    if (n) {
      sh.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
      touched.push(name + ':' + n);
    }
  });
  SpreadsheetApp.flush();
  invalidate_();
  var msg = touched.length ? 'Updated -> ' + touched.join(', ') : 'Nothing to update (already migrated?)';
  Logger.log(msg);
  return msg;
}

/**
 * ONE-CLICK REPAIR — run this from the editor after any upgrade.
 * Idempotent: safe to run as many times as you like.
 */
function repairAll() {
  var log = [];
  log.push('rename: ' + renameSubstationCode('K400', 'J966'));
  log.push('schema: ' + migrateSchema());
  log.push('dedupe: ' + dedupeSubstations());
  invalidate_();
  var msg = log.join('\n');
  Logger.log(msg);
  return msg;
}

/** Remove duplicate Substation rows sharing a code (keeps the oldest). */
function dedupeSubstations() {
  var seen = {}, killed = 0;
  readObjects_('Substations')
    .sort(function (a, b) { return new Date(a.createdAt) - new Date(b.createdAt); })
    .forEach(function (s) {
      if (seen[s.code]) { deleteById_('Substations', s.id); killed++; }
      else seen[s.code] = true;
    });
  invalidate_();
  return killed ? killed + ' duplicate substation row(s) removed' : 'no duplicates';
}

/** Prints what the DB actually contains — paste the output to me if stuck. */
function diagnose() {
  var subs = readObjects_('Substations');
  var lts = readObjects_('LTS');
  var out = ['DB_ID=' + props_().getProperty('DB_ID'),
    'Substations: ' + subs.map(function (s) { return s.code + ' (' + s.name + ', ' + s.status + ')'; }).join(' | '),
    'LTS rows by code: ' + JSON.stringify(lts.reduce(function (a, l) {
      a[l.substationCode] = (a[l.substationCode] || 0) + 1; return a; }, {})),
    'Versions: ' + readObjects_('LTSVersions').length,
    'Operations: ' + readObjects_('Operations').length,
    'Diagrams: ' + readObjects_('Diagrams').length];
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}
