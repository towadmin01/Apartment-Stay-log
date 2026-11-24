/**
 * TOW Apartment Stay System – Google Apps Script backend
 * Paste into the Apartment Stay Log Google Sheet: Extensions → Apps Script.
 * Setup steps: see README.md.
 */

const CONFIG = {
  SITE_URL: 'https://YOUR-GITHUB-USERNAME.github.io/tow-apartment-log/',   // optional: leave as is and the script learns your site address automatically
  CALENDAR_ID: '570299c8872b059e63a109013c0de95228d63a90c299b670f5439a6a9bb5fa60@group.calendar.google.com',
  TZ: 'Asia/Kuala_Lumpur',
  SYNC_CALENDAR: true,
  FOLDER_NAME: 'TOW Apartment Records'
};

const STAY = 'Stay Log', HOST = 'Hosting', APPR = 'Approvals', MIN = 'Ministries';
const HEADS = {};
HEADS[STAY] = ["Stay ID","Status","Category","Name of User","Main Guest","Purpose of Use","Ministry","Ministry Head","Attending Personnel","Contact No.","Unit No.","No. of Guests","Arrival","Departure","Nights","Vehicle Details","Church Transport","Transport PIC","Guest Preferences / Special Notes","Love Gift (RM)","Love Gift Ref","Special Arrangement","Step 1: Assign Cleaners","Step 2: Unit Preparation","Step 3: Guest Essentials","Step 4: Before Check-in","Step 5: Meal Arrangement","Step 6: After Check-out","Final Checked By","Final Check Date","Additional Notes","Source","Approval ID","Approved By","Approved At","Submitted At","Signature","PDF","Acknowledged At","Hosting ID","Calendar Event ID","Last Updated"];
HEADS[HOST] = ["Hosting ID","Status","Guest Name","Church / Ministry From","Event Date(s)","Attending Personnel","PIC Contact","Hosting Pastor","Stay Option","Arrival","Departure","Mode of Travel","Travel Details","Transport Needed","Vehicle Details","Accompanying Persons","Apartment Guests","Dietary & Meal Schedule","Health Notes","Account Holder","Bank Name","Account Number","Mobile Number","Checklist","Linked Stay ID","Created At","Submitted At","Sent to Accounts At","PDF","Admin Notes","Event Date","PIC Email","Needs Apartment","Reminder 14d Sent","Reminder 3d Sent","Calendar Event ID","Secret"];
HEADS[APPR] = ["Approval ID","Ministry","Ministry Head","Member Name","Category","Check-in","Check-out","Created At","Used By Stay","Notes"];
HEADS[MIN] = ["Ministry","Ministry Head","WhatsApp","Notes","Passcode"];
const DATE_FIELDS = ["Arrival","Departure","Final Check Date","Approved At","Submitted At","Acknowledged At","Last Updated","Created At","Sent to Accounts At","Check-in","Check-out","Event Date","Reminder 14d Sent","Reminder 3d Sent"];
const RATES = { "Pastors, Guest Speakers & Mission Pastors": 0, "Member serving that weekend": 50, "Member, personal or non-serving stay": 200, "Non-member": 250 };
const NO_CALENDAR = ['Requested', 'Cancelled'];

/* =================== ONE-TIME SETUP (run in the editor) =================== */

function setup() {
  [STAY, HOST, APPR, MIN].forEach(name => ensureSheet_(name));
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('ADMIN_PASSCODE')) p.setProperty('ADMIN_PASSCODE', 'tow-admin');
  if (!p.getProperty('HEAD_PASSCODE')) p.setProperty('HEAD_PASSCODE', 'tow-heads');
  if (!p.getProperty('ADMIN_EMAIL')) p.setProperty('ADMIN_EMAIL', Session.getEffectiveUser().getEmail());
  folder_();
  Logger.log('Ready. Admin passcode: %s · Ministry Head passcode: %s · Alerts email: %s',
    p.getProperty('ADMIN_PASSCODE'), p.getProperty('HEAD_PASSCODE'), p.getProperty('ADMIN_EMAIL'));
}

/** Edit the values, then run once. */
function setPasscodes() {
  const ADMIN = 'CHANGE-ADMIN-CODE', HEADS_CODE = 'CHANGE-HEAD-CODE', EMAIL = '';
  const p = PropertiesService.getScriptProperties();
  p.setProperty('ADMIN_PASSCODE', ADMIN);
  p.setProperty('HEAD_PASSCODE', HEADS_CODE);
  if (EMAIL) p.setProperty('ADMIN_EMAIL', EMAIL);
  Logger.log('Saved.');
}

/** 1) Paste your bot token below and run. 2) Send any message to your bot (or add it to your admin group and say hi). 3) Run again. */
function setupTelegram() {
  const TOKEN = 'PASTE-BOT-TOKEN-HERE';
  const p = PropertiesService.getScriptProperties();
  if (TOKEN.indexOf('PASTE') < 0) p.setProperty('TELEGRAM_TOKEN', TOKEN);
  const token = p.getProperty('TELEGRAM_TOKEN');
  if (!token) { Logger.log('Paste your bot token first.'); return; }
  const res = JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/getUpdates', { muteHttpExceptions: true }).getContentText());
  const upd = (res.result || []).reverse().find(u => (u.message || u.my_chat_member || {}).chat);
  if (!upd) { Logger.log('No messages yet. Send "hi" to your bot (or in the group), then run setupTelegram again.'); return; }
  const chat = (upd.message || upd.my_chat_member).chat;
  p.setProperty('TELEGRAM_CHAT_ID', String(chat.id));
  telegram_('✅ <b>TOW Apartment alerts connected.</b>\nNew requests and hosting forms will appear here.');
  Logger.log('Connected to chat: %s (%s)', chat.title || chat.first_name, chat.id);
}

/** Turns on the daily automatic backup (runs around 2am). Run once. */
function setupAutoBackup() { return setupDailyJobs(); }
function setupAutoBackupOnly_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'autoBackup').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoBackup').timeBased().everyDays(1).atHour(2).inTimezone(CONFIG.TZ).create();
  const b = backupNow_(true);
  Logger.log('Daily backup is on. First backup: %s', b.url);
}
function autoBackup() {
  try { backupNow_(true); }
  catch (e) { notify_('⚠️ Apartment backup failed', [String(e)], null); }
}
function backupNow_(auto) {
  const ss = SpreadsheetApp.getActive();
  const root = folder_();
  const it = root.getFoldersByName('Backups');
  const bf = it.hasNext() ? it.next() : root.createFolder('Backups');
  const stamp = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HHmm');
  let file;
  try {
    const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName('TOW Apartment Stay Log ' + stamp + '.xlsx');
    file = bf.createFile(blob);
  } catch (e) {
    file = DriveApp.getFileById(ss.getId()).makeCopy('TOW Apartment Stay Log ' + stamp, bf);
  }
  // keep the latest 30 backups
  const files = []; const fi = bf.getFiles(); while (fi.hasNext()) files.push(fi.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated()).slice(30).forEach(f => f.setTrashed(true));
  const info = { at: Utilities.formatDate(new Date(), CONFIG.TZ, "yyyy-MM-dd'T'HH:mm"), url: file.getUrl(), name: file.getName(), auto: !!auto };
  PropertiesService.getScriptProperties().setProperty('LAST_BACKUP', JSON.stringify(info));
  return info;
}

function testAlert() { notify_('🔔 Test alert', ['This is a test from the TOW Apartment system.'], null); }

/** Optional: adds calendar events for upcoming stays entered in the log that have none yet (skips imported and past stays). */
function syncUpcomingToCalendar() {
  const now = new Date();
  table_(STAY).rows.forEach(r => {
    if (NO_CALENDAR.indexOf(r['Status']) < 0 && !r['Calendar Event ID'] && r['Departure'] instanceof Date && r['Departure'] >= now) saveStay_(plain_(r));
  });
}

/* =================== WEB APP =================== */

function doGet() { return json_({ ok: true, app: 'TOW Apartment API' }); }

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    const p = PropertiesService.getScriptProperties();
    upgradeSheets_();
    // Site address: use SITE_URL if you set it; otherwise learn it from the pages themselves (no editing needed)
    if (/YOUR-GITHUB/i.test(CONFIG.SITE_URL)) {
      if (b.site && /^https:\/\/[^\s"'<>]+\/$/.test(b.site) && !/localhost|script\.google/.test(b.site) && b.site !== p.getProperty('SITE_URL')) p.setProperty('SITE_URL', b.site);
      if (p.getProperty('SITE_URL')) CONFIG.SITE_URL = p.getProperty('SITE_URL');
    }
    const isAdmin = b.key && b.key === p.getProperty('ADMIN_PASSCODE');
    const isHead = isAdmin || (b.hkey && b.hkey === p.getProperty('HEAD_PASSCODE'));
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try {
      switch (b.action) {
        // public (token protected)
        case 'getApproval':   return json_({ ok: true, approval: getApproval_(b.t) });
        case 'submitRequest': return json_({ ok: true, stay: submitRequest_(b.t, b.form || {}, b.signature) });
        case 'getHosting':    return json_({ ok: true, hosting: getHostingPublic_(b.h, b.k) });
        case 'saveHosting':   return json_({ ok: true, hosting: saveHostingPublic_(b.h, b.k, b.data || {}, !!b.submit) });
        // Ministry Heads
        case 'headInit':      return json_({ ok: true, ministries: ministries_(false), availability: availability_() });
        case 'createApproval':return json_({ ok: true, approval: createApproval_(b) });
        // Admin
        case 'list':          need_(isAdmin); ensureHeadCodes_(); return json_({ ok: true, stays: rows_(STAY), hosting: hostingList_(), approvals: rows_(APPR).map(x => Object.assign(x, { link: CONFIG.SITE_URL + 'request.html?t=' + x['Approval ID'] })), ministries: ministries_(true), site: CONFIG.SITE_URL, sheet: { name: SpreadsheetApp.getActive().getName(), url: SpreadsheetApp.getActive().getUrl() }, folderUrl: folder_().getUrl(), lastBackup: JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_BACKUP') || 'null') });
        case 'save':          need_(isAdmin); return json_({ ok: true, record: saveStay_(b.record || {}) });
        case 'cancel':        need_(isAdmin); return json_({ ok: true, record: setStay_(b.id, { 'Status': 'Cancelled' }) });
        case 'markAck':       need_(isAdmin); return json_({ ok: true, record: setStay_(b.id, { 'Acknowledged At': new Date() }, true) });
        case 'createHosting': need_(isAdmin); return json_({ ok: true, hosting: createHosting_(b.data || {}) });
        case 'updateHosting': need_(isAdmin); return json_({ ok: true, hosting: updateHosting_(b.id, b.data || {}) });
        case 'newHeadCode':   need_(isAdmin); return json_({ ok: true, ministries: newHeadCode_(b.head) });
        case 'backup':        need_(isAdmin); return json_({ ok: true, backup: backupNow_(false) });
        case 'makePdf':       need_(isAdmin); return json_({ ok: true, url: (b.type === 'hosting' ? hostingPdf_(b.id) : stayPdf_(b.id)).getUrl() });
        default: throw new Error('Unknown action');
      }
    } finally { lock.releaseLock(); }
  } catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
}
/** Adds any new columns after an update, once per version (no need to run setup again). */
const SCHEMA = '3';
function upgradeSheets_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('SCHEMA') === SCHEMA) return;
  [STAY, HOST, APPR, MIN].forEach(n => ensureSheet_(n));
  p.setProperty('SCHEMA', SCHEMA);
}
function need_(ok) { if (!ok) throw new Error('Wrong passcode'); }

/* =================== TABLE HELPERS =================== */

function ensureSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const want = HEADS[name];
  const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
  const missing = want.filter(h => have.indexOf(h) < 0);
  if (!have.filter(String).length) sh.getRange(1, 1, 1, want.length).setValues([want]);
  else if (missing.length) sh.getRange(1, have.length + 1, 1, missing.length).setValues([missing]);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setFontColor('#ffffff').setBackground('#6b4f33').setWrap(true);
  sh.setFrozenRows(1);
  return sh;
}
function table_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name) || ensureSheet_(name);
  const lc = sh.getLastColumn(), lr = sh.getLastRow();
  const head = sh.getRange(1, 1, 1, lc).getValues()[0].map(h => String(h).trim());
  const map = {}; head.forEach((h, i) => map[h] = i);
  const vals = lr > 1 ? sh.getRange(2, 1, lr - 1, lc).getValues() : [];
  const rows = vals.map((r, i) => { const o = { _row: i + 2 }; head.forEach((h, j) => o[h] = r[j]); return o; }).filter(o => o[head[0]] !== '');
  return { sh, head, map, rows };
}
function plain_(o) {
  const out = {};
  Object.keys(o).forEach(k => {
    if (k === '_row' || k === 'Secret') return;
    const v = o[k];
    out[k] = v instanceof Date ? Utilities.formatDate(v, CONFIG.TZ, "yyyy-MM-dd'T'HH:mm") : v;
  });
  return out;
}
function rows_(name) { return table_(name).rows.map(plain_); }
function toDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)) : String(v);
}
function writeRow_(name, rowIdx, obj) {
  const t = table_(name);
  const row = t.head.map(h => {
    if (name === STAY && h === 'Nights') return '';
    const v = obj[h];
    if (DATE_FIELDS.indexOf(h) >= 0) return toDate_(v);
    return v === undefined || v === null ? '' : v;
  });
  t.sh.getRange(rowIdx, 1, 1, t.head.length).setValues([row]);
  DATE_FIELDS.forEach(h => { if (t.map[h] !== undefined) t.sh.getRange(rowIdx, t.map[h] + 1).setNumberFormat('dd/mm/yyyy h:mm am/pm'); });
  if (name === STAY && t.map['Nights'] !== undefined) {
    const n = t.map['Nights'] + 1, a = t.map['Arrival'] + 1 - n, d = t.map['Departure'] + 1 - n;
    t.sh.getRange(rowIdx, n).setFormulaR1C1('=IF(OR(R[0]C[' + a + ']="",R[0]C[' + d + ']=""),"",MAX(0,INT(R[0]C[' + d + '])-INT(R[0]C[' + a + '])))');
  }
}
function nextId_(name, prefix) {
  let max = 0;
  table_(name).rows.forEach(r => { const m = String(r[HEADS[name][0]]).match(/(\d+)$/); if (m) max = Math.max(max, +m[1]); });
  return prefix + String(max + 1).padStart(3, '0');
}
function find_(name, id) { return table_(name).rows.find(r => String(r[HEADS[name][0]]) === String(id)); }
function token_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 18); }

/* =================== MINISTRIES & AVAILABILITY =================== */

function ministries_(admin) {
  return table_(MIN).rows.map(r => {
    const o = { ministry: r['Ministry'], head: r['Ministry Head'], notes: r['Notes'] };
    if (admin) { o.whatsapp = String(r['WhatsApp'] || ''); o.passcode = String(r['Passcode'] || ''); }
    return o;
  });
}
/** Gives every Ministry Head their own 6-digit passcode (same code across all their ministries). */
function ensureHeadCodes_() {
  const t = table_(MIN); if (t.map['Passcode'] === undefined) { ensureSheet_(MIN); return ensureHeadCodes_(); }
  const byHead = {}, used = {};
  t.rows.forEach(r => { const c = String(r['Passcode'] || '').trim(); if (c) { byHead[r['Ministry Head']] = byHead[r['Ministry Head']] || c; used[c] = 1; } });
  t.rows.forEach(r => {
    if (String(r['Passcode'] || '').trim() || r['Ministry'] === 'Accounts') return;
    let code = byHead[r['Ministry Head']];
    if (!code) { do { code = String(Math.floor(100000 + Math.random() * 900000)); } while (used[code]); used[code] = 1; byHead[r['Ministry Head']] = code; }
    t.sh.getRange(r._row, t.map['Passcode'] + 1).setNumberFormat('@').setValue(code);
  });
}
function newHeadCode_(head) {
  const t = table_(MIN); const used = {}; t.rows.forEach(r => used[String(r['Passcode'])] = 1);
  let code; do { code = String(Math.floor(100000 + Math.random() * 900000)); } while (used[code]);
  t.rows.filter(r => r['Ministry Head'] === head).forEach(r => t.sh.getRange(r._row, t.map['Passcode'] + 1).setNumberFormat('@').setValue(code));
  return ministries_(true);
}
function availability_() {
  const from = new Date(); from.setHours(0, 0, 0, 0);
  const to = new Date(from.getTime() + 150 * 864e5);
  return table_(STAY).rows
    .filter(r => r['Status'] !== 'Cancelled' && r['Arrival'] instanceof Date && r['Departure'] instanceof Date && r['Departure'] >= from && r['Arrival'] <= to)
    .map(r => ({ unit: String(r['Unit No.'] || ''), from: plain_({ a: r['Arrival'] }).a, to: plain_({ a: r['Departure'] }).a, status: r['Status'] }));
}

/* =================== MINISTRY HEAD APPROVAL LINK =================== */

function createApproval_(b) {
  ensureHeadCodes_();
  const t0 = table_(MIN);
  const min = t0.rows.find(r => r['Ministry'] === b.ministry);
  if (!min || min['Ministry'] === 'Accounts') throw new Error('Please choose your ministry.');
  const code = String((t0.rows.find(r => r['Ministry Head'] === min['Ministry Head'] && String(r['Passcode'] || '').trim()) || {})['Passcode'] || '').trim();
  if (!code || String(b.code || '').trim() !== code) throw new Error('That passcode doesn\'t match ' + min['Ministry Head'] + '. Please check with the Church Admin.');
  if (!(b.category in RATES) || RATES[b.category] === 0) throw new Error('Please choose a stay category.');
  const cin = toDate_(b.cin), cout = toDate_(b.cout);
  if (!(cin instanceof Date) || !(cout instanceof Date) || cout <= cin) throw new Error('Please check the dates.');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (cin < today) throw new Error('The check-in date has already passed.');
  const earliest = new Date(today); earliest.setDate(earliest.getDate() + 7);
  const shortNotice = cin < earliest;
  const id = token_();
  const t = table_(APPR);
  writeRow_(APPR, t.sh.getLastRow() + 1, { 'Approval ID': id, 'Ministry': min['Ministry'], 'Ministry Head': min['Ministry Head'], 'Member Name': b.member || '', 'Category': b.category, 'Check-in': cin, 'Check-out': cout, 'Created At': new Date(), 'Notes': shortNotice ? 'Short notice (special consideration)' : '' });
  return { id, link: CONFIG.SITE_URL + 'request.html?t=' + id, head: min['Ministry Head'], ministry: min['Ministry'], shortNotice };
}

function getApproval_(t) {
  const a = find_(APPR, t);
  if (!a) throw new Error('This link is not valid. Please ask your Ministry Head for a new link.');
  if (a['Used By Stay']) throw new Error('This link has already been used (request ' + a['Used By Stay'] + ').');
  const o = plain_(a); o.rate = RATES[a['Category']] || 0;
  return o;
}

/* =================== MEMBER STAY REQUEST =================== */

function submitRequest_(t, f, signature) {
  const a = getApproval_(t);
  if (!f.name || !f.agree) throw new Error('Please fill in your name and agree to the Terms & Conditions.');
  if (!signature) throw new Error('Please sign the form.');
  const id = nextId_(STAY, 'APT-');
  const suggested = (a.rate || 0) * nights_({ 'Arrival': a['Check-in'], 'Departure': a['Check-out'] });
  const gift = +f.gift || 0;
  const sigFile = saveSignature_(signature, id + '-guest-signature.png');
  const rec = {
    'Stay ID': id, 'Status': 'Requested', 'Category': a['Category'], 'Name of User': f.name, 'Main Guest': f.mainGuest || f.name,
    'Purpose of Use': f.purpose || '', 'Ministry': a['Ministry'], 'Ministry Head': a['Ministry Head'], 'Contact No.': f.phone || '',
    'No. of Guests': f.pax || '', 'Arrival': a['Check-in'], 'Departure': a['Check-out'], 'Vehicle Details': f.car || '',
    'Love Gift (RM)': gift, 'Love Gift Ref': id, 'Special Arrangement': [a['Notes'], gift < suggested ? 'Below suggested minimum (RM' + suggested + ')' : ''].filter(String).join(' · '),
    'Additional Notes': f.remarks || '', 'Source': 'Online request', 'Approval ID': t, 'Approved By': a['Ministry Head'],
    'Approved At': a['Created At'], 'Submitted At': new Date(), 'Signature': sigFile.getUrl()
  };
  const tb = table_(STAY);
  writeRow_(STAY, tb.sh.getLastRow() + 1, rec);
  const ap = find_(APPR, t), apt = table_(APPR); apt.sh.getRange(ap._row, apt.map['Used By Stay'] + 1).setValue(id);
  const pdf = stayPdf_(id, sigFile);
  notify_('🏠 New stay request · ' + id, [
    '<b>' + esc_(f.name) + '</b> · ' + esc_(a['Category']),
    '📅 ' + fmt_(a['Check-in']) + ' → ' + fmt_(a['Check-out']),
    '👥 ' + (f.pax || '-') + ' pax · ⛪ ' + esc_(a['Ministry']),
    '✍️ Approved by ' + esc_(a['Ministry Head']),
    '💝 Love gift: RM' + gift + ' (suggested RM' + suggested + ')',
    '',
    'Assign a unit: ' + CONFIG.SITE_URL + 'index.html#' + id
  ], pdf);
  return { id };
}

/* =================== ADMIN: STAYS =================== */

function saveStay_(rec) {
  const tb = table_(STAY);
  let row = rec['Stay ID'] ? tb.rows.find(r => r['Stay ID'] === rec['Stay ID']) : null;
  const merged = row ? Object.assign(plain_(row), rec) : Object.assign({}, rec);
  if (!row) { merged['Stay ID'] = nextId_(STAY, 'APT-'); merged['Status'] = merged['Status'] || 'Booked'; merged['Source'] = merged['Source'] || 'Admin'; }
  merged['Last Updated'] = new Date();
  if (CONFIG.SYNC_CALENDAR) merged['Calendar Event ID'] = syncEvent_(merged);
  writeRow_(STAY, row ? row._row : tb.sh.getLastRow() + 1, merged);
  SpreadsheetApp.flush();
  return plain_(find_(STAY, merged['Stay ID']));
}
function setStay_(id, fields, quiet) {
  const r = find_(STAY, id); if (!r) throw new Error('Stay not found: ' + id);
  if (quiet) { const t = table_(STAY); Object.keys(fields).forEach(k => t.sh.getRange(r._row, t.map[k] + 1).setValue(fields[k])); return plain_(find_(STAY, id)); }
  return saveStay_(Object.assign(plain_(r), fields));
}

/* =================== HOSTING (FOC guest speakers) =================== */

const HOST_SCHEDULE = ["Guest Name","Church / Ministry From","Event Date","Event Date(s)","Needs Apartment","Arrival","Departure","Attending Personnel","PIC Contact","PIC Email","Hosting Pastor"];
function createHosting_(d) {
  if (!d['Guest Name']) throw new Error('Please enter the guest speaker name.');
  if (!d['Event Date']) throw new Error('Please enter the preaching / event date.');
  const id = nextId_(HOST, 'HS-'), secret = token_();
  const rec = { 'Hosting ID': id, 'Status': 'Scheduled', 'Created At': new Date(), 'Secret': secret };
  HOST_SCHEDULE.forEach(f => { if (d[f] !== undefined) rec[f] = d[f]; });
  if (!rec['Event Date(s)']) rec['Event Date(s)'] = fmtDay_(rec['Event Date']);
  rec['Calendar Event ID'] = hostEvent_(rec);
  writeRow_(HOST, table_(HOST).sh.getLastRow() + 1, rec);
  return hostingList_().find(x => x['Hosting ID'] === id);
}
function hostAuth_(h, k) { const r = find_(HOST, h); if (!r || String(r['Secret']) !== String(k)) throw new Error('This hosting link is not valid. Please ask the Church Admin for a new link.'); return r; }
function getHostingPublic_(h, k) { return plain_(hostAuth_(h, k)); }
const HOST_EDITABLE = ["Guest Name","Church / Ministry From","Event Date(s)","Attending Personnel","PIC Contact","Hosting Pastor","Stay Option","Arrival","Departure","Mode of Travel","Travel Details","Transport Needed","Vehicle Details","Accompanying Persons","Apartment Guests","Dietary & Meal Schedule","Health Notes","Account Holder","Bank Name","Account Number","Mobile Number","Checklist"];
function saveHostingPublic_(h, k, data, submit) {
  const r = hostAuth_(h, k);
  const rec = Object.assign(plain_(r), { 'Secret': r['Secret'] });
  HOST_EDITABLE.forEach(f => { if (f in data) rec[f] = data[f]; });
  if (submit) { rec['Status'] = 'Submitted'; rec['Submitted At'] = new Date(); }
  else if (['Scheduled', 'Link sent'].indexOf(rec['Status']) >= 0) rec['Status'] = 'In progress';
  // linked apartment stay
  if (submit && rec['Stay Option'] === 'Staying in our apartment' && rec['Arrival'] && rec['Departure']) {
    const stay = { 'Category': 'Pastors, Guest Speakers & Mission Pastors', 'Name of User': rec['Guest Name'], 'Main Guest': rec['Guest Name'],
      'Purpose of Use': 'Guest speaker: ' + (rec['Event Date(s)'] || ''), 'Attending Personnel': rec['Attending Personnel'], 'Contact No.': rec['PIC Contact'],
      'No. of Guests': rec['Apartment Guests'] || '', 'Arrival': rec['Arrival'], 'Departure': rec['Departure'], 'Vehicle Details': rec['Vehicle Details'],
      'Church Transport': /church|pickup|transfer/i.test(rec['Transport Needed'] || '') ? 'Yes' : 'No', 'Transport PIC': rec['Attending Personnel'],
      'Guest Preferences / Special Notes': [rec['Dietary & Meal Schedule'], rec['Health Notes']].filter(String).join('\n'), 'Love Gift (RM)': 0,
      'Source': 'Hosting checklist', 'Hosting ID': h };
    if (rec['Linked Stay ID']) { const s = find_(STAY, rec['Linked Stay ID']); if (s) saveStay_(Object.assign(plain_(s), stay, { 'Status': s['Status'] })); }
    else { stay['Status'] = 'Requested'; stay['Submitted At'] = new Date(); rec['Linked Stay ID'] = saveStay_(stay)['Stay ID']; }
  }
  rec['Calendar Event ID'] = hostEvent_(rec);
  writeRow_(HOST, r._row, rec);
  if (submit) {
    const pdf = hostingPdf_(h);
    notify_('🎤 Hosting checklist submitted · ' + h, [
      '<b>' + esc_(rec['Guest Name']) + '</b>' + (rec['Church / Ministry From'] ? ' (' + esc_(rec['Church / Ministry From']) + ')' : ''),
      '🗓️ ' + esc_(rec['Event Date(s)'] || '-') + ' · 🙋 PIC: ' + esc_(rec['Attending Personnel'] || '-'),
      '🏠 ' + esc_(rec['Stay Option'] || '-') + (rec['Linked Stay ID'] ? ' → assign unit for ' + rec['Linked Stay ID'] : ''),
      '', 'Open: ' + CONFIG.SITE_URL + 'index.html#' + h
    ], pdf);
  }
  return plain_(find_(HOST, h));
}
function hostingList_() {
  return table_(HOST).rows.map(r => { const o = plain_(r); o.link = CONFIG.SITE_URL + 'hosting.html?h=' + r['Hosting ID'] + '&k=' + r['Secret']; return o; });
}
function updateHosting_(id, d) {
  const r = find_(HOST, id); if (!r) throw new Error('Not found: ' + id);
  const before = plain_(r);
  const rec = Object.assign(before, { 'Secret': r['Secret'] }, d);
  // dates moved → reminders start again
  if ((d['Event Date'] && d['Event Date'] !== plain_(r)['Event Date']) || (d['Arrival'] && d['Arrival'] !== plain_(r)['Arrival'])) { rec['Reminder 14d Sent'] = ''; rec['Reminder 3d Sent'] = ''; }
  if (rec['Status'] === 'Closed' && rec['Calendar Event ID']) { /* keep the event; it's history */ }
  rec['Calendar Event ID'] = hostEvent_(rec);
  writeRow_(HOST, r._row, rec);
  return hostingList_().find(x => x['Hosting ID'] === id);
}

/* =================== DRIVE: SIGNATURES & PDFs =================== */

function folder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CONFIG.FOLDER_NAME);
  p.setProperty('FOLDER_ID', f.getId());
  return f;
}
function saveSignature_(dataUrl, name) {
  const b64 = String(dataUrl).split(',')[1];
  return folder_().createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', name));
}
function sigData_(urlOrFile) {
  try {
    const file = typeof urlOrFile === 'object' ? urlOrFile : DriveApp.getFileById(String(urlOrFile).match(/[-\w]{25,}/)[0]);
    return 'data:image/png;base64,' + Utilities.base64Encode(file.getBlob().getBytes());
  } catch (e) { return ''; }
}
function savePdf_(html, name, sheet, idField, id) {
  const blob = Utilities.newBlob(html, 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf');
  const file = folder_().createFile(blob);
  const t = table_(sheet), r = t.rows.find(x => String(x[idField]) === String(id));
  if (r) t.sh.getRange(r._row, t.map['PDF'] + 1).setValue(file.getUrl());
  return file;
}
const LOGO = 'data:image/png;base64,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';
const PDF_CSS = 'body{font-family:Arial,sans-serif;color:#2f251b;font-size:10.5pt;margin:28px 34px}h1{font-family:Georgia,serif;color:#6b4f33;font-size:20pt;margin:2px 0 4px}h2{font-family:Georgia,serif;color:#6b4f33;font-size:12.5pt;margin:16px 0 6px;border-bottom:1px solid #d9cdb8;padding-bottom:3px}.e{letter-spacing:3px;font-size:8pt;color:#9a7d3f}table{width:100%;border-collapse:collapse}td{padding:4px 6px;vertical-align:top;border-bottom:1px solid #eee5d6}td.k{width:32%;color:#8a7a68;font-size:9pt}.stamp{border:1.5px solid #6b4f33;color:#6b4f33;padding:6px 10px;display:inline-block;font-size:9.5pt}.small{font-size:8.5pt;color:#8a7a68}ul{margin:4px 0;padding-left:18px}li{margin-bottom:3px}';
function kv_(pairs) { return '<table>' + pairs.filter(p => p).map(p => '<tr><td class="k">' + esc_(p[0]) + '</td><td>' + esc_(p[1] === undefined || p[1] === null ? '' : p[1]).replace(/\n/g, '<br>') + '</td></tr>').join('') + '</table>'; }
function stayPdf_(id, sigFile) {
  const r = find_(STAY, id); if (!r) throw new Error('Stay not found');
  const sig = sigFile ? sigData_(sigFile) : (r['Signature'] ? sigData_(r['Signature']) : '');
  const rate = RATES[r['Category']] || 0;
  const html = '<html><head><style>' + PDF_CSS + '</style></head><body>' +
    '<img src="' + LOGO + '" style="width:42px;height:42px;float:right">' + '<div class="e">TABERNACLE OF WORSHIP</div><h1>Apartment Stay · Terms &amp; Approval</h1><div class="small">' + esc_(id) + ' · ' + esc_(r['Status']) + ' · Generated ' + fmt_(new Date()) + '</div>' +
    '<h2>Booking details</h2>' + kv_([['Name', r['Name of User']], ['Main guest', r['Main Guest']], ['Contact no.', r['Contact No.']], ['Guest category', r['Category']], ['Ministry / Department', r['Ministry']], ['Purpose of stay', r['Purpose of Use']], ['Unit (assigned)', r['Unit No.'] || 'To be assigned'], ['No. of guests', r['No. of Guests']], ['Check-in', fmt_(r['Arrival']) + ' · ' + keyMode_(r['Arrival'])], ['Check-out', fmt_(r['Departure']) + ' · ' + keyMode_(r['Departure'])], ['Vehicle no.', r['Vehicle Details']], ['Remarks', [r['Special Arrangement'], r['Additional Notes']].filter(String).join(' · ')]]) +
    '<h2>Terms &amp; Conditions</h2><ul><li>Bookings should be made at least 1 week in advance (shorter notice by special consideration only).</li><li>Check-in &amp; check-out: Tue–Fri 9am–6pm and Sun 8am–11am, accompanied by the Ministry Head / Church Admin. At other times it is a self check-in/out, with keys left at the guard house in a sealed envelope.</li><li>Please take time to read the House Rules posted in the unit.</li><li>Keep the unit tidy. Do not move furniture or remove any items.</li><li>At check-out: throw all rubbish into the large bin near the security booth, switch off all appliances, lock up and return the keys as arranged.</li><li>Please report any damage at once. We kindly ask guests to cover any repair or replacement needed.</li><li>Love gifts go towards cleaning and upkeep. They are offerings, not rental fees.</li></ul><div class="small">☑ Agreed to the Terms &amp; Conditions' + (r['Submitted At'] ? ' on ' + fmt_(r['Submitted At']) : '') + '.</div>' +
    (rate ? '<h2>Love gift</h2>' + kv_([['Suggested minimum', 'RM' + rate + ' per night × ' + nights_(r) + ' night(s) = RM' + rate * nights_(r)], ['Love gift pledged', 'RM' + (r['Love Gift (RM)'] || 0)], ['Reference', r['Love Gift Ref'] || id], ['Church account', 'TABERNACLE OF WORSHIP · Hong Leong Bank (Malaysia) · 371 000 128 43']]) : '') +
    '<h2>Signature &amp; approval</h2><table><tr><td style="width:50%">' + (sig ? '<img src="' + sig + '" style="height:60px">' : '<div style="height:60px"></div>') + '<div class="small">Guest signature · ' + esc_(r['Name of User']) + '</div></td><td>' +
    (r['Approved By'] ? '<div class="stamp">✔ Digitally approved by <b>' + esc_(r['Approved By']) + '</b><br>' + esc_(r['Ministry']) + ' · ' + fmt_(r['Approved At']) + '</div>' : '<div style="height:60px"></div><div class="small">Ministry Head / Church Admin</div>') + '</td></tr></table>' +
    '<h2>At check-out · Ministry Head / Church Admin</h2><div class="small">☐ Keys returned in person &nbsp; ☐ Keys left at guard house &nbsp; ☐ Rubbish cleared &nbsp; ☐ Unit in order</div><div style="height:40px;border-bottom:1px solid #d9cdb8"></div><div class="small">Notes / remarks · Signature · Date</div>' +
    '<p class="small" style="margin-top:18px">For Tabernacle of Worship, Seremban · Internal use</p></body></html>';
  return savePdf_(html, 'TOW-Apartment-' + id + '-' + String(r['Name of User']).replace(/[^\w]+/g, '-'), STAY, 'Stay ID', id);
}
function hostingPdf_(id) {
  const r = find_(HOST, id); if (!r) throw new Error('Not found');
  const done = String(r['Checklist'] || '').split('|').filter(String);
  const html = '<html><head><style>' + PDF_CSS + '</style></head><body>' +
    '<img src="' + LOGO + '" style="width:42px;height:42px;float:right">' + '<div class="e">TABERNACLE OF WORSHIP</div><h1>Guest Speaker Hosting · In-House Checklist</h1><div class="small">' + esc_(id) + ' · ' + esc_(r['Status']) + ' · Generated ' + fmt_(new Date()) + '</div>' +
    '<h2>Guest &amp; host</h2>' + kv_([['Guest speaker / pastor', r['Guest Name']], ['Church / ministry (from)', r['Church / Ministry From']], ['Event date(s)', r['Event Date(s)']], ['Attending Personnel (PIC)', r['Attending Personnel'] + (r['PIC Contact'] ? ' · ' + r['PIC Contact'] : '')], ['Hosting Pastor', r['Hosting Pastor']]]) +
    '<h2>Travel &amp; stay</h2>' + kv_([['Stay', r['Stay Option'] + (r['Linked Stay ID'] ? ' (' + r['Linked Stay ID'] + ')' : '')], ['Arrival', fmt_(r['Arrival'])], ['Departure', fmt_(r['Departure'])], ['Mode of travel', [r['Mode of Travel'], r['Travel Details']].filter(String).join(' · ')], ['Transport needed', r['Transport Needed']], ['Vehicle plate / model / colour', r['Vehicle Details']], ['Accompanying persons', r['Accompanying Persons']], ['Guests in apartment', r['Apartment Guests']]]) +
    '<h2>Meals &amp; care</h2>' + kv_([['Dietary preferences & meal schedule', r['Dietary & Meal Schedule']], ['Allergies / health notes', r['Health Notes']]]) + '<div class="small">Meal claims: RM50–RM70 per pax per meal · above RM70 needs prior approval · original receipts · Accounts Submission Deadline SOP.</div>' +
    '<h2>Honorarium / love gift</h2>' + kv_([['Account holder', r['Account Holder']], ['Bank name', r['Bank Name']], ['Account number', r['Account Number']], ['Mobile number', r['Mobile Number']]]) +
    '<h2>Checklist</h2><ul style="list-style:none;padding-left:0">' + HOST_CHECKS.map(c => '<li>' + (done.indexOf(c[0]) >= 0 ? '☑' : '☐') + ' ' + esc_(c[1]) + '</li>').join('') + '</ul>' +
    '<p class="small" style="margin-top:18px">For Tabernacle of Worship, Seremban · Internal use</p></body></html>';
  return savePdf_(html, 'TOW-Hosting-' + id + '-' + String(r['Guest Name']).replace(/[^\w]+/g, '-'), HOST, 'Hosting ID', id);
}
const HOST_CHECKS = [["c1","2 wks: Apartment Preparation Checklist completed (if staying)"],["c2","2 wks: Sermon slides received (deadline 1 week before)"],["c3","2 wks: Bank details forwarded to Accounts"],["c4","2 wks: Gift prepared (if any)"],["c5","2 wks: Travel reconfirmed"],["c6","2 wks: Meal arrangements confirmed"],["c7","Day before: Transport & pickup reconfirmed"],["c8","Day before: Meal schedule reconfirmed"],["c9","Day of: Guest personally welcomed and hosted"],["c10","After: Return transport arranged"],["c11","After: Lodging checkout checked"],["c12","After: Thank-you message sent (cc Senior Pastor)"],["c13","After: Honorarium / love gift processed"]];

/* =================== ALERTS =================== */

function notify_(title, lines, pdfFile) {
  const text = '<b>' + title + '</b>\n' + lines.join('\n');
  try { telegram_(text); } catch (e) { Logger.log('Telegram failed: ' + e); }
  try {
    const to = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (to) {
      const opt = { htmlBody: text.replace(/\n/g, '<br>'), name: 'TOW Apartment' };
      if (pdfFile && pdfFile.getBlob) opt.attachments = [pdfFile.getBlob()];
      MailApp.sendEmail(to, title.replace(/<[^>]+>/g, ''), text.replace(/<[^>]+>/g, ''), opt);
    }
  } catch (e) { Logger.log('Email failed: ' + e); }
}
function telegram_(html) {
  const p = PropertiesService.getScriptProperties();
  const token = p.getProperty('TELEGRAM_TOKEN'), chat = p.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chat) return;
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', { method: 'post', muteHttpExceptions: true,
    payload: { chat_id: chat, text: html, parse_mode: 'HTML', disable_web_page_preview: 'true' } });
}

/* =================== GUEST SPEAKER REMINDERS =================== */

/** Turns on the daily jobs: backup (~2am) and guest speaker reminders (~8am). Run once. */
function setupDailyJobs() {
  ScriptApp.getProjectTriggers().filter(t => ['autoBackup', 'dailyReminders'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoBackup').timeBased().everyDays(1).atHour(2).inTimezone(CONFIG.TZ).create();
  ScriptApp.newTrigger('dailyReminders').timeBased().everyDays(1).atHour(8).inTimezone(CONFIG.TZ).create();
  const b = backupNow_(true);
  Logger.log('Daily backup (2am) and reminders (8am) are on. First backup: %s', b.url);
}
function siteFromProps_() {
  const s = PropertiesService.getScriptProperties().getProperty('SITE_URL');
  if (/YOUR-GITHUB/i.test(CONFIG.SITE_URL) && s) CONFIG.SITE_URL = s;
}
const HOST_DONE = ['Received', 'Sent to Accounts', 'Closed'];
function missingHost_(r) {
  const m = [], apt = r['Stay Option'] === 'Staying in our apartment' || (!r['Stay Option'] && r['Needs Apartment'] === 'Yes');
  if (!r['Stay Option']) m.push('Where the guest is staying');
  if (!r['Arrival']) m.push('Arrival date & time');
  if (apt && !r['Departure']) m.push('Departure date & time');
  if (apt && !r['Apartment Guests']) m.push('Number of guests in apartment');
  if (!r['Dietary & Meal Schedule']) m.push('Dietary needs & meal schedule');
  if (!r['Vehicle Details'] && !/flight|train|bus/i.test(String(r['Mode of Travel'] || ''))) m.push('Vehicle plate / model / colour');
  if (!r['Account Holder'] || !r['Bank Name'] || !r['Account Number']) m.push('Honorarium bank details');
  if (String(r['Checklist'] || '').split('|').indexOf('c2') < 0) m.push('Sermon slides');
  return m;
}
function daysUntil_(d) { const x = toDate_(d); if (!(x instanceof Date)) return null; const t = new Date(); t.setHours(0, 0, 0, 0); const y = new Date(x); y.setHours(0, 0, 0, 0); return Math.round((y - t) / 864e5); }
function picMessage_(r, link) {
  return 'Hi ' + (r['Attending Personnel'] || '') + ' 👋\n\nYou\'re the PIC / host for *' + r['Guest Name'] + '* (' + (r['Event Date(s)'] || fmtDay_(r['Event Date'])) + ').\n\nPlease check with the guest and fill in the Guest Speaker Hosting Checklist here:\n' + link + '\n\nYou can save and come back anytime, then submit to the Church Admin. 🙏';
}
function waUrl_(phone, text) { let d = String(phone || '').replace(/\D/g, ''); if (d.indexOf('0') === 0) d = '6' + d; return 'https://wa.me/' + d + '?text=' + encodeURIComponent(text); }
/** Runs every morning: 14-day and 3-day reminders for scheduled guest speakers. */
function dailyReminders() {
  upgradeSheets_(); siteFromProps_();
  const t = table_(HOST);
  t.rows.forEach(r => {
    if (r['Status'] === 'Closed') return;
    const ref = r['Arrival'] || r['Event Date'];
    const days = daysUntil_(ref);
    if (days === null || days < 0) return;
    const link = CONFIG.SITE_URL + 'hosting.html?h=' + r['Hosting ID'] + '&k=' + r['Secret'];
    const p = plain_(r);
    // 14 days before
    if (days <= 14 && !r['Reminder 14d Sent'] && HOST_DONE.indexOf(r['Status']) < 0 && r['Status'] !== 'Submitted') {
      if (r['PIC Email']) mail_(r['PIC Email'], 'Guest speaker hosting: ' + r['Guest Name'] + ' (' + (r['Event Date(s)'] || fmtDay_(r['Event Date'])) + ')',
        '<p>Hi ' + esc_(r['Attending Personnel'] || '') + ',</p><p>You are the PIC / host for <b>' + esc_(r['Guest Name']) + '</b>' + (r['Church / Ministry From'] ? ' (' + esc_(r['Church / Ministry From']) + ')' : '') + ' on <b>' + esc_(r['Event Date(s)'] || fmtDay_(r['Event Date'])) + '</b>.</p><p>Please check with the guest and fill in the Guest Speaker Hosting Checklist (stay, travel, meals, sermon slides and honorarium details):</p><p><a href="' + link + '">' + link + '</a></p><p>You can save and come back anytime, then submit to the Church Admin. Thank you! 🙏<br>— TOW Church Admin</p>');
      notify_('🎤 14 days to go · ' + r['Guest Name'], [
        '🗓️ ' + esc_(r['Event Date(s)'] || fmtDay_(r['Event Date'])) + ' · 🙋 PIC: ' + esc_(r['Attending Personnel'] || '-'),
        r['PIC Email'] ? '📧 Hosting link emailed to the PIC.' : '⚠️ No PIC email on file. Please send the link on WhatsApp.',
        '', '💬 Forward to PIC on WhatsApp:', waUrl_(r['PIC Contact'], picMessage_(r, link))
      ], null);
      t.sh.getRange(r._row, t.map['Reminder 14d Sent'] + 1).setValue(new Date());
      if (r['Status'] === 'Scheduled') t.sh.getRange(r._row, t.map['Status'] + 1).setValue('Link sent');
    }
    // 3 days before
    if (days <= 3 && !r['Reminder 3d Sent']) {
      const miss = missingHost_(p);
      const apt = p['Stay Option'] === 'Staying in our apartment' || (!p['Stay Option'] && p['Needs Apartment'] === 'Yes');
      const lines = ['🗓️ ' + esc_(r['Event Date(s)'] || fmtDay_(r['Event Date'])) + ' · 🙋 PIC: ' + esc_(r['Attending Personnel'] || '-')];
      if (apt) { const st = r['Linked Stay ID'] ? find_(STAY, r['Linked Stay ID']) : null; lines.push('🏠 <b>Prepare the apartment</b>' + (st ? ' · ' + st['Stay ID'] + (st['Unit No.'] ? ' · Unit ' + st['Unit No.'] : ' · unit not assigned yet') : '') + (r['Arrival'] ? ' · arriving ' + fmt_(r['Arrival']) : '')); }
      lines.push(miss.length ? '⚠️ Still missing: ' + miss.join(', ') : '✅ Hosting checklist details complete');
      if (miss.length) lines.push('', '💬 Follow up with PIC:', waUrl_(r['PIC Contact'], 'Hi ' + (r['Attending Personnel'] || '') + ', a gentle reminder for *' + r['Guest Name'] + '* in 3 days 🙏\nStill needed: ' + miss.join(', ') + '\nPlease update here: ' + link));
      lines.push('', 'Open: ' + CONFIG.SITE_URL + 'index.html#' + r['Hosting ID']);
      notify_((apt ? '🏠 3 days · prepare apartment · ' : '🎤 3 days to go · ') + r['Guest Name'], lines, null);
      if (miss.length && r['PIC Email']) mail_(r['PIC Email'], 'Reminder: details still needed for ' + r['Guest Name'], '<p>Hi ' + esc_(r['Attending Personnel'] || '') + ',</p><p>A gentle reminder that <b>' + esc_(r['Guest Name']) + '</b> is coming in 3 days. We still need:</p><ul>' + miss.map(x => '<li>' + esc_(x) + '</li>').join('') + '</ul><p>Please update the checklist here: <a href="' + link + '">' + link + '</a></p><p>Thank you! 🙏<br>— TOW Church Admin</p>');
      t.sh.getRange(r._row, t.map['Reminder 3d Sent'] + 1).setValue(new Date());
    }
  });
}
function mail_(to, subject, html) { try { MailApp.sendEmail(to, subject, html.replace(/<[^>]+>/g, ''), { htmlBody: html, name: 'TOW Church Admin' }); } catch (e) { Logger.log('Mail failed: ' + e); } }
function hostEvent_(rec) {
  let cal; try { cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID); } catch (e) { cal = null; }
  if (!cal || !CONFIG.SYNC_CALENDAR) return rec['Calendar Event ID'] || '';
  let ev = null; if (rec['Calendar Event ID']) { try { ev = cal.getEventById(rec['Calendar Event ID']); } catch (e) { ev = null; } }
  const a = toDate_(rec['Arrival']), d = toDate_(rec['Departure']), day = toDate_(rec['Event Date']);
  if (!(day instanceof Date) && !(a instanceof Date)) return rec['Calendar Event ID'] || '';
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const endRef = d instanceof Date ? d : (a instanceof Date ? a : day);
  if (!ev && endRef < today) return '';
  const title = '🎤 Guest speaker · ' + rec['Guest Name'] + (rec['Attending Personnel'] ? ' (PIC: ' + rec['Attending Personnel'] + ')' : '');
  const desc = ['Hosting ID: ' + rec['Hosting ID'], 'Event: ' + (rec['Event Date(s)'] || ''), 'From: ' + (rec['Church / Ministry From'] || '-'), 'PIC: ' + (rec['Attending Personnel'] || '-') + (rec['PIC Contact'] ? ' · ' + rec['PIC Contact'] : ''), 'Hosting Pastor: ' + (rec['Hosting Pastor'] || '-'), 'Apartment: ' + (rec['Stay Option'] || rec['Needs Apartment'] || '-'), '', '(TOW Apartment Stay System)'].join('\n');
  const timed = a instanceof Date && d instanceof Date && d > a;
  if (ev) {
    ev.setTitle(title); ev.setDescription(desc);
    try { if (timed) ev.setTime(a, d); else ev.setAllDayDate(day instanceof Date ? day : a); } catch (e) {}
    return ev.getId();
  }
  ev = timed ? cal.createEvent(title, a, d, { description: desc }) : cal.createAllDayEvent(title, day instanceof Date ? day : a, { description: desc });
  return ev.getId();
}
function fmtDay_(v) { const d = toDate_(v); return d instanceof Date ? Utilities.formatDate(d, CONFIG.TZ, 'EEE d MMM yyyy') : (v || ''); }

/* =================== CALENDAR =================== */

function syncEvent_(rec) {
  let cal; try { cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID); } catch (e) { cal = null; }
  if (!cal) return rec['Calendar Event ID'] || '';
  let ev = null;
  if (rec['Calendar Event ID']) { try { ev = cal.getEventById(rec['Calendar Event ID']); } catch (e) { ev = null; } }
  const start = toDate_(rec['Arrival']), end = toDate_(rec['Departure']);
  if (NO_CALENDAR.indexOf(rec['Status']) >= 0 || !(start instanceof Date) || !(end instanceof Date) || end <= start) { if (ev) ev.deleteEvent(); return ''; }
  const title = '🏠 Apartment ' + (rec['Unit No.'] || 'TBC') + ' · ' + (rec['Name of User'] || 'Guest');
  const desc = ['Stay ID: ' + rec['Stay ID'], 'Status: ' + rec['Status'], 'Category: ' + (rec['Category'] || '-'), 'Ministry: ' + (rec['Ministry'] || '-'),
    'Attending Personnel: ' + (rec['Attending Personnel'] || '-'), 'Guests: ' + (rec['No. of Guests'] || '-'), '', '(TOW Apartment Stay Log)'].join('\n');
  // Only NEW stays from today onwards get a calendar event. Past stays and stays imported from the old log are
  // never added, so nothing already in the calendar is duplicated. The script only ever edits events it created itself.
  if (!ev) {
    const today = new Date(); today.setHours(0, 0, 0, 0);
    if (end < today || /imported/i.test(String(rec['Source'] || ''))) return '';
  }
  if (ev) { ev.setTitle(title); ev.setTime(start, end); ev.setDescription(desc); ev.setLocation('TOW Apartment ' + (rec['Unit No.'] || '')); return ev.getId(); }
  return cal.createEvent(title, start, end, { description: desc, location: 'TOW Apartment ' + (rec['Unit No.'] || '') }).getId();
}

/* =================== UTIL =================== */

function nights_(r) { const a = toDate_(r['Arrival']), d = toDate_(r['Departure']); if (!(a instanceof Date) || !(d instanceof Date)) return 0; const A = new Date(a); A.setHours(0,0,0,0); const D = new Date(d); D.setHours(0,0,0,0); return Math.max(0, Math.round((D - A) / 864e5)); }
function keyMode_(v) { const d = toDate_(v); if (!(d instanceof Date)) return ''; const day = d.getDay(), h = d.getHours() + d.getMinutes() / 60; return ((day >= 2 && day <= 5 && h >= 9 && h < 18) || (day === 0 && h >= 8 && h < 11)) ? 'with Ministry Head / Church Admin' : 'self (keys at guard house)'; }
function fmt_(v) { const d = toDate_(v); return d instanceof Date ? Utilities.formatDate(d, CONFIG.TZ, 'EEE d MMM yyyy, h:mm a') : (v || ''); }
function esc_(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
