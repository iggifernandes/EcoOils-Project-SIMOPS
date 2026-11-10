// VERSION: 5.7.0 - faster saves (no register write for plain edits, fresh meeting in the same response); AI check removed
// Setup: Google Sheet > Extensions > Apps Script > paste this file > Project Settings > Script properties > add HOST_EMAILS = the HSSE host email(s) >
// Deploy > Web app (Execute as: Me, Access: Anyone). First run asks for Drive, Sheets and Mail permissions.
// Cost: free (about 100 emails/day on a Gmail account; each signature or host login uses one email).

const HOST_EMAILS_DEFAULT = [];   // leave empty: set the host email(s) once in Project Settings > Script properties > HOST_EMAILS (comma separated). They then survive every re-paste of this file.
function hostEmails_() {
  const p = String(PropertiesService.getScriptProperties().getProperty('HOST_EMAILS') || '');
  return emails_(p).concat(HOST_EMAILS_DEFAULT);
}
const SHEET_NAME = 'SIMOPS';
const FOLDER_NAME = 'SIMOPS Records';
const CODES = ['KKS', 'WEN', 'GSB', 'PBS', 'ECO', 'EOO', 'OT1', 'OT2'];
const ROLES = ['KKS', 'WEN', 'GSB', 'PBS', 'ECO', 'EOO', 'OT1', 'OT2', 'ack', 'noted'];
const NAMES = { KKS: 'KKS', WEN: 'WEN', GSB: 'GSB', PBS: 'PBS', ECO: 'EcoOils Project', EOO: 'EcoOils Operations', OT1: 'Other contractor 1', OT2: 'Other contractor 2', ack: 'Acknowledged (EcoOils Construction Representative)', noted: 'Noted (EcoOils HSSE)' };

// Sheet cells come back as Date objects (date shifted by the time zone, time as 1899): normalise to plain text in the sheet's own time zone
function tz_() { try { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Jakarta'; } catch (e) { return 'Asia/Jakarta'; } }
function fd_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v || '').slice(0, 10); }
function ft_(v) { return v instanceof Date ? Utilities.formatDate(v, tz_(), 'HH:mm') : String(v || '').slice(0, 5); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function folder_() { const it = DriveApp.getFoldersByName(FOLDER_NAME); return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME); }
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['ID', 'Date', 'Time', 'Host', 'Status', 'Parties', 'Clashes', 'Updated', 'PDF', 'Owner', 'HostToken']); }
  return sh;
}
function file_(id) {
  const c = CacheService.getScriptCache(), k = 'fid:' + id, fid = c.get(k);   // remembering the Drive file id avoids a slow folder search on every request
  if (fid) { try { return DriveApp.getFileById(fid); } catch (e) { c.remove(k); } }
  const it = folder_().getFilesByName('M_' + id + '.json');
  if (!it.hasNext()) return null;
  const f = it.next(); c.put(k, f.getId(), 21600); return f;
}
function load_(id) { const f = file_(id); return f ? JSON.parse(f.getBlob().getDataAsString()) : null; }
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function rnd_(n) { return Utilities.getUuid().replace(/-/g, '').slice(0, n); }
function mask_(e) { if (!e) return ''; const p = e.split('@'); return p[0].slice(0, 1) + '***@' + (p[1] || ''); }
const CTL = ['ack', 'noted'];
function emails_(v) { return String(v || '').split(/[,;\s]+/).filter(function (x) { return x; }); }
function okEmails_(v) { const l = emails_(v); return l.length > 0 && l.every(okEmail_); }
function maskList_(v) { return emails_(v).map(mask_).join(', '); }
function okEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e || ''); }

// ---------- clash detection (same area name chosen by two different parties) ----------
function clashList_(m) {
  const rs = [], out = [];
  CODES.forEach(function (c) { ((m.con[c] && m.con[c].rows) || []).forEach(function (r) { rs.push({ c: c, r: r }); }); });
  for (let x = 0; x < rs.length; x++) for (let y = x + 1; y < rs.length; y++) {
    const A = rs[x], B = rs[y]; if (A.c === B.c) continue;
    let why = '';
    const ka = (A.r.area || '').trim().toLowerCase(), kb = (B.r.area || '').trim().toLowerCase();
    if (ka && ka === kb) why = 'same area';
    if (why) out.push(NAMES[A.c] + ' (' + A.r.activity + ') <-> ' + NAMES[B.c] + ' (' + B.r.activity + '): ' + why);
  }
  return out;
}

// ---------- tamper-evident seal ----------
// ---------- meetings ----------
function newMeeting_(f) {
  const m = {
    id: rnd_(8), ht: rnd_(16), tok: {}, owner: f.owner, base: f.base,
    host: f.host, date: f.date, time: f.time, hostEmails: f.hostEmails, ackEmail: f.ackEmail || '', notedEmail: f.notedEmail || '', areas: f.areas || [],
    v: 2, status: 'open', con: {}, att: {}, log: [], parts: [], sign: {}, photos: [], hostc: '', created: new Date().toISOString()
  };
  ROLES.forEach(function (r) { m.tok[r] = rnd_(16); });
  save_(m);
  return m;
}
// as soon as a SIMOPS is fully signed, the next day's meeting opens by itself so contractors can start entering their plans
function openNext_(m) {
  if (m.next) return load_(m.next);
  let d = new Date(String(m.date).slice(0, 10) + 'T00:00:00+07:00');
  if (isNaN(d.getTime())) d = new Date();
  const nd = Utilities.formatDate(new Date(d.getTime() + 86400000), 'Asia/Jakarta', 'yyyy-MM-dd');
  const n = newMeeting_({ owner: m.owner, base: m.base, host: m.host, date: nd, time: m.time, hostEmails: m.hostEmails, ackEmail: m.ackEmail, notedEmail: m.notedEmail, areas: m.areas });
  const att = {}; CODES.forEach(function (c) { const a = (m.att && m.att[c]) || []; if (a.length) att[c] = a.map(function (x) { return { n: x.n, p: x.p, x: false }; }); });
  n.att = att; save_(n);
  m.next = n.id;
  return n;
}

function content_(m) {
  const con = {};
  CODES.forEach(function (c) {
    const e = m.con[c];
    if (e && m.parts.indexOf(c) >= 0) con[c] = { name: e.name, rows: e.rows.map(function (r) { return { area: r.area, activity: r.activity, details: r.details || '', constraint: r.constraint }; }) };
  });
  const out = { id: m.id, host: m.host, date: m.date, time: m.time, parts: m.parts, con: con, hostc: m.hostc || '', photos: (m.photos || []).map(sha_) };
  const att = {};
  CODES.forEach(function (c) { const a = (m.att && m.att[c]) || []; if (a.length) att[c] = a.map(function (x) { return { n: x.n, p: x.p, x: !!x.x }; }); });
  if (Object.keys(att).length) out.att = att;   // only present when an attendance list exists, so older sealed records still verify
  return out;
}
function contentHash_(m) { return sha_(JSON.stringify(content_(m))); }
function sigHash_(who, s, H) { const parts = [who, s.name, s.emailFull, s.at, H, sha_(s.sign)]; if (s.pos) parts.push(s.pos); return sha_(parts.join('|')); }
function required_(m) {
  if (m.v !== 2) return m.parts.concat(['ack', 'noted']);
  return ['ack', 'noted'].filter(function (r) { return !(m.absent && m.absent[r]); });   // a role marked "not present" by the other does not sign
}
function sealOf_(m) {
  const base = m.H + '|' + required_(m).map(function (k) { return m.sign[k] && m.sign[k].sh; }).join(',');
  return sha_(m.absent && Object.keys(m.absent).length ? base + '|' + JSON.stringify(m.absent) : base);
}
function finish_(m) {   // complete the record as soon as everybody who is present has signed
  if (m.status !== 'signing' || !required_(m).length || !required_(m).every(function (r) { return m.sign[r]; })) return false;
  m.seal = sealOf_(m);
  m.sealedAt = new Date().toISOString();
  m.status = 'done';
  const pf = makePdf_(m); m.pdf = pf.getUrl(); m.pdfId = pf.getId();
  openNext_(m);
  save_(m); sendPdf_(m);
  return true;
}
function emailOf_(m, who) {
  const reg = roleEmails_()[who]; if (reg) return reg;
  if (who === 'ack') return m.ackEmail; if (who === 'noted') return m.notedEmail;
  return m.con[who] && m.con[who].email;
}

// ---------- OTP ----------
function otpSend_(key, email, purpose) {
  const code = ('000000' + (parseInt(rnd_(8), 16) % 1000000)).slice(-6);
  const cache = CacheService.getScriptCache();
  cache.put('otp:' + key, code, 600); cache.put('tries:' + key, '0', 600);
  MailApp.sendEmail({ to: emails_(email).join(','), subject: 'SIMOPS verification code', body: 'Your verification code for ' + purpose + ' is ' + code + '. It is valid for 10 minutes. If you did not request it, ignore this email.' });
}
function otpCheck_(key, code) {
  const cache = CacheService.getScriptCache();
  const tries = parseInt(cache.get('tries:' + key) || '0', 10);
  if (tries >= 5) return false;
  cache.put('tries:' + key, String(tries + 1), 600);
  const c = cache.get('otp:' + key);
  if (c && String(code).trim() === c) { cache.remove('otp:' + key); return true; }
  return false;
}
// ---------- long-lived sign-ins (30 days, kept in Script Properties so they survive the 6-hour cache limit) ----------
const SESSION_DAYS = 30;
function lsPut_(kind, id, val) {
  const P = props_(), now = Date.now(), all = P.getProperties();
  Object.keys(all).forEach(function (k) {   // prune expired sign-ins
    if (k.indexOf('LS:') === 0) { try { if (JSON.parse(all[k]).x < now) P.deleteProperty(k); } catch (e) { P.deleteProperty(k); } }
  });
  P.setProperty('LS:' + kind + ':' + id, JSON.stringify({ v: val, x: now + SESSION_DAYS * 86400000 }));
}
function lsGet_(kind, id) {
  if (!id) return null;
  const P = props_(), k = 'LS:' + kind + ':' + id, raw = P.getProperty(k);
  if (!raw) return null;
  let o; try { o = JSON.parse(raw); } catch (e) { return null; }
  const now = Date.now();
  if (o.x < now) { P.deleteProperty(k); return null; }
  if (o.x - now < (SESSION_DAYS - 1) * 86400000) { o.x = now + SESSION_DAYS * 86400000; P.setProperty(k, JSON.stringify(o)); } // stay signed in while in use
  return o.v;
}
function sess_(s) { return lsGet_('h', s); }

// ---------- party sign-in (PIN set only through an emailed code; PINs are stored as salted hashes, never in the sheet) ----------
function props_() { return PropertiesService.getScriptProperties(); }
function roleEmails_() { try { return JSON.parse(props_().getProperty('roleEmails') || '{}'); } catch (e) { return {}; } }
function pinHash_(role, pin) {
  let salt = props_().getProperty('salt:' + role);
  if (!salt) { salt = rnd_(16); props_().setProperty('salt:' + role, salt); }
  return sha_(salt + ':' + role + ':' + pin);
}
function csess_(s) { return lsGet_('c', s); }
function roleAuth_(m, b) { return !!(b.c && ((m.tok[b.c] && m.tok[b.c] === b.t) || (b.s && csess_(b.s) === b.c))); }

// ---------- storage ----------
function save_(m, fast) {   // fast: plain content edits (entries, attendance, photos) skip the register row, which saves a second or two
  m.updated = new Date().toISOString();
  const f = file_(m.id), body = JSON.stringify(m);
  if (f) f.setContent(body); else folder_().createFile('M_' + m.id + '.json', body, MimeType.PLAIN_TEXT);
  if (fast && f) return;
  const sh = sheet_();
  const row = [m.id, m.date, m.time, m.host, m.status, (m.parts || []).map(function (c) { return NAMES[c]; }).join(', '), clashList_(m).join('; '), m.updated, m.pdf || '', m.owner || '', m.ht];
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  const i = ids.indexOf(m.id);
  if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function view_(m, p) {
  const host = !!(p.ht && p.ht === m.ht);
  const who = (p.c && ((m.v === 2 && CODES.indexOf(p.c) >= 0) || (m.tok[p.c] && m.tok[p.c] === p.t) || (p.s && csess_(p.s) === p.c))) ? p.c : null;
  const c = JSON.parse(JSON.stringify(m));
  delete c.ht; delete c.tok; delete c.owner;
  if (!host) {
    delete c.hostEmails; delete c.ackEmail; delete c.notedEmail;
    CODES.forEach(function (k) { if (c.con[k]) delete c.con[k].email; });
    if (who && CODES.indexOf(who) >= 0 && m.con[who]) c.con[who].email = m.con[who].email;
    if (m.status === 'open' && !who && m.v !== 2) c.con = {};
  }
  Object.keys(c.sign || {}).forEach(function (k) { c.sign[k].email = maskList_(c.sign[k].emailFull); delete c.sign[k].emailFull; });
  if (host) c.tok = m.tok;
  return { ok: true, meeting: c, role: host ? 'host' : (who ? 'con' : 'public'), who: who };
}
function verify_(id) {
  const m = load_(id);
  if (!m) return { ok: false, error: 'not found' };
  if (!m.H) return { ok: true, status: m.status, sealed: false };
  const Hnow = contentHash_(m);
  const sigs = Object.keys(m.sign).map(function (k) {
    const s = m.sign[k];
    return { who: k, label: NAMES[k], name: s.name, pos: s.pos || '', email: maskList_(s.emailFull), at: s.at, method: s.method, ok: sigHash_(k, s, m.H) === s.sh };
  });
  const contentOk = Hnow === m.H;
  let sealOk = null;
  if (m.seal) sealOk = sealOf_(m) === m.seal;
  return { ok: true, absent: m.absent || {}, sealed: true, status: m.status, contentOk: contentOk, sealOk: sealOk, H: m.H, seal: m.seal || '', sigs: sigs, required: required_(m).length };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.cc && !p.c) p.c = p.cc;   // the app sends the company as "cc" (a GET parameter named "c" is dropped by the gateway)
  if (p.action === 'get') {
    const m = load_(p.id);
    if (!m) return out_({ ok: false, error: 'not found' });
    return out_(view_(m, p));
  }
  if (p.action === 'verify') return out_(verify_(p.id));
  if (p.action === 'current') {   // the meeting contractors should open: the latest one that is not finished
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
    let best = null;
    rows.forEach(function (r) { if (r[4] !== 'done' && (!best || (fd_(r[1]) + ' ' + ft_(r[2])) >= (fd_(best[1]) + ' ' + ft_(best[2])))) best = r; });
    if (!best) return out_({ ok: true, id: '' });
    const res = { ok: true, id: best[0], date: fd_(best[1]), time: ft_(best[2]), status: best[4] };
    if (p.full) { const mm = load_(best[0]); if (mm) { const v = view_(mm, { id: best[0], c: p.c || '' }); if (v && v.ok) { res.role = v.role; res.meeting = v.meeting; } } }   // one round trip: id + meeting
    return out_(res);
  }
  if (p.action === 'list') {
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 7).getValues() : [];
    const items = rows.filter(function (r) { return r[4] !== 'open'; }).reverse().slice(0, 60).map(function (r) { return { id: r[0], date: fd_(r[1]), time: ft_(r[2]), host: r[3], status: r[4], parties: r[5], clashes: r[6] }; });
    return out_({ ok: true, items: items });
  }
  return out_({ ok: true, version: '4.9.0' });
}

function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return out_(handle_(b)); }
  catch (err) { return out_({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}

function handle_(b) {
  // ----- host login (email OTP, no passwords) -----
  if (b.action === 'hostOtp') {
    const em = String(b.email || '').trim().toLowerCase();
    if (hostEmails_().map(function (x) { return x.toLowerCase(); }).indexOf(em) >= 0) otpSend_('host:' + em, em, 'host login');
    return { ok: true };
  }
  if (b.action === 'hostLogin') {
    const em = String(b.email || '').trim().toLowerCase();
    if (!otpCheck_('host:' + em, b.code)) return { ok: false, error: 'wrong or expired code' };
    const s = Utilities.getUuid();
    lsPut_('h', s, em);
    return { ok: true, session: s, email: em };
  }
  // ----- host (admin) PIN sign-in: PIN is set only through an emailed code -----
  if (b.action === 'hostPinOtp') {
    const em = String(b.email || '').trim().toLowerCase();
    if (hostEmails_().map(function (x) { return x.toLowerCase(); }).indexOf(em) >= 0) otpSend_('hpin:' + em, em, 'setting your SIMOPS host PIN');
    return { ok: true };
  }
  if (b.action === 'hostPinSet') {
    const em = String(b.email || '').trim().toLowerCase();
    if (!/^\d{4,8}$/.test(String(b.pin))) return { ok: false, error: 'PIN must be 4 to 8 digits' };
    if (hostEmails_().map(function (x) { return x.toLowerCase(); }).indexOf(em) < 0 || !otpCheck_('hpin:' + em, b.code)) return { ok: false, error: 'wrong or expired code' };
    props_().setProperty('hpin:' + em, pinHash_('host:' + em, String(b.pin)));
    CacheService.getScriptCache().remove('lock:host:' + em);
    return { ok: true };
  }
  if (b.action === 'hostPinLogin') {
    const em = String(b.email || '').trim().toLowerCase(), cache = CacheService.getScriptCache(), lk = 'lock:host:' + em;
    const tries = parseInt(cache.get(lk) || '0', 10);
    if (tries >= 5) return { ok: false, error: 'too many attempts: wait 10 minutes, or reset the PIN by email' };
    const stored = props_().getProperty('hpin:' + em);
    if (!stored) return { ok: false, error: 'no PIN set yet: use Set or forgot PIN' };
    if (pinHash_('host:' + em, String(b.pin)) !== stored) { cache.put(lk, String(tries + 1), 600); return { ok: false, error: 'wrong PIN' }; }
    cache.remove(lk);
    const sid = Utilities.getUuid(); lsPut_('h', sid, em);
    return { ok: true, session: sid, email: em };
  }
  // ----- party sign-in: one link/QR, party picks its name and enters its PIN -----
  if (b.action === 'pinOtp' || b.action === 'pinSet') return { ok: false, error: 'PINs are set by the Admin' };
  if (b.action === 'pinOtp_old') {
    if (ROLES.indexOf(b.role) < 0) return { ok: false, error: 'unknown party' };
    let e = roleEmails_()[b.role];
    if (!e) {  // first time: the party registers its own email (locked after the first PIN is set; host can reset)
      e = String(b.email || '').trim();
      if (!okEmails_(e)) return { ok: false, error: 'enter your email address' };
      props_().setProperty('pending:' + b.role, e);
    }
    otpSend_('pin:' + b.role, e, 'setting your SIMOPS PIN');
    return { ok: true, to: maskList_(e) };
  }
  if (b.action === 'pinSet_old') {
    if (ROLES.indexOf(b.role) < 0 || !/^\d{4,8}$/.test(String(b.pin))) return { ok: false, error: 'PIN must be 4 to 8 digits' };
    if (!otpCheck_('pin:' + b.role, b.code)) return { ok: false, error: 'wrong or expired code' };
    const pend = props_().getProperty('pending:' + b.role);
    if (!roleEmails_()[b.role] && pend) {
      const em = roleEmails_(); em[b.role] = pend;
      props_().setProperty('roleEmails', JSON.stringify(em)); props_().deleteProperty('pending:' + b.role);
    }
    props_().setProperty('pin:' + b.role, pinHash_(b.role, String(b.pin)));
    CacheService.getScriptCache().remove('lock:' + b.role);
    return { ok: true };
  }
  if (b.action === 'roleLogin') {
    if (ROLES.indexOf(b.role) < 0) return { ok: false, error: 'unknown party' };
    const cache = CacheService.getScriptCache(), lk = 'lock:' + b.role;
    const tries = parseInt(cache.get(lk) || '0', 10);
    if (tries >= 5) return { ok: false, error: 'too many attempts: wait 10 minutes, or reset the PIN by email' };
    const stored = props_().getProperty('pin:' + b.role);
    if (!stored) return { ok: false, error: 'no PIN set yet: use Set or forgot PIN' };
    if (pinHash_(b.role, String(b.pin)) !== stored) { cache.put(lk, String(tries + 1), 600); return { ok: false, error: 'wrong PIN' }; }
    cache.remove(lk);
    const sid = Utilities.getUuid();
    lsPut_('c', sid, b.role);
    return { ok: true, session: sid, role: b.role };
  }
  if (b.action === 'myMeetings') {
    const role = csess_(b.s);
    if (!role) return { ok: false, error: 'session' };
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues().reverse() : [];
    const mk = function (r) { return { id: r[0], date: fd_(r[1]), time: ft_(r[2]), status: r[4] }; };
    const items = rows.filter(function (r) { return r[4] !== 'done'; }).slice(0, 20).map(mk)
      .concat(rows.filter(function (r) { return r[4] === 'done'; }).slice(0, 5).map(mk));
    return { ok: true, role: role, items: items };
  }
  // ----- host registry: party emails and PIN reset -----
  if (b.action === 'regGet') {
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    const em = roleEmails_(), reg = {};
    ROLES.forEach(function (r) { reg[r] = { email: em[r] || '', pinSet: !!props_().getProperty('pin:' + r), code: CTL.indexOf(r) >= 0 ? (props_().getProperty('pinplain:' + r) || '') : '' }; });
    return { ok: true, reg: reg };
  }
  if (b.action === 'regSet') {
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    const clean = {};
    for (let i = 0; i < ROLES.length; i++) {
      const v = String((b.emails || {})[ROLES[i]] || '').trim();
      if (v && !okEmails_(v)) return { ok: false, error: 'invalid email for ' + NAMES[ROLES[i]] };
      if (v) clean[ROLES[i]] = v;
    }
    props_().setProperty('roleEmails', JSON.stringify(clean));
    return { ok: true };
  }
  if (b.action === 'adminSetPin') {   // Admin (host session) sets the EcoOils Rep / EcoOils HSSE email(s) and PIN
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    if (ROLES.indexOf(b.role) < 0) return { ok: false, error: 'unknown role' };
    const em = roleEmails_(), e = String(b.email || '').trim();
    if (e) { if (!okEmails_(e)) return { ok: false, error: 'invalid email' }; em[b.role] = e; props_().setProperty('roleEmails', JSON.stringify(em)); }
    if (b.pin) {
      if (!/^\d{4,8}$/.test(String(b.pin))) return { ok: false, error: 'PIN must be 4 to 8 digits' };
      props_().setProperty('pin:' + b.role, pinHash_(b.role, String(b.pin)));
      props_().setProperty('pinplain:' + b.role, String(b.pin));   // kept in Script properties (owner-only, never in the sheet or code) so the Admin can look the code up again
      CacheService.getScriptCache().remove('lock:' + b.role);
    }
    return { ok: true };
  }
  if (b.action === 'pinReset') {
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    if (ROLES.indexOf(b.role) < 0) return { ok: false, error: 'unknown party' };
    props_().deleteProperty('pin:' + b.role); props_().deleteProperty('pinplain:' + b.role);
    const em = roleEmails_(); delete em[b.role]; props_().setProperty('roleEmails', JSON.stringify(em));
    CacheService.getScriptCache().remove('lock:' + b.role);
    return { ok: true };
  }

  if (b.action === 'delete') {   // Admin removes a meeting: record and PDF go to the Drive bin, the register row is removed
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    const dm = load_(b.id);
    if (!dm) return { ok: false, error: 'not found' };
    try { file_(b.id).setTrashed(true); } catch (e) {}
    try { if (dm.pdfId) DriveApp.getFileById(dm.pdfId).setTrashed(true); } catch (e) {}
    CacheService.getScriptCache().remove('fid:' + b.id);
    const dsh = sheet_(), dl = dsh.getLastRow();
    const dids = dl > 1 ? dsh.getRange(2, 1, dl - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    const di = dids.indexOf(b.id);
    if (di >= 0) dsh.deleteRow(di + 2);
    return { ok: true };
  }
  if (b.action === 'create') {
    const owner = sess_(b.session);
    if (!owner) return { ok: false, error: 'session' };
    const m = newMeeting_({ owner: owner, base: b.base, host: b.host, date: b.date, time: b.time, hostEmails: b.hostEmails || owner, ackEmail: '', notedEmail: '', areas: b.areas || [] });
    return { ok: true, id: m.id, ht: m.ht };
  }
  if (b.action === 'mine') {
    const owner = sess_(b.session);
    if (!owner) return { ok: false, error: 'session' };
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 11).getValues() : [];
    return { ok: true, items: rows.filter(function (r) { return r[9] === owner; }).reverse().slice(0, 40).map(function (r) { return { id: r[0], date: fd_(r[1]), time: ft_(r[2]), status: r[4], ht: r[10] }; }) };
  }

  const m = load_(b.id);
  if (!m) return { ok: false, error: 'not found' };
  const isHost = !!(b.ht && b.ht === m.ht);
  if (b.action === 'get') return view_(m, b);
  if (b.action === 'pdfData') {   // the signed PDF as data, so the phone can attach it straight into WhatsApp (the record is public anyway)
    if (m.status !== 'done') return { ok: false, error: 'not signed yet' };
    let blob;
    try { blob = m.pdfId ? DriveApp.getFileById(m.pdfId).getBlob() : pdfBlob_(m); } catch (err) { blob = pdfBlob_(m); }
    return { ok: true, name: 'SIMOPS_Meeting_Record_' + String(m.date).slice(0, 10) + '.pdf', data: Utilities.base64Encode(blob.getBytes()) };
  }
  // ----- AI wording check (free Gemini tier; key lives in Script Properties as GEMINI_API_KEY, never in code) -----
  if (b.action === 'polish') return { ok: false, error: 'AI check is removed; use the basic correction' };
  if (b.action === 'polish_old') {
    if (!(isHost || roleAuth_(m, b) || (m.v === 2 && CODES.indexOf(b.c) >= 0))) return { ok: false, error: 'choose your company first' };
    if (m.status !== 'open' && !isHost) return { ok: false, error: 'entries are locked' };
    const key = props_().getProperty('GEMINI_API_KEY');
    if (!key) return { ok: false, error: 'AI wording check is not set up yet (ask HSSE)' };
    const multi = Array.isArray(b.texts);
    const parts = (multi ? b.texts : [b.text]).slice(0, 4).map(function (x) { return String(x || '').trim().slice(0, 600); });
    if (parts.every(function (x) { return x.length < 3; })) return multi ? { ok: true, texts: parts } : { ok: true, text: parts[0] };
    const txt = parts.length === 1 ? parts[0] : parts.map(function (x, i) { return '[' + (i + 1) + '] ' + x; }).join('\n');
    const cache = CacheService.getScriptCache(), rk = 'pol:' + (b.c || 'host') + ':' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMddHH');
    const used = parseInt(cache.get(rk) || '0', 10);
    if (used >= 40) return { ok: false, error: 'limit reached for this hour, try again later' };
    cache.put(rk, String(used + 1), 3600);
    const cached = CacheService.getScriptCache().get('pol:model');
    const models = [cached, props_().getProperty('GEMINI_MODEL'), 'gemini-3.8-flash', 'gemini-flash-latest', 'gemini-2.5-flash'].filter(function (x, i, a) { return x && a.indexOf(x) === i; });
    const prompt = 'You edit short notes written by construction site contractors for a SIMOPS (simultaneous operations) meeting record. ' +
      'Rewrite the note below as clear, correct, professional sentence(s). Keep the same language as the note (English, Bahasa Indonesia or Bahasa Malaysia). ' +
      'Keep the meaning exactly. Do not add facts, names, numbers, times or equipment that are not in the note. Keep it short. Output only the rewritten text, nothing else.' + (parts.length > 1 ? ' There are ' + parts.length + ' notes marked [1], [2]...; return them in the same order, each starting with its marker [n] on its own line.' : '') + '\n\nNote:\n' + txt;
    let res = null;
    for (let i = 0; i < models.length; i++) {
      try {
        res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          headers: { 'x-goog-api-key': key },
          payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, maxOutputTokens: 300 } })
        });
      } catch (err) { return { ok: false, error: 'AI service not reachable: ' + String(err && err.message || err).slice(0, 160) }; }
      if (res.getResponseCode() !== 404) { if (res.getResponseCode() === 200) cache.put('pol:model', models[i], 21600); break; }
    }
    if (res.getResponseCode() !== 200) return { ok: false, error: 'AI service error ' + res.getResponseCode() + ': ' + String(res.getContentText()).slice(0, 160) };
    let out = '';
    try { out = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text; } catch (err) { return { ok: false, error: 'AI gave no answer' }; }
    out = String(out || '').trim();
    if (parts.length === 1) { out = out.replace(/^\[1\]\s*/, '').replace(/^["\s]+|["\s]+$/g, '').slice(0, 800); return out ? (multi ? { ok: true, texts: [out] } : { ok: true, text: out }) : { ok: false, error: 'AI gave no answer' }; }
    const got = []; out.split(/^\s*\[(\d)\]\s*/m).forEach(function (seg, i, a) { if (i % 2 === 1) got[parseInt(seg, 10) - 1] = String(a[i + 1] || '').trim().replace(/^["\s]+|["\s]+$/g, '').slice(0, 800); });
    const res2 = parts.map(function (x, i) { return got[i] || x; });
    return { ok: true, texts: res2 };
  }

  const isCtl = isHost || (CTL.indexOf(b.c) >= 0 && roleAuth_(m, b));   // host, HOD (ack) or HSSE (noted) signed in with their PIN
  if (b.action === 'meta') {
    if (!isHost && !(isCtl && b.hostc !== undefined)) return { ok: false, error: 'host only' };
    if (!isHost) b = { action: 'meta', id: b.id, hostc: String(b.hostc || '').slice(0, 2000) };
    if (m.status === 'signing' || m.status === 'done') return { ok: false, error: 'record is locked for signing' };
    ['host', 'date', 'time', 'hostEmails', 'ackEmail', 'notedEmail', 'areas', 'hostc'].forEach(function (k) { if (b[k] !== undefined) m[k] = b[k]; });
    save_(m); return { ok: true };
  }
  if (b.action === 'entry' && m.v === 2) {   // open entry: no PIN, any contractor of the company can update until the discussion starts
    if (CODES.indexOf(b.c) < 0) return { ok: false, error: 'choose your company' };
    if (m.status !== 'open') return { ok: false, error: 'entries are locked' };
    const nm = String(b.name || '').trim().slice(0, 60);
    if (!nm) return { ok: false, error: 'enter your name' };
    m.con[b.c] = { name: nm, rows: b.rows || [], updated: new Date().toISOString() };
    m.log = (m.log || []).concat([{ at: new Date().toISOString(), c: b.c, by: nm }]).slice(-40);
    save_(m, true); return Object.assign({ ok: true }, view_(m, b));   // returns the fresh meeting so the app needs no second request
  }
  if (b.action === 'entry') {
    if (!(CODES.indexOf(b.c) >= 0 && roleAuth_(m, b))) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'open') return { ok: false, error: 'entries are locked' };
    const em = roleEmails_()[b.c] || String(b.email || '').trim();
    if (!okEmail_(em)) return { ok: false, error: 'a valid email is required (used to verify your signature)' };
    m.con[b.c] = { name: b.name || '', email: em, rows: b.rows || [], updated: new Date().toISOString() };
    save_(m); return { ok: true };
  }
  if (b.action === 'people') {
    const w = b.c;
    if (CODES.indexOf(w) < 0 || !(isHost || m.v === 2 || roleAuth_(m, b))) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'open' && m.status !== 'discussion') return { ok: false, error: 'attendance is locked' };
    const list = (Array.isArray(b.list) ? b.list : []).slice(0, 15).map(function (x) {
      return { n: String((x && x.n) || '').trim().slice(0, 60), p: String((x && x.p) || '').trim().slice(0, 60), x: !!(x && x.x) };
    }).filter(function (x) { return x.n; });
    m.att = m.att || {}; m.att[w] = list;
    save_(m, true); return Object.assign({ ok: true }, view_(m, b));
  }
  if (b.action === 'status') {
    if (!isCtl) return { ok: false, error: 'HOD, HSSE or host only' };
    if (['open', 'discussion', 'signing'].indexOf(b.to) < 0) return { ok: false, error: 'bad status' };
    if (m.status === 'done') return { ok: false, error: 'completed' };
    if (m.status === 'signing' && Object.keys(m.sign).length) return { ok: false, error: 'signing already started' };
    if (b.to !== 'signing') delete m.absent;
    if (b.to !== 'open') {
      m.parts = CODES.filter(function (c) { return m.con[c] && m.con[c].rows && m.con[c].rows.length; });
      if (!m.parts.length) return { ok: false, error: 'no entries yet' };
    }
    if (b.to === 'signing') {
      if (!(m.photos || []).length) return { ok: false, error: 'a meeting photo is required before signing' };
      const miss = required_(m).filter(function (r) { return CTL.indexOf(r) < 0 && !okEmails_(emailOf_(m, r)); });   // EcoOils Rep / HSSE sign with their code: no email needed
      if (miss.length) return { ok: false, error: 'email missing for: ' + miss.map(function (r) { return NAMES[r]; }).join(', ') };
      m.H = contentHash_(m); m.frozenAt = new Date().toISOString();
    } else { delete m.H; delete m.frozenAt; }
    m.status = b.to; save_(m); return { ok: true };
  }
  if (b.action === 'photo') {
    if (!(isHost || roleAuth_(m, b) || (m.v === 2 && CODES.indexOf(b.c) >= 0))) return { ok: false, error: 'choose your company first' };
    if (m.status !== 'discussion' && !(m.v === 2 && m.status === 'open')) return { ok: false, error: 'photo can be added before signing only' };
    if (m.photos.length >= 4) return { ok: false, error: 'max 4 photos' };
    m.photos.push(b.data); save_(m, true); return Object.assign({ ok: true }, view_(m, b));
  }
  if (b.action === 'otpSend') {
    const w = b.c;
    if (!roleAuth_(m, b)) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'signing' || required_(m).indexOf(w) < 0) return { ok: false, error: 'not available' };
    if (m.sign[w]) return { ok: false, error: 'already signed' };
    const em = emailOf_(m, w);
    otpSend_('sign:' + m.id + ':' + w, em, 'signing the SIMOPS record of ' + m.date);
    return { ok: true, to: maskList_(em) };
  }
  if (b.action === 'sign') {
    const w = b.who;
    if (!(b.c === w && roleAuth_(m, b))) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'signing' || required_(m).indexOf(w) < 0) return { ok: false, error: 'not available' };
    if (m.sign[w]) return { ok: false, error: 'already signed' };
    if (!b.name || !b.sign) return { ok: false, error: 'name and signature required' };
    if (!String(b.pos || '').trim()) return { ok: false, error: 'position required' };
    const byCode = m.v === 2 && CTL.indexOf(w) >= 0;   // signed in with the sign-off code set by the Admin: no email step
    if (!byCode && !otpCheck_('sign:' + m.id + ':' + w, b.code)) return { ok: false, error: 'wrong or expired code' };
    const s = { name: String(b.name).trim(), emailFull: byCode ? (emailOf_(m, w) || '') : emailOf_(m, w), at: Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd HH:mm 'WIB'"), method: byCode ? 'Sign-off code' : 'Email one-time code', pos: String(b.pos).trim().slice(0, 60), sign: b.sign };
    s.sh = sigHash_(w, s, m.H);
    m.sign[w] = s;
    if (!finish_(m)) save_(m);
    return { ok: true };
  }
  if (b.action === 'absent') {   // the EcoOils Rep or EcoOils HSSE marks the other as not present, so the one who is present can sign alone
    if (m.v !== 2 || m.status !== 'signing') return { ok: false, error: 'not available' };
    const by = b.c, who = b.who;
    if (!(CTL.indexOf(by) >= 0 && roleAuth_(m, b)) || CTL.indexOf(who) < 0 || who === by) return { ok: false, error: 'invalid sign-in' };
    if (m.sign[who]) return { ok: false, error: 'already signed' };
    m.absent = m.absent || {};
    if (b.on) m.absent[who] = { by: by, at: Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd HH:mm 'WIB'") }; else delete m.absent[who];
    if (!finish_(m)) save_(m);
    return { ok: true };
  }
  if (b.action === 'resend') {
    if (!isHost || m.status !== 'done') return { ok: false, error: 'not available' };
    sendPdf_(m); return { ok: true };
  }
  return { ok: false, error: 'unknown action' };
}

// ---------- PDF + email ----------
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function pdfBlob_(m) {
  const cl = clashList_(m);
  const T = 'border="1" cellspacing="0" cellpadding="4" width="100%" style="border-collapse:collapse"';
  let h = '<html><body style="font-family:Arial;font-size:11px;color:#111">';
  h += '<h2 style="letter-spacing:3px;color:#555;margin-bottom:2px">SIMOPS MEETING RECORD</h2><div style="color:#777;font-size:9px;letter-spacing:2px">OFFICIAL ELECTRONIC RECORD</div>';
  h += '<p><b>Host:</b> ' + esc_(m.host) + ' &nbsp; <b>Date:</b> ' + esc_(m.date) + ' &nbsp; <b>Time:</b> ' + esc_(m.time) + '</p>';
  if (cl.length) h += '<p style="background:#fff3cd;padding:6px"><b>Possible clashes:</b><br>' + cl.map(esc_).join('<br>') + '</p>';
  h += '<h3 style="margin-bottom:3px">1. Planned activities and constraints</h3>';
  h += '<table ' + T + '><tr style="background:#eee"><th>Contractor</th><th>Area</th><th>Planned activity</th><th>Agreed constraint</th></tr>';
  CODES.forEach(function (c) {
    const rows = (m.con[c] && m.con[c].rows) || []; if (!rows.length) return;
    rows.forEach(function (r, i) {
      h += '<tr>' + (i === 0 ? '<td rowspan="' + rows.length + '"><b>' + NAMES[c] + '</b>' + (m.con[c].name ? '<br><span style="font-size:9px">updated by ' + esc_(m.con[c].name) + '</span>' : '') + '</td>' : '') +
        '<td>' + esc_(r.area) + '</td><td>' + esc_(r.activity) + (r.details ? '<br><i>' + esc_(r.details) + '</i>' : '') + '</td><td>' + esc_(r.constraint) + '</td></tr>';
    });
  });
  h += '</table>';
  let anyAtt = false, ah = '<h3 style="margin-bottom:3px">2. Attendance</h3><table ' + T + '><tr style="background:#eee"><th>Company</th><th>Name</th><th>Position</th><th>Present</th></tr>';
  CODES.forEach(function (c) {
    const a = (m.att && m.att[c]) || []; if (!a.length) return; anyAtt = true;
    a.forEach(function (x, i) { ah += '<tr>' + (i === 0 ? '<td rowspan="' + a.length + '"><b>' + NAMES[c] + '</b></td>' : '') + '<td>' + esc_(x.n) + '</td><td>' + esc_(x.p) + '</td><td align="center">' + (x.x ? 'Present' : 'Absent') + '</td></tr>'; });
  });
  if (anyAtt) h += ah + '</table>';
  if (m.photos && m.photos.length) h += '<h3 style="margin-bottom:3px">3. Meeting photo</h3><p>' + m.photos.map(function (p) { return '<img src="' + p + '" height="140"> '; }).join('') + '</p>';
  h += '<p><b>Host comments:</b> ' + esc_(m.hostc) + '</p>';
  h += '<h3 style="margin-bottom:3px">4. Sign-off</h3><p style="font-size:9px;color:#555">Each signatory confirmed that they represent the company shown, attended the meeting, and that this record is accurate and agreed. Each signature was verified by a one-time code sent to the signatory\'s registered email and sealed to the content hash below.</p>';
  h += '<table ' + T + '><tr style="background:#eee"><th>Company / role</th><th>Name and position</th><th>Signature</th><th>Verification</th></tr>';
  (m.v === 2 ? ['ack', 'noted'] : required_(m)).forEach(function (k) {
    const s = m.sign[k], ab = m.absent && m.absent[k];
    if (ab) { h += '<tr><td><b>' + esc_(NAMES[k]) + '</b></td><td colspan="3"><i>Not present at the meeting</i> (marked by ' + esc_(NAMES[ab.by]) + ', ' + esc_(ab.at) + ')</td></tr>'; return; }
    h += '<tr><td><b>' + esc_(NAMES[k]) + '</b></td><td>' + (s ? esc_(s.name) + '<br><i>' + esc_(s.pos || '') + '</i>' : '') + '</td><td>' + (s ? '<img src="' + s.sign + '" height="40">' : '') + '</td><td style="font-size:9px">' +
      (s ? '<b style="color:#1b7f3b">DIGITALLY SIGNED</b><br>' + esc_(s.at) + '<br>' + esc_(s.method) + '<br>' + esc_(mask_(s.emailFull)) + '<br>Ref ' + esc_(String(s.sh).slice(0, 12)) : '') + '</td></tr>';
  });
  h += '</table>';
  h += '<hr><p style="font-size:9px;color:#555">Content hash (SHA-256): ' + m.H + '<br>Seal: ' + m.seal + '<br>Verify online: ' + esc_(m.base) + '?v=' + m.id + '</p>';
  h += '</body></html>';
  return Utilities.newBlob(h, 'text/html', 'SIMOPS.html').getAs('application/pdf').setName('SIMOPS_Meeting_Record_' + m.date + '.pdf');
}
function makePdf_(m) {
  const f = folder_().createFile(pdfBlob_(m));
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* sharing may be restricted by the account; the email attachment still works */ }
  return f;
}
function sendPdf_(m) {
  const list = [m.hostEmails, m.ackEmail, m.notedEmail, roleEmails_().ack, roleEmails_().noted];
  m.parts.forEach(function (c) { if (m.con[c] && m.con[c].email) list.push(m.con[c].email); });
  const to = list.join(',').split(/[,;\s]+/).filter(okEmail_).filter(function (x, i, a) { return a.indexOf(x) === i; }).join(',');
  if (!to) return;
  MailApp.sendEmail({
    to: to, subject: 'SIMOPS Meeting Record - ' + m.date + ' ' + (m.time || ''),
    body: 'All parties have signed. The SIMOPS Meeting Record is attached.\n\nHost: ' + m.host + '\nRecord: ' + m.base + '?m=' + m.id + '\nVerify authenticity: ' + m.base + '?v=' + m.id + (m.next ? '\n\nThe next SIMOPS is now open for entries. Sign in at ' + m.base + ' and add your plans for the next day.' : ''),
    attachments: [pdfBlob_(m)]
  });
}

// Run this ONCE from the Apps Script editor (select authorizeAll, click Run, Review permissions, Allow).
// It grants the web app permission to send email, use Drive and call the free Gemini AI service.
function authorizeAll() {
  UrlFetchApp.fetch('https://generativelanguage.googleapis.com/', { muteHttpExceptions: true });
  DriveApp.getRootFolder();
  MailApp.getRemainingDailyQuota();
  SpreadsheetApp.getActive();
  Logger.log('Authorized. Now deploy a New version of the web app.');
}
