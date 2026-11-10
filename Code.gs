// VERSION: 5.12.0 - "Others present" list (guests, visitors) entered by Admin, EcoOils Rep and EcoOils HSSE, sealed and shown on the PDF; Controls are now part of the sealed record
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
    if (e && m.parts.indexOf(c) >= 0) con[c] = { name: e.name, rows: e.rows.map(function (r) { var o = { area: r.area, activity: r.activity, details: r.details || '', constraint: r.constraint }; if (r.controls) o.controls = r.controls; return o; }) };
  });
  const out = { id: m.id, host: m.host, date: m.date, time: m.time, parts: m.parts, con: con, hostc: m.hostc || '', photos: (m.photos || []).map(sha_) };
  const att = {};
  CODES.forEach(function (c) { const a = (m.att && m.att[c]) || []; if (a.length) att[c] = a.map(function (x) { return { n: x.n, p: x.p, x: !!x.x }; }); });
  if (m.others && m.others.length) out.others = m.others.map(function (x) { return { n: x.n, o: x.o || '', p: x.p || '' }; });   // only present when entered, so older sealed records still verify
  if (Object.keys(att).length) out.att = att;   // only present when an attendance list exists, so older sealed records still verify
  return out;
}
function contentHash_(m) { return sha_(JSON.stringify(content_(m))); }
function sigHash_(who, s, H) { const parts = [who, s.name, s.emailFull, s.at, H, sha_(s.sign)]; if (s.pos) parts.push(s.pos); if (s.as) parts.push(s.as); return sha_(parts.join('|')); }
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
    return { who: k, label: NAMES[k], name: s.name, pos: s.pos || '', as: s.as || '', email: maskList_(s.emailFull), at: s.at, method: s.method, ok: sigHash_(k, s, m.H) === s.sh };
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
    return { ok: true, items: rows.filter(function (r) { return r[9] === owner; }).reverse().slice(0, 40).map(function (r) { return { id: r[0], date: fd_(r[1]), time: ft_(r[2]), status: r[4], ht: r[10], pdf: r[8] || '' }; }) };
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
  if (b.action === 'consEdit') {   // Admin (host link), EcoOils Rep and EcoOils HSSE can edit or add constraints until signing starts
    const ctl = CTL.indexOf(b.c) >= 0 && roleAuth_(m, b);
    if (m.v !== 2 || !(isHost || ctl)) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'open' && m.status !== 'discussion') return { ok: false, error: 'constraints are locked once signing starts' };
    const by = isHost ? 'Admin' : (b.c === 'ack' ? 'EcoOils Rep' : 'EcoOils HSSE');
    let n = 0;
    (Array.isArray(b.edits) ? b.edits : []).forEach(function (e) {
      const rows = e && m.con[e.c] && m.con[e.c].rows, i = e ? parseInt(e.i, 10) : -1;
      if (!rows || !(i >= 0 && i < rows.length)) return;
      const v = String(e.constraint || '').trim().slice(0, 300), ct = String(e.controls || '').trim().slice(0, 1000);
      if (v !== String(rows[i].constraint || '') || ct !== String(rows[i].controls || '')) { rows[i].constraint = v; rows[i].controls = ct; rows[i].ed = by; n++; }
    });
    (Array.isArray(b.adds) ? b.adds : []).slice(0, 20).forEach(function (a) {
      if (!a || CODES.indexOf(a.c) < 0) return;
      const act = String(a.activity || '').trim().slice(0, 200), cons = String(a.constraint || '').trim().slice(0, 300);
      const ctl2 = String(a.controls || '').trim().slice(0, 1000);
      if (!act || !(cons || ctl2)) return;
      m.con[a.c] = m.con[a.c] || { name: '', rows: [] };
      m.con[a.c].rows = (m.con[a.c].rows || []).concat([{ area: String(a.area || '').trim().toUpperCase().slice(0, 100), activity: act, details: '', constraint: cons, controls: ctl2, ed: by }]);
      n++;
    });
    if (n) m.log = (m.log || []).concat([{ at: new Date().toISOString(), c: 'constraints', by: by }]).slice(-40);
    save_(m, true); return Object.assign({ ok: true, changed: n }, view_(m, b));
  }
  if (b.action === 'others') {   // Admin, EcoOils Rep and EcoOils HSSE record anyone else who was present (guests, visitors, other departments)
    const ctl = CTL.indexOf(b.c) >= 0 && roleAuth_(m, b);
    if (m.v !== 2 || !(isHost || ctl)) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'open' && m.status !== 'discussion') return { ok: false, error: 'the list is locked once signing starts' };
    m.others = (Array.isArray(b.list) ? b.list : []).slice(0, 30).map(function (x) {
      return { n: String((x && x.n) || '').trim().slice(0, 80), o: String((x && x.o) || '').trim().slice(0, 80), p: String((x && x.p) || '').trim().slice(0, 80) };
    }).filter(function (x) { return x.n; });
    m.log = (m.log || []).concat([{ at: new Date().toISOString(), c: 'others', by: isHost ? 'Admin' : (b.c === 'ack' ? 'EcoOils Rep' : 'EcoOils HSSE') }]).slice(-40);
    save_(m, true); return Object.assign({ ok: true }, view_(m, b));
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
    const asRep = (m.v === 2 && w === 'ack') ? String(b.as || '') : '';
    if (m.v === 2 && w === 'ack' && asRep !== 'ECO' && asRep !== 'EOO') return { ok: false, error: 'choose EcoOils Project Rep or EcoOils Operations Rep' };
    const byCode = m.v === 2 && CTL.indexOf(w) >= 0;   // signed in with the sign-off code set by the Admin: no email step
    if (!byCode && !otpCheck_('sign:' + m.id + ':' + w, b.code)) return { ok: false, error: 'wrong or expired code' };
    const s = { name: String(b.name).trim(), emailFull: byCode ? (emailOf_(m, w) || '') : emailOf_(m, w), at: Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd HH:mm 'WIB'"), method: byCode ? 'Sign-off code' : 'Email one-time code', pos: String(b.pos).trim().slice(0, 60), sign: b.sign };
    if (asRep) s.as = asRep;
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
const LOGO_ = 'data:image/png;base64,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';
function pdfHtml_(m) {
  const cl = clashList_(m);
  const G = '#1b7f3b', GR = '#555555', LN = '#bbbbbb';
  const TH = 'style="background:#eeeeee;border:1px solid ' + LN + ';padding:4px;text-align:left;font-size:9px;letter-spacing:1px"';
  const TD = 'style="border:1px solid ' + LN + ';padding:4px;vertical-align:top;font-size:10px"';
  const T = 'width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse"';
  const H3 = function (n, t) { return '<h3 style="margin:14px 0 4px 0;font-size:12px;letter-spacing:2px;color:' + GR + '">' + (typeof n === 'number' ? n + '. ' : '') + t + '</h3>'; };
  const dt = String(m.date || '').slice(0, 10);
  let h = '<html><body style="font-family:Arial,Helvetica,sans-serif;font-size:10.5px;color:#111111">';
  h += '<table ' + T + '><tr><td style="border-bottom:3px solid ' + G + ';padding-bottom:6px" width="110"><img src="' + LOGO_ + '" width="95"></td><td style="border-bottom:3px solid ' + G + ';padding-bottom:6px"><div style="font-size:19px;letter-spacing:4px;color:#333333"><b>SIMOPS MEETING RECORD</b></div><div style="color:#777777;font-size:9px;letter-spacing:2px;margin-top:3px">EcoOils Sei Mangkei &nbsp;|&nbsp; OFFICIAL ELECTRONIC RECORD</div></td>' +
    '<td align="right" style="border-bottom:3px solid ' + G + ';padding-bottom:6px;font-size:10px"><b style="color:' + G + '">' + (m.status === 'done' ? 'SIGNED AND SEALED' : 'DRAFT') + '</b><br>Ref ' + esc_(String(m.id || '').slice(0, 8).toUpperCase()) + '</td></tr></table>';
  h += '<table ' + T + ' style="margin-top:8px"><tr><td ' + TD + '><span style="color:#777777;font-size:8px">HOST</span><br><b>' + esc_(m.host) + '</b></td><td ' + TD + '><span style="color:#777777;font-size:8px">DATE</span><br><b>' + esc_(dt) + '</b></td><td ' + TD + '><span style="color:#777777;font-size:8px">TIME</span><br><b>' + esc_(m.time) + '</b></td><td ' + TD + '><span style="color:#777777;font-size:8px">CONTRACTORS</span><br><b>' + (m.parts || []).length + '</b></td></tr></table>';
  if (cl.length) h += '<p style="background:#fff3cd;padding:6px;border:1px solid #e6c200;margin:8px 0"><b>Possible clashes to manage:</b><br>' + cl.map(esc_).join('<br>') + '</p>';
  h += H3(1, 'PLANNED ACTIVITIES AND AGREED CONSTRAINTS');
  h += '<table ' + T + '><tr><th ' + TH + '>CONTRACTOR</th><th ' + TH + '>AREA</th><th ' + TH + '>PLANNED ACTIVITY</th><th ' + TH + '>AGREED CONSTRAINT AND CONTROLS</th></tr>';
  CODES.forEach(function (c) {
    const rows = (m.con[c] && m.con[c].rows) || []; if (!rows.length) return;
    rows.forEach(function (r, i) {
      h += '<tr>' + (i === 0 ? '<td ' + TD + ' rowspan="' + rows.length + '"><b>' + esc_(NAMES[c]) + '</b>' + (m.con[c].name ? '<br><span style="font-size:8px;color:#777777">entered by ' + esc_(m.con[c].name) + '</span>' : '') + '</td>' : '') +
        '<td ' + TD + '>' + esc_(r.area) + '</td><td ' + TD + '>' + esc_(r.activity) + (r.details ? '<br><i style="color:#555555">' + esc_(r.details) + '</i>' : '') + '</td><td ' + TD + '>' + esc_(r.constraint) + (r.controls ? '<br><b>Controls:</b> ' + esc_(r.controls).replace(/\n/g, '<br>') : '') + (r.ed ? '<br><span style="font-size:8px;color:#777777">updated by ' + esc_(r.ed) + '</span>' : '') + '</td></tr>';
    });
  });
  h += '</table>';
  let anyAtt = false, ah = H3(2, 'ATTENDANCE') + '<table ' + T + '><tr><th ' + TH + '>CONTRACTOR</th><th ' + TH + '>NAME</th><th ' + TH + '>POSITION</th><th ' + TH + '>STATUS</th></tr>';
  CODES.forEach(function (c) {
    const a = (m.att && m.att[c]) || []; if (!a.length) return; anyAtt = true;
    a.forEach(function (x, i) { ah += '<tr>' + (i === 0 ? '<td ' + TD + ' rowspan="' + a.length + '"><b>' + esc_(NAMES[c]) + '</b></td>' : '') + '<td ' + TD + '>' + esc_(x.n) + '</td><td ' + TD + '>' + esc_(x.p) + '</td><td ' + TD + ' align="center">' + (x.x ? '<b style="color:' + G + '">Present</b>' : 'Absent') + '</td></tr>'; });
  });
  if (anyAtt) h += ah + '</table>';
  if (m.others && m.others.length) {
    h += H3('', 'OTHERS PRESENT') + '<table ' + T + '><tr><th ' + TH + '>NAME</th><th ' + TH + '>ORGANISATION / DEPARTMENT</th><th ' + TH + '>POSITION</th></tr>';
    m.others.forEach(function (x) { h += '<tr><td ' + TD + '>' + esc_(x.n) + '</td><td ' + TD + '>' + esc_(x.o) + '</td><td ' + TD + '>' + esc_(x.p) + '</td></tr>'; });
    h += '</table>';
  }
  if (m.hostc) h += H3(3, 'HOST COMMENTS') + '<p style="margin:2px 0;border-left:3px solid ' + LN + ';padding-left:8px">' + esc_(m.hostc).replace(/\n/g, '<br>') + '</p>';
  h += '<div style="page-break-inside:avoid">' + H3(4, 'SIGN-OFF');
  const byCode = (m.v === 2);
  h += '<p style="font-size:9px;color:' + GR + ';margin:2px 0 6px 0">Each signatory confirmed that they represent the contractor or party shown, attended the meeting, and that this record is accurate and agreed. ' + (byCode ? 'Each signature was made with the personal sign-off code issued by the HSSE Admin' : 'Each signature was verified by a one-time code sent to the signatory\'s registered email') + ' and is sealed to the content hash below.</p>';
  h += '<table ' + T + '><tr><th ' + TH + '>ROLE</th><th ' + TH + '>NAME AND POSITION</th><th ' + TH + '>SIGNATURE</th><th ' + TH + '>VERIFICATION</th></tr>';
  (m.v === 2 ? ['ack', 'noted'] : required_(m)).forEach(function (k) {
    const s = m.sign[k], ab = m.absent && m.absent[k];
    if (ab) { h += '<tr><td ' + TD + '><b>' + esc_(NAMES[k]) + '</b></td><td ' + TD + ' colspan="3"><i>Not present at the meeting</i> (marked by ' + esc_(NAMES[ab.by]) + ', ' + esc_(ab.at) + ')</td></tr>'; return; }
    h += '<tr><td ' + TD + '>' + (s && s.as ? '<b>' + esc_(NAMES[s.as]) + ' Rep</b><br><span style="font-size:9px">Acknowledged</span>' : '<b>' + esc_(NAMES[k]) + '</b>') + '</td><td ' + TD + '>' + (s ? '<b>' + esc_(s.name) + '</b><br><i>' + esc_(s.pos || '') + '</i>' : '') + '</td><td ' + TD + '>' + (s ? '<img src="' + s.sign + '" height="42">' : '') + '</td><td ' + TD + ' style="font-size:8.5px">' +
      (s ? '<b style="color:' + G + '">DIGITALLY SIGNED</b><br>' + esc_(s.at) + '<br>' + esc_(s.method) + (s.emailFull ? '<br>' + esc_(mask_(s.emailFull)) : '') + '<br>Ref ' + esc_(String(s.sh).slice(0, 12)) : '') + '</td></tr>';
  });
  h += '</table></div>';
  h += '<p style="font-size:8px;color:#777777;margin-top:14px;border-top:1px solid ' + LN + ';padding-top:5px">Content hash (SHA-256): ' + esc_(m.H) + '<br>Seal: ' + esc_(m.seal) + '<br>Verify this record online: ' + esc_(m.base) + '?v=' + esc_(m.id) + '</p>';
  if (m.photos && m.photos.length) {
    h += '<div style="page-break-before:always"></div>' + H3('Annex', 'MEETING PHOTO (SUPPORTING PROOF)');
    m.photos.forEach(function (p) { h += '<p style="margin:4px 0"><img src="' + p + '" width="460" style="border:1px solid ' + LN + '"></p>'; });
  }
  h += '</body></html>';
  return h;
}
function pdfBlob_(m) {
  return Utilities.newBlob(pdfHtml_(m), 'text/html', 'SIMOPS.html').getAs('application/pdf').setName('SIMOPS_Meeting_Record_' + String(m.date).slice(0, 10) + '.pdf');
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
