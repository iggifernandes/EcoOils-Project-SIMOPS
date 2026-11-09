// VERSION: 4.5.0 - SIMOPS Meeting Record backend (single sign-in with per-party PIN set by email code, public record, email-OTP signatures, tamper-evident seal, map)
// Setup: Google Sheet > Extensions > Apps Script > paste this file > put the HSSE host email(s) in HOST_EMAILS below >
// Deploy > Web app (Execute as: Me, Access: Anyone). First run asks for Drive, Sheets and Mail permissions.
// Cost: free (about 100 emails/day on a Gmail account; each signature or host login uses one email).

const HOST_EMAILS = ['replace-with-host-email@example.com']; // people allowed to create meetings
const SHEET_NAME = 'SIMOPS';
const FOLDER_NAME = 'SIMOPS Records';
const CODES = ['KKS', 'WEN', 'GSB', 'PBS', 'ECO'];
const ROLES = ['KKS', 'WEN', 'GSB', 'PBS', 'ECO', 'ack', 'noted'];
const NAMES = { KKS: 'KKS', WEN: 'WEN', GSB: 'GSB', PBS: 'PBS', ECO: 'EcoOils', ack: 'Acknowledged (EcoOils Construction Representative)', noted: 'Noted (EcoOils HSSE)' };

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function folder_() { const it = DriveApp.getFoldersByName(FOLDER_NAME); return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME); }
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['ID', 'Date', 'Time', 'Host', 'Status', 'Parties', 'Clashes', 'Updated', 'PDF', 'Owner', 'HostToken']); }
  return sh;
}
function file_(id) { const it = folder_().getFilesByName('M_' + id + '.json'); return it.hasNext() ? it.next() : null; }
function load_(id) { const f = file_(id); return f ? JSON.parse(f.getBlob().getDataAsString()) : null; }
function sha_(s) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function rnd_(n) { return Utilities.getUuid().replace(/-/g, '').slice(0, n); }
function mask_(e) { if (!e) return ''; const p = e.split('@'); return p[0].slice(0, 1) + '***@' + (p[1] || ''); }
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
function content_(m) {
  const con = {};
  CODES.forEach(function (c) {
    const e = m.con[c];
    if (e && m.parts.indexOf(c) >= 0) con[c] = { name: e.name, rows: e.rows.map(function (r) { return { area: r.area, activity: r.activity, details: r.details || '', constraint: r.constraint }; }) };
  });
  return { id: m.id, host: m.host, date: m.date, time: m.time, parts: m.parts, con: con, hostc: m.hostc || '', photos: (m.photos || []).map(sha_) };
}
function contentHash_(m) { return sha_(JSON.stringify(content_(m))); }
function sigHash_(who, s, H) { return sha_([who, s.name, s.emailFull, s.at, H, sha_(s.sign)].join('|')); }
function required_(m) { return m.parts.concat(['ack', 'noted']); }
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
  MailApp.sendEmail({ to: email, subject: 'SIMOPS verification code', body: 'Your verification code for ' + purpose + ' is ' + code + '. It is valid for 10 minutes. If you did not request it, ignore this email.' });
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
function save_(m) {
  m.updated = new Date().toISOString();
  const f = file_(m.id), body = JSON.stringify(m);
  if (f) f.setContent(body); else folder_().createFile('M_' + m.id + '.json', body, MimeType.PLAIN_TEXT);
  const sh = sheet_();
  const row = [m.id, m.date, m.time, m.host, m.status, (m.parts || []).map(function (c) { return NAMES[c]; }).join(', '), clashList_(m).join('; '), m.updated, m.pdf || '', m.owner || '', m.ht];
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  const i = ids.indexOf(m.id);
  if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function view_(m, p) {
  const host = !!(p.ht && p.ht === m.ht);
  const who = (p.c && ((m.tok[p.c] && m.tok[p.c] === p.t) || (p.s && csess_(p.s) === p.c))) ? p.c : null;
  const c = JSON.parse(JSON.stringify(m));
  delete c.ht; delete c.tok; delete c.owner;
  if (!host) {
    delete c.hostEmails; delete c.ackEmail; delete c.notedEmail;
    CODES.forEach(function (k) { if (c.con[k]) delete c.con[k].email; });
    if (who && CODES.indexOf(who) >= 0 && m.con[who]) c.con[who].email = m.con[who].email;
    if (m.status === 'open' && !who) c.con = {};
  }
  Object.keys(c.sign || {}).forEach(function (k) { c.sign[k].email = mask_(c.sign[k].emailFull); delete c.sign[k].emailFull; });
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
    return { who: k, label: NAMES[k], name: s.name, email: mask_(s.emailFull), at: s.at, method: s.method, ok: sigHash_(k, s, m.H) === s.sh };
  });
  const contentOk = Hnow === m.H;
  let sealOk = null;
  if (m.seal) sealOk = sha_(m.H + '|' + required_(m).map(function (k) { return m.sign[k] && m.sign[k].sh; }).join(',')) === m.seal;
  return { ok: true, sealed: true, status: m.status, contentOk: contentOk, sealOk: sealOk, H: m.H, seal: m.seal || '', sigs: sigs, required: required_(m).length };
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'get') {
    const m = load_(p.id);
    if (!m) return out_({ ok: false, error: 'not found' });
    return out_(view_(m, p));
  }
  if (p.action === 'verify') return out_(verify_(p.id));
  if (p.action === 'list') {
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 7).getValues() : [];
    const items = rows.filter(function (r) { return r[4] !== 'open'; }).reverse().slice(0, 60).map(function (r) { return { id: r[0], date: r[1], time: r[2], host: r[3], status: r[4], parties: r[5], clashes: r[6] }; });
    return out_({ ok: true, items: items });
  }
  return out_({ ok: true, version: '4.5.0' });
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
    if (HOST_EMAILS.map(function (x) { return x.toLowerCase(); }).indexOf(em) >= 0) otpSend_('host:' + em, em, 'host login');
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
    if (HOST_EMAILS.map(function (x) { return x.toLowerCase(); }).indexOf(em) >= 0) otpSend_('hpin:' + em, em, 'setting your SIMOPS host PIN');
    return { ok: true };
  }
  if (b.action === 'hostPinSet') {
    const em = String(b.email || '').trim().toLowerCase();
    if (!/^\d{4,8}$/.test(String(b.pin))) return { ok: false, error: 'PIN must be 4 to 8 digits' };
    if (HOST_EMAILS.map(function (x) { return x.toLowerCase(); }).indexOf(em) < 0 || !otpCheck_('hpin:' + em, b.code)) return { ok: false, error: 'wrong or expired code' };
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
  if (b.action === 'pinOtp') {
    if (ROLES.indexOf(b.role) < 0) return { ok: false, error: 'unknown party' };
    let e = roleEmails_()[b.role];
    if (!e) {  // first time: the party registers its own email (locked after the first PIN is set; host can reset)
      e = String(b.email || '').trim();
      if (!okEmail_(e)) return { ok: false, error: 'enter your email address' };
      props_().setProperty('pending:' + b.role, e);
    }
    otpSend_('pin:' + b.role, e, 'setting your SIMOPS PIN');
    return { ok: true, to: mask_(e) };
  }
  if (b.action === 'pinSet') {
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
    const mk = function (r) { return { id: r[0], date: r[1], time: r[2], status: r[4] }; };
    const items = rows.filter(function (r) { return r[4] !== 'done'; }).slice(0, 20).map(mk)
      .concat(rows.filter(function (r) { return r[4] === 'done'; }).slice(0, 5).map(mk));
    return { ok: true, role: role, items: items };
  }
  // ----- host registry: party emails and PIN reset -----
  if (b.action === 'regGet') {
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    const em = roleEmails_(), reg = {};
    ROLES.forEach(function (r) { reg[r] = { email: em[r] || '', pinSet: !!props_().getProperty('pin:' + r) }; });
    return { ok: true, reg: reg };
  }
  if (b.action === 'regSet') {
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    const clean = {};
    for (let i = 0; i < ROLES.length; i++) {
      const v = String((b.emails || {})[ROLES[i]] || '').trim();
      if (v && !okEmail_(v)) return { ok: false, error: 'invalid email for ' + NAMES[ROLES[i]] };
      if (v) clean[ROLES[i]] = v;
    }
    props_().setProperty('roleEmails', JSON.stringify(clean));
    return { ok: true };
  }
  if (b.action === 'pinReset') {
    if (!sess_(b.session)) return { ok: false, error: 'session' };
    if (ROLES.indexOf(b.role) < 0) return { ok: false, error: 'unknown party' };
    props_().deleteProperty('pin:' + b.role);
    const em = roleEmails_(); delete em[b.role]; props_().setProperty('roleEmails', JSON.stringify(em));
    CacheService.getScriptCache().remove('lock:' + b.role);
    return { ok: true };
  }

  if (b.action === 'create') {
    const owner = sess_(b.session);
    if (!owner) return { ok: false, error: 'session' };
    const m = {
      id: rnd_(8), ht: rnd_(16), tok: {}, owner: owner, base: b.base,
      host: b.host, date: b.date, time: b.time, hostEmails: b.hostEmails || owner, ackEmail: '', notedEmail: '', areas: b.areas || [],
      status: 'open', con: {}, parts: [], sign: {}, photos: [], hostc: '', created: new Date().toISOString()
    };
    ROLES.forEach(function (r) { m.tok[r] = rnd_(16); });
    save_(m);
    return { ok: true, id: m.id, ht: m.ht };
  }
  if (b.action === 'mine') {
    const owner = sess_(b.session);
    if (!owner) return { ok: false, error: 'session' };
    const sh = sheet_(), last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, 11).getValues() : [];
    return { ok: true, items: rows.filter(function (r) { return r[9] === owner; }).reverse().slice(0, 40).map(function (r) { return { id: r[0], date: r[1], time: r[2], status: r[4], ht: r[10] }; }) };
  }

  const m = load_(b.id);
  if (!m) return { ok: false, error: 'not found' };
  const isHost = !!(b.ht && b.ht === m.ht);
  if (b.action === 'get') return view_(m, b);
  // ----- AI wording check (free Gemini tier; key lives in Script Properties as GEMINI_API_KEY, never in code) -----
  if (b.action === 'polish') {
    if (!(isHost || roleAuth_(m, b))) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'open' && !isHost) return { ok: false, error: 'entries are locked' };
    const key = props_().getProperty('GEMINI_API_KEY');
    if (!key) return { ok: false, error: 'AI wording check is not set up yet (ask HSSE)' };
    const txt = String(b.text || '').trim().slice(0, 600);
    if (txt.length < 3) return { ok: true, text: txt };
    const cache = CacheService.getScriptCache(), rk = 'pol:' + (b.c || 'host') + ':' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMddHH');
    const used = parseInt(cache.get(rk) || '0', 10);
    if (used >= 40) return { ok: false, error: 'limit reached for this hour, try again later' };
    cache.put(rk, String(used + 1), 3600);
    const model = props_().getProperty('GEMINI_MODEL') || 'gemini-2.0-flash';
    const prompt = 'You edit short notes written by construction site contractors for a SIMOPS (simultaneous operations) meeting record. ' +
      'Rewrite the note below as clear, correct, professional sentence(s). Keep the same language as the note (English, Bahasa Indonesia or Bahasa Malaysia). ' +
      'Keep the meaning exactly. Do not add facts, names, numbers, times or equipment that are not in the note. Keep it short. Output only the rewritten text, nothing else.\n\nNote:\n' + txt;
    let res;
    try {
      res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-goog-api-key': key },
        payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.2, maxOutputTokens: 300 } })
      });
    } catch (err) { return { ok: false, error: 'AI service not reachable' }; }
    if (res.getResponseCode() !== 200) return { ok: false, error: 'AI service busy, try again later' };
    let out = '';
    try { out = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text; } catch (err) { return { ok: false, error: 'AI gave no answer' }; }
    out = String(out || '').trim().replace(/^["\s]+|["\s]+$/g, '').slice(0, 800);
    return out ? { ok: true, text: out } : { ok: false, error: 'AI gave no answer' };
  }

  if (b.action === 'meta') {
    if (!isHost) return { ok: false, error: 'host only' };
    if (m.status === 'signing' || m.status === 'done') return { ok: false, error: 'record is locked for signing' };
    ['host', 'date', 'time', 'hostEmails', 'ackEmail', 'notedEmail', 'areas', 'hostc'].forEach(function (k) { if (b[k] !== undefined) m[k] = b[k]; });
    save_(m); return { ok: true };
  }
  if (b.action === 'entry') {
    if (!(CODES.indexOf(b.c) >= 0 && roleAuth_(m, b))) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'open') return { ok: false, error: 'entries are locked' };
    const em = roleEmails_()[b.c] || String(b.email || '').trim();
    if (!okEmail_(em)) return { ok: false, error: 'a valid email is required (used to verify your signature)' };
    m.con[b.c] = { name: b.name || '', email: em, rows: b.rows || [], updated: new Date().toISOString() };
    save_(m); return { ok: true };
  }
  if (b.action === 'status') {
    if (!isHost) return { ok: false, error: 'host only' };
    if (['open', 'discussion', 'signing'].indexOf(b.to) < 0) return { ok: false, error: 'bad status' };
    if (m.status === 'done') return { ok: false, error: 'completed' };
    if (m.status === 'signing' && Object.keys(m.sign).length) return { ok: false, error: 'signing already started' };
    if (b.to === 'open' && m.status === 'signing') return { ok: false, error: 'go back to discussion first' };
    if (b.to !== 'open') {
      m.parts = CODES.filter(function (c) { return m.con[c] && m.con[c].rows && m.con[c].rows.length; });
      if (!m.parts.length) return { ok: false, error: 'no entries yet' };
    }
    if (b.to === 'signing') {
      const miss = required_(m).filter(function (r) { return !okEmail_(emailOf_(m, r)); });
      if (miss.length) return { ok: false, error: 'email missing for: ' + miss.map(function (r) { return NAMES[r]; }).join(', ') };
      m.H = contentHash_(m); m.frozenAt = new Date().toISOString();
    } else { delete m.H; delete m.frozenAt; }
    m.status = b.to; save_(m); return { ok: true };
  }
  if (b.action === 'photo') {
    if (!isHost) return { ok: false, error: 'host only' };
    if (m.status !== 'discussion') return { ok: false, error: 'photo can be added during discussion only' };
    if (m.photos.length >= 4) return { ok: false, error: 'max 4 photos' };
    m.photos.push(b.data); save_(m); return { ok: true };
  }
  if (b.action === 'otpSend') {
    const w = b.c;
    if (!roleAuth_(m, b)) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'signing' || required_(m).indexOf(w) < 0) return { ok: false, error: 'not available' };
    if (m.sign[w]) return { ok: false, error: 'already signed' };
    const em = emailOf_(m, w);
    otpSend_('sign:' + m.id + ':' + w, em, 'signing the SIMOPS record of ' + m.date);
    return { ok: true, to: mask_(em) };
  }
  if (b.action === 'sign') {
    const w = b.who;
    if (!(b.c === w && roleAuth_(m, b))) return { ok: false, error: 'invalid sign-in' };
    if (m.status !== 'signing' || required_(m).indexOf(w) < 0) return { ok: false, error: 'not available' };
    if (m.sign[w]) return { ok: false, error: 'already signed' };
    if (!b.name || !b.sign) return { ok: false, error: 'name and signature required' };
    if (!otpCheck_('sign:' + m.id + ':' + w, b.code)) return { ok: false, error: 'wrong or expired code' };
    const s = { name: String(b.name).trim(), emailFull: emailOf_(m, w), at: Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd HH:mm 'WIB'"), method: 'Email one-time code', sign: b.sign };
    s.sh = sigHash_(w, s, m.H);
    m.sign[w] = s;
    if (required_(m).every(function (r) { return m.sign[r]; })) {
      m.seal = sha_(m.H + '|' + required_(m).map(function (r) { return m.sign[r].sh; }).join(','));
      m.sealedAt = new Date().toISOString();
      m.status = 'done';
      m.pdf = makePdf_(m).getUrl();
      save_(m); sendPdf_(m);
    } else save_(m);
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
  let h = '<html><body style="font-family:Arial;font-size:11px">';
  h += '<h2 style="letter-spacing:3px;color:#555">SIMOPS MEETING RECORD</h2>';
  h += '<p><b>Host:</b> ' + esc_(m.host) + ' &nbsp; <b>Date:</b> ' + esc_(m.date) + ' &nbsp; <b>Time:</b> ' + esc_(m.time) + '</p>';
  if (cl.length) h += '<p style="background:#fff3cd;padding:6px"><b>Possible clashes:</b><br>' + cl.map(esc_).join('<br>') + '</p>';
  h += '<table border="1" cellspacing="0" cellpadding="4" width="100%"><tr style="background:#eee"><th>Contractor</th><th>Area</th><th>Planned activity</th><th>Agreed constraint</th><th>Signed by</th></tr>';
  CODES.forEach(function (c) {
    const rows = (m.con[c] && m.con[c].rows) || []; if (!rows.length) return;
    const s = m.sign[c];
    const sig = s ? esc_(s.name) + '<br><img src="' + s.sign + '" height="40"><br>' + esc_(s.at) + '<br>' + esc_(s.method) : '';
    rows.forEach(function (r, i) {
      h += '<tr>' + (i === 0 ? '<td rowspan="' + rows.length + '"><b>' + NAMES[c] + '</b></td>' : '') +
        '<td>' + esc_(r.area) + '</td><td>' + esc_(r.activity) + (r.details ? '<br><i>' + esc_(r.details) + '</i>' : '') + '</td><td>' + esc_(r.constraint) + '</td>' +
        (i === 0 ? '<td rowspan="' + rows.length + '">' + sig + '</td>' : '') + '</tr>';
    });
  });
  h += '</table>';
  if (m.photos && m.photos.length) h += '<p>' + m.photos.map(function (p) { return '<img src="' + p + '" height="140"> '; }).join('') + '</p>';
  h += '<p><b>Host comments:</b> ' + esc_(m.hostc) + '</p>';
  ['ack', 'noted'].forEach(function (k) {
    const s = m.sign[k];
    h += '<p><b>' + NAMES[k] + ':</b> ' + (s ? esc_(s.name) + ' <img src="' + s.sign + '" height="40"> ' + esc_(s.at) + ' (' + esc_(s.method) + ')' : '') + '</p>';
  });
  h += '<hr><p style="font-size:9px;color:#555">Electronic record. Content hash (SHA-256): ' + m.H + '<br>Seal: ' + m.seal + '<br>Verify online: ' + esc_(m.base) + '?v=' + m.id + '</p>';
  h += '</body></html>';
  return Utilities.newBlob(h, 'text/html', 'SIMOPS.html').getAs('application/pdf').setName('SIMOPS_Meeting_Record_' + m.date + '.pdf');
}
function makePdf_(m) { return folder_().createFile(pdfBlob_(m)); }
function sendPdf_(m) {
  const list = [m.hostEmails, m.ackEmail, m.notedEmail];
  m.parts.forEach(function (c) { if (m.con[c] && m.con[c].email) list.push(m.con[c].email); });
  const to = list.join(',').split(/[,;\s]+/).filter(okEmail_).filter(function (x, i, a) { return a.indexOf(x) === i; }).join(',');
  if (!to) return;
  MailApp.sendEmail({
    to: to, subject: 'SIMOPS Meeting Record - ' + m.date + ' ' + (m.time || ''),
    body: 'All parties have signed. The SIMOPS Meeting Record is attached.\n\nHost: ' + m.host + '\nRecord: ' + m.base + '?m=' + m.id + '\nVerify authenticity: ' + m.base + '?v=' + m.id,
    attachments: [pdfBlob_(m)]
  });
}
