'use strict';
/* זיכוי חלקים – PWA
 * הנתונים יושבים בגיליון (Apps Script). בטלפון: עותק אחרון של הרשימה + תור פעולות שלא סונכרנו.
 * עבודה בלי קליטה: סורקים ומוסיפים, הכול נכנס לתור ומסונכרן כשחוזרת קליטה.
 * שליחה: רק אחרי סנכרון מלא, בשני שלבים (ספירה → אישור), והשרת בודק שהרשימה לא השתנתה.
 */

const VERSION = '1.0.0';
const LS = 'zikui.v1.';
const FIELDS = ['caseId', 'pn', 'rma'];
const LABELS = { caseId: 'Case ID', pn: 'Part Number', rma: 'RMA' };
// תבניות לפי דוגמאות אמיתיות: Case 5166539780 | PN N36076-001 | RMA 0619940586010200B
const PATTERNS = {
  caseId: /^\d{10}$/,
  pn: /^[A-Z0-9]{5,7}-[A-Z0-9]{3}$/i,
  rma: /^\d{14,18}[A-Z]$/i
};
const PREFIXES = [
  [/^(case\s*id|case)\s*[:=]\s*/i, 'caseId'],
  [/^(part\s*number|part|p\/n|pn)\s*[:=]\s*/i, 'pn'],
  [/^rma\s*[:=]\s*/i, 'rma']
];
// סורק שמתנהג כמקלדת כשהמקלדת בעברית: אותיות עבריות ← האות האנגלית שבאותו מקש
const HEB = { 'ש': 'A', 'נ': 'B', 'ב': 'C', 'ג': 'D', 'ק': 'E', 'כ': 'F', 'ע': 'G', 'י': 'H', 'ן': 'I', 'ח': 'J', 'ל': 'K', 'ך': 'L', 'צ': 'M', 'מ': 'N', 'ם': 'O', 'פ': 'P', '/': 'Q', 'ר': 'R', 'ד': 'S', 'א': 'T', 'ו': 'U', 'ה': 'V', "'": 'W', 'ס': 'X', 'ט': 'Y', 'ז': 'Z' };

const $ = s => document.querySelector(s);
const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };

/* ---------- אחסון מקומי ---------- */

function load(k, d) { try { const v = localStorage.getItem(LS + k); return v == null ? d : JSON.parse(v); } catch (_) { return d; } }
function save(k, v) { try { localStorage.setItem(LS + k, JSON.stringify(v)); } catch (_) { /* מלא/חסום */ } }

const S = {
  cfg: load('cfg', { url: '', key: '' }),
  server: load('server', null),        // { mode, to, test, nextShipNo, items[] }
  queue: load('queue', []),            // [{ type:'add'|'update'|'delete', ... }]
  net: 'idle',                         // idle | syncing | ok | offline | error | auth
  netErr: '',
  invalid: new Set()
};
function persist() { save('server', S.server); save('queue', S.queue); }
const configured = () => !!(S.cfg.url && (S.cfg.key || S.cfg.url === 'demo'));

function uid() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return 'x' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
}

/** הרשימה כפי שהיא בטלפון: מה שהגיע מהשרת + פעולות שעוד לא סונכרנו */
function viewItems() {
  const items = (S.server ? S.server.items : []).map(i => Object.assign({}, i));
  for (const op of S.queue) {
    if (op.type === 'add') {
      if (!items.some(i => i.id === op.item.id)) items.push(Object.assign({ pending: true }, op.item));
    } else if (op.type === 'update') {
      const i = items.find(x => x.id === op.id);
      if (i) Object.assign(i, op.item, { pending: true });
    } else if (op.type === 'delete') {
      const k = items.findIndex(x => x.id === op.id);
      if (k >= 0) items.splice(k, 1);
    }
  }
  return items;
}

const dupKey = i => [i.caseId, i.pn, i.rma].map(s => String(s).trim()).join('|');
function dupIds(items) {
  const seen = new Set(), d = new Set();
  items.forEach(i => { const k = dupKey(i); if (seen.has(k)) d.add(i.id); else seen.add(k); });
  return d;
}
function counts(items) {
  return {
    total: items.length,
    defective: items.filter(i => i.status === 'Defective').length,
    unused: items.filter(i => i.status === 'Unused').length
  };
}
// ⁨…⁩ = בידוד כיווניות, כדי שהחלק האנגלי עם הסוגריים לא יתהפך בתוך טקסט עברי
const summaryText = c => `${c.total} פריטים ⁨(Defective: ${c.defective}, Unused: ${c.unused})⁩`;

/* ---------- רשת ---------- */

class NetError extends Error {}

async function api(action, data) {
  if (S.cfg.url === 'demo') return Demo.handle(action, data || {});
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), action === 'confirm' ? 90000 : 45000);
  let res;
  try {
    res = await fetch(S.cfg.url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },  // בלי preflight – Apps Script לא תומך ב-OPTIONS
      body: JSON.stringify(Object.assign({ action, key: S.cfg.key }, data || {})),
      redirect: 'follow',
      signal: ctrl.signal
    });
  } catch (e) {
    throw new NetError(e.name === 'AbortError' ? 'אין תשובה מהשרת (זמן קצוב)' : 'אין חיבור לשרת');
  } finally { clearTimeout(t); }
  let j;
  try { j = await res.json(); } catch (_) { throw new Error('תשובה לא תקינה מהשרת – בדוק את כתובת ה-Web App'); }
  if (!j.ok) { const e = new Error(j.error || 'שגיאה'); e.auth = j.auth === false; throw e; }
  return j;
}

let syncing = null, syncAgain = false, syncTimer = null;

function scheduleSync(ms) {
  clearTimeout(syncTimer);
  syncTimer = setTimeout(sync, ms == null ? 400 : ms);
}

/** שולח את התור (בסדר) ומקבל את המצב העדכני מהגיליון */
function sync() {
  if (!configured()) { S.net = 'idle'; renderNet(); return Promise.resolve(false); }
  if (syncing) { syncAgain = true; return syncing; }
  syncing = (async () => {
    S.net = 'syncing'; renderNet();
    try {
      do {
        syncAgain = false;
        const ops = S.queue.slice();
        const res = await api('sync', { ops });
        S.queue = S.queue.slice(ops.length);   // מה שנוסף בזמן הבקשה נשאר בתור
        S.server = res.state;
        persist();
      } while (syncAgain);
      S.net = 'ok'; S.netErr = '';
      return true;
    } catch (e) {
      S.net = e instanceof NetError ? 'offline' : (e.auth ? 'auth' : 'error');
      S.netErr = e.message;
      return false;
    } finally {
      syncing = null;
      render();
    }
  })();
  return syncing;
}

function enqueue(op) {
  S.queue.push(op);
  persist();
  render();
  scheduleSync();
}

/* ---------- סריקה ---------- */

const inputs = {};
let awaitingStatus = false;

function cleanScan(raw) {
  let v = String(raw == null ? '' : raw)
    .replace(/[\u0000-\u001F\u007F​-‏‪-‮⁦-⁩﻿]/g, '')
    .trim();
  let note = '';
  if (/[א-ת]/.test(v)) {
    v = [...v].map(c => HEB[c] || c).join('').toUpperCase();
    note = 'המקלדת בעברית – התווים הומרו לאנגלית. כדאי להחליף שפה במקלדת.';
  }
  let hinted = null;
  for (const [re, f] of PREFIXES) if (re.test(v)) { v = v.replace(re, '').trim(); hinted = f; break; }
  return { v, note, hinted };
}

const detect = v => FIELDS.find(f => PATTERNS[f].test(v)) || null;

function setHint(text, kind) {
  const h = $('#scanHint');
  h.textContent = text || '';
  h.className = 'hint' + (kind ? ' ' + kind : '');
}

function markField(f) {
  const v = inputs[f].value.trim();
  const box = inputs[f].closest('.field');
  box.classList.toggle('filled', !!v && PATTERNS[f].test(v));
  box.classList.toggle('warn', !!v && !PATTERNS[f].test(v));
}

/**
 * סריקה (או הקלדה + Enter) נכנסת לשדה שהתבנית שלה מתאימה לו; אם אין התאמה – לשדה שבו היא נסרקה.
 * אחר כך הסמן עובר לשדה הריק הבא, וכשהשלושה מלאים – בחירת מצב.
 */
function acceptScan(focused, raw, prevValue) {
  const { v, note, hinted } = cleanScan(raw);
  inputs[focused].value = prevValue || '';
  if (!v) { markField(focused); return; }
  const detected = hinted || detect(v);
  const target = detected || focused;
  const cur = inputs[target].value.trim();
  if (cur && cur !== v && target !== focused) {
    setHint(`נסרק ${LABELS[target]} (${v}), אבל כבר יש ${LABELS[target]}. בחר מצב לפריט הנוכחי, או נקה שדות.`, 'warn');
    buzz([60, 60, 60]);
    markField(focused);
    focusNext(focused);
    return;
  }
  inputs[target].value = v;
  markField(target); markField(focused);
  const msgs = [];
  if (note) msgs.push(note);
  if (detected && detected !== focused) msgs.push(`זוהה ${LABELS[detected]} – הוכנס לשדה הנכון.`);
  if (!detected) msgs.push(`לא נראה כמו ${LABELS[target]} רגיל – בדוק שזה הברקוד הנכון.`);
  setHint(msgs.join(' '), (note || !detected) ? 'warn' : (msgs.length ? 'ok' : ''));
  focusNext(target);
}

function focusNext(from) {
  const order = FIELDS.slice(FIELDS.indexOf(from) + 1).concat(FIELDS.slice(0, FIELDS.indexOf(from) + 1));
  const next = order.find(f => !inputs[f].value.trim());
  if (next) { setAwait(false); inputs[next].focus(); }
  else setAwait(true);
}

function setAwait(on) {
  awaitingStatus = on;
  $('#statusPick').classList.toggle('await', on);
  if (on) {
    if (document.activeElement && document.activeElement.blur) document.activeElement.blur();
    if (!$('#scanHint').textContent) setHint('בחר מצב – מקש 1 ל-⁨Defective⁩, מקש 2 ל-⁨Unused⁩');
  }
}

function clearScan(keepHint) {
  FIELDS.forEach(f => { inputs[f].value = ''; markField(f); });
  if (!keepHint) setHint('');
  setAwait(false);
  inputs.caseId.focus();
}

function buzz(p) { try { navigator.vibrate && navigator.vibrate(p); } catch (_) { /* */ } }

function addItem(status) {
  const vals = {};
  for (const f of FIELDS) {
    vals[f] = cleanScan(inputs[f].value).v;
    if (!vals[f]) { toast(`חסר ${LABELS[f]}`, 'warn'); setAwait(false); inputs[f].focus(); return; }
  }
  const item = Object.assign({ id: 'a-' + uid(), status, at: new Date().toISOString() }, vals);
  const items = viewItems();
  const twin = items.findIndex(i => dupKey(i) === dupKey(item));
  enqueue({ type: 'add', item });
  const n = items.length + 1;
  buzz(40);
  if (twin >= 0) toast(`⚠ נוסף #${n} – זהה לפריט #${twin + 1} שכבר ברשימה`, 'warn', 4000);
  else toast(`נוסף #${n} · ${status}`, 'ok', 1500);
  clearScan();
}

function bindScanInputs() {
  FIELDS.forEach(f => {
    const inp = inputs[f] = $('#f-' + f);
    let before = '';
    inp.addEventListener('focus', () => { before = inp.value; setAwait(false); });
    inp.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        const raw = inp.value;
        const prev = before && raw.startsWith(before) && raw !== before ? before : '';
        // סריקה לתוך שדה שכבר היה מלא מצטרפת לתוכן הקיים – מפרידים את החלק החדש
        acceptScan(f, prev ? raw.slice(prev.length) : raw, prev);
        before = inputs[f].value;
      }
    });
    inp.addEventListener('blur', () => { const c = cleanScan(inp.value).v; if (c !== inp.value) inp.value = c; markField(f); });
  });

  document.querySelectorAll('.st').forEach(b => b.addEventListener('click', () => addItem(b.dataset.status)));
  $('#clearScan').addEventListener('click', () => clearScan());

  // בחירת מצב מהמקלדת/סורק: 1 / 2. סורק מקליד תווים ברצף מהיר – אז מחכים רגע לוודא שזו לחיצה בודדת.
  let keyTimer = null, burst = 0, burstWarned = false;
  document.addEventListener('keydown', e => {
    if (!awaitingStatus || document.querySelector('dialog[open]')) return;
    if (e.target && e.target.tagName === 'INPUT') return;
    if (e.key === 'Enter') {
      e.preventDefault();
      if (burst > 1) { setHint('נסרק ברקוד לפני שנבחר מצב – בחר מצב לפריט הנוכחי ואז סרוק שוב.', 'warn'); buzz([60, 60, 60]); }
      burst = 0; burstWarned = false;
      return;
    }
    if (e.key.length !== 1) return;
    e.preventDefault();
    burst++;
    if (keyTimer) { clearTimeout(keyTimer); keyTimer = null; }
    if (burst === 1 && /^[12duDU]$/.test(e.key)) {
      const k = e.key.toLowerCase();
      keyTimer = setTimeout(() => {
        keyTimer = null;
        if (burst === 1) addItem(k === '1' || k === 'd' ? 'Defective' : 'Unused');
        burst = 0;
      }, 150);
    } else if (!burstWarned && burst > 1) {
      burstWarned = true;
    }
  });

  document.querySelectorAll('.cam').forEach(b => b.addEventListener('click', () => openCamera(b.dataset.cam)));
}

/* ---------- מצלמה ---------- */

async function openCamera(field) {
  const dlg = $('#camDlg');
  $('#camTitle').textContent = 'סריקה במצלמה – ' + LABELS[field];
  $('#camMsg').textContent = 'מכוון את המצלמה לברקוד…';
  dlg.showModal();
  try {
    const value = await Camera.scan($('#camVideo'), PATTERNS[field], m => { $('#camMsg').textContent = m; });
    dlg.close();
    if (value != null) { buzz(40); acceptScan(field, value, inputs[field].value); }
  } catch (e) {
    $('#camMsg').textContent = e.message || 'המצלמה לא זמינה';
  }
}

/* ---------- תצוגה ---------- */

function render() {
  const items = viewItems();
  const d = dupIds(items);
  const c = counts(items);
  $('#summary').textContent = items.length ? `ברשימה ${summaryText(c)}` : 'הרשימה ריקה';
  const dn = $('#dupNote');
  dn.hidden = !d.size;
  dn.textContent = `⚠ ${d.size} זהים`;

  const ol = $('#items');
  ol.textContent = '';
  for (let k = items.length - 1; k >= 0; k--) {   // החדש למעלה, המספור כמו ב-PDF
    const i = items[k];
    const li = el('li', 'item' + (d.has(i.id) ? ' dup' : '') + (S.invalid.has(i.id) ? ' invalid' : ''));
    li.appendChild(el('span', 'no', String(k + 1)));
    const codes = el('div', 'codes');
    [['Case', i.caseId], ['PN', i.pn], ['RMA', i.rma]].forEach(([k2, v]) => {
      const row = el('div');
      row.appendChild(el('span', 'k', k2));
      row.appendChild(document.createTextNode(v || '—'));
      codes.appendChild(row);
    });
    li.appendChild(codes);
    const meta = el('div', 'meta');
    meta.appendChild(el('span', 'pill ' + (i.status || ''), i.status || 'חסר מצב'));
    if (i.pending) meta.appendChild(el('span', 'pending', '⏳ לא סונכרן'));
    if (d.has(i.id)) meta.appendChild(el('span', 'pending', 'זהה לפריט קודם'));
    li.appendChild(meta);
    li.addEventListener('click', () => openEdit(i, k + 1));
    ol.appendChild(li);
  }
  $('#emptyList').hidden = items.length > 0;

  const chip = $('#modeChip');
  if (S.server) {
    chip.textContent = S.server.mode;
    chip.className = 'chip ' + (S.server.test ? 'test' : 'prod');
  } else {
    chip.textContent = '—';
    chip.className = 'chip';
  }
  $('#setupBanner').hidden = configured();
  renderNet();
}

function renderNet() {
  const s = $('#syncStatus');
  const q = S.queue.length;
  const map = {
    idle: configured() ? '' : 'לא מחובר',
    syncing: '⟳ מסנכרן…',
    ok: q ? `⏳ ${q} ממתינים` : '✓ מסונכרן',
    offline: q ? `⏳ ${q} ממתינים · אין קליטה` : '✕ אין קליטה',
    error: '⚠ שגיאה',
    auth: '⚠ מפתח שגוי'
  };
  s.textContent = map[S.net] || '';
  s.title = S.netErr || '';
}

let toastTimer = null;
function toast(text, kind, ms) {
  const t = $('#toast');
  t.textContent = text;
  t.className = 'toast' + (kind ? ' ' + kind : '');
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, ms || 2500);
}

/* ---------- עריכה ---------- */

let editing = null;
function openEdit(item, no) {
  editing = item;
  $('#editNo').textContent = '#' + no;
  FIELDS.forEach(f => { $('#e-' + f).value = item[f] || ''; });
  document.querySelectorAll('input[name=e-status]').forEach(r => { r.checked = r.value === item.status; });
  $('#editErr').textContent = '';
  $('#editDlg').showModal();
}

function bindEdit() {
  $('#editSave').addEventListener('click', () => {
    const item = {};
    for (const f of FIELDS) {
      item[f] = cleanScan($('#e-' + f).value).v;
      if (!item[f]) { $('#editErr').textContent = `חסר ${LABELS[f]}`; return; }
    }
    const st = document.querySelector('input[name=e-status]:checked');
    if (!st) { $('#editErr').textContent = 'בחר מצב'; return; }
    item.status = st.value;
    S.invalid.delete(editing.id);
    enqueue({ type: 'update', id: editing.id, item });
    $('#editDlg').close();
    toast('נשמר', 'ok', 1200);
  });
  $('#editDelete').addEventListener('click', () => {
    if (!confirm('למחוק את הפריט מהרשימה?')) return;
    enqueue({ type: 'delete', id: editing.id });
    $('#editDlg').close();
    toast('נמחק', 'ok', 1200);
  });
}

/* ---------- שליחה דו-שלבית ---------- */

const send = { arm: null, expiresAt: 0, timer: null, busy: false, pdf: null };

function sendView(title, body, actions) {
  $('#sendTitle').textContent = title;
  const b = $('#sendBody');
  b.textContent = '';
  (Array.isArray(body) ? body : [body]).forEach(n => n && b.appendChild(typeof n === 'string' ? el('p', '', n) : n));
  const a = $('#sendActions');
  a.textContent = '';
  a.appendChild(el('span', 'spacer'));
  (actions || []).forEach(([label, cls, fn]) => {
    const btn = el('button', 'btn ' + cls, label);
    btn.type = 'button';
    btn.addEventListener('click', fn);
    a.appendChild(btn);
  });
}

function closeSend(disarm) {
  clearInterval(send.timer);
  if (disarm && send.arm) api('disarm').catch(() => {});
  send.arm = null;
  $('#sendDlg').close();
}

async function startSend() {
  if (!configured()) { openSettings(); return; }
  const dlg = $('#sendDlg');
  if (!dlg.open) dlg.showModal();
  sendView('שליחת משלוח', 'מסנכרן עם הגיליון…', [['ביטול', 'ghost', () => closeSend(false)]]);
  const ok = await sync();
  if (!ok || S.queue.length) {
    sendView('אי אפשר לשלוח עכשיו',
      [S.net === 'offline' ? 'אין קליטה. הפריטים שמורים בטלפון וייכנסו לגיליון כשתחזור קליטה.' : (S.netErr || 'הסנכרון לא הושלם.'),
        'שליחה אפשרית רק אחרי שכל הפריטים נמצאים בגיליון.'],
      [['סגור', 'ghost', () => closeSend(false)], ['נסה שוב', 'primary', startSend]]);
    return;
  }
  sendView('שליחת משלוח', 'סופר את הרשימה בגיליון…', []);
  try {
    const res = await api('arm');
    if (res.state) { S.server = res.state; persist(); }
    if (!res.arm) {
      S.invalid = new Set(res.invalidIds || []);
      render();
      sendView('אי אפשר לשלוח', [res.error || 'הרשימה לא תקינה', res.invalidIds ? 'הפריטים החסרים מסומנים באדום ברשימה – הקש עליהם להשלמה.' : ''],
        [['סגור', 'primary', () => closeSend(false)]]);
      return;
    }
    S.invalid = new Set();
    render();
    showArm(res.arm, false);
  } catch (e) {
    showSendError(e);
  }
}

function showArm(arm, changed) {
  send.arm = arm;
  send.expiresAt = Date.now() + arm.ttlMs;
  const nodes = [];
  if (changed) nodes.push(el('p', 'hint warn', '🔄 הרשימה השתנתה מאז הספירה. ספור שוב.'));
  const mode = el('div', 'send-mode ' + (arm.test ? 'test' : 'prod'), `מצב: ${arm.mode}${arm.test ? '' : ' – משלוח אמיתי'}`);
  mode.appendChild(el('span', 'to', arm.to));
  nodes.push(mode);
  nodes.push(el('p', 'send-count', `${arm.counts.total} פריטים`));
  nodes.push(el('p', 'send-split', `Defective: ${arm.counts.defective} · Unused: ${arm.counts.unused}`));
  nodes.push(el('p', 'muted', `משלוח מס' ${arm.shipNo}`));
  if (arm.dupIds && arm.dupIds.length) {
    const items = viewItems();
    const nos = arm.dupIds.map(id => items.findIndex(i => i.id === id) + 1).filter(n => n > 0);
    nodes.push(el('p', 'hint warn', `⚠ פריטים זהים לפריט קודם: ${nos.map(n => '#' + n).join(', ')}. לא חוסם – רק לבדוק.`));
  }
  nodes.push(el('p', 'send-ask', 'ספור את הכלוב. אם הכמות תואמת – אשר.'));
  const timer = el('p', 'timer');
  nodes.push(timer);
  sendView('אישור ספירה', nodes, [
    ['ביטול', 'ghost', () => closeSend(true)],
    ['הכלוב תואם – שלח', 'primary', confirmSend]
  ]);
  clearInterval(send.timer);
  const tick = () => {
    const left = send.expiresAt - Date.now();
    if (left <= 0) {
      clearInterval(send.timer);
      sendView('⏱ פג הזמן', 'עברו 15 דקות מהספירה. צריך לספור מחדש.',
        [['סגור', 'ghost', () => closeSend(true)], ['ספור מחדש', 'primary', startSend]]);
      return;
    }
    const m = Math.floor(left / 60000), s = Math.floor(left / 1000) % 60;
    timer.textContent = `האישור בתוקף עוד ${m}:${String(s).padStart(2, '0')}`;
  };
  tick();
  send.timer = setInterval(tick, 1000);
}

async function confirmSend() {
  if (send.busy || !send.arm) return;
  send.busy = true;
  clearInterval(send.timer);
  sendView('שולח…', 'יוצר PDF ושולח מייל. לא לסגור את האפליקציה.', []);
  try {
    const res = await api('confirm', { fp: send.arm.fp });
    if (res.state) { S.server = res.state; persist(); render(); }
    if (res.status === 'sent') {
      send.arm = null;
      send.pdf = res.pdf ? { b64: res.pdf, name: res.pdfName } : null;
      const nodes = [el('div', 'big-ok', '✅'),
        el('p', 'send-count', `${res.mode === 'בדיקה' ? '[בדיקה] ' : ''}משלוח ${res.shipNo} נשלח`),
        el('p', '', summaryText(res.counts)),
        el('p', 'muted', 'נשלח ל: ' + res.to),
        el('p', 'muted', 'הרשימה נוקתה ונשמרה בהיסטוריה.')];
      const acts = [['סגור', 'primary', () => closeSend(false)]];
      if (send.pdf) acts.unshift(['שתף PDF…', 'ghost', () => sharePdf(send.pdf)]);
      sendView('נשלח', nodes, acts);
      buzz([40, 60, 40]);
    } else if (res.status === 'changed') {
      showArm(res.arm || send.arm, true);
      if (!res.arm) showSendError(new Error(res.error || 'הרשימה השתנתה'));
    } else if (res.status === 'expired') {
      sendView('⏱ פג הזמן', 'עברו 15 דקות מהספירה. צריך לספור מחדש.',
        [['סגור', 'ghost', () => closeSend(false)], ['ספור מחדש', 'primary', startSend]]);
    } else {
      sendView('השליחה בוטלה', 'הספירה בוטלה (למשל בגלל החלפת מצב עבודה). צריך לספור מחדש.',
        [['סגור', 'ghost', () => closeSend(false)], ['ספור מחדש', 'primary', startSend]]);
    }
  } catch (e) {
    send.arm = null;
    showSendError(e, true);
  } finally {
    send.busy = false;
  }
}

function showSendError(e, duringSend) {
  const msg = [e.message || 'שגיאה'];
  if (duringSend && e instanceof NetError) {
    msg.push('לא ברור אם המייל יצא. לפני שליחה חוזרת – בדוק בהיסטוריה (או בתיבת "נשלחו" ב-Gmail).');
  } else if (duringSend) {
    msg.push('לא נמחק כלום מהרשימה.');
  }
  sendView('השליחה נכשלה', msg, [['סגור', 'ghost', () => closeSend(false)], ['התחל מחדש', 'primary', startSend]]);
}

/* ---------- PDF ---------- */

function pdfBlob(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type: 'application/pdf' });
}

async function sharePdf(p) {
  const file = new File([pdfBlob(p.b64)], p.name, { type: 'application/pdf' });
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: p.name }); } catch (_) { /* בוטל */ }
  } else {
    showPdf(p, 'PDF');
  }
}

let pdfUrl = null;
function showPdf(p, title) {
  if (pdfUrl) URL.revokeObjectURL(pdfUrl);
  pdfUrl = URL.createObjectURL(pdfBlob(p.b64));
  $('#pdfTitle').textContent = title;
  $('#pdfName').textContent = p.name;
  const a = $('#pdfOpen');
  a.href = pdfUrl;
  a.download = p.name;
  $('#pdfShare').hidden = !navigator.canShare;
  $('#pdfShare').onclick = () => sharePdf(p);
  const d = $('#pdfDlg');
  if (!d.open) d.showModal();
}

async function preview() {
  if (!configured()) { openSettings(); return; }
  toast('מסנכרן ויוצר תצוגה מקדימה…', '', 8000);
  const ok = await sync();
  if (!ok) { toast(S.net === 'offline' ? 'אין קליטה – תצוגה מקדימה דורשת חיבור' : S.netErr, 'err', 3500); return; }
  if (S.queue.length) { toast('יש פריטים שלא סונכרנו עדיין', 'warn'); return; }
  try {
    const res = await api('preview');
    $('#toast').hidden = true;
    showPdf({ b64: res.pdf, name: res.pdfName }, 'תצוגה מקדימה (לא נשלח)');
  } catch (e) {
    toast(e.message, 'err', 3500);
  }
}

/* ---------- היסטוריה ---------- */

async function openHistory() {
  if (!configured()) { openSettings(); return; }
  const body = $('#histBody');
  body.textContent = 'טוען…';
  $('#histMode').textContent = '';
  $('#historyDlg').showModal();
  try {
    const res = await api('history');
    $('#histMode').textContent = '(' + res.mode + ')';
    body.textContent = '';
    if (!res.shipments.length) { body.textContent = 'עוד לא נשלחו משלוחים במצב הזה.'; return; }
    const t = el('table', 'hist');
    const hr = el('tr');
    ['#', 'תאריך', 'סה"כ', 'Defective', 'Unused'].forEach(h => hr.appendChild(el('th', '', h)));
    t.appendChild(hr);
    res.shipments.forEach(s => {
      const tr = el('tr');
      const d = new Date(s.date);
      [s.shipNo, isNaN(d) ? s.date : d.toLocaleDateString('he-IL'), s.total, s.defective, s.unused]
        .forEach(v => tr.appendChild(el('td', '', String(v))));
      t.appendChild(tr);
    });
    body.appendChild(t);
  } catch (e) {
    body.textContent = e instanceof NetError ? 'אין קליטה – היסטוריה דורשת חיבור.' : e.message;
  }
}

/* ---------- מצב עבודה ---------- */

function openMode() {
  if (!configured()) { openSettings(); return; }
  if (!S.server) { toast('עוד לא התקבל מצב מהגיליון', 'warn'); sync(); return; }
  $('#modeErr').textContent = '';
  $('#toTest').textContent = S.server.test ? S.server.to : '';
  document.querySelectorAll('[data-mode]').forEach(b => b.classList.toggle('current', b.dataset.mode === S.server.mode));
  $('#modeDlg').showModal();
}

function bindMode() {
  document.querySelectorAll('[data-mode]').forEach(b => b.addEventListener('click', async () => {
    const mode = b.dataset.mode;
    if (S.server && mode === S.server.mode) { $('#modeDlg').close(); return; }
    if (mode === 'אמיתי' && !confirm('לעבור למצב אמיתי?\nהמשלוח הבא יישלח למחסן ולטלדור.')) return;
    $('#modeErr').textContent = '';
    try {
      const res = await api('setMode', { mode });
      S.server = res.state; persist(); render();
      $('#modeDlg').close();
      toast('מצב עבודה: ' + mode, mode === 'אמיתי' ? 'ok' : '', 2500);
    } catch (e) {
      $('#modeErr').textContent = e instanceof NetError ? 'אין קליטה – החלפת מצב דורשת חיבור.' : e.message;
    }
  }));
}

/* ---------- הגדרות ---------- */

function openSettings() {
  $('#s-url').value = S.cfg.url || '';
  $('#s-key').value = S.cfg.key || '';
  $('#settingsMsg').textContent = '';
  $('#versionLine').textContent = 'גרסה ' + VERSION;
  $('#demoOfflineRow').hidden = S.cfg.url !== 'demo';
  $('#demoOffline').checked = Demo.offline;
  $('#settingsDlg').showModal();
}

function readSettings() {
  return { url: $('#s-url').value.trim(), key: $('#s-key').value.trim() };
}

function bindSettings() {
  $('#settingsBtn').addEventListener('click', openSettings);
  $('#s-url').addEventListener('input', () => { $('#demoOfflineRow').hidden = $('#s-url').value.trim() !== 'demo'; });
  $('#demoOffline').addEventListener('change', e => { Demo.setOffline(e.target.checked); if (!e.target.checked) sync(); else { S.net = 'offline'; renderNet(); } });
  $('#openSettingsFromBanner').addEventListener('click', openSettings);
  $('#settingsTest').addEventListener('click', async () => {
    const msg = $('#settingsMsg');
    const old = S.cfg;
    S.cfg = readSettings();
    msg.style.color = '';
    msg.textContent = 'בודק…';
    try {
      const res = await api('state');
      msg.style.color = 'var(--ok)';
      msg.textContent = `✓ מחובר. מצב: ${res.state.mode}, ${res.state.items.length} פריטים בגיליון.`;
    } catch (e) {
      msg.textContent = '✕ ' + e.message;
    } finally {
      S.cfg = old;
    }
  });
  $('#settingsSave').addEventListener('click', () => {
    const c = readSettings();
    if (c.url && c.url !== 'demo' && !/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(c.url)) {
      $('#settingsMsg').textContent = 'הכתובת צריכה להיות בצורה https://script.google.com/macros/s/…/exec';
      return;
    }
    const changedTarget = c.url !== S.cfg.url;
    if (changedTarget && S.queue.length && !confirm(`יש ${S.queue.length} פעולות שלא סונכרנו לגיליון הנוכחי. הן יישלחו לגיליון החדש. להמשיך?`)) return;
    S.cfg = c;
    save('cfg', c);
    if (changedTarget) { S.server = null; persist(); }
    $('#settingsDlg').close();
    render();
    sync();
  });
}

/* ---------- הפעלה ---------- */

function init() {
  bindScanInputs();
  bindEdit();
  bindMode();
  bindSettings();
  $('#modeChip').addEventListener('click', openMode);
  $('#sendBtn').addEventListener('click', startSend);
  $('#previewBtn').addEventListener('click', preview);
  $('#historyBtn').addEventListener('click', openHistory);
  $('#camClose').addEventListener('click', () => { Camera.stop(); $('#camDlg').close(); });
  $('#camDlg').addEventListener('close', () => Camera.stop());
  // אי אפשר לסגור את חלון השליחה באמצע שליחה
  $('#sendDlg').addEventListener('cancel', e => { if (send.busy) e.preventDefault(); else { clearInterval(send.timer); if (send.arm) api('disarm').catch(() => {}); send.arm = null; } });

  window.addEventListener('online', () => sync());
  window.addEventListener('offline', () => { S.net = 'offline'; renderNet(); });
  document.addEventListener('visibilitychange', () => { if (!document.hidden) sync(); });
  setInterval(() => { if (!document.hidden && (S.queue.length || S.net !== 'syncing')) sync(); }, 60000);

  render();
  if (configured()) sync();
  inputs.caseId.focus();

  if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1')) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
}

document.addEventListener('DOMContentLoaded', init);
