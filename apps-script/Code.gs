/**
 * זיכוי חלקים – משלוח למחסן
 * גיליון עבודה: A Case ID | B Part Number | C RMA | D מצב (Defective/Unused) | E תאריך | F מזהה (מוסתר, לאפליקציה)
 *
 * מצב עבודה – תא H3 (רשימה נפתחת): "בדיקה" / "אמיתי"
 *   בדיקה: מייל ל-TEST_RECIPIENTS, היסטוריה/מספור/תיקייה נפרדים. התהליך המלא (כולל ניקוי) – כמו אמיתי.
 *   אמיתי: מייל ל-PROD_RECIPIENTS, היסטוריה/מספור/תיקייה נקיים, בלי שאריות מהבדיקות.
 *
 * שליחה בשני שלבים (עובד גם מהטלפון):
 *   1. V ב-H1 "שלח משלוח" → ספירה + מצב + נמענים ב-G2
 *   2. V ב-H2 "אישור"     → PDF → מייל → היסטוריה → ניקוי
 *
 * אפליקציה (PWA): doPost – API מוגן במפתח (Script Properties → API_KEY, נוצר ב-setupApi).
 *   הגיליון והאפליקציה עובדים על אותם נתונים, אותה נעילה ואותם מונים.
 *
 * נמענים ושם טכנאי – בקובץ Config.gs (TEST_RECIPIENTS, PROD_RECIPIENTS, SENDER_NAME).
 */

const SHEET_NAME  = 'גיליון1';
const ARM_CELL = 'H1', CONFIRM_CELL = 'H2', MODE_CELL = 'H3', MSG_CELL = 'G2';
const ID_COL = 6;                       // F – מזהה פריט לאפליקציה
const ARM_TIMEOUT_MIN = 15;
const STATUS_OPTIONS = ['Defective', 'Unused'];
const MODE_OPTIONS = ['בדיקה', 'אמיתי'];
const AUTO_JUMP = false;
const API_KEY_PROP = 'API_KEY';         // Script Properties
const APP_ARM_PROP = 'appArm';          // Document Properties – שלב 1 שנפתח מהאפליקציה

// getters – הנמענים מוגדרים ב-Config.gs, נקראים רק בזמן ריצה
const MODES = {
  'בדיקה': { get to() { return TEST_RECIPIENTS; }, hist: 'היסטוריה - בדיקות', folder: 'זיכוי חלקים - בדיקות', key: 'shipNo_test', tag: '[בדיקה] ' },
  'אמיתי': { get to() { return PROD_RECIPIENTS; }, hist: 'היסטוריה',          folder: 'זיכוי חלקים - משלוחים', key: 'shipNo',      tag: '' }
};

/** להריץ פעם אחת ידנית (בטוח להריץ שוב – לא מוחק נתונים) */
function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  const n = sh.getMaxRows() - 1;

  // ניקוי שאריות מגרסאות קודמות (תיבות סימון/כותרות ישנות) – לא נוגע בנתונים ב-A:C
  if (sh.getRange('F1').getValue() === 'נשלח ב-' || sh.getRange('E1').getValue() === 'שלח') {
    sh.getRange(2, 1, n, 8).clearDataValidations();
    sh.getRange(2, 5, n, 2).clearContent();   // E:F – שאריות FALSE / "נשלח ב-"
    sh.getRange('F1').clearContent();
  }

  sh.getRange('A1:F1').setValues([['Case ID', 'Part Number', 'RMA', 'מצב', 'תאריך', 'מזהה']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, n, 3).setNumberFormat('@');
  sh.getRange(2, ID_COL, n, 1).setNumberFormat('@');
  sh.hideColumns(ID_COL);
  sh.getRange(2, 4, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).setAllowInvalid(false).build());

  const dupRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"",COUNTIFS($A$2:$A,$A2,$B$2:$B,$B2,$C$2:$C,$C2)>1)')
    .setBackground('#fce8b2').setRanges([sh.getRange(2, 1, n, 5)]).build();
  const testRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('בדיקה').setBackground('#cfe2f3').setRanges([sh.getRange(MODE_CELL)]).build();
  const prodRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('אמיתי').setBackground('#d9ead3').setRanges([sh.getRange(MODE_CELL)]).build();
  sh.setConditionalFormatRules([dupRule, testRule, prodRule]);

  sh.getRange('G1').setValue('שלח משלוח ←').setFontWeight('bold');
  sh.getRange('G3').setValue('מצב עבודה ←').setFontWeight('bold');
  sh.getRange(ARM_CELL).insertCheckboxes().uncheck();
  sh.getRange(CONFIRM_CELL).insertCheckboxes().uncheck();
  const mode = sh.getRange(MODE_CELL);
  mode.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(MODE_OPTIONS, true).setAllowInvalid(false).build());
  if (!MODES[mode.getValue()]) mode.setValue('בדיקה');
  sh.getRange(MSG_CELL).setWrap(true);
  sh.setColumnWidth(7, 260);

  Object.values(MODES).forEach(m => getHist_(m));

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'handleEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
}

/** להריץ פעם אחת ידנית: setup + יצירת מפתח גישה לאפליקציה. המפתח מודפס ביומן (View → Logs). */
function setupApi() {
  setup();
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty(API_KEY_PROP);
  if (!k) {
    k = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
    p.setProperty(API_KEY_PROP, k);
  }
  Logger.log('מפתח גישה לאפליקציה: ' + k);
}

/** אם המפתח דלף – להריץ, ולהזין את החדש באפליקציה */
function rotateApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(API_KEY_PROP);
  setupApi();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('זיכוי חלקים')
    .addItem('שלח משלוח (עם אישור ספירה)', 'sendFromMenu')
    .addItem('תצוגה מקדימה של PDF (בלי שליחה)', 'previewPdf')
    .addToUi();
}

function handleEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;
  const a1 = e.range.getA1Notation();

  if (a1 === ARM_CELL) { if (e.value === 'TRUE') arm_(sh); else disarm_(sh); return; }
  if (a1 === CONFIRM_CELL) { if (e.value === 'TRUE') confirm_(sh); return; }
  if (a1 === MODE_CELL) { // החלפת מצב מבטלת שליחה פתוחה – גם בגיליון וגם באפליקציה
    sh.getRange(ARM_CELL).uncheck(); disarm_(sh);
    PropertiesService.getDocumentProperties().deleteProperty(APP_ARM_PROP);
    return;
  }

  const row = e.range.getRow(), col = e.range.getColumn();
  if (row < 2 || col > 5 || e.range.getNumRows() > 1) return;
  if (col <= 3 && typeof e.value === 'string') {
    const clean = clean_(e.value);
    if (clean !== e.value) e.range.setValue(clean);
  }
  if (col <= 4 && !sh.getRange(row, 5).getValue()) {
    sh.getRange(row, 5).setValue(new Date()).setNumberFormat('dd/MM/yyyy');
  }
  if (AUTO_JUMP && col <= 3) sh.getRange(row, col + 1).activate();
}

/** ניקוי קלט מסורק: תווי בקרה, סימוני כיווניות, תווים ברוחב אפס, רווחים בקצוות */
function clean_(s) {
  return String(s == null ? '' : s)
    .replace(/[\u0000-\u001F\u007F​-‏‪-‮⁦-⁩﻿]/g, '')
    .trim();
}

/* ---------- מצב ---------- */

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
}

function mode_(sh) {
  const name = String(sh.getRange(MODE_CELL).getValue());
  if (!MODES[name]) throw new Error('בחר מצב עבודה בתא H3');
  return Object.assign({ name }, MODES[name]);
}

function getHist_(m) {
  const ss = SpreadsheetApp.getActive();
  let h = ss.getSheetByName(m.hist);
  if (!h) {
    h = ss.insertSheet(m.hist);
    h.getRange('A1:G1').setValues([['משלוח #', 'תאריך שליחה', 'Case ID', 'Part Number', 'RMA', 'מצב', 'תאריך הזנה']])
      .setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

/* ---------- שליחה דו-שלבית (גיליון) ---------- */

function counts_(rows) {
  const c = s => rows.filter(o => o.v[3] === s).length;
  return { total: rows.length, defective: c('Defective'), unused: c('Unused') };
}

function summary_(rows) {
  const k = counts_(rows);
  return `${k.total} פריטים (Defective: ${k.defective}, Unused: ${k.unused})`;
}

function fingerprint_(rows, modeName) {
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5, modeName + JSON.stringify(rows.map(o => o.v.slice(0, 4)))));
}

function arm_(sh) {
  let m;
  try { m = mode_(sh); } catch (err) { sh.getRange(ARM_CELL).uncheck(); sh.getRange(MSG_CELL).setValue('❌ ' + err.message); return; }
  const rows = getRows_(sh);
  const errs = validate_(rows);
  if (!rows.length || errs.length) {
    sh.getRange(ARM_CELL).uncheck();
    sh.getRange(MSG_CELL).setValue(!rows.length ? '❌ אין חלקים ברשימה' : '❌ ' + errs.join(' | '));
    return;
  }
  const dups = duplicates_(rows);
  PropertiesService.getDocumentProperties().setProperties({ armFp: fingerprint_(rows, m.name), armAt: String(Date.now()) });
  sh.getRange(MSG_CELL).setValue(
    `מצב: ${m.name} → ${m.to}\n` +
    `ברשימה ${summary_(rows)}.` +
    (dups.length ? `\n⚠ שורות זהות: ${dups.join(', ')}` : '') +
    `\nספור את הכלוב – אם תואם, סמן V ב"אישור" ←`);
}

function disarm_(sh) {
  PropertiesService.getDocumentProperties().deleteProperty('armFp');
  sh.getRange(CONFIRM_CELL).uncheck();
  sh.getRange(MSG_CELL).setValue('');
}

function confirm_(sh) {
  const props = PropertiesService.getDocumentProperties();
  const armFp = props.getProperty('armFp');
  const armAt = Number(props.getProperty('armAt') || 0);

  if (!armFp || sh.getRange(ARM_CELL).getValue() !== true) {
    sh.getRange(CONFIRM_CELL).uncheck();
    sh.getRange(MSG_CELL).setValue('קודם סמן V ב"שלח משלוח"');
    return;
  }
  if (Date.now() - armAt > ARM_TIMEOUT_MIN * 60000) {
    sh.getRange(ARM_CELL).uncheck(); disarm_(sh);
    sh.getRange(MSG_CELL).setValue('⏱ פג הזמן – סמן שוב "שלח משלוח"');
    return;
  }
  const m = mode_(sh);
  const rows = getRows_(sh);
  if (fingerprint_(rows, m.name) !== armFp) {
    sh.getRange(CONFIRM_CELL).uncheck();
    arm_(sh);
    sh.getRange(MSG_CELL).setValue('🔄 הרשימה השתנתה. ' + sh.getRange(MSG_CELL).getValue());
    return;
  }
  try {
    const res = doSend_(rows, m);
    sh.getRange(MSG_CELL).setValue(`✅ ${m.tag}משלוח ${res.shipNo} נשלח ל-${m.to} – ${summary_(rows)}`);
  } catch (err) {
    sh.getRange(MSG_CELL).setValue('❌ השליחה נכשלה, לא נמחק כלום: ' + err.message);
  } finally {
    props.deleteProperty('armFp');
    sh.getRange(ARM_CELL).uncheck();
    sh.getRange(CONFIRM_CELL).uncheck();
  }
}

function sendFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const sh = sheet_();
  const m = mode_(sh);
  const rows = getRows_(sh);
  if (!rows.length) return ui.alert('אין חלקים ברשימה');
  const errs = validate_(rows);
  if (errs.length) return ui.alert('לא נשלח:\n' + errs.join('\n'));
  const dups = duplicates_(rows);
  const ans = ui.alert('אישור משלוח',
    `מצב: ${m.name}\nנמענים: ${m.to}\n\nברשימה ${summary_(rows)}.` +
    (dups.length ? `\n⚠ שורות זהות: ${dups.join(', ')}` : '') +
    '\n\nזה תואם למה שבכלוב? לשלוח?', ui.ButtonSet.YES_NO);
  if (ans !== ui.Button.YES) return;
  const res = doSend_(rows, m);
  sh.getRange(MSG_CELL).setValue(`✅ ${m.tag}משלוח ${res.shipNo} נשלח ל-${m.to} – ${summary_(rows)}`);
}

/* ---------- ליבה ---------- */

/** שורות עם תוכן ב-A:D. o.v = [Case, PN, RMA, מצב, תאריך, מזהה] */
function getRows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ID_COL).getValues()
    .map((v, i) => ({ row: i + 2, v }))
    .filter(o => o.v.slice(0, 4).some(x => String(x).trim()));
}

function validate_(rows) {
  const missing = rows.filter(o => o.v.slice(0, 4).some(x => !String(x).trim())).map(o => o.row);
  return missing.length ? [`חסרים נתונים (כולל מצב) בשורות: ${missing.join(', ')}`] : [];
}

/** שורות שהן חזרה על שורה קודמת זהה לגמרי (Case+PN+RMA). RMA לבד לא נחשב כפול. */
function duplicates_(rows) {
  return duplicateRows_(rows).map(o => o.row);
}

function duplicateRows_(rows) {
  const seen = {}, dups = [];
  rows.forEach(o => {
    const k = o.v.slice(0, 3).map(x => String(x).trim()).join('|');
    if (seen[k]) dups.push(o); else seen[k] = true;
  });
  return dups;
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) throw new Error('פעולה אחרת כבר רצה – נסה שוב');
  try { return fn(); } finally { lock.releaseLock(); }
}

function doSend_(rows, m) {
  return withLock_(() => doSendLocked_(rows, m));
}

/** סדר בטוח: מייל קודם; רק אם הצליח → ארכיון, היסטוריה, ניקוי, קידום מונה. לקרוא רק כשהנעילה מוחזקת. */
function doSendLocked_(rows, m) {
  const sh = sheet_();
  const props = PropertiesService.getDocumentProperties();
  const shipNo = Number(props.getProperty(m.key) || 0) + 1;
  const now = new Date();
  const pdf = buildPdf_(rows, shipNo, now, m);

  MailApp.sendEmail({
    to: m.to,
    subject: `${m.tag}זיכוי חלקים – משלוח ${shipNo} (${rows.length} פריטים)`,
    htmlBody: `<div dir="rtl" style="font-family:Arial">שלום,<br>מצורפת רשימת החלקים שנשלחו לזיכוי (${rows.length} פריטים).<br><br>תודה,<br>${SENDER_NAME}</div>`,
    attachments: [pdf]
  });

  getFolder_(m.folder).createFile(pdf);
  const hist = getHist_(m);
  hist.getRange(hist.getLastRow() + 1, 1, rows.length, 7)
    .setValues(rows.map(o => [shipNo, now, o.v[0], o.v[1], o.v[2], o.v[3], o.v[4]]));
  sh.getRange(2, 1, sh.getMaxRows() - 1, ID_COL).clearContent();
  props.setProperty(m.key, String(shipNo));
  return { shipNo, pdf };
}

function buildPdf_(rows, shipNo, date, m) {
  const tz = Session.getScriptTimeZone();
  const d = Utilities.formatDate(date, tz, 'dd/MM/yyyy');
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const c = s => rows.filter(o => o.v[3] === s).length;
  const html = `
  <html><head><meta charset="utf-8"><style>
    body{font-family:Arial,sans-serif;direction:rtl;font-size:12px}
    h2{margin:0 0 4px} .meta{margin-bottom:12px;color:#444} .test{color:#c00;font-weight:bold}
    table{border-collapse:collapse;width:100%} th,td{border:1px solid #888;padding:6px;text-align:center}
    th{background:#e8e8e8} td.ltr{direction:ltr}
  </style></head><body>
    ${m.tag ? '<div class="test">*** בדיקה – לא משלוח אמיתי ***</div>' : ''}
    <h2>רשימת חלקים לזיכוי</h2>
    <div class="meta">משלוח מס' ${shipNo} &nbsp;|&nbsp; תאריך: ${d} &nbsp;|&nbsp; טכנאי: ${SENDER_NAME}<br>
      סה"כ פריטים: ${rows.length} &nbsp;|&nbsp; Defective: ${c('Defective')} &nbsp;|&nbsp; Unused: ${c('Unused')}</div>
    <table><tr><th>#</th><th>Case ID</th><th>Part Number</th><th>RMA</th><th>מצב</th></tr>
    ${rows.map((o, i) => `<tr><td>${i + 1}</td><td class="ltr">${esc(o.v[0])}</td><td class="ltr">${esc(o.v[1])}</td><td class="ltr">${esc(o.v[2])}</td><td class="ltr">${esc(o.v[3])}</td></tr>`).join('')}
    </table></body></html>`;
  const name = `${m.tag}זיכוי חלקים - משלוח ${shipNo} - ${Utilities.formatDate(date, tz, 'yyyy-MM-dd')}.pdf`;
  return Utilities.newBlob(html, 'text/html', 'tmp.html').getAs('application/pdf').setName(name);
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function previewPdf() {
  const ss = SpreadsheetApp.getActive();
  const sh = sheet_();
  const rows = getRows_(sh);
  if (!rows.length) { ss.toast('אין חלקים ברשימה'); return; }
  const m = mode_(sh);
  const file = getFolder_(MODES['בדיקה'].folder).createFile(buildPdf_(rows, 'טיוטה', new Date(), m));
  const out = HtmlService.createHtmlOutput(
    `<div dir="rtl" style="font-family:Arial"><a href="${file.getUrl()}" target="_blank">פתח תצוגה מקדימה ▸</a></div>`
  ).setWidth(300).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(out, 'PDF');
}

/* ======================================================================
 *  API לאפליקציה (PWA)
 *  פריסה: Deploy → New deployment → Web app | Execute as: Me | Who has access: Anyone
 *  כל בקשה: POST, גוף JSON (Content-Type: text/plain) { action, key, ... }
 * ====================================================================== */

function doGet() {
  return json_({ ok: true, app: 'zikui' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (_) { return json_({ ok: false, error: 'בקשה לא תקינה' }); }
  const key = PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
  if (!key || req.key !== key) return json_({ ok: false, auth: false, error: 'מפתח גישה שגוי' });
  const fn = API_[req.action];
  if (!fn) return json_({ ok: false, error: 'פעולה לא מוכרת: ' + req.action });
  try {
    return json_(Object.assign({ ok: true }, fn(req)));
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

const API_ = {
  /** מצב נוכחי: מצב עבודה, נמענים, פריטים */
  state: () => withLock_(() => ({ state: state_(sheet_()) })),

  /** החלת פעולות מהטלפון (בסדר) + החזרת מצב. אידמפוטנטי: הוספה עם מזהה קיים מדולגת. */
  sync: req => withLock_(() => {
    const sh = sheet_();
    const results = (req.ops || []).map(op => applyOp_(sh, op));
    return { results, state: state_(sh) };
  }),

  /** שלב 1: ספירה. שומר טביעת אצבע של הרשימה+המצב */
  arm: () => withLock_(() => {
    const sh = sheet_();
    const m = mode_(sh);
    const rows = ensureIds_(sh, getRows_(sh));
    const res = armRows_(rows, m);
    return Object.assign(res, { state: state_(sh) });
  }),

  /** שלב 2: אישור. רק אם הרשימה והמצב לא השתנו ולא עברו 15 דקות */
  confirm: req => withLock_(() => {
    const sh = sheet_();
    const props = PropertiesService.getDocumentProperties();
    const a = JSON.parse(props.getProperty(APP_ARM_PROP) || 'null');
    if (!a || a.fp !== req.fp) return { status: 'notArmed', state: state_(sh) };
    if (Date.now() - a.at > ARM_TIMEOUT_MIN * 60000) {
      props.deleteProperty(APP_ARM_PROP);
      return { status: 'expired', state: state_(sh) };
    }
    const m = mode_(sh);
    const rows = ensureIds_(sh, getRows_(sh));
    if (fingerprint_(rows, m.name) !== a.fp) {
      return Object.assign({ status: 'changed' }, armRows_(rows, m), { state: state_(sh) });
    }
    props.deleteProperty(APP_ARM_PROP);
    let res;
    try {
      res = doSendLocked_(rows, m);
    } catch (err) {
      throw new Error('השליחה נכשלה, לא נמחק כלום: ' + err.message);
    }
    return {
      status: 'sent', shipNo: res.shipNo, mode: m.name, to: m.to, counts: counts_(rows),
      pdf: Utilities.base64Encode(res.pdf.getBytes()), pdfName: res.pdf.getName(),
      state: state_(sh)
    };
  }),

  /** ביטול שלב 1 פתוח */
  disarm: () => {
    PropertiesService.getDocumentProperties().deleteProperty(APP_ARM_PROP);
    return {};
  },

  /** החלפת מצב עבודה – מבטלת כל שליחה פתוחה (גיליון ואפליקציה) */
  setMode: req => withLock_(() => {
    if (!MODES[req.mode]) throw new Error('מצב לא מוכר');
    const sh = sheet_();
    sh.getRange(MODE_CELL).setValue(req.mode);
    sh.getRange(ARM_CELL).uncheck();
    disarm_(sh);
    PropertiesService.getDocumentProperties().deleteProperty(APP_ARM_PROP);
    return { state: state_(sh) };
  }),

  /** תצוגה מקדימה – PDF בלי שליחה ובלי שמירה */
  preview: () => {
    const sh = sheet_();
    const rows = getRows_(sh);
    if (!rows.length) throw new Error('אין חלקים ברשימה');
    const pdf = buildPdf_(rows, 'טיוטה', new Date(), mode_(sh));
    return { pdf: Utilities.base64Encode(pdf.getBytes()), pdfName: pdf.getName() };
  },

  /** משלוחים אחרונים במצב הנוכחי */
  history: () => {
    const sh = sheet_();
    const m = mode_(sh);
    const h = getHist_(m);
    const last = h.getLastRow();
    if (last < 2) return { mode: m.name, shipments: [] };
    const by = {};
    h.getRange(2, 1, last - 1, 6).getValues().forEach(v => {
      const no = v[0];
      if (no === '' || no == null) return;
      const s = by[no] || (by[no] = { shipNo: no, date: v[1] instanceof Date ? v[1].toISOString() : String(v[1]), total: 0, defective: 0, unused: 0 });
      s.total++;
      if (v[5] === 'Defective') s.defective++;
      if (v[5] === 'Unused') s.unused++;
    });
    const shipments = Object.keys(by).map(k => by[k]).sort((a, b) => Number(b.shipNo) - Number(a.shipNo)).slice(0, 30);
    return { mode: m.name, shipments };
  }
};

function armRows_(rows, m) {
  const props = PropertiesService.getDocumentProperties();
  if (!rows.length) { props.deleteProperty(APP_ARM_PROP); return { error: 'אין חלקים ברשימה' }; }
  const invalid = rows.filter(o => o.v.slice(0, 4).some(x => !String(x).trim()) || STATUS_OPTIONS.indexOf(o.v[3]) < 0);
  if (invalid.length) {
    props.deleteProperty(APP_ARM_PROP);
    return { error: `חסרים נתונים (כולל מצב) ב-${invalid.length} פריטים`, invalidIds: invalid.map(o => o.v[5]) };
  }
  const fp = fingerprint_(rows, m.name);
  props.setProperty(APP_ARM_PROP, JSON.stringify({ fp, at: Date.now() }));
  return {
    arm: {
      fp, mode: m.name, to: m.to, test: !!m.tag, counts: counts_(rows),
      dupIds: duplicateRows_(rows).map(o => o.v[5]),
      ttlMs: ARM_TIMEOUT_MIN * 60000,
      shipNo: Number(props.getProperty(m.key) || 0) + 1
    }
  };
}

/** מצמיד מזהה לכל שורה שאין לה (שורות שהוזנו ידנית בגיליון) */
function ensureIds_(sh, rows) {
  rows.forEach(o => {
    if (!String(o.v[5] || '').trim()) {
      o.v[5] = 's-' + Utilities.getUuid();
      sh.getRange(o.row, ID_COL).setValue(o.v[5]);
    }
  });
  return rows;
}

function state_(sh) {
  const m = mode_(sh);
  const rows = ensureIds_(sh, getRows_(sh));
  const props = PropertiesService.getDocumentProperties();
  return {
    mode: m.name, to: m.to, test: !!m.tag,
    nextShipNo: Number(props.getProperty(m.key) || 0) + 1,
    items: rows.map(o => ({
      id: String(o.v[5]), caseId: String(o.v[0]), pn: String(o.v[1]), rma: String(o.v[2]), status: String(o.v[3]),
      date: o.v[4] instanceof Date ? o.v[4].toISOString() : String(o.v[4] || '')
    }))
  };
}

/** שורת הנתונים האחרונה (לפי A:E) – לא מושפעת מתאי הבקרה ב-G:H */
function lastDataRow_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return 1;
  const v = sh.getRange(2, 1, last - 1, 5).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i].some(x => String(x).trim())) return i + 2;
  return 1;
}

function cleanItem_(it) {
  it = it || {};
  const status = clean_(it.status);
  return {
    id: clean_(it.id), caseId: clean_(it.caseId), pn: clean_(it.pn), rma: clean_(it.rma),
    status: STATUS_OPTIONS.indexOf(status) >= 0 ? status : ''
  };
}

function applyOp_(sh, op) {
  const rows = getRows_(sh);
  const find = id => rows.filter(o => String(o.v[5]) === id)[0];

  if (op.type === 'add') {
    const it = cleanItem_(op.item);
    if (!it.id) return 'bad';
    if (find(it.id)) return 'exists';
    const r = lastDataRow_(sh) + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    const at = new Date(op.item && op.item.at);
    sh.getRange(r, 1, 1, 3).setNumberFormat('@');
    sh.getRange(r, ID_COL).setNumberFormat('@');
    sh.getRange(r, 1, 1, ID_COL).setValues([[it.caseId, it.pn, it.rma, it.status, isNaN(at) ? new Date() : at, it.id]]);
    sh.getRange(r, 5).setNumberFormat('dd/MM/yyyy');
    return 'added';
  }
  if (op.type === 'update') {
    const o = find(clean_(op.id));
    if (!o) return 'missing';
    const it = cleanItem_(op.item);
    sh.getRange(o.row, 1, 1, 4).setValues([[it.caseId, it.pn, it.rma, it.status]]);
    return 'updated';
  }
  if (op.type === 'delete') {
    const o = find(clean_(op.id));
    if (!o) return 'missing';
    // מוחק רק את A:F ומזיז למעלה – לא נוגע בתאי הבקרה ב-G:H
    sh.getRange(o.row, 1, 1, ID_COL).deleteCells(SpreadsheetApp.Dimension.ROWS);
    return 'deleted';
  }
  return 'bad';
}
