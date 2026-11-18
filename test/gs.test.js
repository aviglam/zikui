'use strict';
/* בדיקות ל-Code.gs (צד שרת). הרצה: node --test test/ */
const test = require('node:test');
const assert = require('node:assert/strict');
const { makeEnv } = require('./gs-harness');

function ready(opts) {
  const env = makeEnv(opts);
  env.gs.setupApi();
  env.key = env.scriptProps.getProperty('API_KEY');
  env.api = (action, data) => env.post(Object.assign({ action, key: env.key }, data || {}));
  return env;
}
const item = (id, caseId, pn, rma, status) => ({ type: 'add', item: { id, caseId, pn, rma, status, at: '2026-11-18T10:00:00Z' } });
const SAMPLE = ['5166539780', 'N36076-001', '0619940586010200B'];

test('setupApi: מפתח, כותרות, מצב ברירת מחדל בדיקה, עמודת מזהה מוסתרת', () => {
  const env = ready();
  assert.match(env.key, /^[0-9a-f]{64}$/);
  assert.ok(env.logs.some(l => l.includes(env.key)));
  assert.deepEqual(env.sheet.getRange('A1:F1').getValues()[0], ['Case ID', 'Part Number', 'RMA', 'מצב', 'תאריך', 'מזהה']);
  assert.equal(env.sheet.getRange('H3').getValue(), 'בדיקה');
  assert.ok(env.sheet.hidden.has(6));
  env.gs.setupApi();
  assert.equal(env.scriptProps.getProperty('API_KEY'), env.key, 'הרצה חוזרת לא מחליפה מפתח');
});

test('מפתח שגוי / חסר נדחה', () => {
  const env = ready();
  assert.deepEqual(env.post({ action: 'state', key: 'nope' }), { ok: false, auth: false, error: 'מפתח גישה שגוי' });
  assert.equal(env.post({ action: 'state' }).auth, false);
  assert.equal(JSON.parse(env.gs.doPost({ postData: { contents: 'garbage' } }).getContent()).ok, false);
});

test('sync: הוספה, ניקוי תווים נסתרים, אפסים מובילים, אידמפוטנטיות', () => {
  const env = ready();
  const r = env.api('sync', { ops: [
    item('a1', ' 5166539780‎', 'N36076-001\r', '​0619940586010200B ', 'Defective'),
    item('a2', '0012345678', 'L12345-B21', '0619940586010200C', 'Unused'),
    item('a1', 'x', 'x', 'x', 'Unused')   // שליחה חוזרת אחרי תשובה שאבדה
  ] });
  assert.equal(r.ok, true);
  assert.deepEqual(r.results, ['added', 'added', 'exists']);
  assert.equal(r.state.items.length, 2);
  assert.deepEqual([r.state.items[0].caseId, r.state.items[0].pn, r.state.items[0].rma], SAMPLE);
  assert.equal(r.state.items[1].caseId, '0012345678');
  assert.equal(env.sheet.getRange('A3').getValue(), '0012345678');
  assert.ok(env.sheet.getRange('E2').getValue() && Object.prototype.toString.call(env.sheet.getRange('E2').getValue()) === '[object Date]');
  // סטטוס לא חוקי לא נכתב
  env.api('sync', { ops: [item('a3', '1', '2', '3', 'Broken')] });
  assert.equal(env.sheet.getRange('D4').getValue(), '');
});

test('שורה שהוזנה ידנית בגיליון מקבלת מזהה; עדכון ומחיקה לפי מזהה', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective'), item('a2', '1111111111', 'P1', 'R1', 'Unused')] });
  env.sheet.getRange('A4:D4').setValues([['2222222222', 'P2', 'R2', 'Unused']]);
  let st = env.api('state').state;
  assert.equal(st.items.length, 3);
  assert.match(st.items[2].id, /^s-/);
  const manualId = st.items[2].id;

  st = env.api('sync', { ops: [{ type: 'update', id: 'a2', item: { caseId: '1111111111', pn: 'P1', rma: 'R1', status: 'Defective' } }] }).state;
  assert.equal(st.items[1].status, 'Defective');

  const h = ['H1', 'H2', 'H3', 'G1', 'G3'].map(c => env.sheet.getRange(c).getValue());
  const r = env.api('sync', { ops: [{ type: 'delete', id: 'a1' }, { type: 'delete', id: 'nope' }] });
  assert.deepEqual(r.results, ['deleted', 'missing']);
  assert.deepEqual(r.state.items.map(i => i.id), ['a2', manualId]);
  assert.deepEqual(['H1', 'H2', 'H3', 'G1', 'G3'].map(c => env.sheet.getRange(c).getValue()), h, 'תאי הבקרה לא זזו');
  assert.equal(env.sheet.getRange('A2').getValue(), '1111111111', 'השורות עלו למעלה');
});

test('הוספה אחרי סוף הגיליון מרחיבה שורות', () => {
  const env = ready({ maxRows: 4 });
  const r = env.api('sync', { ops: [1, 2, 3, 4, 5].map(n => item('a' + n, String(n), 'P', 'R', 'Unused')) });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.state.items.length, 5);
});

test('שליחה דו-שלבית מהאפליקציה – מצב בדיקה', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective'), item('a2', ...SAMPLE, 'Defective'), item('a3', '1', 'P', 'R', 'Unused')] });
  const a = env.api('arm');
  assert.equal(a.ok, true);
  assert.deepEqual(a.arm.counts, { total: 3, defective: 2, unused: 1 });
  assert.deepEqual(a.arm.dupIds, ['a2'], 'כפול = אזהרה, לא חוסם');
  assert.equal(a.arm.to, 'you@example.com');
  assert.equal(a.arm.shipNo, 1);
  assert.equal(a.arm.test, true);

  assert.equal(env.api('confirm', { fp: 'wrong' }).status, 'notArmed');
  assert.equal(env.mails.length, 0);

  const c = env.api('confirm', { fp: a.arm.fp });
  assert.equal(c.status, 'sent', JSON.stringify(c));
  assert.equal(c.shipNo, 1);
  assert.equal(env.mails.length, 1);
  assert.equal(env.mails[0].to, 'you@example.com');
  assert.equal(env.mails[0].subject, '[בדיקה] זיכוי חלקים – משלוח 1 (3 פריטים)');
  assert.match(env.mails[0].attachments[0].html, /\*\*\* בדיקה – לא משלוח אמיתי \*\*\*/);
  assert.match(env.mails[0].attachments[0].html, /שם הטכנאי/);
  assert.equal(env.files.length, 1);
  assert.equal(env.files[0].folder, 'זיכוי חלקים - בדיקות');
  assert.equal(c.state.items.length, 0, 'הרשימה נוקתה');
  assert.equal(env.sheet.getRange('F2').getValue(), '', 'גם עמודת המזהה נוקתה');
  assert.equal(env.sheet.getRange('H3').getValue(), 'בדיקה', 'תאי הבקרה נשארו');
  const hist = env.ss.getSheetByName('היסטוריה - בדיקות');
  assert.deepEqual(hist.getRange('A2:F2').getValues()[0].filter((_, i) => i !== 1), [1, ...SAMPLE, 'Defective']);
  assert.equal(env.docProps.getProperty('shipNo_test'), '1');
  assert.equal(env.docProps.getProperty('shipNo'), null, 'המונה האמיתי לא זז');
  assert.ok(c.pdf && c.pdfName.startsWith('[בדיקה] זיכוי חלקים - משלוח 1 - '));
  assert.equal(env.api('confirm', { fp: a.arm.fp }).status, 'notArmed', 'אי אפשר לאשר פעמיים');
  assert.equal(env.lockHeld, false);
});

test('הרשימה השתנתה בין השלבים → לא נשלח, ספירה מחודשת', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective')] });
  const a = env.api('arm');
  env.sheet.getRange('D2').setValue('Unused');           // עריכה ידנית בגיליון
  const c = env.api('confirm', { fp: a.arm.fp });
  assert.equal(c.status, 'changed');
  assert.equal(env.mails.length, 0);
  assert.deepEqual(c.arm.counts, { total: 1, defective: 0, unused: 1 });
  assert.equal(env.api('confirm', { fp: c.arm.fp }).status, 'sent');
});

test('אישור פג אחרי 15 דקות', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective')] });
  const a = env.api('arm');
  const armed = JSON.parse(env.docProps.getProperty('appArm'));
  armed.at -= 15 * 60000 + 1000;
  env.docProps.setProperty('appArm', JSON.stringify(armed));
  assert.equal(env.api('confirm', { fp: a.arm.fp }).status, 'expired');
  assert.equal(env.mails.length, 0);
});

test('החלפת מצב מבטלת שליחה פתוחה (מהאפליקציה ומהגיליון)', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective')] });
  let a = env.api('arm');
  env.api('setMode', { mode: 'אמיתי' });
  assert.equal(env.api('confirm', { fp: a.arm.fp }).status, 'notArmed');
  a = env.api('arm');
  env.edit('H3', 'בדיקה');
  assert.equal(env.api('confirm', { fp: a.arm.fp }).status, 'notArmed');
  assert.equal(env.mails.length, 0);
  assert.equal(env.api('setMode', { mode: 'xxx' }).ok, false);
});

test('מצב אמיתי: נמענים, בלי תגית, מונה/היסטוריה/תיקייה נפרדים', () => {
  const env = ready();
  env.api('sync', { ops: [item('t1', ...SAMPLE, 'Defective')] });
  env.api('confirm', { fp: env.api('arm').arm.fp });          // בדיקה #1
  env.api('setMode', { mode: 'אמיתי' });
  env.api('sync', { ops: [item('p1', ...SAMPLE, 'Unused')] });
  const a = env.api('arm');
  assert.equal(a.arm.shipNo, 1, 'מונה אמיתי מתחיל מ-1');
  assert.equal(a.arm.test, false);
  const c = env.api('confirm', { fp: a.arm.fp });
  assert.equal(c.status, 'sent');
  const m = env.mails[1];
  assert.equal(m.to, 'warehouse@example.com, office@example.com');
  assert.equal(m.subject, 'זיכוי חלקים – משלוח 1 (1 פריטים)');
  assert.doesNotMatch(m.attachments[0].html, /בדיקה – לא משלוח/);
  assert.equal(env.files[1].folder, 'זיכוי חלקים - משלוחים');
  assert.equal(env.ss.getSheetByName('היסטוריה').getLastRow(), 2);
  assert.equal(env.ss.getSheetByName('היסטוריה - בדיקות').getLastRow(), 2);
});

test('כשל במייל → לא נמחק כלום, המונה לא זז', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective')] });
  const a = env.api('arm');
  env.mailFail = true;
  const c = env.api('confirm', { fp: a.arm.fp });
  assert.equal(c.ok, false);
  assert.match(c.error, /לא נמחק כלום/);
  assert.equal(env.gs.getRows_(env.sheet).length, 1);
  assert.equal(env.docProps.getProperty('shipNo_test'), null);
  assert.equal(env.files.length, 0);
  assert.equal(env.lockHeld, false, 'הנעילה שוחררה');
});

test('ולידציה: חסר שדה חוסם שליחה ומחזיר מזהים', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective')] });
  env.sheet.getRange('A3:C3').setValues([['3333333333', '', 'R']]);
  const a = env.api('arm');
  assert.equal(a.arm, undefined);
  assert.match(a.error, /חסרים נתונים/);
  assert.equal(a.invalidIds.length, 1);
  assert.equal(env.api('arm').arm, undefined);
  const empty = ready();
  assert.equal(empty.api('arm').error, 'אין חלקים ברשימה');
});

test('נעילה תפוסה (שליחה מהגיליון רצה) → שגיאה ברורה', () => {
  const env = ready();
  env.lockHeld = true;
  const r = env.api('sync', { ops: [] });
  assert.equal(r.ok, false);
  assert.match(r.error, /פעולה אחרת/);
});

test('הגיליון עדיין עובד: H1 → ספירה, H2 → שליחה, אותו מונה', () => {
  const env = ready();
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective')] });
  env.sheet.getRange('A3:D3').setValues([['4444444444', 'Q12345-001', '0619940586010200X', 'Unused']]);
  env.edit('H1', 'TRUE');
  assert.match(env.sheet.getRange('G2').getValue(), /ברשימה 2 פריטים \(Defective: 1, Unused: 1\)/);
  env.edit('H2', 'TRUE');
  assert.match(env.sheet.getRange('G2').getValue(), /✅ \[בדיקה\] משלוח 1 נשלח/);
  assert.equal(env.mails.length, 1);
  assert.equal(env.api('state').state.nextShipNo, 2, 'האפליקציה רואה את המונה של הגיליון');
});

test('ניקוי בעריכה ידנית בגיליון (handleEdit)', () => {
  const env = ready();
  env.edit('A2', '‏ 5166539780\t');
  assert.equal(env.sheet.getRange('A2').getValue(), '5166539780');
  assert.ok(env.sheet.getRange('E2').getValue() && Object.prototype.toString.call(env.sheet.getRange('E2').getValue()) === '[object Date]');
});

test('תצוגה מקדימה והיסטוריה', () => {
  const env = ready();
  assert.equal(env.api('preview').ok, false);
  env.api('sync', { ops: [item('a1', ...SAMPLE, 'Defective'), item('a2', '1', 'P', 'R', 'Unused')] });
  const p = env.api('preview');
  assert.ok(p.pdf && p.pdfName.includes('משלוח טיוטה'));
  assert.equal(env.mails.length, 0, 'תצוגה מקדימה לא שולחת');
  assert.equal(env.files.length, 0, 'ולא שומרת קובץ');
  env.api('confirm', { fp: env.api('arm').arm.fp });
  env.api('sync', { ops: [item('b1', ...SAMPLE, 'Unused')] });
  env.api('confirm', { fp: env.api('arm').arm.fp });
  const h = env.api('history');
  assert.equal(h.mode, 'בדיקה');
  assert.deepEqual(h.shipments.map(s => [s.shipNo, s.total, s.defective, s.unused]), [[2, 1, 0, 1], [1, 2, 1, 1]]);
});
