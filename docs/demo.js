'use strict';
/* מצב הדגמה (כתובת "demo" בהגדרות): שרת מדומה בתוך הדפדפן, אותה התנהגות כמו Code.gs.
 * לא נשלח שום מייל. מאפשר לנסות את הממשק, הסריקה והשליחה הדו-שלבית לפני חיבור לגיליון.
 * הדמיית אין קליטה: Demo.setOffline(true)
 */
const Demo = (() => {
  const KEY = 'zikui.demo.db';
  const TTL = 15 * 60000;
  const MODES = {
    'בדיקה': { to: 'test@example.com', key: 'shipNo_test', tag: '[בדיקה] ' },
    'אמיתי': { to: 'warehouse@example.com, office@example.com', key: 'shipNo', tag: '' }
  };
  const db = () => {
    try { return JSON.parse(localStorage.getItem(KEY)) || null; } catch (_) { return null; }
  };
  let D = db() || { mode: 'בדיקה', items: [], shipNo: 0, shipNo_test: 0, hist: { 'בדיקה': [], 'אמיתי': [] }, arm: null, offline: false };
  const put = () => { try { localStorage.setItem(KEY, JSON.stringify(D)); } catch (_) { /* */ } };

  const clean = s => String(s == null ? '' : s).replace(/[\u0000-\u001F\u007F​-‏‪-‮⁦-⁩﻿]/g, '').trim();
  const fp = () => D.mode + JSON.stringify(D.items.map(i => [i.caseId, i.pn, i.rma, i.status]));
  const counts = items => ({ total: items.length, defective: items.filter(i => i.status === 'Defective').length, unused: items.filter(i => i.status === 'Unused').length });
  const state = () => {
    const m = MODES[D.mode];
    return { mode: D.mode, to: m.to, test: !!m.tag, nextShipNo: D[m.key] + 1, items: D.items.map(i => Object.assign({}, i)) };
  };
  function dupIds() {
    const seen = new Set(), d = [];
    D.items.forEach(i => { const k = [i.caseId, i.pn, i.rma].join('|'); if (seen.has(k)) d.push(i.id); else seen.add(k); });
    return d;
  }
  function arm() {
    if (!D.items.length) { D.arm = null; return { error: 'אין חלקים ברשימה' }; }
    const bad = D.items.filter(i => !i.caseId || !i.pn || !i.rma || ['Defective', 'Unused'].indexOf(i.status) < 0);
    if (bad.length) { D.arm = null; return { error: `חסרים נתונים (כולל מצב) ב-${bad.length} פריטים`, invalidIds: bad.map(i => i.id) }; }
    const m = MODES[D.mode];
    D.arm = { fp: fp(), at: Date.now() };
    return { arm: { fp: D.arm.fp, mode: D.mode, to: m.to, test: !!m.tag, counts: counts(D.items), dupIds: dupIds(), ttlMs: TTL, shipNo: D[m.key] + 1 } };
  }

  const H = {
    state: () => ({ state: state() }),
    sync: req => {
      const results = (req.ops || []).map(op => {
        if (op.type === 'add') {
          const it = op.item || {};
          if (!it.id) return 'bad';
          if (D.items.some(i => i.id === it.id)) return 'exists';
          D.items.push({ id: it.id, caseId: clean(it.caseId), pn: clean(it.pn), rma: clean(it.rma), status: clean(it.status), date: it.at || new Date().toISOString() });
          return 'added';
        }
        const i = D.items.findIndex(x => x.id === op.id);
        if (i < 0) return 'missing';
        if (op.type === 'update') { Object.assign(D.items[i], { caseId: clean(op.item.caseId), pn: clean(op.item.pn), rma: clean(op.item.rma), status: clean(op.item.status) }); return 'updated'; }
        if (op.type === 'delete') { D.items.splice(i, 1); return 'deleted'; }
        return 'bad';
      });
      return { results, state: state() };
    },
    arm: () => Object.assign(arm(), { state: state() }),
    confirm: req => {
      if (!D.arm || D.arm.fp !== req.fp) return { status: 'notArmed', state: state() };
      if (Date.now() - D.arm.at > TTL) { D.arm = null; return { status: 'expired', state: state() }; }
      if (fp() !== D.arm.fp) return Object.assign({ status: 'changed' }, arm(), { state: state() });
      D.arm = null;
      const m = MODES[D.mode];
      const shipNo = D[m.key] + 1;
      const items = D.items;
      const name = `${m.tag}זיכוי חלקים - משלוח ${shipNo} - ${new Date().toISOString().slice(0, 10)}.pdf`;
      const pdf = makePdf(items, shipNo, m);
      D.hist[D.mode].push({ shipNo, date: new Date().toISOString(), c: counts(items) });
      D.items = [];
      D[m.key] = shipNo;
      return { status: 'sent', shipNo, mode: D.mode, to: m.to, counts: counts(items), pdf, pdfName: name, state: state() };
    },
    disarm: () => { D.arm = null; return {}; },
    setMode: req => {
      if (!MODES[req.mode]) throw new Error('מצב לא מוכר');
      D.mode = req.mode; D.arm = null;
      return { state: state() };
    },
    preview: () => {
      if (!D.items.length) throw new Error('אין חלקים ברשימה');
      const m = MODES[D.mode];
      return { pdf: makePdf(D.items, 'DRAFT', m), pdfName: `${m.tag}זיכוי חלקים - טיוטה.pdf` };
    },
    history: () => ({
      mode: D.mode,
      shipments: D.hist[D.mode].slice().reverse().map(s => ({ shipNo: s.shipNo, date: s.date, total: s.c.total, defective: s.c.defective, unused: s.c.unused }))
    })
  };

  /** PDF מינימלי (ASCII בלבד) – רק להדגמה. ה-PDF האמיתי נוצר ב-Apps Script */
  function makePdf(items, shipNo, m) {
    const esc = s => String(s).replace(/[\\()]/g, c => '\\' + c).replace(/[^\x20-\x7E]/g, '?');
    const c = counts(items);
    const lines = [
      (m.tag ? '*** TEST - not a real shipment ***   ' : '') + 'Parts credit - shipment ' + shipNo + '  (DEMO)',
      'Date: ' + new Date().toLocaleDateString('en-GB') + '   Items: ' + c.total + '   Defective: ' + c.defective + '   Unused: ' + c.unused,
      '',
      '#    Case ID       Part Number     RMA                   Status'
    ].concat(items.slice(0, 48).map((i, k) =>
      String(k + 1).padEnd(5) + i.caseId.padEnd(14) + i.pn.padEnd(16) + i.rma.padEnd(22) + i.status));
    const content = 'BT /F1 10 Tf 40 800 Td 14 TL ' + lines.map(l => '(' + esc(l) + ') Tj T*').join(' ') + ' ET';
    const objs = [
      '<< /Type /Catalog /Pages 2 0 R >>',
      '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
      '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
      '<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>',
      '<< /Length ' + content.length + ' >>\nstream\n' + content + '\nendstream'
    ];
    let out = '%PDF-1.4\n';
    const offs = [];
    objs.forEach((o, k) => { offs.push(out.length); out += (k + 1) + ' 0 obj\n' + o + '\nendobj\n'; });
    const x = out.length;
    out += 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n' + offs.map(o => String(o).padStart(10, '0') + ' 00000 n \n').join('');
    out += 'trailer\n<< /Size ' + (objs.length + 1) + ' /Root 1 0 R >>\nstartxref\n' + x + '\n%%EOF';
    return btoa(out);
  }

  async function handle(action, data) {
    await new Promise(r => setTimeout(r, action === 'confirm' ? 900 : 250));
    if (D.offline) throw new NetError('אין חיבור לשרת');
    const fn = H[action];
    if (!fn) throw new Error('פעולה לא מוכרת: ' + action);
    const res = fn(data);
    put();
    return Object.assign({ ok: true }, JSON.parse(JSON.stringify(res)));
  }

  return {
    handle,
    setOffline(v) { D.offline = !!v; put(); },
    get offline() { return D.offline; },
    reset() { localStorage.removeItem(KEY); D = { mode: 'בדיקה', items: [], shipNo: 0, shipNo_test: 0, hist: { 'בדיקה': [], 'אמיתי': [] }, arm: null, offline: false }; },
    /** הדמיית עריכה ידנית בגיליון בין שלב 1 לשלב 2 */
    sheetEdit() { if (D.items[0]) { D.items[0].status = D.items[0].status === 'Defective' ? 'Unused' : 'Defective'; put(); } }
  };
})();
