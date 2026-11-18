'use strict';
/* סביבת Google Apps Script מדומה – מספיקה כדי להריץ את Code.gs בנוד ולבדוק את הלוגיקה. */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

function colNum(letters) { return letters.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); }
function a1(r, c) { let s = ''; let n = c; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s + r; }

const chain = () => new Proxy({}, { get: (t, p) => (p === 'build' ? () => ({}) : () => chain()) });

class FakeSheet {
  constructor(name, ss, maxRows = 200) { this.name = name; this.ss = ss; this.maxRows = maxRows; this.cells = {}; this.hidden = new Set(); }
  key(r, c) { return r + ',' + c; }
  get(r, c) { const v = this.cells[this.key(r, c)]; return v === undefined ? '' : v; }
  set(r, c, v) {
    if (r > this.maxRows) throw new Error(`Range out of bounds: row ${r} > ${this.maxRows}`);
    if (v === '' || v == null) delete this.cells[this.key(r, c)]; else this.cells[this.key(r, c)] = v;
  }
  getName() { return this.name; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(_after, n) { this.maxRows += n; }
  getLastRow() { let m = 0; for (const k of Object.keys(this.cells)) { const r = Number(k.split(',')[0]); if (r > m) m = r; } return m; }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      const r1 = Number(m[2]), c1 = colNum(m[1]);
      const r2 = m[3] ? Number(m[4]) : r1, c2 = m[3] ? colNum(m[3]) : c1;
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, a, b, c || 1, d || 1);
  }
  setFrozenRows() {} hideColumns(c) { this.hidden.add(c); } setColumnWidth() {} setConditionalFormatRules() {}
}

class FakeRange {
  constructor(sh, r, c, nr, nc) {
    if (r < 1 || c < 1 || nr < 1 || nc < 1) throw new Error('bad range');
    if (r + nr - 1 > sh.maxRows) throw new Error(`Range out of bounds: rows ${r}..${r + nr - 1} > ${sh.maxRows}`);
    Object.assign(this, { sh, r, c, nr, nc });
  }
  getSheet() { return this.sh; }
  getRow() { return this.r; } getColumn() { return this.c; } getNumRows() { return this.nr; }
  getA1Notation() { return this.nr === 1 && this.nc === 1 ? a1(this.r, this.c) : a1(this.r, this.c) + ':' + a1(this.r + this.nr - 1, this.c + this.nc - 1); }
  getValue() { return this.sh.get(this.r, this.c); }
  getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.get(this.r + i, this.c + j)); out.push(row); } return out; }
  setValue(v) { this.sh.set(this.r, this.c, v); return this; }
  setValues(vals) {
    if (vals.length !== this.nr || vals.some(r => r.length !== this.nc)) throw new Error('setValues dimension mismatch');
    vals.forEach((row, i) => row.forEach((v, j) => this.sh.set(this.r + i, this.c + j, v)));
    return this;
  }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); return this; }
  deleteCells(dim) {
    if (dim !== 'ROWS') throw new Error('only ROWS');
    const last = this.sh.maxRows;
    for (let j = 0; j < this.nc; j++) {
      const col = this.c + j;
      for (let r = this.r; r <= last; r++) {
        const src = r + this.nr;
        this.sh.set(r, col, src <= last ? this.sh.get(src, col) : '');
      }
    }
  }
  uncheck() { return this.setValue(false); }
  insertCheckboxes() { if (this.getValue() === '') this.setValue(false); return this; }
  activate() { return this; }
  setNumberFormat() { return this; } setFontWeight() { return this; } setDataValidation() { return this; }
  clearDataValidations() { return this; } setWrap() { return this; } setBackground() { return this; }
}

class FakeSpreadsheet {
  constructor() { this.sheets = {}; }
  getSheetByName(n) { return this.sheets[n] || null; }
  insertSheet(n) { return (this.sheets[n] = new FakeSheet(n, this)); }
}

function props() {
  const m = {};
  return {
    _m: m,
    getProperty: k => (k in m ? m[k] : null),
    setProperty: (k, v) => { m[k] = String(v); },
    setProperties: o => Object.keys(o).forEach(k => { m[k] = String(o[k]); }),
    deleteProperty: k => { delete m[k]; }
  };
}

function makeEnv({ maxRows = 200 } = {}) {
  const ss = new FakeSpreadsheet();
  ss.insertSheet('גיליון1');
  ss.sheets['גיליון1'].maxRows = maxRows;
  const docProps = props(), scriptProps = props();
  const mails = [], files = [], logs = [];
  const env = { ss, docProps, scriptProps, mails, files, logs, mailFail: false, lockHeld: false };
  const folders = {};

  const g = {
    SpreadsheetApp: {
      getActive: () => ss,
      newDataValidation: chain,
      newConditionalFormatRule: chain,
      Dimension: { ROWS: 'ROWS', COLUMNS: 'COLUMNS' },
      getUi: () => { throw new Error('no UI in tests'); }
    },
    PropertiesService: { getDocumentProperties: () => docProps, getScriptProperties: () => scriptProps },
    LockService: {
      getDocumentLock: () => ({
        tryLock: () => { if (env.lockHeld) return false; env.lockHeld = true; return true; },
        releaseLock: () => { env.lockHeld = false; }
      })
    },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      DigestAlgorithm: { MD5: 'md5' },
      computeDigest: (alg, s) => [...crypto.createHash(alg).update(String(s), 'utf8').digest()].map(b => (b > 127 ? b - 256 : b)),
      base64Encode: x => Buffer.from(typeof x === 'string' ? x : x.map(b => (b < 0 ? b + 256 : b))).toString('base64'),
      formatDate: (d, tz, f) => {
        const p = n => String(n).padStart(2, '0');
        return f.replace('yyyy', d.getFullYear()).replace('MM', p(d.getMonth() + 1)).replace('dd', p(d.getDate()));
      },
      newBlob: (html) => ({
        getAs: () => {
          let name = 'x.pdf';
          const blob = { html, setName: n => { name = n; return blob; }, getName: () => name, getBytes: () => [...Buffer.from('%PDF-fake ' + html.length)] };
          return blob;
        }
      })
    },
    MailApp: { sendEmail: o => { if (env.mailFail) throw new Error('Service invoked too many times'); mails.push(o); } },
    DriveApp: {
      getFoldersByName: n => { const f = folders[n]; let used = false; return { hasNext: () => !!f && !used, next: () => { used = true; return f; } }; },
      createFolder: n => (folders[n] = { name: n, createFile: b => { files.push({ folder: n, name: b.getName() }); return { getUrl: () => 'https://drive/x' }; } })
    },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ content: s, setMimeType() { return this; }, getContent() { return this.content; } }) },
    Session: { getScriptTimeZone: () => 'Asia/Jerusalem' },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger: () => {}, newTrigger: chain },
    HtmlService: { createHtmlOutput: chain },
    Logger: { log: s => logs.push(String(s)) },
    console
  };
  const ctx = vm.createContext(g);
  const dir = path.join(__dirname, '..', 'apps-script');
  // כמו ב-Apps Script: כל הקבצים באותו scope גלובלי. Code.gs לפני Config.gs – כדי לוודא שסדר הקבצים לא משנה.
  const src = fs.readFileSync(path.join(dir, 'Code.gs'), 'utf8') + '\n' + fs.readFileSync(path.join(dir, 'Config.example.gs'), 'utf8') +
    '\n;globalThis.__api = { setup, setupApi, doPost, doGet, handleEdit, getRows_, MODES };';
  vm.runInContext(src, ctx, { filename: 'Code.gs+Config.example.gs' });
  env.gs = ctx.__api;
  env.sheet = ss.sheets['גיליון1'];
  env.post = body => JSON.parse(env.gs.doPost({ postData: { contents: JSON.stringify(body) } }).getContent());
  env.edit = (a1n, value) => { const r = env.sheet.getRange(a1n); r.setValue(value === 'TRUE' ? true : value); env.gs.handleEdit({ range: r, value }); };
  return env;
}

module.exports = { makeEnv };
