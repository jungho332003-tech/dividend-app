// 아주 작은 Apps Script 흉내: dist/Code.gs를 그대로 불러와 서버 함수를 시험한다
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const chain = new Proxy(function () {}, { get: (t, k) => (k === 'build' ? () => ({}) : () => chain), apply: () => chain });

function a1ToRC(a1) {
  const m = /^([A-Z]+)(\d*)(?::([A-Z]+)(\d*))?$/.exec(a1);
  const col = s => s.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const r1 = m[2] ? +m[2] : 1, c1 = col(m[1]);
  const r2 = m[3] ? (m[4] ? +m[4] : 1000) : r1, c2 = m[3] ? col(m[3]) : c1;
  return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
}

class Sheet {
  constructor(name, id) { this.name = name; this.id = id; this.d = []; this.prot = []; }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getLastRow() { for (let r = this.d.length; r > 0; r--) if ((this.d[r - 1] || []).some(v => v !== '' && v != null)) return r; return 0; }
  getLastColumn() { let m = 0; this.d.forEach(row => { for (let c = row.length; c > 0; c--) if (row[c - 1] !== '' && row[c - 1] != null) { m = Math.max(m, c); break; } }); return m; }
  getMaxRows() { return Math.max(1000, this.d.length); }
  get(r, c) { return (this.d[r - 1] || [])[c - 1] ?? ''; }
  set(r, c, v) { while (this.d.length < r) this.d.push([]); const row = this.d[r - 1]; while (row.length < c) row.push(''); row[c - 1] = v; }
  getRange(a, b, nr, nc) {
    if (typeof a === 'string') [a, b, nr, nc] = a1ToRC(a);
    return new Range(this, a, b, nr || 1, nc || 1);
  }
  appendRow(vals) { const r = this.getLastRow() + 1; vals.forEach((v, i) => this.set(r, i + 1, v)); }
  deleteRow(r) { this.d.splice(r - 1, 1); }
  getProtections() { return this.prot; }
  protect() { const p = { warn: false, eds: [], setDescription() { return p; }, setWarningOnly(w) { p.warn = w; return p; }, isWarningOnly() { return p.warn; }, remove: () => { this.prot = this.prot.filter(x => x !== p); }, addEditor(e) { p.eds.push(e); }, addEditors(l) { p.eds.push(...l); }, removeEditors() { p.eds = []; }, getEditors() { return p.eds.map(e => ({ getEmail: () => e })); }, canDomainEdit() { return false; }, setDomainEdit() {}, setUnprotectedRanges() {} }; this.prot.push(p); return p; }
}
['setColumnWidth', 'setFrozenRows', 'setFrozenColumns', 'setRowHeight', 'hideColumns', 'hideSheet'].forEach(k => { Sheet.prototype[k] = function () { return this; }; });
Sheet.prototype.clear = function () { this.d = []; return this; };
Sheet.prototype.setName = function (n) { this.name = n; return this; };
Sheet.prototype.getDataRange = function () { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); };

class Range {
  constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
  getValues() { return Array.from({ length: this.nr }, (_, i) => Array.from({ length: this.nc }, (_, j) => this.sh.get(this.r + i, this.c + j))); }
  getValue() { return this.sh.get(this.r, this.c); }
  setValues(v) { assert.strictEqual(v.length, this.nr, 'rows mismatch'); v.forEach((row, i) => { assert.strictEqual(row.length, this.nc, 'cols mismatch'); row.forEach((x, j) => this.sh.set(this.r + i, this.c + j, x)); }); return this; }
  setValue(v) { this.sh.set(this.r, this.c, v); return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.set(this.r + i, this.c + j, ''); return this; }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  protect() { return this.sh.protect(); }
}
['clearDataValidations', 'setNumberFormat', 'setWrap', 'setBackground', 'setFontWeight', 'setHorizontalAlignment', 'setBorder', 'setNote', 'setDataValidation', 'setFontSize', 'setFontColor', 'merge', 'breakApart', 'clear', 'setVerticalAlignment', 'insertCheckboxes'].forEach(k => { Range.prototype[k] = function () { return this; }; });

function makeBook(name, id) {
  const sheets = [new Sheet('Sheet1', 0)];
  let sid = 1;
  return {
    sheets, getSheets: () => sheets, getId: () => id, getUrl: () => 'https://docs/' + id, getName: () => name,
    getSheetByName: n => sheets.find(s => s.name === n) || null,
    insertSheet: n => { const s = new Sheet(n, sid++); sheets.push(s); return s; },
    setActiveSheet() {},
  };
}

function makeEnv() {
  const sheets = [];
  let sid = 1;
  const books = {};
  const ss = {
    getSheetByName: n => sheets.find(s => s.name === n) || null,
    insertSheet: n => { const s = new Sheet(n, sid++); sheets.push(s); return s; },
    getUrl: () => 'https://sheet', getName: () => '연말정산 검토', getId: () => 'SSID',
    setSpreadsheetTimeZone() {}, getOwner: () => ({ getEmail: () => 'owner@x.com' }),
  };
  const cacheStore = {};
  const userProps = {};
  const state = { user: 'owner@x.com', files: {}, fseq: 0, editors: {}, trashed: [] };
  const mkFolder = (name) => {
    const subs = {};
    const f = {
      getName: () => name,
      getFoldersByName: n => { const has = !!subs[n]; return { hasNext: () => has, next: () => subs[n] }; },
      createFolder: n => (subs[n] = mkFolder(n)),
      createFile: blob => { const id = 'F' + (++state.fseq); const file = { getId: () => id, getName: () => blob.name, getUrl: () => 'https://drive/' + id, getMimeType: () => blob.mime, setDescription() {}, setTrashed(t) { if (t) state.trashed.push(id); } }; state.files[id] = file; return file; },
      getOwner: () => ({ getEmail: () => 'owner@x.com' }), getEditors: () => [], getViewers: () => [], addEditor() {}, removeEditor() {}, removeViewer() {},
    };
    return f;
  };
  const root = mkFolder('root');
  const g = {
    console,
    SpreadsheetApp: {
      getActive: () => ss, getUi: () => ({ alert() {}, createMenu: () => chain, showModalDialog() {} }),
      newDataValidation: () => chain, ProtectionType: { SHEET: 'S', RANGE: 'R' }, BorderStyle: {}, flush() {},
      create: n => { const id = 'BOOK' + (Object.keys(books).length + 1); books[id] = makeBook(n, id); return books[id]; },
      openById: id => books[id],
    },
    Session: { getActiveUser: () => ({ getEmail: () => state.user }), getEffectiveUser: () => ({ getEmail: () => state.user }) },
    CacheService: { getDocumentCache: () => ({
      get: k => cacheStore[k] ?? null, put: (k, v) => { assert(v.length <= 100000, 'cache value too big'); cacheStore[k] = v; },
      getAll: ks => Object.fromEntries(ks.filter(k => k in cacheStore).map(k => [k, cacheStore[k]])),
      putAll: o => Object.entries(o).forEach(([k, v]) => { assert(v.length <= 100000); cacheStore[k] = v; }),
      remove: k => { delete cacheStore[k]; }, removeAll: ks => ks.forEach(k => delete cacheStore[k]),
    }) },
    PropertiesService: { getUserProperties: () => ({ getProperty: k => (userProps[state.user] || {})[k] ?? null, setProperty: (k, v) => { (userProps[state.user] = userProps[state.user] || {})[k] = v; } }) },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      formatDate: (d, tz, p) => { const k = new Date(d.getTime() + 9 * 3600e3); const pad = n => String(n).padStart(2, '0'); return p.replace('yyyy', k.getUTCFullYear()).replace('MM', pad(k.getUTCMonth() + 1)).replace('dd', pad(k.getUTCDate())).replace('HH', pad(k.getUTCHours())).replace('mm', pad(k.getUTCMinutes())); },
      base64Decode: s => Array.from(Buffer.from(s, 'base64')), newBlob: (bytes, mime, name) => ({ bytes, mime, name }),
    },
    DriveApp: { getFolderById: id => { if (id !== 'FOLDER') throw new Error('no folder'); return root; }, getFileById: id => books[id] ? { isTrashed: () => false } : id === 'SSID' ? { getOwner: () => ({ getEmail: () => 'owner@x.com' }), getEditors: () => [], getViewers: () => [], addEditor(e) { state.editors[e] = 1; }, removeEditor() {}, removeViewer() {} } : (state.files[id] || (() => { throw new Error('nf'); })()) },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://webapp' }) },
    HtmlService: { createHtmlOutputFromFile: () => chain, createHtmlOutput: () => chain },
  };
  vm.createContext(g);
  vm.runInContext(fs.readFileSync(require('path').join(__dirname, '..', 'dist', 'Code.gs'), 'utf8') + '\n;this.__api = { COLS, initialize, apiBootstrap, apiSavePerson, apiImportPeople, apiAddPerson, apiDeletePerson, apiAddLog, apiSaveNotice, apiDeleteNotice, apiAddNoticeComment, apiMarkRead, apiMarkUnread, apiSaveRule, apiDeleteRule, apiUploadFile, apiDeleteFile, apiAdminData, apiSaveMembers, apiSaveMenuAccess, apiSaveSettings, apiApplyPermissions, apiRevokeAccess, putCache_, cached_, setupProtections, apiTemplateLink, CtxDate: Date };', g);
  return { api: g.__api, ss, state, cacheStore, books };
}

module.exports = { makeEnv, assert };
