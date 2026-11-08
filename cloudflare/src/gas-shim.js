// ============================================================================================
// Google Apps Script stand-ins for Cloudflare Workers (Durable Object + SQLite storage).
// Code.gs runs unchanged on top of these: SpreadsheetApp, PropertiesService, CacheService,
// LockService, Utilities, Session, HtmlService, ContentService, ScriptApp, DriveApp, MailApp,
// UrlFetchApp (only the "export as .xlsx" call Code.gs makes), Logger.
// Everything is synchronous, like Apps Script: a Durable Object's SQLite storage is synchronous.
// ============================================================================================

/** Per-request runtime context. */
const RT = {
  env: {},              // Worker env (vars + secrets)
  origin: '',           // https://your-worker.workers.dev
  tz: 'UTC',
  main: null,           // the tracker spreadsheet, a working copy over M for this request
  temp: new Map(),      // id -> temporary spreadsheet (exports)
  outbox: [],           // emails to send after the request
  needed: []            // D1: data the request asked for that was not loaded yet
};

/** Starts a request (or a re-run of it): fresh journal and working copy. */
function shimBegin_(env, origin) {
  RT.env = env || {};
  RT.tz = String(RT.env.TIMEZONE || 'UTC');
  if (origin) RT.origin = origin;
  RT.temp = new Map();
  RT.outbox = [];
  RT.needed = [];
  J = journalNew_();
  RT.main = new GSpreadsheet_('main', 'Cabin TV Tracker', true);
}

// --------------------------------------------------------------------------------------------
// Values: Dates are kept as Dates (stored as {"$d": ms})
// --------------------------------------------------------------------------------------------
function encRow_(row) { return JSON.stringify(row.map(v => v instanceof Date ? { $d: v.getTime() } : v)); }
function decRow_(text) {
  return JSON.parse(text).map(v => (v && typeof v === 'object' && '$d' in v) ? new Date(v.$d) : v);
}
function cellIn_(v) {
  if (v === undefined || v === null) return '';
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : new Date(v.getTime());
  if (typeof v === 'number') return isFinite(v) ? v : '';
  if (typeof v === 'boolean' || typeof v === 'string') return v;
  return String(v);
}
function cellOut_(v) { return v instanceof Date ? new Date(v.getTime()) : (v === undefined || v === null ? '' : v); }
function isBlank_(v) { return v === '' || v === null || v === undefined; }

function displayOf_(v, fmt) {
  if (v instanceof Date) return gasFormatDate_(v, RT.tz, fmt && fmt !== '@' && /[yMdHhms]/.test(fmt) ? fmt : 'yyyy-MM-dd HH:mm:ss');
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  if (v === '' || v == null) return '';
  return String(v);
}

// --------------------------------------------------------------------------------------------
// Spreadsheet / Sheet / Range
// --------------------------------------------------------------------------------------------
function chainable_(target) {
  // Any formatting call we don't model (setFontWeight, setBackground, createFilter, …) is a no-op
  // that returns the object, so Code.gs chains keep working.
  return new Proxy(target, {
    get(t, prop) {
      if (prop in t) return t[prop];
      if (typeof prop === 'symbol' || prop === 'toJSON' || prop === 'then' || prop === 'constructor') return undefined;
      return function () { return this; };
    }
  });
}

class GSpreadsheet_ {
  constructor(id, name, persisted) {
    this._id = id; this._name = name; this._persisted = persisted; this._sheets = null; this._dropped = new Set(); this._orderDirty = false;
    return chainable_(this);
  }
  _load() {
    if (this._sheets) return;
    this._sheets = [];
    if (!this._persisted) return;
    for (const m of ST_.sheetList()) {
      const sh = new GSheet_(this, m.name, JSON.parse(JSON.stringify(m.meta || {})));
      sh._known = true;
      this._sheets.push(sh);
    }
  }
  getId() { return this._id; }
  getName() { return this._name; }
  getUrl() { return RT.origin + '/'; }
  getSpreadsheetTimeZone() { return RT.tz; }
  getSheets() { this._load(); return this._sheets.slice(); }
  getSheetByName(name) { this._load(); return this._sheets.find(s => s._name === String(name)) || null; }
  getActiveSheet() { this._load(); return this._sheets[0] || null; }
  insertSheet(name, index) {
    this._load();
    name = name == null ? 'Sheet' + (this._sheets.length + 1) : String(name);
    if (this.getSheetByName(name)) throw new Error('A sheet with the name "' + name + '" already exists. Please enter another name.');
    const sh = new GSheet_(this, name, {});
    sh._rows = []; sh._loaded = true; sh._isNew = true;
    if (typeof index === 'number' && index >= 0 && index <= this._sheets.length) this._sheets.splice(index, 0, sh);
    else this._sheets.push(sh);
    this._orderDirty = true;
    return sh;
  }
  deleteSheet(sheet) {
    this._load();
    const raw = sheet && sheet._self ? sheet._self : sheet;
    const target = this._sheets.find(s => s === raw) || this._sheets.find(s => s._name === (sheet && sheet.getName && sheet.getName()));
    if (!target) return;
    this._sheets.splice(this._sheets.indexOf(target), 1);
    if (target._known) this._dropped.add(target._name);
    this._orderDirty = true;
  }
  toast() { return this; }
  /** Moves this request's sheet changes into the journal J (nothing is written yet). */
  _stage() {
    if (!this._persisted || !this._sheets) return;
    this._dropped.forEach(n => { if (!this._sheets.some(s => s._name === n && s._known)) J.drop.add(n); });
    this._sheets.forEach((sh, pos) => {
      const m = M.sheets.get(sh._name);
      if (sh._isNew || sh._metaDirty || this._orderDirty || !m || m.pos !== pos) J.sheets.set(sh._name, { pos, meta: JSON.parse(JSON.stringify(sh._meta)) });
      sh._stageRows();
      sh._metaDirty = false; sh._isNew = false; sh._known = true;
    });
    this._dropped.clear();
    this._orderDirty = false;
  }
}

class GSheet_ {
  constructor(ss, name, meta) {
    this._ss = ss; this._name = name; this._meta = Object.assign({ formats: {}, frozen: 0 }, meta || {});
    this._rows = null; this._loaded = false; this._dirty = new Set(); this._owned = new Set(); this._dims = null; this._self = this;
    this._isNew = false; this._known = false;
    return chainable_(this);
  }
  _load() {
    if (this._loaded) return;
    // Working copy: the row arrays are shared with M until a row is changed (then it is copied).
    this._rows = this._ss._persisted && this._known ? ST_.sheetRows(this._name).slice() : [];
    for (let i = 0; i < this._rows.length; i++) if (!this._rows[i]) this._rows[i] = [];
    this._loaded = true;
  }
  _own(r) {
    if (!this._owned.has(r)) { this._rows[r] = this._rows[r] ? this._rows[r].slice() : []; this._owned.add(r); }
    return this._rows[r];
  }
  _touch(r) { this._dirty.add(r); this._dims = null; }
  _stageRows() {
    if (!this._loaded) return;
    if (this._isNew) { for (let r = 0; r < this._rows.length; r++) this._dirty.add(r); }
    if (!this._dirty.size) return;
    let jr = J.rows.get(this._name);
    if (!jr) { jr = new Map(); J.rows.set(this._name, jr); }
    for (const r of this._dirty) {
      const row = this._rows[r];
      if (!row || !row.some(v => !isBlank_(v))) jr.set(r, null);
      else { let n = row.length; while (n > 0 && isBlank_(row[n - 1])) n--; jr.set(r, row.slice(0, n)); }
    }
    this._dirty.clear();
  }
  _calcDims() {
    if (this._dims) return this._dims;
    this._load();
    let lastRow = 0, lastCol = 0;
    for (let r = 0; r < this._rows.length; r++) {
      const row = this._rows[r];
      if (!row) continue;
      let n = row.length; while (n > 0 && isBlank_(row[n - 1])) n--;
      if (n) { lastRow = r + 1; if (n > lastCol) lastCol = n; }
    }
    this._dims = { lastRow, lastCol };
    return this._dims;
  }
  _get(r, c) { const row = this._rows[r]; return row ? row[c] : undefined; }
  _set(r, c, v) {
    while (this._rows.length <= r) { this._rows.push([]); this._owned.add(this._rows.length - 1); }
    const row = this._own(r);
    while (row.length < c) row.push('');
    row[c] = v;
    this._touch(r);
  }
  getName() { return this._name; }
  setName(name) {
    name = String(name);
    if (name === this._name) return this;
    if (this._ss._persisted) {
      this._load();
      if (this._known) this._ss._dropped.add(this._name);
      this._isNew = true; this._known = false; this._ss._orderDirty = true;
    }
    this._name = name; this._metaDirty = true;
    return this;
  }
  getParent() { return this._ss; }
  getSheetId() { return this._name.length; }
  getLastRow() { return this._calcDims().lastRow; }
  getLastColumn() { return this._calcDims().lastCol; }
  getMaxRows() { return Math.max(this.getLastRow(), 1); }
  getMaxColumns() { return Math.max(this.getLastColumn(), 1); }
  getFrozenRows() { return this._meta.frozen || 0; }
  setFrozenRows(n) { this._meta.frozen = Number(n) || 0; this._metaDirty = true; return this; }
  getDataRange() { const d = this._calcDims(); return new GRange_(this, 1, 1, Math.max(1, d.lastRow), Math.max(1, d.lastCol)); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') return a1Range_(this, a);
    if (!(a >= 1) || !(b >= 1)) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
    const nr = c === undefined ? 1 : c, nc = d === undefined ? 1 : d;
    if (!(nr >= 1) || !(nc >= 1)) throw new Error('The number of rows in the range must be at least 1.');
    return new GRange_(this, a, b, nr, nc);
  }
  appendRow(values) {
    const r = this.getLastRow();
    values.forEach((v, j) => this._set(r, j, cellIn_(v)));
    if (!values.length) this._touch(r);
    return this;
  }
  deleteRows(start, n) { this._load(); this._rows.splice(start - 1, n); this._reindexOwned(start - 1, -n); this._markAllFrom(start - 1, n); return this; }
  deleteRow(r) { return this.deleteRows(r, 1); }
  insertRows(before, n) {
    this._load(); n = n || 1;
    const add = []; for (let i = 0; i < n; i++) add.push([]);
    this._rows.splice(before - 1, 0, ...add); this._reindexOwned(before - 1, n);
    for (let i = 0; i < n; i++) this._owned.add(before - 1 + i);
    this._markAllFrom(before - 1, 0); return this;
  }
  _reindexOwned(from, delta) {
    const next = new Set();
    this._owned.forEach(r => { if (r < from) next.add(r); else if (delta > 0) next.add(r + delta); else if (r >= from - delta) next.add(r + delta); });
    this._owned = next;
  }
  insertRowBefore(r) { return this.insertRows(r, 1); }
  insertRowAfter(r) { return this.insertRows(r + 1, 1); }
  insertRowsAfter(r, n) { return this.insertRows(r + 1, n); }
  insertRowsBefore(r, n) { return this.insertRows(r, n); }
  _markAllFrom(from, removed) {
    const end = this._rows.length + (removed || 0) + 1;
    for (let r = Math.max(0, from); r < end; r++) this._touch(r);
  }
  deleteColumn(c) { this._load(); this._rows.forEach((row, r) => { if (row && row.length >= c) { this._own(r).splice(c - 1, 1); this._touch(r); } }); return this; }
  insertColumnAfter(c) { this._load(); this._rows.forEach((row, r) => { if (row && row.length > c) { this._own(r).splice(c, 0, ''); this._touch(r); } }); return this; }
  insertColumnBefore(c) { return this.insertColumnAfter(c - 1); }
  insertColumnsAfter(c, n) { for (let i = 0; i < n; i++) this.insertColumnAfter(c); return this; }
  clear() { this._load(); for (let r = 0; r < this._rows.length; r++) this._touch(r); this._rows = []; this._owned = new Set(); return this; }
  clearContents() { return this.clear(); }
  getSheetValues(r, c, nr, nc) { return this.getRange(r, c, nr, nc).getValues(); }
  sort(col, asc) { this.getRange(2, 1, Math.max(1, this.getLastRow() - 1), Math.max(1, this.getLastColumn())).sort({ column: col, ascending: asc !== false }); return this; }
  activate() { return this; }
  showSheet() { return this; }
  hideSheet() { return this; }
}

function colNum_(letters) { let n = 0; for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function colLetters_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function a1Range_(sh, a1) {
  const m = String(a1).replace(/^.*!/, '').match(/^([A-Z]+)?(\d+)?(?::([A-Z]+)?(\d+)?)?$/i);
  if (!m) throw new Error('Range not found: ' + a1);
  const c1 = m[1] ? colNum_(m[1]) : 1, r1 = m[2] ? Number(m[2]) : 1;
  const c2 = m[3] ? colNum_(m[3]) : (m[4] ? c1 : (m[1] && !m[2] ? c1 : Math.max(c1, sh.getLastColumn() || 1)));
  const r2 = m[4] ? Number(m[4]) : (m[3] || m[1] && !m[2] ? Math.max(r1, sh.getLastRow() || 1) : r1);
  const rg = new GRange_(sh, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  if (m[1] && !m[2]) rg._wholeColumn = true;
  return rg;
}

class GRange_ {
  constructor(sh, r, c, nr, nc) {
    this._sh = sh._self || sh; this._r = r; this._c = c; this._nr = nr; this._nc = nc;
    return chainable_(this);
  }
  getRow() { return this._r; }
  getColumn() { return this._c; }
  getLastRow() { return this._r + this._nr - 1; }
  getLastColumn() { return this._c + this._nc - 1; }
  getNumRows() { return this._nr; }
  getNumColumns() { return this._nc; }
  getSheet() { return this._sh; }
  getA1Notation() { return colLetters_(this._c) + this._r + (this._nr > 1 || this._nc > 1 ? ':' + colLetters_(this.getLastColumn()) + this.getLastRow() : ''); }
  getCell(r, c) { return new GRange_(this._sh, this._r + r - 1, this._c + c - 1, 1, 1); }
  offset(dr, dc, nr, nc) { return new GRange_(this._sh, this._r + dr, this._c + dc, nr || this._nr, nc || this._nc); }
  getValues() {
    this._sh._load();
    const out = new Array(this._nr);
    for (let i = 0; i < this._nr; i++) {
      const row = this._sh._rows[this._r - 1 + i], o = new Array(this._nc);
      for (let j = 0; j < this._nc; j++) o[j] = cellOut_(row ? row[this._c - 1 + j] : '');
      out[i] = o;
    }
    return out;
  }
  getDisplayValues() {
    const f = this._sh._meta.formats || {};
    return this.getValues().map(row => row.map((v, j) => displayOf_(v, f[this._c + j])));
  }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getFormulas() { return this.getValues().map(r => r.map(() => '')); }
  setValues(values) {
    if (!Array.isArray(values) || values.length !== this._nr) throw new Error('The number of rows in the data does not match the number of rows in the range. The data has ' + (values ? values.length : 0) + ' but the range has ' + this._nr + '.');
    this._sh._load();
    for (let i = 0; i < this._nr; i++) {
      const row = values[i];
      if (!Array.isArray(row) || row.length !== this._nc) throw new Error('The number of columns in the data does not match the number of columns in the range. The data has ' + (row ? row.length : 0) + ' but the range has ' + this._nc + '.');
      for (let j = 0; j < this._nc; j++) this._sh._set(this._r - 1 + i, this._c - 1 + j, cellIn_(row[j]));
    }
    return this;
  }
  setValue(v) {
    const grid = []; for (let i = 0; i < this._nr; i++) { const r = []; for (let j = 0; j < this._nc; j++) r.push(v); grid.push(r); }
    return this.setValues(grid);
  }
  setFormula(f) { return this.setValue(String(f)); }
  clearContent() {
    this._sh._load();
    const last = this._sh.getLastRow();
    for (let i = 0; i < this._nr; i++) {
      const r = this._r - 1 + i;
      if (r >= Math.max(last, this._sh._rows.length)) break;
      for (let j = 0; j < this._nc; j++) if (this._sh._rows[r] && !isBlank_(this._sh._rows[r][this._c - 1 + j])) this._sh._set(r, this._c - 1 + j, '');
    }
    return this;
  }
  clear() { return this.clearContent(); }
  clearFormat() { return this; }
  setNumberFormat(fmt) {
    const f = this._sh._meta.formats || (this._sh._meta.formats = {});
    for (let j = 0; j < this._nc; j++) f[this._c + j] = String(fmt);
    this._sh._metaDirty = true;
    return this;
  }
  setNumberFormats(grid) { if (grid && grid[0]) grid[0].forEach((fmt, j) => this.getCell(1, j + 1).setNumberFormat(fmt)); return this; }
  getNumberFormat() { return (this._sh._meta.formats || {})[this._c] || 'General'; }
  sort(spec) {
    const specs = (Array.isArray(spec) ? spec : [spec]).map(s => typeof s === 'number' ? { column: s, ascending: true } : s);
    const vals = this.getValues();
    const cmp = (a, b) => {
      for (const s of specs) {
        const i = s.column - this._c;
        let x = a[i], y = b[i];
        if (x instanceof Date) x = x.getTime(); if (y instanceof Date) y = y.getTime();
        if (x === y) continue;
        if (isBlank_(x)) return 1; if (isBlank_(y)) return -1;
        const r = (typeof x === 'number' && typeof y === 'number') ? x - y : String(x).localeCompare(String(y), undefined, { numeric: true });
        if (r) return s.ascending === false ? -r : r;
      }
      return 0;
    };
    this.setValues(vals.slice().sort(cmp));
    return this;
  }
  createTextFinder(text) {
    const range = this, opt = { whole: false, matchCase: false };
    let pos = -1, hits = null;
    const find = () => {
      if (hits) return hits;
      hits = [];
      const want = String(text);
      range.getDisplayValues().forEach((row, i) => row.forEach((v, j) => {
        const a = opt.matchCase ? v : v.toLowerCase(), b = opt.matchCase ? want : want.toLowerCase();
        if (opt.whole ? a === b : a.indexOf(b) >= 0) hits.push(new GRange_(range._sh, range._r + i, range._c + j, 1, 1));
      }));
      return hits;
    };
    const tf = {
      matchEntireCell(b) { opt.whole = !!b; hits = null; return tf; },
      matchCase(b) { opt.matchCase = !!b; hits = null; return tf; },
      useRegularExpression() { return tf; },
      ignoreDiacritics() { return tf; },
      findNext() { const h = find(); pos++; return pos < h.length ? h[pos] : null; },
      findAll() { return find().slice(); }
    };
    return tf;
  }
  getDataValidation() { return null; }
  getMergedRanges() { return []; }
  activate() { return this; }
}

// --------------------------------------------------------------------------------------------
// SpreadsheetApp
// --------------------------------------------------------------------------------------------
const SpreadsheetApp = new Proxy({
  getActiveSpreadsheet() { return RT.main; },
  getActive() { return RT.main; },
  openById(id) { if (id === 'main' || id === RT.main.getId()) return RT.main; const t = RT.temp.get(id); if (!t) throw new Error('Spreadsheet not found: ' + id); return t; },
  create(name) { const id = 'tmp_' + gasUuid_(); const ss = new GSpreadsheet_(id, String(name || 'Untitled'), false); ss.insertSheet('Sheet1'); RT.temp.set(id, ss); return ss; },
  flush() { shimFlush_(); },
  getUi() { throw new Error('The spreadsheet menu is not available on Cloudflare. Use the web app instead.'); },
  newDataValidation() { const b = new Proxy({}, { get: (t, p) => p === 'build' ? () => ({}) : () => b }); return b; },
  newRichTextValue() { const b = new Proxy({}, { get: () => () => b }); return b; }
}, { get(t, p) { if (p in t) return t[p]; return new Proxy({}, { get: (_, k) => String(k) }); } });   // enums (BorderStyle.SOLID …)

/** Moves sheet changes into the journal. The backend commits the journal at the end of the request. */
function shimFlush_() { if (RT.main) RT.main._stage(); }

// --------------------------------------------------------------------------------------------
// PropertiesService / CacheService / LockService
// --------------------------------------------------------------------------------------------
const ENV_PROPERTIES_ = { TRACKER_ADMIN_SETUP_KEY: 'ADMIN_SETUP_KEY' };
const scriptProperties_ = {
  getProperty(k) {
    k = String(k);
    const v = ST_.propGet(k);
    if (v !== null) return v;
    const envName = ENV_PROPERTIES_[k];
    return envName && RT.env[envName] ? String(RT.env[envName]) : null;
  },
  setProperty(k, v) { ST_.propSet(String(k), String(v)); return scriptProperties_; },
  deleteProperty(k) { ST_.propDel(String(k)); return scriptProperties_; },
  getProperties() { return ST_.propsAll(); },
  getKeys() { return Object.keys(ST_.propsAll()); },
  setProperties(obj, deleteAll) {
    if (deleteAll) scriptProperties_.deleteAllProperties();
    Object.keys(obj || {}).forEach(k => scriptProperties_.setProperty(k, obj[k]));
    return scriptProperties_;
  },
  deleteAllProperties() { ST_.propsClear(); return scriptProperties_; }
};
const PropertiesService = {
  getScriptProperties: () => scriptProperties_,
  getUserProperties: () => scriptProperties_,
  getDocumentProperties: () => scriptProperties_
};

const scriptCache_ = {
  get(k) { return ST_.cacheGet(String(k)); },
  put(k, v, ttl) { ST_.cachePut(String(k), String(v), Date.now() + Math.min(21600, Number(ttl) || 600) * 1000); },
  remove(k) { ST_.cacheDel(String(k)); },
  getAll(keys) { const o = {}; (keys || []).forEach(k => { const v = scriptCache_.get(k); if (v !== null) o[k] = v; }); return o; },
  putAll(obj, ttl) { Object.keys(obj || {}).forEach(k => scriptCache_.put(k, obj[k], ttl)); },
  removeAll(keys) { (keys || []).forEach(k => scriptCache_.remove(k)); }
};
const CacheService = { getScriptCache: () => scriptCache_, getUserCache: () => scriptCache_, getDocumentCache: () => scriptCache_ };

// A Durable Object handles one request at a time and Code.gs never waits on I/O, so a lock is always free.
function makeLock_() { let held = false; return { tryLock() { held = true; return true; }, waitLock() { held = true; }, hasLock() { return held; }, releaseLock() { held = false; } }; }
const LockService = { getScriptLock: makeLock_, getUserLock: makeLock_, getDocumentLock: makeLock_ };

// --------------------------------------------------------------------------------------------
// Utilities / Session / Logger
// --------------------------------------------------------------------------------------------
function gasUuid_() { return crypto.randomUUID(); }

const utf8Enc_ = new TextEncoder(), utf8Dec_ = new TextDecoder();
function toBytes_(data) {
  if (typeof data === 'string') return utf8Enc_.encode(data);
  if (data instanceof Uint8Array) return data;
  if (Array.isArray(data)) return Uint8Array.from(data, b => b & 255);
  if (data && typeof data.getBytes === 'function') return toBytes_(data.getBytes());
  return utf8Enc_.encode(String(data));
}
function toSigned_(u8) { const out = new Array(u8.length); for (let i = 0; i < u8.length; i++) out[i] = u8[i] > 127 ? u8[i] - 256 : u8[i]; return out; }

// SHA-256 (synchronous, like Utilities.computeDigest)
const SHA_K_ = new Uint32Array([0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2]);
function sha256_(bytes) {
  const len = bytes.length, total = ((len + 9 + 63) >> 6) << 6, m = new Uint8Array(total);
  m.set(bytes); m[len] = 0x80;
  const bits = len * 8, dv = new DataView(m.buffer);
  dv.setUint32(total - 8, Math.floor(bits / 0x100000000)); dv.setUint32(total - 4, bits >>> 0);
  const H = new Uint32Array([0x6a09e667,0xbb67ae85,0x3c6ef372,0xa54ff53a,0x510e527f,0x9b05688c,0x1f83d9ab,0x5be0cd19]), W = new Uint32Array(64);
  for (let off = 0; off < total; off += 64) {
    for (let i = 0; i < 16; i++) W[i] = dv.getUint32(off + i * 4);
    for (let i = 16; i < 64; i++) {
      const a = W[i - 15], b = W[i - 2];
      const s0 = ((a >>> 7) | (a << 25)) ^ ((a >>> 18) | (a << 14)) ^ (a >>> 3);
      const s1 = ((b >>> 17) | (b << 15)) ^ ((b >>> 19) | (b << 13)) ^ (b >>> 10);
      W[i] = (W[i - 16] + s0 + W[i - 7] + s1) >>> 0;
    }
    let a = H[0], b = H[1], c = H[2], d = H[3], e = H[4], f = H[5], g = H[6], h = H[7];
    for (let i = 0; i < 64; i++) {
      const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
      const t1 = (h + S1 + ((e & f) ^ (~e & g)) + SHA_K_[i] + W[i]) >>> 0;
      const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
      const t2 = (S0 + ((a & b) ^ (a & c) ^ (b & c))) >>> 0;
      h = g; g = f; f = e; e = (d + t1) >>> 0; d = c; c = b; b = a; a = (t1 + t2) >>> 0;
    }
    H[0] = (H[0] + a) >>> 0; H[1] = (H[1] + b) >>> 0; H[2] = (H[2] + c) >>> 0; H[3] = (H[3] + d) >>> 0;
    H[4] = (H[4] + e) >>> 0; H[5] = (H[5] + f) >>> 0; H[6] = (H[6] + g) >>> 0; H[7] = (H[7] + h) >>> 0;
  }
  const out = new Uint8Array(32), odv = new DataView(out.buffer);
  for (let i = 0; i < 8; i++) odv.setUint32(i * 4, H[i]);
  return out;
}

const B64_ = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
function b64enc_(u8) {
  let s = '', i = 0;
  for (; i + 2 < u8.length; i += 3) { const n = (u8[i] << 16) | (u8[i + 1] << 8) | u8[i + 2]; s += B64_[n >> 18 & 63] + B64_[n >> 12 & 63] + B64_[n >> 6 & 63] + B64_[n & 63]; }
  if (i < u8.length) { const n = (u8[i] << 16) | ((u8[i + 1] || 0) << 8); s += B64_[n >> 18 & 63] + B64_[n >> 12 & 63] + (i + 1 < u8.length ? B64_[n >> 6 & 63] : '=') + '='; }
  return s;
}
function b64dec_(s) {
  s = String(s).replace(/-/g, '+').replace(/_/g, '/').replace(/[^A-Za-z0-9+/]/g, '');
  const out = []; let buf = 0, bits = 0;
  for (const ch of s) { buf = (buf << 6) | B64_.indexOf(ch); bits += 6; if (bits >= 8) { bits -= 8; out.push((buf >> bits) & 255); } }
  return Uint8Array.from(out);
}

class GBlob_ {
  constructor(bytes, type, name) { this._b = toBytes_(bytes == null ? '' : bytes); this._type = type || 'application/octet-stream'; this._name = name || ''; }
  getBytes() { return toSigned_(this._b); }
  _raw() { return this._b; }
  getDataAsString() { return utf8Dec_.decode(this._b); }
  getContentType() { return this._type; }
  setContentType(t) { this._type = t; return this; }
  getName() { return this._name; }
  setName(n) { this._name = String(n); return this; }
  copyBlob() { return new GBlob_(this._b.slice(), this._type, this._name); }
  getBlob() { return this; }
}

const MONTHS_ = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const DAYS_ = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
function tzParts_(date, tz) {
  let zone = String(tz || 'UTC');
  const m = zone.match(/^(?:GMT|UTC)([+-])(\d{1,2})(?::?(\d{2}))?$/i);
  if (m) {   // "GMT+2" (Apps Script style: real offset) -> shift manually
    const off = (m[1] === '-' ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3] || 0));
    const d = new Date(date.getTime() + off * 60000);
    return { y: d.getUTCFullYear(), M: d.getUTCMonth() + 1, d: d.getUTCDate(), H: d.getUTCHours(), m: d.getUTCMinutes(), s: d.getUTCSeconds(), wd: d.getUTCDay(), off };
  }
  let f;
  try { f = new Intl.DateTimeFormat('en-US', { timeZone: zone, hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric', weekday: 'short' }); }
  catch (_) { f = new Intl.DateTimeFormat('en-US', { timeZone: 'UTC', hourCycle: 'h23', year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric', weekday: 'short' }); }
  const p = {}; f.formatToParts(date).forEach(x => p[x.type] = x.value);
  const r = { y: +p.year, M: +p.month, d: +p.day, H: +p.hour % 24, m: +p.minute, s: +p.second, wd: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].indexOf(p.weekday) };
  r.off = Math.round((Date.UTC(r.y, r.M - 1, r.d, r.H, r.m, r.s) - Math.floor(date.getTime() / 1000) * 1000) / 60000);
  return r;
}
/** Utilities.formatDate with Java SimpleDateFormat patterns (yyyy MM MMM dd d HH H hh h mm ss a EEE EEEE z Z X '…'). */
function gasFormatDate_(date, tz, pattern) {
  if (!(date instanceof Date)) date = new Date(date);
  if (isNaN(date.getTime())) throw new Error('Invalid date for formatDate.');
  const p = tzParts_(date, tz), pad = (n, w) => String(n).padStart(w, '0');
  const offStr = (sep) => { const o = p.off, s = o < 0 ? '-' : '+', a = Math.abs(o); return s + pad(Math.floor(a / 60), 2) + sep + pad(a % 60, 2); };
  return String(pattern).replace(/'([^']*)'|y+|M+|d+|H+|h+|m+|s+|S+|a+|E+|z+|Z+|X+|u+/g, (t, lit) => {
    if (lit !== undefined) return lit === '' ? "'" : lit;
    const n = t.length;
    switch (t[0]) {
      case 'y': return n === 2 ? pad(p.y % 100, 2) : pad(p.y, n);
      case 'M': return n >= 4 ? MONTHS_[p.M - 1] : n === 3 ? MONTHS_[p.M - 1].slice(0, 3) : pad(p.M, n);
      case 'd': return pad(p.d, n);
      case 'H': return pad(p.H, n);
      case 'h': return pad(p.H % 12 || 12, n);
      case 'm': return pad(p.m, n);
      case 's': return pad(p.s, n);
      case 'S': return pad(date.getUTCMilliseconds(), 3).slice(0, n);
      case 'a': return p.H < 12 ? 'AM' : 'PM';
      case 'E': return n >= 4 ? DAYS_[p.wd] : DAYS_[p.wd].slice(0, 3);
      case 'u': return String(p.wd || 7);
      case 'z': return 'GMT' + offStr(':');
      case 'Z': return offStr('');
      case 'X': return p.off === 0 ? 'Z' : offStr(n >= 3 ? ':' : '');
    }
    return t;
  });
}

const Utilities = {
  DigestAlgorithm: { SHA_256: 'SHA_256', MD5: 'MD5', SHA_1: 'SHA_1', SHA_384: 'SHA_384', SHA_512: 'SHA_512', MD2: 'MD2' },
  Charset: { UTF_8: 'UTF_8', US_ASCII: 'US_ASCII' },
  MacAlgorithm: { HMAC_SHA_256: 'HMAC_SHA_256' },
  getUuid: gasUuid_,
  computeDigest(alg, data) {
    if (alg !== 'SHA_256') throw new Error('Only SHA_256 digests are available on Cloudflare.');
    return toSigned_(sha256_(toBytes_(data)));
  },
  base64Encode(data) { return b64enc_(toBytes_(data)); },
  base64EncodeWebSafe(data) { return b64enc_(toBytes_(data)).replace(/\+/g, '-').replace(/\//g, '_'); },
  base64Decode(s) { return toSigned_(b64dec_(s)); },
  base64DecodeWebSafe(s) { return toSigned_(b64dec_(s)); },
  newBlob(data, type, name) { return new GBlob_(data, type, name); },
  formatDate: gasFormatDate_,
  formatString(fmt, ...args) { let i = 0; return String(fmt).replace(/%[sd]/g, () => String(args[i++])); },
  sleep() {},
  parseCsv(text) {
    const rows = []; let row = [], cell = '', q = false; text = String(text);
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  },
  jsonStringify: o => JSON.stringify(o),
  jsonParse: s => JSON.parse(s)
};

const Session = {
  getScriptTimeZone: () => RT.tz,
  getActiveUser: () => ({ getEmail: () => '' }),
  getEffectiveUser: () => ({ getEmail: () => String(RT.env.MAIL_FROM || '') }),
  getTemporaryActiveUserKey: () => ''
};
const Logger = { log(...a) { console.log(...a); return Logger; }, clear() {}, getLog: () => '' };

// --------------------------------------------------------------------------------------------
// HtmlService / ContentService (pages are built into the worker at deploy time)
// --------------------------------------------------------------------------------------------
class GHtmlOutput_ {
  constructor(html) { this._html = String(html || ''); this._title = ''; this._meta = []; this._w = 0; this._h = 0; }
  getContent() { return this._html; }
  setContent(h) { this._html = String(h); return this; }
  append(h) { this._html += String(h); return this; }
  setTitle(t) { this._title = String(t); return this; }
  getTitle() { return this._title; }
  addMetaTag(name, content) { this._meta.push([name, content]); return this; }
  setXFrameOptionsMode() { return this; }
  setSandboxMode() { return this; }
  setFaviconUrl() { return this; }
  setWidth(w) { this._w = w; return this; }
  setHeight(h) { this._h = h; return this; }
  getAs() { return new GBlob_(this._html, 'text/html'); }
}
const HtmlService = {
  XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
  SandboxMode: { IFRAME: 'IFRAME', NATIVE: 'NATIVE' },
  createHtmlOutput(html) { return new GHtmlOutput_(html && html.getContent ? html.getContent() : html); },
  createHtmlOutputFromFile(name) {
    const f = HTML_FILES_[String(name).replace(/\.html$/, '')];
    if (f === undefined) throw new Error('No HTML file named ' + name + ' was found.');
    return new GHtmlOutput_(f);
  },
  createTemplateFromFile(name) {
    const out = HtmlService.createHtmlOutputFromFile(name);
    return { evaluate: () => out, getRawContent: () => out.getContent() };
  },
  createTemplate(html) { const out = new GHtmlOutput_(html); return { evaluate: () => out }; }
};

class GTextOutput_ {
  constructor(s) { this._s = String(s == null ? '' : s); this._mime = 'text/plain'; }
  getContent() { return this._s; }
  setContent(s) { this._s = String(s); return this; }
  append(s) { this._s += String(s); return this; }
  setMimeType(m) { this._mime = m; return this; }
  getMimeType() { return this._mime; }
  downloadAsFile() { return this; }
}
const ContentService = {
  MimeType: { JSON: 'application/json', TEXT: 'text/plain', JAVASCRIPT: 'application/javascript', CSV: 'text/csv', XML: 'application/xml', ICAL: 'text/calendar', RSS: 'application/rss+xml', ATOM: 'application/atom+xml', VCARD: 'text/vcard' },
  createTextOutput(s) { return new GTextOutput_(s); }
};

// --------------------------------------------------------------------------------------------
// ScriptApp: time triggers are kept in the "__TRIGGERS" property and run by the Worker cron (every minute).
// --------------------------------------------------------------------------------------------
function readTriggers_() { try { return JSON.parse(scriptProperties_.getProperty('__TRIGGERS') || '[]'); } catch (_) { return []; } }
function writeTriggers_(list) { scriptProperties_.setProperty('__TRIGGERS', JSON.stringify(list)); }
function triggerObj_(t) {
  return { getHandlerFunction: () => t.handler, getUniqueId: () => t.id, getEventType: () => t.event || 'CLOCK',
    getTriggerSource: () => t.event && t.event !== 'CLOCK' ? 'SPREADSHEETS' : 'CLOCK', getTriggerSourceId: () => 'main' };
}
const ScriptApp = new Proxy({
  getProjectTriggers() { return readTriggers_().map(triggerObj_); },
  getUserTriggers() { return ScriptApp.getProjectTriggers(); },
  deleteTrigger(t) { const id = t && t.getUniqueId ? t.getUniqueId() : t; writeTriggers_(readTriggers_().filter(x => x.id !== id)); },
  newTrigger(handler) {
    const t = { id: gasUuid_(), handler: String(handler), event: 'CLOCK', every: 0, unit: '', atHour: null, at: 0, last: Date.now(), ranDay: 0 };
    const b = {
      timeBased() { return b; },
      forSpreadsheet() { return b; },
      forDocument() { return b; },
      onChange() { t.event = 'ON_CHANGE'; return b; }, onEdit() { t.event = 'ON_EDIT'; return b; },
      onOpen() { t.event = 'ON_OPEN'; return b; }, onFormSubmit() { t.event = 'ON_FORM_SUBMIT'; return b; },
      everyMinutes(n) { t.every = n; t.unit = 'm'; return b; },
      everyHours(n) { t.every = n; t.unit = 'h'; return b; },
      everyDays(n) { t.every = n; t.unit = 'd'; return b; },
      everyWeeks(n) { t.every = n * 7; t.unit = 'd'; return b; },
      atHour(h) { t.atHour = Number(h); return b; },
      nearMinute(m) { t.nearMinute = Number(m) || 0; return b; },
      onWeekDay(d) { t.weekDay = d; return b; },
      onMonthDay(d) { t.monthDay = d; return b; },
      inTimezone() { return b; },
      after(ms) { t.at = Date.now() + Number(ms); t.unit = 'once'; return b; },
      at(date) { t.at = new Date(date).getTime(); t.unit = 'once'; return b; },
      create() { const list = readTriggers_(); list.push(t); writeTriggers_(list); return triggerObj_(t); }
    };
    return b;
  },
  getOAuthToken() { return ''; },
  getService() { return { getUrl: () => RT.origin + '/', isEnabled: () => true }; },
  getScriptId() { return 'cloudflare'; },
  invalidateAuth() {},
  requireScopes() {},
  AuthMode: { FULL: 'FULL', LIMITED: 'LIMITED', NONE: 'NONE' },
  EventType: { CLOCK: 'CLOCK', ON_CHANGE: 'ON_CHANGE', ON_EDIT: 'ON_EDIT', ON_OPEN: 'ON_OPEN' },
  TriggerSource: { CLOCK: 'CLOCK', SPREADSHEETS: 'SPREADSHEETS' },
  WeekDay: { MONDAY: 1, TUESDAY: 2, WEDNESDAY: 3, THURSDAY: 4, FRIDAY: 5, SATURDAY: 6, SUNDAY: 0 }
}, { get(t, p) { return p in t ? t[p] : undefined; } });

/** Called by the cron every minute: runs the time triggers that are due. Returns the handler names that ran. */
function runDueTriggers_(now, call) {
  const ran = [];
  const p = tzParts_(new Date(now), RT.tz), today = p.y * 10000 + p.M * 100 + p.d;
  for (const t of readTriggers_()) {
    if (t.event !== 'CLOCK') continue;
    let due = false;
    if (t.unit === 'once') due = now >= t.at;
    else if (t.unit === 'm') due = now - (t.last || 0) >= t.every * 60000 - 15000;
    else if (t.unit === 'h') due = now - (t.last || 0) >= t.every * 3600000 - 30000;
    else if (t.unit === 'd') due = (t.atHour == null || p.H === t.atHour) && t.ranDay !== today &&
      (t.every <= 1 || !t.ranDay || now - (t.last || 0) >= (t.every - 1) * 86400000 - 3600000);
    if (!due) continue;
    // Save "ran" before running (re-read: an earlier job may have changed the trigger list), so a failing job
    // is not retried every minute.
    const fresh = readTriggers_(), mine = fresh.find(x => x.id === t.id);
    if (!mine) continue;
    if (t.unit === 'once') fresh.splice(fresh.indexOf(mine), 1); else { mine.last = now; mine.ranDay = today; }
    writeTriggers_(fresh);
    try { call(t.handler, { triggerUid: t.id }); ran.push(t.handler); }
    catch (err) { console.error('Trigger ' + t.handler + ' failed: ' + (err && err.stack || err)); }
  }
  return ran;
}

// --------------------------------------------------------------------------------------------
// DriveApp: backup copies of the tracker are kept in SQLite (split in parts), downloadable as .xlsx.
// --------------------------------------------------------------------------------------------
function snapshotOf_(ss) {
  return JSON.stringify(ss.getSheets().map(sh => {
    const s = sh._self || sh; s._load();
    return { name: s._name, rows: s._rows.map(r => r ? JSON.parse(encRow_(r)) : []) };
  }));
}
function saveFile_(id, name, folder, mime, text) { ST_.fileSave({ id, name, folder, created: Date.now(), mime }, text); }
function readFileText_(id) { return ST_.fileText(id); }
function fileRow_(id) { return ST_.fileMeta(id); }
function driveFile_(row) {
  return {
    getId: () => row.id, getName: () => row.name, setName(n) { ST_.fileRename(row.id, String(n)); row.name = String(n); return this; },
    getUrl: () => RT.origin + '/files/' + row.id + '.xlsx', getDownloadUrl: () => RT.origin + '/files/' + row.id + '.xlsx',
    getDateCreated: () => new Date(row.created), getLastUpdated: () => new Date(row.created), getMimeType: () => row.mime, getSize: () => 0,
    setTrashed(b) { if (b) ST_.fileDelete(row.id); return this; },
    isTrashed: () => false,
    makeCopy() { throw new Error('Only the tracker spreadsheet itself can be copied.'); },
    getBlob: () => new GBlob_(readFileText_(row.id), row.mime, row.name)
  };
}
function iterator_(arr) { let i = 0; return { hasNext: () => i < arr.length, next: () => { if (i >= arr.length) throw new Error('No more items.'); return arr[i++]; } }; }
function driveFolder_(name) {
  return {
    getId: () => 'folder:' + name, getName: () => name, getUrl: () => RT.origin + '/backups',
    getFiles: () => iterator_(ST_.fileList(name).map(driveFile_)),
    getFilesByName: n => iterator_(ST_.fileList(name, String(n)).map(driveFile_)),
    createFile(blobOrName, content, mime) {
      const blob = typeof blobOrName === 'string' ? new GBlob_(content || '', mime || 'text/plain', blobOrName) : blobOrName;
      const id = gasUuid_(); saveFile_(id, blob.getName() || 'file', name, blob.getContentType(), JSON.stringify({ blob: b64enc_(toBytes_(blob.getBytes())) }));
      return driveFile_(fileRow_(id));
    }
  };
}
function folderNames_() { try { return JSON.parse(scriptProperties_.getProperty('__FOLDERS') || '[]'); } catch (_) { return []; } }
const DriveApp = {
  getFileById(id) {
    if (id === 'main' || (RT.main && id === RT.main.getId())) {
      return {
        getId: () => 'main', getName: () => RT.main.getName(), getUrl: () => RT.origin + '/',
        makeCopy(name, folder) {
          shimFlush_();
          const fid = gasUuid_(), fname = String(name || 'Copy of tracker');
          saveFile_(fid, fname, folder && folder.getName ? folder.getName() : '', 'application/x-tracker-snapshot', snapshotOf_(RT.main));
          return driveFile_(fileRow_(fid));
        },
        setTrashed() { throw new Error('The tracker itself cannot be deleted.'); }
      };
    }
    if (RT.temp.has(id)) return { getId: () => id, setTrashed() { RT.temp.delete(id); return this; }, getUrl: () => '' };
    const row = fileRow_(id);
    if (!row) throw new Error('File not found: ' + id);
    return driveFile_(row);
  },
  getFoldersByName(name) { return iterator_(folderNames_().includes(String(name)) ? [driveFolder_(String(name))] : []); },
  createFolder(name) { const list = folderNames_(); if (!list.includes(String(name))) { list.push(String(name)); scriptProperties_.setProperty('__FOLDERS', JSON.stringify(list)); } return driveFolder_(String(name)); },
  getRootFolder() { return driveFolder_(''); },
  getFolderById(id) { return driveFolder_(String(id).replace(/^folder:/, '')); },
  createFile(blob) { return driveFolder_('').createFile(blob); }
};

// --------------------------------------------------------------------------------------------
// MailApp: sends through Resend (https://resend.com) when RESEND_API_KEY and MAIL_FROM are set.
// Without them the quota is 0, so Code.gs reports "email not available" and sends nothing.
// --------------------------------------------------------------------------------------------
const MailApp = {
  getRemainingDailyQuota() { return RT.env.RESEND_API_KEY && RT.env.MAIL_FROM ? 100 : 0; },
  sendEmail(a, subject, body, opts) {
    const m = typeof a === 'object' && a !== null ? Object.assign({}, a) : Object.assign({ to: a, subject, body }, opts || {});
    if (!(RT.env.RESEND_API_KEY && RT.env.MAIL_FROM)) throw new Error('Email is not set up: add RESEND_API_KEY and MAIL_FROM to the Worker.');
    RT.outbox.push(m);
  }
};
const GmailApp = { sendEmail: (to, subject, body, opts) => MailApp.sendEmail(to, subject, body, opts) };

async function sendOutbox_(list) {
  list = list || RT.outbox.splice(0);
  for (const m of list) {
    const payload = {
      from: (m.name ? m.name + ' <' + RT.env.MAIL_FROM + '>' : RT.env.MAIL_FROM),
      to: String(m.to || '').split(/[,;]\s*/).filter(Boolean),
      subject: String(m.subject || ''), text: String(m.body || ''),
    };
    if (m.htmlBody) payload.html = String(m.htmlBody);
    if (m.cc) payload.cc = String(m.cc).split(/[,;]\s*/).filter(Boolean);
    if (m.replyTo) payload.reply_to = m.replyTo;
    if (m.attachments && m.attachments.length) payload.attachments = m.attachments.map(b => ({ filename: b.getName() || 'attachment', content: b64enc_(toBytes_(b.getBytes())) }));
    try {
      const r = await fetch('https://api.resend.com/emails', { method: 'POST', headers: { Authorization: 'Bearer ' + RT.env.RESEND_API_KEY, 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      if (!r.ok) console.error('Email to ' + payload.to + ' failed: HTTP ' + r.status + ' ' + (await r.text()).slice(0, 300));
    } catch (err) { console.error('Email failed: ' + err); }
  }
}

// --------------------------------------------------------------------------------------------
// UrlFetchApp: only the Google Sheets "export as .xlsx" link Code.gs uses is supported.
// --------------------------------------------------------------------------------------------
const UrlFetchApp = {
  fetch(url) {
    const m = String(url).match(/\/spreadsheets\/d\/([^/]+)\/export\?format=xlsx/);
    if (!m) throw new Error('UrlFetchApp is not available on Cloudflare (' + String(url).slice(0, 80) + ').');
    const ss = SpreadsheetApp.openById(decodeURIComponent(m[1]));
    const bytes = xlsxFromSheets_(ss.getSheets().map(sh => { const s = sh._self || sh; s._load(); return { name: s._name, rows: s._rows, formats: s._meta.formats || {} }; }));
    const blob = new GBlob_(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    return { getResponseCode: () => 200, getBlob: () => blob, getContent: () => blob.getBytes(), getContentText: () => '', getHeaders: () => ({}) };
  },
  fetchAll(reqs) { return reqs.map(r => UrlFetchApp.fetch(typeof r === 'string' ? r : r.url)); }
};

// --------------------------------------------------------------------------------------------
// Minimal .xlsx writer (stored zip, inline strings). Dates are written as text in the tracker time zone.
// --------------------------------------------------------------------------------------------
const CRC_TABLE_ = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32_(u8) { let c = 0xffffffff; for (let i = 0; i < u8.length; i++) c = CRC_TABLE_[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function zipStore_(files) {
  const chunks = [], central = []; let offset = 0;
  const u16 = n => [n & 255, (n >>> 8) & 255], u32 = n => [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255];
  for (const f of files) {
    const name = utf8Enc_.encode(f.name), data = typeof f.data === 'string' ? utf8Enc_.encode(f.data) : f.data, crc = crc32_(data);
    const head = Uint8Array.from([...u32(0x04034b50), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0)]);
    chunks.push(head, name, data);
    central.push(Uint8Array.from([...u32(0x02014b50), ...u16(20), ...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(data.length), ...u32(data.length), ...u16(name.length), ...u16(0), ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(offset)]), name);
    offset += head.length + name.length + data.length;
  }
  const cdSize = central.reduce((n, c) => n + c.length, 0);
  const end = Uint8Array.from([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(files.length), ...u16(files.length), ...u32(cdSize), ...u32(offset), ...u16(0)]);
  const all = chunks.concat(central, [end]), total = all.reduce((n, c) => n + c.length, 0), out = new Uint8Array(total);
  let p = 0; for (const c of all) { out.set(c, p); p += c.length; }
  return out;
}
function xmlEsc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, ''); }
function xlsxFromSheets_(sheets) {
  const used = new Set();
  const safeName = n => { let s = String(n || 'Sheet').replace(/[\\\/?*\[\]:]/g, ' ').slice(0, 31) || 'Sheet', base = s, i = 2; while (used.has(s.toLowerCase())) s = (base.slice(0, 28) + ' ' + i++); used.add(s.toLowerCase()); return s; };
  const list = sheets.length ? sheets : [{ name: 'Sheet1', rows: [] }];
  const files = [];
  const sheetXml = list.map((sh, si) => {
    const rows = sh.rows || [], fm = sh.formats || {}; let x = '';
    rows.forEach((row, r) => {
      if (!row || !row.length) return;
      let cells = '';
      row.forEach((v, c) => {
        if (isBlank_(v)) return;
        const ref = colLetters_(c + 1) + (r + 1), style = r === 0 ? ' s="1"' : '';
        if (typeof v === 'number' && isFinite(v) && fm[c + 1] !== '@') cells += '<c r="' + ref + '"' + style + '><v>' + v + '</v></c>';
        else if (typeof v === 'boolean') cells += '<c r="' + ref + '"' + style + ' t="b"><v>' + (v ? 1 : 0) + '</v></c>';
        else cells += '<c r="' + ref + '"' + style + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc_(v instanceof Date ? displayOf_(v, fm[c + 1]) : v) + '</t></is></c>';
      });
      x += '<row r="' + (r + 1) + '">' + cells + '</row>';
    });
    const maxC = rows.reduce((n, r) => Math.max(n, r ? r.length : 0), 1), dim = 'A1:' + colLetters_(maxC) + Math.max(1, rows.length);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="' + dim + '"/>' +
      (rows.length > 1 ? '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' : '') +
      '<sheetData>' + x + '</sheetData></worksheet>';
  });
  files.push({ name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    list.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>' });
  files.push({ name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' });
  files.push({ name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
    list.map((sh, i) => '<sheet name="' + xmlEsc_(safeName(sh.name)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>' });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    list.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
    '<Relationship Id="rId' + (list.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' });
  files.push({ name: 'xl/styles.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF26174A"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' });
  sheetXml.forEach((x, i) => files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: x }));
  return zipStore_(files);
}
/** Backup snapshot (DriveApp.makeCopy) -> .xlsx bytes, for the /files/<id>.xlsx download. */
function xlsxFromSnapshot_(text) {
  const sheets = JSON.parse(text).map(s => ({ name: s.name, rows: s.rows.map(r => (r || []).map(v => (v && typeof v === 'object' && '$d' in v) ? new Date(v.$d) : v)) }));
  return xlsxFromSheets_(sheets);
}
