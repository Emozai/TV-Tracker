// ============================================================================================
// Storage layer shared by both Cloudflare versions.
//   M  = committed data, kept in memory between requests (decoded rows, properties, cache, file list)
//   J  = the current request's changes; the backend commits J and then applies it to M
// Backends:
//   DO (Durable Object, src/entry.js)  – synchronous SQLite, rows load on demand, commit at the end
//   D1 (paste-in worker, src/entry-d1.js) – async; data is synced before the request runs, missing
//      sheets throw NeedData_ and the request is re-run once they are loaded; commit is one guarded batch
// ============================================================================================

class NeedData_ extends Error {
  constructor(what) { super('Loading ' + JSON.stringify(what)); this.name = 'NeedData_'; this.what = what; }
}

const M = {
  ready: false, seq: 0, floor: 0,
  sheets: new Map(),     // name -> { pos, meta }
  rows: new Map(),       // name -> array of decoded rows (only sheets that were loaded)
  props: new Map(),      // k -> v
  cache: new Map(),      // k -> { v, exp }
  files: new Map(),      // id -> { id, name, folder, created, mime, parts }
  fileText: new Map()    // id -> text (only files that were read)
};
const BACKEND = { syncLoadSheet: null, syncLoadFile: null };   // DO fills these; D1 leaves them null
let J = null;

function journalNew_() {
  return { sheets: new Map(), drop: new Set(), rows: new Map(), props: new Map(), propsClear: false,
    cache: new Map(), files: new Map(), fileText: new Map() };
}
function journalEmpty_(j) {
  return !j.sheets.size && !j.drop.size && !j.rows.size && !j.props.size && !j.propsClear && !j.cache.size && !j.files.size;
}

const ST_ = {
  sheetList() { return Array.from(M.sheets.entries()).sort((a, b) => a[1].pos - b[1].pos).map(([name, s]) => ({ name, meta: s.meta })); },
  sheetRows(name) {
    if (!M.rows.has(name)) {
      if (!M.sheets.has(name)) return [];
      if (BACKEND.syncLoadSheet) BACKEND.syncLoadSheet(name);
      else { RT.needed.push({ sheet: name }); throw new NeedData_({ sheet: name }); }
    }
    return M.rows.get(name) || [];
  },
  propGet(k) {
    if (J.props.has(k)) return J.props.get(k);
    if (J.propsClear) return null;
    return M.props.has(k) ? M.props.get(k) : null;
  },
  propSet(k, v) { J.props.set(k, v); },
  propDel(k) { J.props.set(k, null); },
  propsAll() {
    const o = {};
    if (!J.propsClear) M.props.forEach((v, k) => o[k] = v);
    J.props.forEach((v, k) => { if (v === null) delete o[k]; else o[k] = v; });
    return o;
  },
  propsClear() { J.propsClear = true; J.props.clear(); },
  cacheGet(k) {
    const e = J.cache.has(k) ? J.cache.get(k) : M.cache.get(k);
    if (!e || e.v === null) return null;
    if (e.exp < Date.now()) return null;
    return e.v;
  },
  cachePut(k, v, exp) { J.cache.set(k, { v, exp }); },
  cacheDel(k) { J.cache.set(k, { v: null, exp: 0 }); },
  fileMeta(id) { if (J.files.has(id)) return J.files.get(id); return M.files.get(id) || null; },
  fileList(folder, name) {
    const all = new Map(M.files);
    J.files.forEach((m, id) => { if (m === null) all.delete(id); else all.set(id, m); });
    return Array.from(all.values()).filter(f => f.folder === folder && (name === undefined || f.name === name)).sort((a, b) => a.created - b.created);
  },
  fileSave(meta, text) { J.files.set(meta.id, meta); J.fileText.set(meta.id, text); },
  fileRename(id, name) { const m = ST_.fileMeta(id); if (m) J.files.set(id, Object.assign({}, m, { name })); },
  fileDelete(id) { J.files.set(id, null); J.fileText.delete(id); },
  fileText(id) {
    if (J.fileText.has(id)) return J.fileText.get(id);
    if (!M.fileText.has(id)) {
      if (BACKEND.syncLoadFile) BACKEND.syncLoadFile(id);
      else { RT.needed.push({ file: id }); throw new NeedData_({ file: id }); }
    }
    return M.fileText.get(id) || '';
  }
};

/** After a successful commit: make the request's changes part of M. */
function applyJournal_(j) {
  j.drop.forEach(name => { M.sheets.delete(name); M.rows.delete(name); });
  j.sheets.forEach((s, name) => { M.sheets.set(name, s); if (!M.rows.has(name) && j.drop.has(name)) M.rows.set(name, []); });
  j.rows.forEach((rows, name) => {
    if (!M.rows.has(name)) { if (!M.sheets.has(name)) return; M.rows.set(name, []); }
    const arr = M.rows.get(name);
    rows.forEach((row, r) => { if (row === null) { if (r < arr.length) arr[r] = []; } else { while (arr.length < r) arr.push([]); arr[r] = row; } });
    let n = arr.length; while (n > 0 && !(arr[n - 1] && arr[n - 1].length)) n--; arr.length = n;
  });
  if (j.propsClear) M.props.clear();
  j.props.forEach((v, k) => { if (v === null) M.props.delete(k); else M.props.set(k, v); });
  j.cache.forEach((e, k) => { if (e.v === null) M.cache.delete(k); else M.cache.set(k, e); });
  j.files.forEach((m, id) => { if (m === null) { M.files.delete(id); M.fileText.delete(id); } else M.files.set(id, m); });
  j.fileText.forEach((t, id) => { if (M.files.has(id)) M.fileText.set(id, t); });
}
