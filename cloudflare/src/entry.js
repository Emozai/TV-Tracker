// ============================================================================================
// Cloudflare Worker entry (wrangler deploy). All tracker data lives in ONE Durable Object ("main")
// with SQLite storage: synchronous, one request at a time, so nothing can clash.
// ============================================================================================

function doSchema_(sql) {
  sql.exec(`CREATE TABLE IF NOT EXISTS gs_props (k TEXT PRIMARY KEY, v TEXT)`);
  sql.exec(`CREATE TABLE IF NOT EXISTS gs_sheets (name TEXT PRIMARY KEY, pos INTEGER, meta TEXT)`);
  sql.exec(`CREATE TABLE IF NOT EXISTS gs_rows (sheet TEXT, r INTEGER, data TEXT, PRIMARY KEY (sheet, r))`);
  sql.exec(`CREATE TABLE IF NOT EXISTS gs_files (id TEXT PRIMARY KEY, name TEXT, folder TEXT, created INTEGER, mime TEXT, parts INTEGER)`);
  sql.exec(`CREATE TABLE IF NOT EXISTS gs_file_parts (id TEXT, i INTEGER, data TEXT, PRIMARY KEY (id, i))`);
}
function doColdLoad_(sql) {
  M.sheets.clear(); M.rows.clear(); M.props.clear(); M.files.clear(); M.fileText.clear();
  for (const r of sql.exec('SELECT name, pos, meta FROM gs_sheets').toArray()) M.sheets.set(r.name, { pos: r.pos, meta: JSON.parse(r.meta || '{}') });
  for (const r of sql.exec('SELECT k, v FROM gs_props').toArray()) M.props.set(r.k, r.v);
  for (const r of sql.exec('SELECT id, name, folder, created, mime FROM gs_files').toArray()) M.files.set(r.id, { id: r.id, name: r.name, folder: r.folder, created: r.created, mime: r.mime });
  BACKEND.syncLoadSheet = name => {
    const arr = [];
    for (const r of sql.exec('SELECT r, data FROM gs_rows WHERE sheet = ? ORDER BY r', name).toArray()) arr[r.r] = decRow_(r.data);
    for (let i = 0; i < arr.length; i++) if (!arr[i]) arr[i] = [];
    M.rows.set(name, arr);
  };
  BACKEND.syncLoadFile = id => M.fileText.set(id, sql.exec('SELECT data FROM gs_file_parts WHERE id = ? ORDER BY i', id).toArray().map(r => r.data).join(''));
  M.ready = true;
}
const DO_PART_SIZE_ = 900000;
function doCommit_(sql, j) {
  j.drop.forEach(n => { sql.exec('DELETE FROM gs_rows WHERE sheet = ?', n); sql.exec('DELETE FROM gs_sheets WHERE name = ?', n); });
  j.sheets.forEach((s, n) => sql.exec('INSERT OR REPLACE INTO gs_sheets (name, pos, meta) VALUES (?, ?, ?)', n, s.pos, JSON.stringify(s.meta)));
  j.rows.forEach((rows, n) => rows.forEach((row, r) => {
    if (row === null) sql.exec('DELETE FROM gs_rows WHERE sheet = ? AND r = ?', n, r);
    else sql.exec('INSERT OR REPLACE INTO gs_rows (sheet, r, data) VALUES (?, ?, ?)', n, r, encRow_(row));
  }));
  if (j.propsClear) sql.exec('DELETE FROM gs_props');
  j.props.forEach((v, k) => { if (v === null) sql.exec('DELETE FROM gs_props WHERE k = ?', k); else sql.exec('INSERT OR REPLACE INTO gs_props (k, v) VALUES (?, ?)', k, v); });
  j.files.forEach((m, id) => {
    if (m === null) { sql.exec('DELETE FROM gs_file_parts WHERE id = ?', id); sql.exec('DELETE FROM gs_files WHERE id = ?', id); return; }
    sql.exec('INSERT OR REPLACE INTO gs_files (id, name, folder, created, mime, parts) VALUES (?, ?, ?, ?, ?, 0)', id, m.name, m.folder, m.created, m.mime);
  });
  j.fileText.forEach((t, id) => {
    if (!j.files.get(id)) return;
    sql.exec('DELETE FROM gs_file_parts WHERE id = ?', id);
    const parts = Math.max(1, Math.ceil(t.length / DO_PART_SIZE_));
    for (let i = 0; i < parts; i++) sql.exec('INSERT INTO gs_file_parts (id, i, data) VALUES (?, ?, ?)', id, i, t.slice(i * DO_PART_SIZE_, (i + 1) * DO_PART_SIZE_));
  });
  applyJournal_(j);   // cache entries live in memory only (this object is the only instance)
}

export class TrackerStore extends DurableObject {
  constructor(ctx, env) { super(ctx, env); this.ctx = ctx; this.env = env; doSchema_(ctx.storage.sql); doColdLoad_(ctx.storage.sql); }

  async fetch(request) {
    const url = new URL(request.url);
    const body = request.method === 'POST' ? await request.text() : null;   // read before any tracker code runs
    const origin = request.headers.get('x-tracker-origin') || url.origin;
    const out = runOnce_(this.env, origin, url, request.method, request.headers.get('content-type'), body);
    if (out.need) return json_({ error: 'Data not available: ' + JSON.stringify(out.need) }, 500);   // cannot happen with synchronous loading
    try { if (!journalEmpty_(out.journal)) doCommit_(this.ctx.storage.sql, out.journal); }
    catch (err) { console.error('Save failed: ' + (err && err.stack || err)); doColdLoad_(this.ctx.storage.sql); return json_({ error: 'Save failed: ' + String(err && err.message || err) }, 500); }
    if (out.mail.length) await sendOutbox_(out.mail);
    return out.res;
  }
}

function trackerStub_(env) { return env.TRACKER.get(env.TRACKER.idFromName('main')); }

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/favicon.ico') return new Response(null, { status: 204 });
    if (url.pathname === '/__cron') return new Response('Not found', { status: 404 });
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type' } });
    const headers = new Headers(request.headers);
    headers.set('x-tracker-origin', url.origin);
    return trackerStub_(env).fetch(new Request(request.url, { method: request.method, headers, body: request.method === 'POST' ? await request.arrayBuffer() : undefined }));
  },
  async scheduled(event, env, ctx) {
    ctx.waitUntil(trackerStub_(env).fetch('https://tracker.internal/__cron', { method: 'POST' }));
  }
};
