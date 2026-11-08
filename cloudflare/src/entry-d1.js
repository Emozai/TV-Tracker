// ============================================================================================
// Paste-in Worker (Cloudflare dashboard → your Worker → Edit code). Data is stored in a D1 database
// bound as DB (Settings → Bindings → D1 database, variable name DB).
//
// How saving works: each Worker copy keeps the data in memory and, before every request, reads only the
// rows changed since it last looked (a change counter "seq"). The request runs on that copy; its changes
// are written in ONE batch whose statements only apply if seq is still the same. If another request saved
// in between, nothing is written: the Worker reads the new data and runs the request again.
// ============================================================================================

const D1_CHUNK_ = 80000;     // bytes per JSON batch parameter (D1: 100 KB per statement)
const D1_PART_ = 30000;      // characters per backup-file part (≤ 90 KB even if every character is 3 bytes)
const utf8Len_ = s => /[^\x00-\x7f]/.test(s) ? new TextEncoder().encode(s).length : s.length;
let d1SchemaReady_ = false;

async function d1Schema_(db) {
  if (d1SchemaReady_) return;
  await db.batch([
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_meta (k TEXT PRIMARY KEY, v INTEGER NOT NULL)`),
    db.prepare(`INSERT OR IGNORE INTO gs_meta (k, v) VALUES ('seq', 0), ('floor', 0)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_sheets (name TEXT PRIMARY KEY, pos INTEGER, meta TEXT, seq INTEGER NOT NULL DEFAULT 0, deleted INTEGER NOT NULL DEFAULT 0)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_rows (sheet TEXT NOT NULL, r INTEGER NOT NULL, data TEXT, seq INTEGER NOT NULL DEFAULT 0, PRIMARY KEY (sheet, r))`),
    db.prepare(`CREATE INDEX IF NOT EXISTS gs_rows_seq ON gs_rows (seq)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_props (k TEXT PRIMARY KEY, v TEXT, seq INTEGER NOT NULL DEFAULT 0)`),
    db.prepare(`CREATE INDEX IF NOT EXISTS gs_props_seq ON gs_props (seq)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_cache (k TEXT PRIMARY KEY, v TEXT, exp INTEGER NOT NULL DEFAULT 0, seq INTEGER NOT NULL DEFAULT 0)`),
    db.prepare(`CREATE INDEX IF NOT EXISTS gs_cache_seq ON gs_cache (seq)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_files (id TEXT PRIMARY KEY, name TEXT, folder TEXT, created INTEGER, mime TEXT, seq INTEGER NOT NULL DEFAULT 0, deleted INTEGER NOT NULL DEFAULT 0)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS gs_file_parts (id TEXT NOT NULL, i INTEGER NOT NULL, data TEXT, PRIMARY KEY (id, i))`)
  ]);
  d1SchemaReady_ = true;
}

function d1ApplySheets_(rows) { for (const r of rows) { if (r.deleted) { M.sheets.delete(r.name); M.rows.delete(r.name); } else M.sheets.set(r.name, { pos: r.pos, meta: JSON.parse(r.meta || '{}') }); } }
function d1ApplyProps_(rows) { for (const r of rows) { if (r.v === null) M.props.delete(r.k); else M.props.set(r.k, r.v); } }
function d1ApplyCache_(rows) { for (const r of rows) { if (r.v === null) M.cache.delete(r.k); else M.cache.set(r.k, { v: r.v, exp: r.exp }); } }
function d1ApplyFiles_(rows) { for (const r of rows) { if (r.deleted) { M.files.delete(r.id); M.fileText.delete(r.id); } else { M.files.set(r.id, { id: r.id, name: r.name, folder: r.folder, created: r.created, mime: r.mime }); M.fileText.delete(r.id); } } }
function d1ApplyRows_(rows) {
  for (const r of rows) {
    const arr = M.rows.get(r.sheet);
    if (!arr) continue;                         // sheet not loaded in this copy: it is read in full when needed
    if (r.data === null) { if (r.r < arr.length) arr[r.r] = []; }
    else { while (arr.length < r.r) arr.push([]); arr[r.r] = decRow_(r.data); }
  }
  M.rows.forEach(arr => { let n = arr.length; while (n > 0 && !(arr[n - 1] && arr[n - 1].length)) n--; arr.length = n; });
}

/** Brings M up to date with the database (full load the first time, only changes afterwards). */
async function d1Sync_(db) {
  await d1Schema_(db);
  const since = M.ready ? M.seq : -1;
  const meta = db.prepare(`SELECT k, v FROM gs_meta`);
  if (since < 0) {
    const [m, sh, pr, ca, fi] = await db.batch([meta,
      db.prepare(`SELECT name, pos, meta, deleted FROM gs_sheets WHERE deleted = 0`),
      db.prepare(`SELECT k, v FROM gs_props WHERE v IS NOT NULL`),
      db.prepare(`SELECT k, v, exp FROM gs_cache WHERE v IS NOT NULL AND exp > ?`).bind(Date.now()),
      db.prepare(`SELECT id, name, folder, created, mime, deleted FROM gs_files WHERE deleted = 0`)]);
    M.sheets.clear(); M.rows.clear(); M.props.clear(); M.cache.clear(); M.files.clear(); M.fileText.clear();
    d1ApplySheets_(sh.results); d1ApplyProps_(pr.results); d1ApplyCache_(ca.results); d1ApplyFiles_(fi.results);
    const mv = Object.fromEntries(m.results.map(r => [r.k, r.v]));
    M.seq = mv.seq || 0; M.floor = mv.floor || 0; M.ready = true;
    return;
  }
  // One round trip: the change counter plus everything changed since we last looked (indexed on seq).
  const [m2, sh, pr, ca, fi, ro] = await db.batch([meta,
    db.prepare(`SELECT name, pos, meta, deleted FROM gs_sheets WHERE seq > ?`).bind(since),
    db.prepare(`SELECT k, v FROM gs_props WHERE seq > ?`).bind(since),
    db.prepare(`SELECT k, v, exp FROM gs_cache WHERE seq > ?`).bind(since),
    db.prepare(`SELECT id, name, folder, created, mime, deleted FROM gs_files WHERE seq > ?`).bind(since),
    db.prepare(`SELECT sheet, r, data FROM gs_rows WHERE seq > ?`).bind(since)]);
  const mv = Object.fromEntries(m2.results.map(r => [r.k, r.v]));
  if ((mv.floor || 0) > since) { M.ready = false; return d1Sync_(db); }   // too far behind: reload everything
  d1ApplySheets_(sh.results); d1ApplyProps_(pr.results); d1ApplyCache_(ca.results); d1ApplyFiles_(fi.results); d1ApplyRows_(ro.results);
  M.seq = mv.seq || 0; M.floor = mv.floor || 0;
}

async function d1Load_(db, needs) {
  const stmts = [], kinds = [];
  for (const n of needs) {
    if (n.sheet !== undefined && !M.rows.has(n.sheet) && !kinds.some(k => k.sheet === n.sheet)) { stmts.push(db.prepare(`SELECT r, data FROM gs_rows WHERE sheet = ? AND data IS NOT NULL`).bind(n.sheet)); kinds.push(n); }
    if (n.file !== undefined && !M.fileText.has(n.file) && !kinds.some(k => k.file === n.file)) { stmts.push(db.prepare(`SELECT data FROM gs_file_parts WHERE id = ? ORDER BY i`).bind(n.file)); kinds.push(n); }
  }
  if (!stmts.length) return;
  const res = await db.batch(stmts);
  res.forEach((r, i) => {
    const k = kinds[i];
    if (k.sheet !== undefined) {
      const arr = [];
      for (const row of r.results) arr[row.r] = decRow_(row.data);
      for (let x = 0; x < arr.length; x++) if (!arr[x]) arr[x] = [];
      M.rows.set(k.sheet, arr);
    } else M.fileText.set(k.file, r.results.map(x => x.data).join(''));
  });
}

/** Splits [..] items into JSON strings of at most D1_CHUNK_ characters. */
function d1Chunks_(items) {
  const out = []; let cur = [], size = 2;
  for (const it of items) {
    const s = JSON.stringify(it), n = utf8Len_(s);
    if (cur.length && size + n + 1 > D1_CHUNK_) { out.push('[' + cur.join(',') + ']'); cur = []; size = 2; }
    cur.push(s); size += n + 1;
  }
  if (cur.length) out.push('[' + cur.join(',') + ']');
  return out;
}

/** Writes the journal in one guarded batch. Returns true if saved, false if someone else saved first. */
async function d1Commit_(db, j) {
  const old = M.seq, seq = old + 1;
  const G = `(SELECT v FROM gs_meta WHERE k = 'seq') = ${Number(old)}`;
  const st = [];
  j.drop.forEach(n => {
    st.push(db.prepare(`UPDATE gs_rows SET data = NULL, seq = ? WHERE sheet = ? AND ${G}`).bind(seq, n));
    st.push(db.prepare(`UPDATE gs_sheets SET deleted = 1, seq = ? WHERE name = ? AND ${G}`).bind(seq, n));
  });
  const sheetItems = []; j.sheets.forEach((s, n) => sheetItems.push([n, s.pos, JSON.stringify(s.meta)]));
  for (const c of d1Chunks_(sheetItems)) st.push(db.prepare(
    `INSERT INTO gs_sheets (name, pos, meta, seq, deleted) SELECT json_extract(value, '$[0]'), json_extract(value, '$[1]'), json_extract(value, '$[2]'), ?, 0 FROM json_each(?) WHERE ${G}
     ON CONFLICT(name) DO UPDATE SET pos = excluded.pos, meta = excluded.meta, seq = excluded.seq, deleted = 0`).bind(seq, c));
  const rowItems = []; j.rows.forEach((rows, n) => rows.forEach((row, r) => rowItems.push([n, r, row === null ? null : encRow_(row)])));
  for (const c of d1Chunks_(rowItems)) st.push(db.prepare(
    `INSERT INTO gs_rows (sheet, r, data, seq) SELECT json_extract(value, '$[0]'), json_extract(value, '$[1]'), json_extract(value, '$[2]'), ? FROM json_each(?) WHERE ${G}
     ON CONFLICT(sheet, r) DO UPDATE SET data = excluded.data, seq = excluded.seq`).bind(seq, c));
  if (j.propsClear) st.push(db.prepare(`UPDATE gs_props SET v = NULL, seq = ? WHERE ${G}`).bind(seq));
  const propItems = []; j.props.forEach((v, k) => propItems.push([k, v]));
  for (const c of d1Chunks_(propItems)) st.push(db.prepare(
    `INSERT INTO gs_props (k, v, seq) SELECT json_extract(value, '$[0]'), json_extract(value, '$[1]'), ? FROM json_each(?) WHERE ${G}
     ON CONFLICT(k) DO UPDATE SET v = excluded.v, seq = excluded.seq`).bind(seq, c));
  const cacheItems = []; j.cache.forEach((e, k) => cacheItems.push([k, e.v, e.exp]));
  for (const c of d1Chunks_(cacheItems)) st.push(db.prepare(
    `INSERT INTO gs_cache (k, v, exp, seq) SELECT json_extract(value, '$[0]'), json_extract(value, '$[1]'), json_extract(value, '$[2]'), ? FROM json_each(?) WHERE ${G}
     ON CONFLICT(k) DO UPDATE SET v = excluded.v, exp = excluded.exp, seq = excluded.seq`).bind(seq, c));
  j.files.forEach((m, id) => {
    st.push(db.prepare(`DELETE FROM gs_file_parts WHERE id = ? AND ${G}`).bind(id));
    if (m === null) st.push(db.prepare(`UPDATE gs_files SET deleted = 1, seq = ? WHERE id = ? AND ${G}`).bind(seq, id));
    else st.push(db.prepare(`INSERT INTO gs_files (id, name, folder, created, mime, seq, deleted) SELECT ?, ?, ?, ?, ?, ?, 0 WHERE ${G}
      ON CONFLICT(id) DO UPDATE SET name = excluded.name, folder = excluded.folder, created = excluded.created, mime = excluded.mime, seq = excluded.seq, deleted = 0`)
      .bind(id, m.name, m.folder, m.created, m.mime, seq));
  });
  j.fileText.forEach((t, id) => {
    if (!j.files.get(id)) return;
    const parts = Math.max(1, Math.ceil(t.length / D1_PART_));
    for (let i = 0; i < parts; i++) st.push(db.prepare(`INSERT OR REPLACE INTO gs_file_parts (id, i, data) SELECT ?, ?, ? WHERE ${G}`).bind(id, i, t.slice(i * D1_PART_, (i + 1) * D1_PART_)));
  });
  // Housekeeping every 200 saves: drop old "deleted" markers and expired cache entries.
  let floor = M.floor;
  if (seq % 200 === 0) {
    floor = seq - 100;
    st.push(db.prepare(`DELETE FROM gs_rows WHERE data IS NULL AND seq < ? AND ${G}`).bind(floor));
    st.push(db.prepare(`DELETE FROM gs_props WHERE v IS NULL AND seq < ? AND ${G}`).bind(floor));
    st.push(db.prepare(`DELETE FROM gs_cache WHERE (v IS NULL OR exp < ?) AND seq < ? AND ${G}`).bind(Date.now(), floor));
    st.push(db.prepare(`DELETE FROM gs_sheets WHERE deleted = 1 AND seq < ? AND ${G}`).bind(floor));
    st.push(db.prepare(`DELETE FROM gs_files WHERE deleted = 1 AND seq < ? AND ${G}`).bind(floor));
    st.push(db.prepare(`UPDATE gs_meta SET v = ? WHERE k = 'floor' AND ${G}`).bind(floor));
  }
  st.push(db.prepare(`UPDATE gs_meta SET v = ? WHERE k = 'seq' AND v = ?`).bind(seq, old));
  const res = await db.batch(st);
  const ok = res[res.length - 1].meta.changes === 1;
  if (ok) { applyJournal_(j); M.seq = seq; M.floor = floor; }
  return ok;
}

async function handleD1_(env, url, method, contentType, body) {
  const db = env.DB;
  if (!db || typeof db.prepare !== 'function') {
    return new Response('<h2>Almost there</h2><p>Add a D1 database binding named <b>DB</b> to this Worker: Settings → Bindings → Add → D1 database (variable name <b>DB</b>), then reload.</p>',
      { status: 503, headers: { 'Content-Type': 'text/html; charset=utf-8' } });
  }
  await d1Sync_(db);
  for (let attempt = 0; attempt < 12; attempt++) {
    const out = runOnce_(env, url.hostname === 'tracker.internal' ? (env.PUBLIC_URL || '') : url.origin, url, method, contentType, body);
    if (out.need) { await d1Load_(db, out.need); continue; }
    if (journalEmpty_(out.journal)) { if (out.mail.length) await sendOutbox_(out.mail); return out.res; }
    if (await d1Commit_(db, out.journal)) { if (out.mail.length) await sendOutbox_(out.mail); return out.res; }
    await d1Sync_(db);   // someone else saved first: read their changes and run the request again
  }
  return json_({ error: 'The tracker is busy. Please try again in a few seconds.' }, 503);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/favicon.ico') return new Response(null, { status: 204 });
    if (url.pathname === '/__cron') return new Response('Not found', { status: 404 });
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET, POST', 'Access-Control-Allow-Headers': 'Content-Type' } });
    const body = request.method === 'POST' ? await request.text() : null;
    try { return await handleD1_(env, url, request.method, request.headers.get('content-type'), body); }
    catch (err) { console.error(err && err.stack || err); return json_({ error: String(err && err.message || err) }, 500); }
  },
  async scheduled(event, env, ctx) {
    ctx.waitUntil(handleD1_(env, new URL('https://tracker.internal/__cron'), 'POST', '', '').catch(err => console.error('cron: ' + (err && err.stack || err))));
  }
};
