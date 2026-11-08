import { DatabaseSync } from 'node:sqlite';
export async function load(dbFile) {
  const mod = await import(new URL('../dist/worker.test.mjs', import.meta.url).href + '?' + Date.now());
  const db = new DatabaseSync(dbFile);
  const sql = { exec(q, ...b) {
    const st = db.prepare(q);
    if (/^\s*(select|pragma|with)/i.test(q)) { const rows = st.all(...b); return { toArray: () => rows, one: () => rows[0] }; }
    st.run(...b); return { toArray: () => [], one: () => undefined };
  } };
  const env = { ADMIN_SETUP_KEY: 'test-setup-key-123', TIMEZONE: 'Etc/GMT+2', TRACKER: null };
  const store = new mod.TrackerStore({ storage: { sql } }, env);
  const origin = 'https://tracker.example.workers.dev';
  const req = (path, opts = {}) => store.fetch(new Request(origin + path, { method: opts.method || 'GET', headers: Object.assign({ 'x-tracker-origin': origin }, opts.headers || {}), body: opts.body }));
  const run = async (fn, ...args) => { const r = await req('/api/run', { method: 'POST', body: JSON.stringify({ fn, args }), headers: { 'content-type': 'application/json' } }); const j = await r.json(); if (j.error) throw new Error(fn + ': ' + j.error); return j.result; };
  return { mod, db, store, req, run, env };
}
