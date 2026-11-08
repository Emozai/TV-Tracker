// Stand-in for a Cloudflare D1 binding on node:sqlite. Enforces D1's 100 KB per-statement limit and counts rows read.
import { DatabaseSync } from 'node:sqlite';
export function fakeD1(file) {
  const db = new DatabaseSync(file);
  const stats = { statements: 0, rowsRead: 0, rowsWritten: 0, batches: 0, maxStatementBytes: 0 };
  const size = (q, b) => Buffer.byteLength(q) + b.reduce((n, v) => n + (typeof v === 'string' ? Buffer.byteLength(v) : 8), 0);
  function exec(q, b) {
    const bytes = size(q, b);
    stats.maxStatementBytes = Math.max(stats.maxStatementBytes, bytes);
    if (bytes > 100000) throw new Error('D1_ERROR: statement too long (' + bytes + ' bytes)');
    stats.statements++;
    const st = db.prepare(q);
    if (/^\s*(select|with)/i.test(q)) { const rows = st.all(...b); stats.rowsRead += rows.length; return { results: rows, success: true, meta: { changes: 0 } }; }
    const r = st.run(...b); stats.rowsWritten += Number(r.changes); return { results: [], success: true, meta: { changes: Number(r.changes) } };
  }
  class Stmt {
    constructor(q, b) { this.q = q; this.b = b || []; }
    bind(...b) { return new Stmt(this.q, b.map(v => v === undefined ? null : v)); }
    async all() { return exec(this.q, this.b); }
    async run() { return exec(this.q, this.b); }
    async first(col) { const r = exec(this.q, this.b).results[0]; return r ? (col ? r[col] : r) : null; }
  }
  return {
    stats, raw: db,
    prepare: q => new Stmt(q),
    async batch(list) {
      stats.batches++;
      await new Promise(r => setTimeout(r, 1));   // let other requests interleave, like the network does
      db.exec('BEGIN');
      try { const out = list.map(s => exec(s.q, s.b)); db.exec('COMMIT'); return out; }
      catch (e) { db.exec('ROLLBACK'); throw e; }
    },
    async exec(q) { db.exec(q); return { count: 1 }; }
  };
}
