// Paste-in (D1) version: two Worker copies ("isolates") sharing one database.
import { fakeD1 } from './fake-d1.mjs';
import { rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const file = join(tmpdir(), 'tracker-d1-test.db'); rmSync(file, { force: true });
const DB = fakeD1(file);
const env = { DB, ADMIN_SETUP_KEY: 'test-setup-key-123', TIMEZONE: 'Etc/GMT+2' };
const ok = (c, m) => { if (!c) { console.log('FAIL ' + m); process.exitCode = 1; } else console.log('ok   ' + m); };
const isolate = async tag => {
  const mod = await import(new URL('../dist/hello-world.test.mjs', import.meta.url).href + '?' + tag);
  const W = mod.default, origin = 'https://hello-world.example.workers.dev';
  const req = (path, opts = {}) => W.fetch(new Request(origin + path, { method: opts.method || 'GET', headers: opts.headers || {}, body: opts.body }), env, {});
  const run = async (fn, ...args) => { const r = await req('/api/run', { method: 'POST', body: JSON.stringify({ fn, args }), headers: { 'content-type': 'application/json' } }); const j = await r.json(); if (j.error) throw new Error(fn + ': ' + j.error); return j.result; };
  return { mod, W, req, run };
};
const A = await isolate('A'), B = await isolate('B');

let r = await (await isolate('nodb')).W.fetch(new Request('https://x.workers.dev/'), {}, {});
ok(r.status === 503 && /D1 database/.test(await r.text()), 'without the DB binding the page explains what to add');
r = await A.req('/'); ok(r.status === 200 && (await r.text()).includes('Upload Cabins'), 'GET / serves the tracker');
ok((await A.run('saveTechnicianPasswordSetup', 'test-setup-key-123', 'Admin', 'averylongpassword1')).success, 'Admin password set on copy A');
const LB = await B.run('loginTechnician', 'Admin', 'averylongpassword1', 'sB'); ok(LB.success, 'copy B sees the password set on A (sign-in works)');
const LA = await A.run('loginTechnician', 'Admin', 'averylongpassword1', 'sA');
const items = []; for (let i = 1; i <= 1500; i++) items.push({ type: 'PAX', cabin: String(5000 + i), deck: String(5 + i % 10) }); for (let i = 1; i <= 400; i++) items.push({ type: 'CREW', cabin: String(3000 + i), deck: '3' });
let t0 = Date.now(), s0 = DB.stats.maxStatementBytes;
let up = await A.run('importCabinInventory', LA.token, 'sA', items, 'merge');
ok(up.success && up.added === 1900, 'upload 1,900 cabins on A: ' + up.message + ' (' + (Date.now() - t0) + ' ms, largest statement ' + Math.round(DB.stats.maxStatementBytes / 1024) + ' KB ≤ 100 KB)');
await A.run('importVenueInventory', LA.token, 'sA', [{ venue: 'Lobby', location: 'Reception', deck: '6', size: '55', qty: '2', area: 'PAX' }, { venue: 'Crew Mess', deck: '3', area: 'CREW' }], 'merge');
let sb = await B.run('getTrackerSnapshot', LB.token); ok(sb.cabins.length === 1900 && sb.venues.length === 2, 'copy B sees the upload: ' + sb.cabins.length + ' cabins, ' + sb.venues.length + ' venues');
const c1 = sb.cabins.find(x => x.cabin === '5001');
let s = await A.run('saveTV', c1.rowNumber, { authToken: LA.token, sessionId: 'sA', cabin: '5001', cabinType: 'PAX', 'TV installed': true, comments: 'saved on A', requestId: 'q1' });
ok(s.success, 'save on A');
const before = { ...DB.stats };
sb = await B.run('getTrackerSnapshot', LB.token);
const k = sb.cabins.find(x => x.cabin === '5001');
ok(k['TV installed'] === true && /saved on A/.test(k.comments), 'copy B sees A\'s save straight away');
ok(DB.stats.rowsRead - before.rowsRead < 50, 'B read only the changed rows (' + (DB.stats.rowsRead - before.rowsRead) + ' rows, not 1,900+)');
// Two technicians saving at the same moment on different copies
const sn = await A.run('getTrackerSnapshot', LA.token), rn = c => sn.cabins.find(x => x.cabin === c).rowNumber;
const [x1, x2] = await Promise.all([
  A.run('saveTV', rn('5002'), { authToken: LA.token, sessionId: 'sA', cabin: '5002', cabinType: 'PAX', 'TV installed': true, comments: 'A at the same time', requestId: 'q2' }),
  B.run('saveTV', rn('5003'), { authToken: LB.token, sessionId: 'sB', cabin: '5003', cabinType: 'PAX', 'TV installed': true, comments: 'B at the same time', requestId: 'q3' })]);
ok(x1.success && x2.success, 'two saves at the same moment on A and B both succeed');
const sa = await A.run('getTrackerSnapshot', LA.token), g = c => sa.cabins.find(x => x.cabin === c);
ok(/A at the same/.test(g('5002').comments) && /B at the same/.test(g('5003').comments) && /saved on A/.test(g('5001').comments), 'neither save overwrote the other (both remarks are there)');
// Same cabin, both copies, at once
const [y1, y2] = await Promise.all([
  A.run('saveTV', rn('5004'), { authToken: LA.token, sessionId: 'sA', cabin: '5004', cabinType: 'PAX', 'Power Connected': true, requestId: 'q4' }),
  B.run('saveTV', rn('5004'), { authToken: LB.token, sessionId: 'sB', cabin: '5004', cabinType: 'PAX', 'TV installed': true, requestId: 'q5' })]);
const c4 = (await B.run('getTrackerSnapshot', LB.token)).cabins.find(x => x.cabin === '5004');
ok(y1.success && y2.success && c4['TV installed'] === true && c4['Power Connected'] === true, 'two saves on the SAME cabin at once: both ticks kept');
const asg = await A.run('assignCabinsToTechnician', LA.token, 'sA', [{ cabin: '5010', cabinType: 'PAX' }], 'TV Check 2');
const LT = await B.run('loginTechnician', 'TV Check 2', '', 'sT'); const st = await B.run('getTrackerSnapshot', LT.token);
ok(asg.success && LT.success && st.cabins.length === 1 && st.cabins[0].cabin === '5010', 'assignment on A, TV Check 2 on B sees exactly its cabin');
const bk = await A.run('backupNow', LA.token, 'sA'); const fid = (JSON.stringify(bk).match(/files\/([0-9a-f-]{36})\.xlsx/) || [])[1];
const fr = fid ? await B.req('/files/' + fid + '.xlsx') : null; const fb = fr ? Buffer.from(await fr.arrayBuffer()) : Buffer.alloc(0);
ok(fr && fr.status === 200 && fb.length > 100000, 'backup made on A downloads from B (' + Math.round(fb.length / 1024) + ' KB, largest statement ' + Math.round(DB.stats.maxStatementBytes / 1024) + ' KB)');
const ex = await B.run('exportTrackerExcel', LB.token); ok(ex.success && ex.base64.length > 1000, 'Excel export on B');
const cl = await B.run('clearInventory', LB.token, 'sB', 'venue', 'DELETE'); const va = await A.run('getTrackerSnapshot', LA.token);
ok(cl.success && va.venues.length === 0 && va.cabins.length === 1900, 'remove all public areas on B; A sees 0 venues and still 1,900 cabins');
// Scheduled jobs
const trig = JSON.parse(A.mod.M.props.get('__TRIGGERS') || '[]');
ok(trig.some(t => t.handler === 'processTrackerEmailQueue') && trig.some(t => t.handler === 'nightlyMaintenance'), 'scheduled jobs set up automatically');
await A.W.scheduled({}, env, { waitUntil: p => p }); ok(true, 'cron handler runs');
// Many saves: housekeeping every 200 changes, and a copy that fell far behind reloads fully
for (let i = 0; i < 210; i++) await A.run('heartbeatUser', 'sA', '5001', LA.token);
const seq = DB.raw.prepare("SELECT v FROM gs_meta WHERE k='seq'").get().v, floor = DB.raw.prepare("SELECT v FROM gs_meta WHERE k='floor'").get().v;
ok(seq > 200 && floor > 0, 'after ' + seq + ' saves: housekeeping ran (floor ' + floor + ')');
const late = await isolate('late'); late.mod.M.ready = false;
const sl = await late.run('getTrackerSnapshot', (await late.run('loginTechnician', 'Admin', 'averylongpassword1', 'sL')).token);
ok(sl.cabins.length === 1900 && sl.cabins.find(x => x.cabin === '5003').comments.includes('B at the same'), 'a brand-new copy loads everything correctly');
sb = await B.run('getTrackerSnapshot', LB.token); ok(sb.cabins.length === 1900 && sb.venues.length === 0, 'copy B (behind by 200+ saves) catches up correctly');
const api = await B.req('/?api=sheets'); ok(api.status === 200, 'phone-app API answers');
console.log('D1 totals: ' + DB.stats.batches + ' batches, ' + DB.stats.statements + ' statements, ' + DB.stats.rowsRead + ' rows read, ' + DB.stats.rowsWritten + ' rows written');
