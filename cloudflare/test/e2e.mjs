import { load } from './harness.mjs';
import { rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os'; import { join } from 'node:path';
const out = p => join(tmpdir(), p);
const dbFile = out('tracker-test.db'); rmSync(dbFile, { force: true });
const ok = (c, m) => { if (!c) { console.log('FAIL ' + m); process.exitCode = 1; } else console.log('ok   ' + m); };
let t = await load(dbFile);
let r = await t.req('/'); let html = await r.text();
ok(r.status === 200 && html.includes('google.script') && html.includes('Upload Cabins') && !html.includes('<?'), 'GET / serves the tracker page (' + Math.round(html.length/1024) + ' KB)');
ok(!html.includes('__SERVICE_URL__') && !html.includes('finalCheckLink"'), 'page links resolved');
r = await t.req('/setup'); ok(r.status === 200 && (await r.text()).includes('saveTechnicianPasswordSetup'), 'GET /setup serves the password page');
let bad = ''; try { await t.run('saveTechnicianPasswordSetup', 'wrong', 'Admin', 'averylongpassword1'); } catch (e) { bad = e.message; }
ok(/Incorrect admin setup key/.test(bad), 'wrong setup key refused');
const set = await t.run('saveTechnicianPasswordSetup', 'test-setup-key-123', 'Admin', 'averylongpassword1');
ok(set && set.success !== false, 'Admin password set: ' + JSON.stringify(set).slice(0, 90));
const blocked = await (await t.req('/api/run', { method: 'POST', body: JSON.stringify({ fn: 'setup', args: [] }) })).json();
ok(blocked.error && /not found/.test(blocked.error), 'setup() cannot be called from the browser');
let L = await t.run('loginTechnician', 'Admin', 'wrongpassword', 's1'); ok(!L.success, 'wrong Admin password refused');
L = await t.run('loginTechnician', 'Admin', 'averylongpassword1', 's1'); ok(L.success && L.token, 'Admin signs in');
const tok = L.token;
ok((await t.run('verifyTechnicianSession', tok)).success !== false, 'session verifies');
const items = [];
for (let i = 1; i <= 1500; i++) items.push({ type: 'PAX', cabin: String(5000 + i), deck: String(5 + (i % 10)), fz: String(1 + i % 4), tvCount: i % 50 === 0 ? '2' : '1' });
for (let i = 1; i <= 400; i++) items.push({ type: 'CREW', cabin: String(3000 + i), deck: '3' });
let t0 = Date.now(); let up = await t.run('importCabinInventory', tok, 's1', items, 'merge'); ok(up.success && up.added === 1900, 'upload 1,900 cabins: ' + up.message + ' (' + (Date.now() - t0) + ' ms)');
up = await t.run('importVenueInventory', tok, 's1', [{ venue: 'Lobby', location: 'Reception', deck: '6', size: '55', qty: '2', area: 'PAX', source: 'IPTV' }, { venue: 'Crew Mess', deck: '3', size: '43', area: 'CREW' }], 'merge');
ok(up.success && up.added === 2, 'upload public areas: ' + up.message);
t0 = Date.now(); let snap = await t.run('getTrackerSnapshot', tok); ok(snap.cabins.length === 1900 && snap.venues.length === 2, 'snapshot: ' + snap.cabins.length + ' cabins, ' + snap.venues.length + ' venues (' + (Date.now() - t0) + ' ms)');
const c = snap.cabins.find(x => x.cabin === '5001');
let s = await t.run('saveTV', c.rowNumber, { authToken: tok, sessionId: 's1', cabin: c.cabin, cabinType: c.cabinType, 'TV installed': true, 'Power Connected': true, comments: 'Bracket loose', requestId: 'r1', baseCheckedAt: c.checkedAt || '' });
ok(s.success, 'saveTV ticks a cabin: ' + (s.message || s.status || '').toString().slice(0, 80));
const ver1 = await t.run('getDataVersion'); 
const asg = await t.run('assignCabinsToTechnician', tok, 's1', [{ cabin: '5002', cabinType: 'PAX' }, { cabin: '5003', cabinType: 'PAX' }], 'TV Check 1');
ok(asg && asg.success !== false, 'assign 2 cabins to TV Check 1: ' + JSON.stringify(asg).slice(0, 80));
const L2 = await t.run('loginTechnician', 'TV Check 1', '', 's2'); ok(L2.success, 'TV Check 1 signs in without password');
const snap2 = await t.run('getTrackerSnapshot', L2.token); ok(snap2.cabins.length === 2, 'TV Check 1 sees only its 2 cabins');
const c2 = snap2.cabins.find(x => x.cabin === '5002');
s = await t.run('saveTV', c2.rowNumber, { authToken: L2.token, sessionId: 's2', cabin: c2.cabin, cabinType: c2.cabinType, 'TV installed': true, requestId: 'r2' }); ok(s.success, 'TV Check 1 saves its cabin');
const c3 = snap.cabins.find(x => x.cabin === '5010');
s = await t.run('saveTV', c3.rowNumber, { authToken: L2.token, sessionId: 's2', cabin: c3.cabin, cabinType: c3.cabinType, 'TV installed': true, requestId: 'r3' }); ok(!s.success && s.forbidden, 'TV Check 1 cannot save a cabin not assigned to it');
const hist = await t.run('getUpdateHistory', 20); ok(Array.isArray(hist) ? hist.length > 0 : (hist && JSON.stringify(hist).includes('5001')), 'update history has entries');
await t.run('registerTechnicianLogin', tok); await t.run('heartbeatUser', 's1', '5001', tok);
const act = await t.run('getActiveUsers', tok); ok(JSON.stringify(act).includes('Admin'), 'active users list works');
t0 = Date.now(); const ex = await t.run('exportTrackerExcel', tok); ok(ex.success && ex.base64 && ex.base64.length > 1000, 'Excel export built (' + Math.round(ex.base64.length * 0.75 / 1024) + ' KB, ' + (Date.now() - t0) + ' ms)');
writeFileSync(out('export.xlsx'), Buffer.from(ex.base64, 'base64'));
const ex2 = await t.run('exportAssignmentTemplate', tok); ok(ex2.success, 'assignment Excel built'); writeFileSync(out('assign.xlsx'), Buffer.from(ex2.base64, 'base64'));
const bk = await t.run('backupNow', tok, 's1'); ok(bk && bk.success !== false, 'backup now: ' + JSON.stringify(bk).slice(0, 120));
const fileId = (JSON.stringify(bk).match(/files\/([0-9a-f-]{36})\.xlsx/) || [])[1];
if (fileId) { const fr = await t.req('/files/' + fileId + '.xlsx'); const b = Buffer.from(await fr.arrayBuffer()); writeFileSync(out('backup.xlsx'), b); ok(fr.status === 200 && b.length > 1000, 'backup downloads as .xlsx (' + Math.round(b.length/1024) + ' KB)'); }
else ok(false, 'backup link present');
const trig = JSON.parse(t.mod.M.props.get('__TRIGGERS') || '[]'); ok(trig.some(x => x.handler === 'processTrackerEmailQueue') && trig.some(x => x.handler === 'nightlyMaintenance'), 'triggers installed: ' + trig.map(x => x.handler + '/' + x.unit + (x.atHour != null ? '@' + x.atHour : '')).join(', '));
r = await t.req('/__cron', { method: 'POST' }); ok(r.status === 200, 'cron runs: ' + await r.text());
// "eviction": brand-new instance must see the same data
// (new module instance below = fresh memory)
const t2 = await load(dbFile);
const L3 = await t2.run('loginTechnician', 'Admin', 'averylongpassword1', 's9'); ok(L3.success, 'after restart: Admin signs in again');
const snap3 = await t2.run('getTrackerSnapshot', L3.token); const k = snap3.cabins.find(x => x.cabin === '5001');
ok(snap3.cabins.length === 1900 && k['TV installed'] === true && /Bracket/.test(k.comments || ''), 'after restart: 1,900 cabins and the saved ticks/remark are still there');
const cl = await t2.run('clearInventory', L3.token, 's9', 'venue', 'DELETE'); ok(cl.success && cl.removed === 2, 'remove all public areas: ' + cl.message.slice(0, 70));
const api = await t2.req('/?api=sheets&key=x'); ok(api.status === 200, 'phone-app API answers (' + (await api.text()).slice(0, 70) + ')');
// scheduled jobs: pretend the email-queue job last ran 2 minutes ago, and it is 23:00 ship time for the nightly job
{
  const P = t2.mod.M.props, list = JSON.parse(P.get('__TRIGGERS') || '[]');
  ok(list.some(x => x.handler === 'processTrackerEmailQueue') && list.some(x => x.handler === 'nightlyMaintenance') && list.some(x => x.handler === 'dailySummaryEmail'),
    'scheduled jobs set up automatically: ' + list.map(x => x.handler + '(' + x.unit + (x.atHour != null ? ' @' + x.atHour + 'h' : '') + ')').join(', '));
  list.forEach(x => { x.last = Date.now() - 120000; });
  const tz = t2.mod.RT.tz, H = Number(t2.mod.gasFormatDate_(new Date(), tz, 'H'));
  list.forEach(x => { if (x.handler === 'nightlyMaintenance') x.atHour = H; });
  t2.mod.M.props.set('__TRIGGERS', JSON.stringify(list));
  let rr = await t2.req('/__cron', { method: 'POST' }); let ran = (await rr.json()).ran;
  ok(ran.includes('processTrackerEmailQueue') && ran.includes('nightlyMaintenance'), 'cron ran due jobs: ' + ran.join(', '));
  rr = await t2.req('/__cron', { method: 'POST' }); ran = (await rr.json()).ran;
  ok(!ran.includes('nightlyMaintenance'), 'nightly job does not run twice the same day (ran now: ' + (ran.join(', ') || 'nothing') + ')');
}
