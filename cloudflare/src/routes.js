// ============================================================================================
// Routes shared by both Cloudflare versions (Durable Object: entry.js, D1 paste-in: entry-d1.js).
//   GET  /                 the tracker web app (doGet)        GET/POST /?api=…  phone app API (doGet/doPost)
//   GET  /setup            Admin password page (replaces the spreadsheet menu)
//   POST /api/run          google.script.run calls from the page: { fn, args } -> { result } | { error }
//   GET  /files/<id>.xlsx  backup copies made by the nightly backup / "Back up now"
//   cron (every minute)    time triggers: email queue, nightly backup + clean-up, daily summary
// ============================================================================================

// Functions the browser may not call (spreadsheet menu items, triggers, page handlers).
const BLOCKED_FUNCTIONS_ = new Set(['doGet', 'doPost', 'onOpen', 'onEdit', 'onInstall', 'onChange', 'handleTrackerChange',
  'setup', 'nightlyMaintenance', 'processTrackerEmailQueue', 'dailySummaryEmail', 'include']);
function callable_(name) {
  return typeof name === 'string' && /^[A-Za-z][A-Za-z0-9]*$/.test(name) && !BLOCKED_FUNCTIONS_.has(name) &&
    !/FromMenu$/.test(name) && Object.prototype.hasOwnProperty.call(GAS_FUNCTIONS_, name) && typeof GAS_FUNCTIONS_[name] === 'function';
}
function callGas_(name, args) {
  const f = GAS_FUNCTIONS_[name];
  if (typeof f !== 'function') throw new Error('Script function not found: ' + name);
  return f.apply(null, args || []);
}
function json_(obj, status) {
  return new Response(JSON.stringify(obj), { status: status || 200, headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } });
}
function htmlPage_(out) {
  let html = out.getContent().split('__SERVICE_URL__').join(RT.origin + '/');
  const head = [];
  if (out.getTitle && out.getTitle() && !/<title>/i.test(html)) head.push('<title>' + xmlEsc_(out.getTitle()) + '</title>');
  (out._meta || []).forEach(([n, c]) => { if (!new RegExp('<meta[^>]+name=["\']' + n + '["\']', 'i').test(html)) head.push('<meta name="' + xmlEsc_(n) + '" content="' + xmlEsc_(c) + '">'); });
  if (head.length) html = /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, m => m + head.join('')) : head.join('') + html;
  return new Response(html, { headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store',
    'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'same-origin' } });
}
function eventFrom_(url, body, contentType) {
  const parameter = {}, parameters = {};
  url.searchParams.forEach((v, k) => { if (!(k in parameter)) parameter[k] = v; (parameters[k] = parameters[k] || []).push(v); });
  if (body && /application\/x-www-form-urlencoded/i.test(contentType || '')) {
    new URLSearchParams(body).forEach((v, k) => { if (!(k in parameter)) parameter[k] = v; (parameters[k] = parameters[k] || []).push(v); });
  }
  const e = { parameter, parameters, queryString: url.search.replace(/^\?/, ''), contextPath: '', contentLength: body ? body.length : -1, pathInfo: '' };
  if (body != null) e.postData = { contents: body, type: contentType || '', length: body.length, name: 'postData' };
  return e;
}
function respondOutput_(out) {
  if (out instanceof GTextOutput_) return new Response(out.getContent(), { headers: { 'Content-Type': out.getMimeType() + '; charset=utf-8', 'Cache-Control': 'no-store', 'Access-Control-Allow-Origin': '*' } });
  if (out instanceof GHtmlOutput_) return htmlPage_(out);
  return json_(out === undefined ? null : out);
}

/** Handles one request synchronously on the working copy. NeedData_ (D1: data not loaded yet) is re-thrown. */
function route_(url, method, contentType, body) {
  const path = url.pathname;
  // First visit after deploy: run setup() once (creates the sheets and the scheduled jobs). Apps Script needs this by hand.
  if (!ST_.propGet('__CF_SETUP_DONE')) {
    callGas_('setup', []);
    scriptProperties_.setProperty('__CF_SETUP_DONE', new Date().toISOString());
  }
  if (path === '/__cron') return json_({ ran: runDueTriggers_(Date.now(), (h, e) => callGas_(h, [e])) });

  if (path === '/api/run') {
    if (method !== 'POST') return json_({ error: 'Use POST.' }, 405);
    let msg;
    try { msg = JSON.parse(body || '{}'); } catch (_) { return json_({ error: 'Bad request.' }, 400); }
    if (!callable_(msg.fn)) return json_({ error: 'Script function not found: ' + String(msg.fn).slice(0, 60) });
    try {
      const result = callGas_(msg.fn, Array.isArray(msg.args) ? msg.args : []);
      return json_({ result: result === undefined ? null : result });
    } catch (err) {
      if (err instanceof NeedData_) throw err;
      console.warn(msg.fn + ': ' + (err && err.stack || err));
      return json_({ error: String(err && err.message || err) });
    }
  }

  if (path === '/setup') {
    const out = HtmlService.createHtmlOutputFromFile('PasswordSetup').setTitle('Tracker — Admin Password Setup');
    out.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return htmlPage_(out);
  }

  const file = path.match(/^\/files\/([0-9a-f-]{36})\.xlsx$/i);
  if (file) {
    const row = fileRow_(file[1]);
    if (!row) return new Response('Not found', { status: 404 });
    const text = readFileText_(row.id);
    const bytes = row.mime === 'application/x-tracker-snapshot' ? xlsxFromSnapshot_(text) : b64dec_(JSON.parse(text).blob || '');
    const name = String(row.name || 'backup').replace(/[^\w .()-]+/g, '_') + '.xlsx';
    return new Response(bytes, { headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Content-Disposition': 'attachment; filename="' + name + '"', 'Cache-Control': 'no-store' } });
  }

  if (path === '/' || path === '/exec' || path === '/index.html') {
    const e = eventFrom_(url, body, contentType);
    if (method === 'POST') {
      if (typeof GAS_FUNCTIONS_.doPost !== 'function') return json_({ error: 'POST is not supported.' }, 405);
      return respondOutput_(GAS_FUNCTIONS_.doPost(e));
    }
    return respondOutput_(GAS_FUNCTIONS_.doGet(e));
  }
  return new Response('Not found', { status: 404 });
}

/** Runs route_ and moves the changes into the journal. Returns { res, journal, mail } or { need }. */
function runOnce_(env, origin, url, method, contentType, body) {
  shimBegin_(env, origin);
  let res;
  try { res = route_(url, method, contentType, body); }
  catch (err) {
    if (err instanceof NeedData_) return { need: RT.needed.slice() };
    console.error(err && err.stack || err);
    res = json_({ error: String(err && err.message || err) }, 500);
  }
  if (RT.needed.length) return { need: RT.needed.slice() };   // Code.gs caught NeedData_ itself: discard and re-run
  shimFlush_();
  return { res, journal: J, mail: RT.outbox.slice() };
}
