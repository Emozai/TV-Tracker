// google.script.run for the Cloudflare version: each call becomes POST /api/run { fn, args }.
(function () {
  function runner(opts) {
    return new Proxy({}, {
      get: function (_, fn) {
        if (fn === 'withSuccessHandler') return function (f) { return runner(Object.assign({}, opts, { ok: f })); };
        if (fn === 'withFailureHandler') return function (f) { return runner(Object.assign({}, opts, { fail: f })); };
        if (fn === 'withUserObject') return function (u) { return runner(Object.assign({}, opts, { user: u })); };
        if (typeof fn !== 'string' || fn === 'then') return undefined;
        return function () {
          var args = Array.prototype.slice.call(arguments);
          var fail = function (e) { if (opts.fail) opts.fail(e, opts.user); else console.error(e); };
          fetch('/api/run', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fn: fn, args: args }), credentials: 'same-origin' })
            .then(function (r) {
              return r.text().then(function (t) {
                var j; try { j = JSON.parse(t); } catch (_) { throw new Error('Server error (HTTP ' + r.status + ').'); }
                return j;
              });
            })
            .then(function (j) {
              if (j && j.error) { var e = new Error(j.error); e.name = 'ScriptError'; fail(e); return; }
              if (opts.ok) opts.ok(j ? j.result : null, opts.user);
            })
            .catch(function (e) { fail(e instanceof Error ? e : new Error(String(e))); });
        };
      }
    });
  }
  window.google = window.google || {};
  window.google.script = {
    run: runner({}),
    host: { close: function () { try { window.close(); } catch (_) {} }, setHeight: function () {}, setWidth: function () {}, editor: { focus: function () {} }, origin: location.origin },
    url: { getLocation: function (cb) {
      var p = {}, ps = {};
      new URLSearchParams(location.search).forEach(function (v, k) { if (!(k in p)) p[k] = v; (ps[k] = ps[k] || []).push(v); });
      cb({ parameter: p, parameters: ps, hash: location.hash.replace(/^#/, '') });
    } },
    history: { push: function () {}, replace: function () {}, setChangeHandler: function () {} }
  };
})();
