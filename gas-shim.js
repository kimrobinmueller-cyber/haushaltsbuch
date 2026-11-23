// Ersetzt google.script.run, wenn die Seite nicht in Apps Script läuft
// (z. B. auf GitHub Pages). Ruft stattdessen die Web-App per fetch auf.
(function () {
  // In Apps Script gibt es google.script.run bereits: dann nichts tun.
  if (window.google && window.google.script && window.google.script.run) return;

  const TOKEN_KEY = 'hb_token';

  function getToken() {
    let t = null;
    try { t = localStorage.getItem(TOKEN_KEY); } catch (e) {}
    if (!t) {
      t = (prompt('Zugangs-Token eingeben (einmalig pro Gerät):') || '').trim();
      if (t) { try { localStorage.setItem(TOKEN_KEY, t); } catch (e) {} }
    }
    return t;
  }

  async function callApi(fn, args) {
    const cfg = window.HB_CONFIG || {};
    if (!cfg.API_URL || cfg.API_URL.indexOf('DEINE_ID') >= 0) {
      throw new Error('Bitte in config.js die Web-App-URL eintragen.');
    }
    const res = await fetch(cfg.API_URL, {
      method: 'POST',
      body: JSON.stringify({fn: fn, args: args, token: getToken()})
    });
    const data = await res.json();
    if (!data.ok) {
      if (data.error === 'UNAUTHORIZED') {
        try { localStorage.removeItem(TOKEN_KEY); } catch (e) {}
        throw new Error('Token falsch. Bitte App neu öffnen und das Token erneut eingeben.');
      }
      throw new Error(data.error);
    }
    return data.result;
  }

  function runner(ok, fail) {
    return new Proxy({}, {
      get: function (_, name) {
        if (name === 'withSuccessHandler') return function (f) { return runner(f, fail); };
        if (name === 'withFailureHandler') return function (f) { return runner(ok, f); };
        return function () {
          const args = Array.prototype.slice.call(arguments);
          callApi(String(name), args).then(
            function (r) { if (ok) ok(r); },
            function (e) { if (fail) fail(e); else console.error(e); }
          );
        };
      }
    });
  }

  window.google = {script: {run: runner(null, null)}};
})();
