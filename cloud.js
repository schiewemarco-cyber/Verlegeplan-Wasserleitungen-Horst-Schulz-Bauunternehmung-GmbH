// Cloud-Anbindung: Pläne in OneDrive oder SharePoint speichern und öffnen (Microsoft Graph).
// Wird nur aktiv, wenn in config.js eine clientId eingetragen ist.
(function () {
  'use strict';
  var C = window.VP_CLOUD || {};
  if (!C.clientId) return;

  var EXT = '.vplan.json', G = 'https://graph.microsoft.com/v1.0';
  var SP = C.mode === 'sharepoint';
  var SC = (C.scopes && C.scopes.length) ? C.scopes : (SP ? ['Files.ReadWrite.All'] : ['Files.ReadWrite']);
  var FOLDER = (C.folder || 'Verlegeplaene').replace(/^\/+|\/+$/g, '');
  var RU = C.redirectUri || (location.origin + location.pathname.replace(/[^\/]*$/, ''));
  // MSAL zuerst aus dem eigenen Ordner laden (falls abgelegt), sonst von jsDelivr (feste Version)
  var MSAL_LOCAL = 'msal-browser.min.js';
  var MSAL_CDN = 'https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.37.1/lib/msal-browser.min.js';
  var app = null, acct = null, baseP = null, initP = null, CE = {};

  function $(i) { return document.getElementById(i); }
  function enc(p) { return p.split('/').map(encodeURIComponent).join('/'); }
  function who(m) { return m && m.lastModifiedBy && m.lastModifiedBy.user ? m.lastModifiedBy.user.displayName : '?'; }
  function when(m) { return new Date(m.lastModifiedDateTime).toLocaleString('de-DE'); }
  function say(t) {
    var m = $('cmsg'); if (!m) return;
    m.textContent = t; if (typeof msg === 'function') msg(t); clearTimeout(say.t);
    say.t = setTimeout(function () { m.textContent = ''; }, 7000);
  }
  function err(e) {
    if (!e) return 'Unbekannter Fehler';
    if (e.status === 401 || e.status === 403) return 'Keine Berechtigung für die Cloud (' + e.status + ')';
    if (!navigator.onLine || /failed to fetch|networkerror|load failed/i.test(e.message || '')) return 'Keine Verbindung zur Cloud';
    return e.message || String(e);
  }
  function ok(r) {
    if (r.ok) return Promise.resolve(r);
    return r.json().catch(function () { return {}; }).then(function (j) {
      var e = new Error((j.error && j.error.message) || ('HTTP ' + r.status)); e.status = r.status; throw e;
    });
  }
  function json(r) { return r.json(); }

  function loadScript(src) {
    return new Promise(function (res, rej) {
      var s = document.createElement('script'); s.src = src; s.onload = res;
      s.onerror = function () { s.remove(); rej(new Error('Skript nicht ladbar: ' + src)); };
      document.head.appendChild(s);
    });
  }
  function getMsal() {
    if (window.msal) return Promise.resolve();
    return loadScript(MSAL_LOCAL).then(function () { if (!window.msal) throw 0; })
      .catch(function () { return loadScript(MSAL_CDN); })
      .then(function () { if (!window.msal) throw new Error('Anmelde-Bibliothek (MSAL) nicht verfügbar'); });
  }

  function init() {
    if (initP) return initP;
    initP = getMsal().then(function () {
      app = new msal.PublicClientApplication({
        auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + (C.tenantId || 'organizations'), redirectUri: RU, postLogoutRedirectUri: RU },
        cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: false }
      });
      return app.initialize ? app.initialize() : null;
    }).then(function () { return app.handleRedirectPromise(); })
      .then(function (r) {
        acct = (r && r.account) || app.getAllAccounts()[0] || null;
        if (acct) app.setActiveAccount(acct);
        ui(); if (acct) refresh();
      }).catch(function (e) { initP = null; say(err(e)); });
    return initP;
  }
  function login() { init().then(function () { if (app && !acct) return app.loginRedirect({ scopes: SC }); }); }
  function logout() { if (app && acct) app.logoutRedirect({ account: acct, postLogoutRedirectUri: RU }); }

  function token() {
    var req = { scopes: SC, account: acct };
    return app.acquireTokenSilent(req).then(function (r) { return r.accessToken; }).catch(function (e) {
      if (window.msal && e instanceof msal.InteractionRequiredAuthError) { app.acquireTokenRedirect(req); return new Promise(function () {}); }
      throw e;
    });
  }
  function gr(method, path, body, hdr) {
    return token().then(function (t) {
      var h = { Authorization: 'Bearer ' + t };
      if (body !== undefined) h['Content-Type'] = 'application/json';
      if (hdr) Object.keys(hdr).forEach(function (k) { h[k] = hdr[k]; });
      return fetch(G + path, { method: method, headers: h, body: body });
    });
  }

  function spDrive() {
    var s = C.site || {};
    if (!s.host || !s.path) return Promise.reject(new Error('SharePoint-Site ist in config.js nicht eingetragen'));
    var sp = '/sites/' + s.host + ':' + (s.path.charAt(0) === '/' ? '' : '/') + s.path;
    if (!s.library) return gr('GET', sp + ':/drive?$select=id').then(ok).then(json).then(function (j) { return '/drives/' + j.id; });
    return gr('GET', sp + ':/drives?$select=id,name').then(ok).then(json).then(function (j) {
      var d = (j.value || []).filter(function (x) { return x.name === s.library; })[0];
      if (!d) throw new Error('Bibliothek „' + s.library + '“ nicht gefunden');
      return '/drives/' + d.id;
    });
  }
  function base() {
    if (baseP) return baseP;
    baseP = (SP ? spDrive() : Promise.resolve('/me/drive')).then(function (d) { return d + '/root:/' + enc(FOLDER); });
    baseP.catch(function () { baseP = null; });
    return baseP;
  }

  function fname() {
    var n = ($('bv').value || $('str').value || 'Plan').replace(/[\\\/:*?"<>|#%~]+/g, '_').replace(/[. ]+$/, '').trim();
    return (n || 'Plan').slice(0, 80) + EXT;
  }
  function label(name) { return name.slice(0, -EXT.length); }

  var CL = [], IDXN = '_index.json';
  function idxRead() {
    return base().then(function (B) {
      return gr('GET', B + '/' + IDXN + '?$select=eTag').then(function (r) {
        if (r.status === 404) return { et: null, d: { v: 1, plans: {} } };
        return ok(r).then(json).then(function (m) {
          return gr('GET', B + '/' + IDXN + ':/content').then(ok).then(json).then(function (d) {
            return { et: m.eTag, d: (d && d.plans) ? d : { v: 1, plans: {} } };
          });
        });
      });
    });
  }
  function idxUpdate(fn) {
    function once() {
      return idxRead().then(function (x) {
        fn(x.d.plans);
        return base().then(function (B) {
          return gr('PUT', B + '/' + IDXN + ':/content', JSON.stringify(x.d), x.et ? { 'If-Match': x.et } : null);
        });
      }).then(function (r) { if (r.status === 412) throw { retry: 1 }; return ok(r); });
    }
    return once().catch(function (e) { if (e && e.retry) return once(); throw e; }).catch(function () {});
  }
  function renderCloud(sel) {
    var q = ($('cq').value || '').toLowerCase().trim(), st = $('cfs').value, cur = sel || $('cls').value, o = '', n = 0;
    CL.forEach(function (f) {
      var m = f.m, txt = [label(f.name), m.str, m.ks, m.zr, m.st, f.by].join(' ').toLowerCase();
      if (q && txt.indexOf(q) < 0) return;
      if (st && m.st !== st) return;
      n++;
      o += '<option value="' + esc(f.name) + '">' + esc(label(f.name)) + (m.str && m.str !== label(f.name) ? ' · ' + esc(m.str) : '') +
        (m.st ? ' · ' + esc(m.st) : '') + ' · ' + new Date(f.t).toLocaleDateString('de-DE') + (f.by && f.by !== '?' ? ' · ' + esc(f.by) : '') + '</option>';
    });
    $('cls').innerHTML = o; $('cls').value = cur; $('ccnt').textContent = n + ' von ' + CL.length + ' Plänen';
  }
  function refresh() {
    if (!acct) return Promise.resolve();
    return base().then(function (B) {
      return gr('GET', B + ':/children?$select=name,eTag,lastModifiedDateTime,lastModifiedBy&$top=200');
    }).then(function (r) { return r.status === 404 ? { value: [] } : ok(r).then(json); })
      .then(function (j) {
        return idxRead().catch(function () { return { d: { plans: {} } }; }).then(function (x) {
          CL = (j.value || []).filter(function (f) { return f.name && f.name.slice(-EXT.length) === EXT; })
            .sort(function (a, b) { return a.lastModifiedDateTime < b.lastModifiedDateTime ? 1 : -1; })
            .map(function (f) { return { name: f.name, t: f.lastModifiedDateTime, by: who(f), m: x.d.plans[f.name] || {} }; });
          renderCloud();
        });
      }).catch(function (e) { say(err(e)); });
  }

  function cSave() {
    if (!acct) return;
    if (!navigator.onLine) { say('Offline – Cloud nicht erreichbar'); return; }
    var name = fname(), B, ss = state(), d = JSON.parse(ss);
    d.v = 1; d.savedBy = acct.name || acct.username || ''; d.savedAt = new Date().toISOString();
    base().then(function (b) { B = b; return gr('GET', B + '/' + enc(name) + '?$select=eTag,lastModifiedDateTime,lastModifiedBy'); })
      .then(function (r) {
        if (r.status === 404) return true;
        return ok(r).then(json).then(function (m) {
          if (CE[name] === undefined) return confirm('In der Cloud gibt es schon einen Plan „' + label(name) + '“ (zuletzt: ' + who(m) + ', ' + when(m) + ').\nÜberschreiben?');
          if (m.eTag !== CE[name]) return confirm('Der Plan wurde in der Cloud inzwischen geändert (' + who(m) + ', ' + when(m) + ').\nTrotzdem überschreiben? Die Änderungen der anderen gehen dann verloren.');
          return true;
        });
      })
      .then(function (go) { if (!go) throw { cancel: 1 }; return gr('PUT', B + '/' + enc(name) + ':/content', JSON.stringify(d)); })
      .then(ok).then(json)
      .then(function (j) { CE[name] = j.eTag; if (window.markSaved) markSaved(ss); say('In Cloud gespeichert ✓'); toggle(false); return idxUpdate(function (p) { p[name] = { bv: d.bv || '', str: d.str || '', ks: d.ks || '', zr: d.zr || '', st: d.st || '', by: acct.name || acct.username || '' }; }).then(refresh).then(function () { renderCloud(name); }); })
      .catch(function (e) { if (!e || !e.cancel) say(err(e)); });
  }

  function cOpen() {
    var name = $('cls').value; if (!name || !acct) return;
    if (S.items && S.items.length && !confirm('Den aktuellen Plan durch „' + label(name) + '“ aus der Cloud ersetzen?\nNicht in der Cloud gespeicherte Änderungen gehen verloren.')) return;
    var B, et;
    base().then(function (b) { B = b; return gr('GET', B + '/' + enc(name) + '?$select=eTag'); })
      .then(ok).then(json)
      .then(function (m) { et = m.eTag; return gr('GET', B + '/' + enc(name) + ':/content'); })
      .then(ok).then(json)
      .then(function (d) { applyState(d); CE[name] = et; if (window.markSaved) markSaved(state()); say('Plan geöffnet ✓'); toggle(false); })
      .catch(function (e) { say(err(e)); });
  }

  function cDel() {
    var name = $('cls').value; if (!name || !acct) return;
    if (!confirm('Plan „' + label(name) + '“ in der Cloud löschen?')) return;
    base().then(function (B) { return gr('DELETE', B + '/' + enc(name)); }).then(ok)
      .then(function () { delete CE[name]; say('Gelöscht'); return idxUpdate(function (p) { delete p[name]; }).then(refresh); })
      .catch(function (e) { say(err(e)); });
  }

  var CLOUD_SVG = '<svg width="22" height="22" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true"><path d="M19.35 10.04A7.49 7.49 0 0 0 12 4C9.11 4 6.6 5.64 5.35 8.04A5.994 5.994 0 0 0 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96z"/></svg>';
  function toggle(open) {
    var p = $('cpan'); if (!p) return;
    p.classList.toggle('open', open === undefined ? !p.classList.contains('open') : open);
  }
  function ui() {
    var on = !!acct, nm = on ? (acct.name || acct.username) : '';
    $('cbtn').classList.toggle('on', on);
    $('cbtn').title = on ? 'Cloud: angemeldet als ' + nm : 'Cloud: nicht angemeldet';
    $('cst').textContent = on ? 'Angemeldet als ' + nm : 'Nicht angemeldet';
    $('cin').style.display = on ? 'none' : '';
    $('cwork').style.display = on ? 'flex' : 'none';
    if (window.autoSave) autoSave();
  }

  window.cloudState = function () { return !!acct; };
  window.cloudDirty = function (d) {
    var b = $('cbtn'); if (!b) return;
    var on = !!d && !!acct, t = $('cdt');
    b.classList.toggle('dirty', on);
    if (t) t.textContent = on ? '⚠ Änderungen noch nicht in der Cloud gesichert' : '';
    b.title = !acct ? 'Cloud: nicht angemeldet' : (on ? 'Cloud: angemeldet – Änderungen nicht gesichert' : 'Cloud: angemeldet als ' + (acct.name || acct.username));
  };

  var slot = $('cloudslot');
  if (!slot) return;
  slot.innerHTML = '<button id="cbtn" class="cloudb" title="Cloud">' + CLOUD_SVG + '</button>' +
    '<div id="cpan" class="cmenu"><b id="cst"></b><small id="cdt" style="color:#b45f06"></small>' +
    '<button id="cin" class="p">Mit Firmenkonto anmelden</button>' +
    '<div id="cwork"><button id="csv" class="p">In Cloud speichern</button>' +
    '<input id="cq" type="search" placeholder="Suchen: Name, Straße, Person …"><select id="cfs"><option value="">Alle Status</option><option>In Planung</option><option>In Ausführung</option><option>Gebaut</option><option>Abgenommen</option></select><select id="cls" size="6"></select><small id="ccnt" style="color:var(--mu)"></small>' +
    '<div class="row"><button id="cld">Öffnen</button><button id="cdel" title="Plan in der Cloud löschen">🗑 Löschen</button><button id="crf" title="Liste aktualisieren">↻</button></div>' +
    '<button id="cout">Abmelden</button></div><span id="cmsg"></span></div>';
  $('cbtn').onclick = function () { toggle(); if (acct && $('cpan').classList.contains('open')) refresh(); };
  document.addEventListener('click', function (e) {
    var t = e.target; if (t && t.closest && !t.closest('#cpan') && !t.closest('#cbtn')) toggle(false);
  });
  $('cin').onclick = login; $('cout').onclick = logout; $('csv').onclick = cSave;
  $('cld').onclick = cOpen; $('cdel').onclick = cDel; $('crf').onclick = function () { refresh(); };
  $('cq').oninput = function () { renderCloud(); }; $('cfs').onchange = function () { renderCloud(); };
  ui();

  // Rückkehr von der Microsoft-Anmeldung oder früher angemeldet: sofort initialisieren
  var back = false;
  try { back = /[#&](code|error)=/.test(location.hash) || Object.keys(localStorage).some(function (k) { return k.indexOf('msal.') === 0; }); } catch (e) {}
  if (back) init();
})();
