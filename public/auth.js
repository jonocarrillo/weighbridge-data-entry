// After login the server sets an HttpOnly cookie. Browser sends it automatically.
// This file only: (1) include credentials on fetch, (2) send to login on 401, (3) logout.
(function () {
  'use strict';

  function sameOrigin(input) {
    try {
      const url = typeof input === 'string' ? input : (input && input.url) || '';
      if (!url || url.startsWith('/')) return !!url || url === '';
      if (url.startsWith('/')) return true;
      return new URL(url, location.origin).origin === location.origin;
    } catch { return false; }
  }

  function goLogin() {
    const p = location.pathname;
    if (p === '/login.html' || p === '/login' || p === '/sign.html' || p === '/sign') return;
    if (document.body && document.body.dataset.aisAuthRedirecting) return;
    if (document.body) document.body.dataset.aisAuthRedirecting = '1';
    location.replace('/login.html?next=' + encodeURIComponent((p + location.search) || '/'));
  }

  const rawFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {
    const url = typeof input === 'string' ? input : (input && input.url) || '';
    const local = !url || url.startsWith('/') || (function () {
      try { return new URL(url, location.origin).origin === location.origin; }
      catch { return false; }
    })();
    if (!local) return rawFetch(input, init);

    const opts = init ? Object.assign({}, init) : {};
    opts.credentials = opts.credentials || 'same-origin';
    return rawFetch(input, opts).then(function (res) {
      if (res.status === 401 && url.indexOf('/api/login') === -1) goLogin();
      return res;
    });
  };

  window.aisDownload = async function (url, suggestedName) {
    const res = await window.fetch(url);
    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: 'Download failed' }));
      throw new Error(err.error || 'Download failed');
    }
    const blob = await res.blob();
    const objectUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objectUrl;
    a.download = suggestedName || 'download.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(objectUrl); }, 1000);
  };

  window.aisLogout = async function () {
    try {
      await rawFetch('/api/logout', { method: 'POST', credentials: 'same-origin' });
    } catch { /* ignore */ }
    location.replace('/login.html');
  };

  window.aisAuth = { logout: window.aisLogout, loginUrl: '/login.html' };
})();
