// LAN token glue. Wraps the global fetch() so every API call from any page
// automatically includes the X-AIS-Token header that the server requires for
// write + export endpoints. Read endpoints don't require the header, so this is
// safe to set on all fetches.
//
// Token lifecycle on a tablet:
//   1. First visit: user is prompted for the token (whatever was set via
//      AIS_DASHBOARD_TOKEN on the server). Token is stored in localStorage.
//   2. Subsequent visits: token is read from localStorage. No prompt.
//   3. 401 from server: token cleared from localStorage; user re-prompted.
//
// The server probes its own auth on startup; if no token was configured, this
// script is harmless — the server accepts all callers and our header is just
// ignored.
(function () {
    'use strict';

    const STORAGE_KEY = 'aisDashboardToken';
    const HEADER = 'X-AIS-Token';

    function getStoredToken() {
        try { return localStorage.getItem(STORAGE_KEY) || ''; }
        catch { return ''; }
    }
    function setStoredToken(value) {
        try {
            if (value) localStorage.setItem(STORAGE_KEY, value);
            else localStorage.removeItem(STORAGE_KEY);
        } catch { /* private mode etc */ }
    }

    function promptForToken(message) {
        // Synchronous prompt — fine for a small internal LAN tool. Comment
        // the existing token (if any) so a forgetful operator doesn't get
        // stuck staring at a blank box.
        const existing = getStoredToken();
        const banner = message || 'Enter the dashboard access token (set on the server as AIS_DASHBOARD_TOKEN).';
        const entered = window.prompt(banner, existing);
        if (entered == null) return null; // user hit Cancel
        const trimmed = entered.trim();
        setStoredToken(trimmed);
        return trimmed;
    }

    // Only same-origin /api/* calls get the header. External fetches (none today,
    // but a defensive guard) are untouched so we never leak the token off-host.
    function shouldAttach(input) {
        try {
            const url = typeof input === 'string' ? input : (input && input.url) || '';
            if (!url) return false;
            if (url.startsWith('/')) return true;
            const u = new URL(url, location.origin);
            return u.origin === location.origin;
        } catch { return false; }
    }

    const originalFetch = window.fetch.bind(window);
    window.fetch = async function (input, init) {
        if (!shouldAttach(input)) {
            return originalFetch(input, init);
        }

        const opts = Object.assign({}, init || {});
        const headers = new Headers(opts.headers || (input instanceof Request ? input.headers : undefined));
        const token = getStoredToken();
        if (token) headers.set(HEADER, token);
        opts.headers = headers;

        let response = await originalFetch(input, opts);

        // On 401, retry once with a fresh prompt. If the user cancels or the new
        // token is also wrong, propagate the original 401 so the caller can show
        // an error rather than spinning forever.
        if (response.status === 401) {
            setStoredToken('');
            const refreshed = promptForToken('Access denied. Enter the dashboard token to continue.');
            if (refreshed) {
                headers.set(HEADER, refreshed);
                opts.headers = headers;
                response = await originalFetch(input, opts);
            }
        }
        return response;
    };

    // Helper exposed for pages that need to download a protected resource
    // (e.g. the AR billing CSV). <a href> can't carry custom headers, so we
    // fetch as a blob and trigger a save via a temporary <a download>.
    window.aisDownload = async function (url, suggestedName) {
        const res = await window.fetch(url);
        if (!res.ok) {
            const err = await res.json().catch(() => ({ error: 'Download failed (HTTP ' + res.status + ')' }));
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
        // Revoke after a short delay so the browser has time to start the download.
        setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
    };

    // Exposed so pages can offer a "change token" affordance if they want one.
    window.aisAuth = {
        getToken: getStoredToken,
        setToken: setStoredToken,
        prompt: promptForToken,
    };
})();
