/* ============================================================
   ProCapSys AUTH-COMMON — shared session guard + logout module
   ============================================================
   Loaded by every authenticated dashboard (Super Admin, College Admin,
   Trainer, Student). This is the SINGLE implementation of:

     1. Single Active Session guard
        Wraps window.fetch so any /api/* call that comes back 401
        (this account got logged in elsewhere, or the session was
        otherwise invalidated) automatically clears local state and
        bounces the user back to Login with a one-time notice.

     2. EIP.logout()
        The one logout implementation every role's Logout button
        calls. POSTs to /api/auth/logout (clears the httpOnly JWT
        cookie + the server-side session record), clears any
        client-side cached/auth-adjacent state, then redirects to
        the Login page. If the request can't reach the server, the
        user is told clearly and is NOT silently left thinking they
        logged out when they didn't.

     3. Back/forward-cache (bfcache) guard
        If a browser restores this page from bfcache (e.g. pressing
        Back right after logging out), force a real reload so every
        on-load data fetch runs again — which re-triggers guard #1
        and bounces back to Login if the session is gone. Prevents
        "Back" from ever re-showing a stale authenticated page.

   Previously this ~15-line session-guard block was hand-copied into
   super_admin.html, college_admin.html, trainer.html and student.html
   independently (4 separate copies to keep in sync). It now lives
   here once. Include with:

       <script src="auth-common.js"></script>

   as EARLY as possible in <body>, before any other script that
   calls fetch() — it needs to install the wrapper first.
============================================================ */
(function () {
  'use strict';

  var LOGIN_PAGE = 'login.html';
  var nativeFetch = window.fetch.bind(window);
  var redirectingForSessionLoss = false;

  function goToLogin(params) {
    var qs = '';
    if (params) {
      var usp = new URLSearchParams(params);
      qs = '?' + usp.toString();
    }
    // replace(), not href — the login page should not be reachable
    // via the Back button from here, only re-entered fresh.
    window.location.replace(LOGIN_PAGE + qs);
  }

  function clearClientState() {
    try { sessionStorage.clear(); } catch (e) { /* ignore */ }
    try {
      // Auth itself lives entirely in the httpOnly JWT cookie the
      // backend clears via unset_jwt_cookies — nothing to remove
      // there from JS. localStorage on this app is only ever used
      // for cosmetic prefs (e.g. theme), which are safe/expected to
      // persist across accounts on a shared machine, so only
      // namespaced "eip_" keys other than the theme are swept, in
      // case any page ever starts caching per-user data there.
      Object.keys(localStorage).forEach(function (k) {
        if (k.indexOf('eip_') === 0 && k !== 'eip_theme') {
          localStorage.removeItem(k);
        }
      });
    } catch (e) { /* ignore */ }
  }

  // ------------------------------------------------------------
  // 1. SINGLE ACTIVE SESSION GUARD
  // ------------------------------------------------------------
  window.fetch = function (input, init) {
    return nativeFetch(input, init).then(function (res) {
      var url = typeof input === 'string' ? input : (input && input.url) || '';
      if (res.status === 401 && url.indexOf('/api/') !== -1 && !redirectingForSessionLoss) {
        redirectingForSessionLoss = true;
        var message = 'Your account has been logged in on another device.';
        res.clone().json().then(function (cloned) {
          if (cloned && cloned.message) message = cloned.message;
        }).catch(function () { /* ignore, use default message */ }).finally(function () {
          try { sessionStorage.setItem('eip_session_message', message); } catch (e) { /* ignore */ }
          nativeFetch('/api/auth/logout', { method: 'POST', credentials: 'include' })
            .catch(function () { /* ignore — we're leaving anyway */ })
            .finally(function () { goToLogin({ sessionExpired: '1' }); });
        });
      }
      return res;
    });
  };

  // ------------------------------------------------------------
  // 2. SHARED LOGOUT
  // ------------------------------------------------------------
  // opts.button: an HTMLElement to disable while the request is in
  //   flight and re-enable on failure (avoids double-submits and
  //   gives the user a visible "working on it" state).
  // opts.onError: optional custom error surface. If omitted, this
  //   falls back to a global showToast(msg, true) if the page
  //   defines one (every dashboard does), else window.alert.
  async function logout(opts) {
    opts = opts || {};
    var btn = opts.button;
    if (btn) { btn.disabled = true; }

    var reached = false;
    try {
      var res = await nativeFetch('/api/auth/logout', {
        method: 'POST',
        credentials: 'include',
      });
      // The backend's /logout handler is defensive by design (jwt
      // optional, wrapped so it always returns a JSON body) — a
      // non-OK response here means something genuinely went wrong
      // server-side, not just "already logged out".
      reached = !!res && res.ok;
    } catch (e) {
      reached = false;
    }

    if (!reached) {
      var msg = 'Could not reach the server to complete logout. Please check your connection and try again.';
      if (typeof opts.onError === 'function') {
        opts.onError(msg);
      } else if (typeof window.showToast === 'function') {
        try { window.showToast(msg, true); } catch (e) { window.alert(msg); }
      } else {
        window.alert(msg);
      }
      if (btn) { btn.disabled = false; }
      return false;
    }

    clearClientState();
    goToLogin();
    return true;
  }

  // ------------------------------------------------------------
  // 3. BACK/FORWARD-CACHE GUARD
  // ------------------------------------------------------------
  window.addEventListener('pageshow', function (event) {
    if (event.persisted) {
      window.location.reload();
    }
  });

  window.EIP = window.EIP || {};
  window.EIP.logout = logout;
})();
