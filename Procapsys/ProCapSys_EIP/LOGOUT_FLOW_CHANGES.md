# Logout Flow — What Changed

## New file
- **`frontend/auth-common.js`** — the single, shared logout/session
  implementation used by all four dashboards. Provides:
  1. The Single Active Session `fetch` guard (previously copy-pasted
     identically into all 4 HTML files — now defined once).
  2. `window.EIP.logout()` — the one logout function every role's
     button now calls. It POSTs to `/api/auth/logout`, clears
     client-side state, and redirects to `login.html`. If the request
     can't reach the server it surfaces a clear error instead of
     silently pretending the logout worked.
  3. A `pageshow`/`bfcache` guard that forces a reload if the page is
     restored from the back/forward cache, so the Back button can
     never show a stale authenticated view after logout.

  Included via `<script src="auth-common.js"></script>` near the top
  of the `<body>` script block in all four dashboards.

## Per-role bugs found and fixed

### Trainer — was throwing an actual error (the reported bug)
`signOut()` never called the backend at all. It showed a toast and
then did `window.location.href = 'index.html'` — a page that doesn't
exist anywhere in `frontend/`, so the browser hit Flask's catch-all
static route, got a 404, and rendered the app's JSON error response.
That's the "clicking Logout throws an error" bug.
**Fix:** `signOut()` now calls `EIP.logout()`, which invalidates the
session server-side and redirects to the real `login.html`.

### Super Admin — no reachable Logout button
The only related control was buried in Settings → "Log out of all
sessions," which called `DELETE /auth/sessions` — an endpoint that
does not exist anywhere in the backend (`login.py`, `superadmin.py`,
`app.py`), so it always failed silently.
**Fix:**
- Added a visible Logout button to the topbar (next to the profile
  chip), wired through the existing `data-action` dispatcher.
- Fixed the Settings page button to call `EIP.logout()` (the real,
  working endpoint) instead of the nonexistent sessions-DELETE route.
- Removed the dead `logOutAllSessions()` function.

### College Admin — logout button was a no-op
The sidebar Logout link was wired up, but its handler just closed the
confirm modal and showed a "Logged out successfully" toast — it never
called the backend, never cleared the auth cookie, and never
redirected. The admin stayed fully logged in.
**Fix:** `onConfirm` now calls `EIP.logout()`.

### Student — mostly correct already
`logout()` was already calling the real endpoint and redirecting.
Consolidated onto the shared `EIP.logout()` for consistency and to
pick up the new error handling + bfcache guard.

## Backend
No backend changes were required. `login.py`'s `POST /api/auth/logout`
and the single-active-session mechanism (JWT `sid` claim + blocklist
loader checked on every `@jwt_required()` route) were already correct:
auth is stored entirely in an httpOnly cookie (not localStorage), and
logout clears it via `unset_jwt_cookies` plus wipes the server-side
`currentSessionId` / `super_admin_sessions` record. Every protected API
route already 401s once that happens, which is what the shared
`fetch` guard listens for.

## Result
For all four roles (Super Admin, College Admin, Trainer, Student):
- A visible, working Logout control exists.
- Clicking it calls the real backend endpoint, clears the httpOnly
  JWT cookie and server-side session record.
- The user is redirected to `login.html`.
- Any further API call (including the very next page load) 401s and
  bounces back to Login automatically.
- Pressing Back after logout forces a reload instead of showing a
  cached authenticated page.
- Network/server failures during logout show a clear error instead of
  silently leaving the user in an inconsistent state.
