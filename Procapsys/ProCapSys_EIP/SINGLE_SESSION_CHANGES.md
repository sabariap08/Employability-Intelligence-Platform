# Single Active Session — Implementation Summary

Only **7 files** were touched. No schema redesign, no endpoint URL/UI changes.

## Backend

### `login.py` (all the logic lives here, per the project's existing convention)
- New `sid` (session id) claim embedded in every JWT, generated fresh on
  every successful login (`_new_session_id()` / `_start_new_session()`).
- **Students / Trainers / College Admins**: session id is stored on the
  existing `db.users` document as `currentSessionId` (plus `lastLoginAt`,
  `lastLoginDevice` — the "minimal fields" the brief allowed). Reuses the
  existing collection, no schema redesign.
- **Super Admin**: since it's a hardcoded account with no `db.users` row,
  its session id lives in one tiny new single-document collection,
  `db.super_admin_sessions`.
- `login()` and `google_login()` now call `_start_new_session(...)` before
  issuing the token — every new login immediately overwrites the previous
  session id, which is exactly what invalidates the old token.
- `logout()` now clears the stored session id (super admin or user) in
  addition to unsetting cookies, and is `@jwt_required(optional=True)` so
  logout still works even if the token was already superseded.
- **The actual enforcement** is two JWTManager callbacks registered at the
  bottom of `init_auth(...)`:
  - `token_in_blocklist_loader` — compares the token's `sid` claim against
    the current value in the DB; also checks the user still exists and is
    `approved`. Returns "revoked" on any mismatch.
  - `revoked_token_loader` — returns `401` with
    `"Your account has been logged in from another device."`

Because these are registered on the app's single `JWTManager` instance,
**every** `@jwt_required()`-protected route in the whole app is covered
automatically — `superadmin.py`, `trainer.py`, `student.py`,
`collegeadmin.py`, `colleges.py`, `quiz_module.py`, `question_bank.py` all
reuse `role_required()` (from `login.py` or `quiz_common.py`), which wraps
`jwt_required()`. **None of those files needed to change.**

### `app.py`
- `app.config["JWT_BLOCKLIST_ENABLED"] = True` and
  `JWT_BLOCKLIST_TOKEN_CHECKS = ["access"]` (only access tokens exist in
  this app).
- `init_auth(...)` now also receives the `jwt` manager so it can register
  the callbacks above.

## Frontend
No markup, styling, or layout changed anywhere. One small `<script>`
addition per page:

- **`student.html` / `trainer.html` / `super_admin.html` /
  `college_admin.html`**: a self-invoking "session guard" wraps
  `window.fetch` once, at the very top of the page's script. Any `/api/*`
  call that comes back `401` (this only happens for a revoked/replaced
  session, since these pages never call the login endpoint) triggers:
  clear the stored session server-side (`POST /api/auth/logout`), remember
  the server's message, then redirect to `login.html?sessionExpired=1`. A
  `redirecting` flag prevents any repeat/loop.
- **`login.html`**: on load, checks for that flag/message and shows it in
  the existing alert UI (`showFormAlert`), then clears it — a manual
  refresh of the login page won't keep re-showing the notice.

Because every existing `apiGet`/`apiPost`/`apiCall`/`authApiGet`/etc.
helper on every dashboard is just a thin wrapper around `fetch(...)`,
wrapping `window.fetch` once covers **all** of them without editing each
call site individually.

## Behavior verified
A standalone test (Flask test client + mongomock, not shipped with the
project) confirmed:
- Student logs in on Device A → works. Same student logs in on Device B →
  Device A's next request gets `401` with the "logged in from another
  device" message; Device B keeps working.
- Same for the hardcoded Super Admin account.
- Logging out clears the session so the same token can't be reused
  afterward either.
