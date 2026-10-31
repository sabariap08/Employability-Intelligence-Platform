# Employability Intelligence Platform — Auth Backend

Flask + MongoDB authentication system wired to the existing, unmodified
`login.html` UI. Only two Python files, as requested:

- **`app.py`** — Flask app setup: Mongo, JWT, CORS, rate limiting, Firebase
  Admin init, blueprint registration, static file serving.
- **`login.py`** — all authentication logic: register, login, OTP,
  forgot/reset password, Google Sign-In, approval workflow.

Future role-specific modules (`superadmin.py`, `trainer.py`, `student.py`,
`collegeadmin.py`) can each expose their own `init_*(db=db, ...)` blueprint
factory and be registered in `app.py` next to `login.py`, without touching
this authentication code.

## 1. Prerequisites

- Python 3.10+
- A running MongoDB instance (local or Atlas)
- (Optional, for Google Sign-In) A Firebase project with a service-account key

## 2. Setup

```bash
cd eip-backend
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp .env.example .env
# then edit .env:
#   - set MONGO_URI to your MongoDB connection string
#   - set JWT_SECRET_KEY to a long random string
#     (python -c "import secrets; print(secrets.token_hex(32))")
#   - leave OTP_DEV_CODE=0000 for now (per spec)
#   - leave FIREBASE_CREDENTIALS_PATH blank unless you have a service-account key
```

## 3. Run

```bash
python app.py
```

The server starts on `http://localhost:5000` and serves the existing
frontend directly — visiting `http://localhost:5000/` loads `login.html`.
No separate frontend server or build step is required.

## 4. What's implemented

- Email/password registration for Student, Trainer, College Admin with
  full validation (email format, password complexity, mobile format,
  duplicate-email check) and a dynamic ID field (`rollNumber` /
  `employeeId` / `tneaCode`) matched to the selected role.
- Email OTP flow — `send-otp` / `verify-otp`, using the fixed development
  code `0000` (see `OTP_DEV_CODE` in `.env`). Verification issues a
  short-lived token that `register` / `reset-password` require, so the
  OTP step can't be skipped by calling those endpoints directly.
- Login — validates email → password → **selected role matches the
  registered role** → approval status, in that order, with the exact
  messages from the spec ("Wrong role selected.", "Your account is
  waiting for approval.", "Your account has been rejected.").
- Single hardcoded Super Admin account (`SUPER_ADMIN_EMAIL` /
  `SUPER_ADMIN_PASSWORD` in `.env`), checked directly with no DB lookup
  and no registration path.
- Forgot / reset password, same OTP-token pattern as registration.
- Google Sign-In via Firebase Admin (`/api/auth/google`) — verifies the
  ID token server-side, logs in existing users, and auto-creates new
  ones as `pending`. Returns `501` with a clear message if
  `FIREBASE_CREDENTIALS_PATH` isn't configured, instead of crashing.
- JWT auth, issued both as an httpOnly cookie and in the JSON response
  body, with role claims for role-based route protection
  (`@role_required(...)`).
- Approval workflow: College Admin / Trainer → only Super Admin can
  approve/reject; Student → Super Admin **or** Trainer can. Endpoints:
  `GET /api/auth/pending/college-admins`, `GET /api/auth/pending/trainers`,
  `GET /api/auth/pending/students`, `PUT /api/auth/approve/<id>`,
  `PUT /api/auth/reject/<id>`.
- Passwords hashed with bcrypt, never stored in plaintext.
- Rate limiting on all auth endpoints (Flask-Limiter), CORS restricted to
  `FRONTEND_ORIGIN`, all secrets/config in `.env`.

## 5. API summary

| Method | Path                              | Purpose                                |
|--------|-----------------------------------|-----------------------------------------|
| POST   | /api/auth/send-otp                | Send registration OTP                   |
| POST   | /api/auth/verify-otp              | Verify OTP, get a one-time otpToken     |
| POST   | /api/auth/register                | Create account (status: pending)        |
| POST   | /api/auth/login                   | Log in, get JWT + redirect target       |
| POST   | /api/auth/logout                  | Clear JWT cookie                        |
| POST   | /api/auth/forgot-password         | Send password-reset OTP                 |
| POST   | /api/auth/reset-password          | Reset password using otpToken           |
| POST   | /api/auth/google                  | Google Sign-In (Firebase ID token)      |
| GET    | /api/auth/me                      | Current user (JWT-protected)            |
| GET    | /api/auth/pending/college-admins  | Super Admin only                        |
| GET    | /api/auth/pending/trainers        | Super Admin only                        |
| GET    | /api/auth/pending/students        | Super Admin or Trainer                  |
| GET    | /api/auth/stats                   | Super Admin (full) or Trainer (students-only) |
| PUT    | /api/auth/approve/<user_id>       | Super Admin or Trainer (role-checked)   |
| PUT    | /api/auth/reject/<user_id>        | Super Admin or Trainer (role-checked)   |

## 6. Frontend discrepancies I did NOT silently paper over

Per your instructions I made **zero** changes to CSS, layout, animations,
or HTML markup. Two things in the uploaded files were worth flagging
instead of quietly working around invisibly:

1. **Role value mismatch.** `login.html`'s login-form role `<select>` uses
   `value="knowledge_admin"` for "College Admin", while the register form
   uses `value="college_admin"` for the same role. Left as-is in the HTML;
   `login.py` normalizes `knowledge_admin` → `college_admin` server-side
   (`ROLE_ALIASES`) so login still works correctly for that role.
2. **No Google button / no Firebase SDK in `login.html`.** The original
   brief asked for a "Continue with Google" button, but the uploaded page
   has neither the button nor the Firebase JS SDK included. Adding either
   would change the HTML structure, which you said not to do — so the
   backend endpoint (`POST /api/auth/google`) is fully built and ready,
   but nothing was inserted into the page. Tell me if you'd like the
   button + Firebase SDK `<script>` added and I'll wire it up.
3. **Redirect filenames.** The original brief said `superadmin.html` /
   `collegeadmin.html`; your actually-uploaded files are named
   `super_admin.html` / `college_admin.html`. The backend redirects to the
   real filenames you uploaded.

## 7. Not yet built (by design — out of scope for this pass)

- `superadmin.py`, `trainer.py`, `student.py`, `collegeadmin.py` — the
  dashboards themselves are still on mock in-page data (`localStorage` /
  in-memory arrays), as you can see from the `TODO`/`future endpoint`
  comments already in those files. Wiring those up is the natural next
  step once this auth layer is confirmed working end-to-end.
- A real transactional email service for OTPs (currently fixed at `0000`
  per your instructions).
