"""
============================================================
 login.py — Authentication module
============================================================
Everything related to identity lives here:

  - Register / Login / Logout
  - Email OTP send + verify (dev OTP = "0000", see .env)
  - Forgot password / Reset password
  - Google Sign-In (Firebase Authentication)
  - Password / email / mobile validation helpers
  - Duplicate email checking
  - Password hashing (bcrypt)
  - JWT issuing + role-based access control
  - Approval status checks at login time
  - The College Admin / Trainer / Student approval workflow

This module exposes a single factory, `init_auth(...)`, which app.py
calls once at startup. That keeps login.py fully self-contained and
testable, and means future modules (superadmin.py, trainer.py,
student.py, collegeadmin.py) can be added later by following the same
factory pattern — none of them need to modify this file.
============================================================
"""

import os
import re
import secrets
from datetime import datetime, timedelta, timezone

from flask import Blueprint, request, jsonify
from flask_jwt_extended import (
    create_access_token,
    set_access_cookies,
    unset_jwt_cookies,
    jwt_required,
    get_jwt,
    get_jwt_identity,
)
from bson import ObjectId
from bson.errors import InvalidId
from pymongo.errors import DuplicateKeyError

try:
    from firebase_admin import auth as firebase_auth
except ImportError:
    firebase_auth = None

from colleges import resolve_active_college, resolve_active_department


# ------------------------------------------------------------
# CONSTANTS / CONFIG (env-driven, dev defaults kept explicit)
# ------------------------------------------------------------
OTP_DEV_CODE = os.getenv("OTP_DEV_CODE", "2026")
OTP_EXPIRY_MINUTES = int(os.getenv("OTP_EXPIRY_MINUTES", "10"))
OTP_TOKEN_EXPIRY_MINUTES = int(os.getenv("OTP_TOKEN_EXPIRY_MINUTES", "15"))

SUPER_ADMIN_EMAIL = os.getenv("SUPER_ADMIN_EMAIL", "login@superadmin.in")
SUPER_ADMIN_PASSWORD = os.getenv("SUPER_ADMIN_PASSWORD", "Superadmin@123")

DEFAULT_GOOGLE_SIGNUP_ROLE = os.getenv("DEFAULT_GOOGLE_SIGNUP_ROLE", "student")

# Roles that self-register through the Create Account form.
REGISTERABLE_ROLES = {"student", "trainer", "college_admin"}

# The frontend's login-page <select> has a mismatched option value
# ("knowledge_admin") for College Admin compared to the register form
# ("college_admin"). Rather than edit the shipped HTML, we normalize
# it here so login still works correctly for that role.
ROLE_ALIASES = {
    "knowledge_admin": "college_admin",
}

# Redirect targets match the ACTUAL uploaded filenames
# (college_admin.html / super_admin.html / trainer.html / student.html),
# not the collegeadmin.html / superadmin.html names in the original brief.
ROLE_REDIRECTS = {
    "super_admin": "super_admin.html",
    "college_admin": "college_admin.html",
    "trainer": "trainer.html",
    "student": "student.html",
}

# Which role's dynamic "ID field" maps to which DB column.
ID_FIELD_BY_ROLE = {
    "student": "rollNumber",
    "trainer": "employeeId",
    "college_admin": "tneaCode",
}

EMAIL_REGEX = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
MOBILE_REGEX = re.compile(r"^\d{10}$")
# Min 8 chars, at least 1 upper, 1 lower, 1 digit, 1 special char.
PASSWORD_REGEX = re.compile(
    r"^(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$"
)


def validate_password(password):
    """Shared password rule — used by register / reset-password /
    complete-first-login and the Super Admin's direct trainer creation."""
    if not password or not PASSWORD_REGEX.match(password):
        return (
            "Password must be at least 8 characters and include an "
            "uppercase letter, a lowercase letter, a number, and a "
            "special character."
        )
    return None


def validate_email(email):
    if not email or not EMAIL_REGEX.match(email):
        return "Enter a valid email address."
    return None


def validate_mobile(mobile):
    if not mobile or not MOBILE_REGEX.match(mobile):
        return "Mobile number must be exactly 10 digits."
    return None


def _now():
    return datetime.utcnow()


# Fixed key for the single-document collection that tracks the hardcoded
# Super Admin's active session (Super Admin has no row in db.users, since
# its credentials come from the environment, not the database — see the
# "Super Admin: hardcoded" branch in login() below).
SUPER_ADMIN_SESSION_KEY = "super_admin"


def _new_session_id():
    """Opaque, unguessable identifier for a single login session (Single
    Active Session feature). A fresh one is minted on every successful
    login and stored server-side; it travels inside the JWT as the "sid"
    claim so every subsequent request can be checked against it."""
    return secrets.token_hex(20)


def normalize_role(role):
    if not role:
        return role
    role = role.strip().lower()
    return ROLE_ALIASES.get(role, role)


def init_auth(bcrypt, db, limiter, jwt, firebase_ready=False):
    """
    Build and return the auth Blueprint.

    Parameters
    ----------
    bcrypt : flask_bcrypt.Bcrypt      — for hashing / checking passwords
    db     : pymongo.database.Database — the MongoDB database handle
    limiter: flask_limiter.Limiter    — for per-route rate limiting
    jwt    : flask_jwt_extended.JWTManager — used to register the
             single-active-session callbacks (token_in_blocklist_loader /
             revoked_token_loader) so every @jwt_required()-protected
             route across the whole app (not just this blueprint) enforces
             single-session automatically, with zero changes required in
             superadmin.py / trainer.py / student.py / collegeadmin.py /
             colleges.py / quiz_module.py / question_bank.py.
    firebase_ready : bool             — whether Firebase Admin was initialized
    """
    auth_bp = Blueprint("auth", __name__)

    users = db.users
    otp_verifications = db.otp_verifications
    # Single-document collection: tracks the current session id for the
    # hardcoded Super Admin account only (it has no db.users row).
    super_admin_sessions = db.super_admin_sessions

    # --------------------------------------------------------
    # Small internal helpers
    # --------------------------------------------------------
    def error(message, status=400):
        return jsonify(success=False, message=message), status

    def ok(payload=None, message=None, status=200):
        body = {"success": True}
        if message:
            body["message"] = message
        if payload:
            body.update(payload)
        return jsonify(body), status

    def find_user_by_email(email):
        return users.find_one({"email": email.strip().lower()})

    def public_user(doc):
        """Strip sensitive/internal fields before sending a user back to the client."""
        if not doc:
            return None
        created_at = doc.get("createdAt")
        return {
            "id": str(doc["_id"]),
            "fullName": doc.get("fullName"),
            "email": doc.get("email"),
            "mobile": doc.get("mobile"),
            "role": doc.get("role"),
            "college": doc.get("college"),
            "collegeId": str(doc["collegeId"]) if doc.get("collegeId") else None,
            "department": doc.get("department"),
            "departmentId": str(doc["departmentId"]) if doc.get("departmentId") else None,
            "district": doc.get("district"),
            "rollNumber": doc.get("rollNumber"),
            "employeeId": doc.get("employeeId"),
            "tneaCode": doc.get("tneaCode"),
            "cohort": doc.get("cohort") or "entry_level",
            "approvalStatus": doc.get("approvalStatus"),
            "googleLogin": doc.get("googleLogin", False),
            "createdAt": created_at.isoformat() if created_at else None,
        }

    def issue_token_response(role, identity, session_id, extra_claims=None):
        # "sid" is the single-active-session marker: every request re-checks
        # this value against the latest one stored server-side (see the
        # token_in_blocklist_loader registered below). A brand new login
        # always mints a brand new sid, which immediately makes every
        # previously issued token for this account fail that check.
        claims = {"role": role, "sid": session_id}
        if extra_claims:
            claims.update(extra_claims)
        token = create_access_token(identity=identity, additional_claims=claims)
        return token

    def _start_new_session(role, identity):
        """Mint a new session id and persist it as THE current session for
        this account, replacing whatever was there before. Returns the new
        session id (to embed in the JWT claims)."""
        session_id = _new_session_id()
        if role == "super_admin":
            super_admin_sessions.update_one(
                {"_id": SUPER_ADMIN_SESSION_KEY},
                {"$set": {"sessionId": session_id, "lastLoginAt": _now()}},
                upsert=True,
            )
        else:
            users.update_one(
                {"_id": ObjectId(identity)},
                {
                    "$set": {
                        "currentSessionId": session_id,
                        "lastLoginAt": _now(),
                        "lastLoginDevice": (request.headers.get("User-Agent") or "")[:255],
                    }
                },
            )
        return session_id

    def role_required(*allowed_roles):
        """Decorator: restrict a route to one or more JWT roles."""
        def decorator(fn):
            @jwt_required()
            def wrapper(*args, **kwargs):
                claims = get_jwt()
                if claims.get("role") not in allowed_roles:
                    return error("You are not authorized to perform this action.", 403)
                return fn(*args, **kwargs)
            wrapper.__name__ = fn.__name__
            return wrapper
        return decorator

    # ==========================================================
    # OTP — SEND
    # ==========================================================
    @auth_bp.route("/send-otp", methods=["POST"])
    @limiter.limit("5 per minute")
    def send_otp():
        data = request.get_json(silent=True) or {}
        email = (data.get("email") or "").strip().lower()
        purpose = (data.get("purpose") or "register").strip().lower()

        if purpose not in ("register", "forgot", "first_login"):
            return error("Invalid OTP purpose.")

        email_err = validate_email(email)
        if email_err:
            return error(email_err)

        existing_user = find_user_by_email(email)
        if purpose == "register" and existing_user:
            return error("An account with this email already exists.")
        if purpose in ("forgot", "first_login") and not existing_user:
            return error("No account found with this email.")

        # Upsert a fresh OTP record. In production this is where a real
        # email-sending service would be called; for now the OTP is the
        # fixed development code below.
        otp_verifications.update_one(
            {"email": email, "purpose": purpose},
            {
                "$set": {
                    "email": email,
                    "purpose": purpose,
                    "otp": OTP_DEV_CODE,
                    "verified": False,
                    "token": None,
                    "createdAt": _now(),
                    "expiresAt": _now() + timedelta(minutes=OTP_EXPIRY_MINUTES),
                }
            },
            upsert=True,
        )
        return ok(message="OTP sent.")

    # ==========================================================
    # OTP — VERIFY
    # ==========================================================
    @auth_bp.route("/verify-otp", methods=["POST"])
    @limiter.limit("10 per minute")
    def verify_otp():
        data = request.get_json(silent=True) or {}
        email = (data.get("email") or "").strip().lower()
        purpose = (data.get("purpose") or "register").strip().lower()
        otp = (data.get("otp") or "").strip()

        record = otp_verifications.find_one({"email": email, "purpose": purpose})
        if not record or record["expiresAt"] < _now():
            return error("OTP expired or not requested. Please resend.")

        if otp != record["otp"]:
            return error("Incorrect OTP.")

        token = secrets.token_urlsafe(24)
        otp_verifications.update_one(
            {"_id": record["_id"]},
            {
                "$set": {
                    "verified": True,
                    "token": token,
                    "tokenExpiresAt": _now() + timedelta(minutes=OTP_TOKEN_EXPIRY_MINUTES),
                }
            },
        )
        return ok({"otpToken": token}, message="OTP verified.")

    def _consume_otp_token(email, purpose, token):
        """Validate + burn a one-time OTP verification token. Returns error string or None."""
        record = otp_verifications.find_one({"email": email, "purpose": purpose})
        if (
            not record
            or not record.get("verified")
            or record.get("token") != token
            or record.get("tokenExpiresAt") is None
            or record["tokenExpiresAt"] < _now()
        ):
            return "OTP verification required or has expired. Please verify again."
        otp_verifications.delete_one({"_id": record["_id"]})
        return None

    # ==========================================================
    # REGISTER
    # ==========================================================
    @auth_bp.route("/register", methods=["POST"])
    @limiter.limit("10 per minute")
    def register():
        data = request.get_json(silent=True) or {}

        full_name = (data.get("fullName") or "").strip()
        role = normalize_role(data.get("role"))
        email = (data.get("email") or "").strip().lower()
        password = data.get("password") or ""
        confirm_password = data.get("confirmPassword") or ""
        mobile = (data.get("mobile") or "").strip()
        otp_token = data.get("otpToken") or ""

        if not full_name:
            return error("Full name is required.")
        if role not in REGISTERABLE_ROLES:
            return error("Invalid role selected.")

        email_err = validate_email(email)
        if email_err:
            return error(email_err)

        password_err = validate_password(password)
        if password_err:
            return error(password_err)
        if password != confirm_password:
            return error("Passwords do not match.")

        mobile_err = validate_mobile(mobile)
        if mobile_err:
            return error(mobile_err)

        # --------------------------------------------------------
        # Role-specific fields. The frontend form is different per
        # role, so validation branches here instead of one shared
        # shape (district no longer exists anywhere; trainers no
        # longer pick a college/department at registration time).
        # --------------------------------------------------------
        role_fields = {
            "college": None,
            "collegeId": None,
            "department": None,
            "departmentId": None,
            "district": None,
            "rollNumber": None,
            "employeeId": None,
            "tneaCode": None,
        }

        if role == "student":
            college_id = (data.get("collegeId") or data.get("college") or "").strip()
            department_id = (data.get("departmentId") or data.get("department") or "").strip()
            roll_number = (data.get("rollNumber") or data.get("idValue") or "").strip()

            if not college_id:
                return error("College is required.")
            if not department_id:
                return error("Department is required.")
            if not roll_number:
                return error("Roll Number is required.")

            college_doc = resolve_active_college(db, college_id)
            if not college_doc:
                return error("Selected college is invalid or inactive.")
            department_doc = resolve_active_department(db, department_id, college_id)
            if not department_doc:
                return error("Selected department is invalid or inactive for this college.")

            if users.find_one({"role": "student", "rollNumber": roll_number}):
                return error("An account with this Roll Number already exists.")

            role_fields.update({
                "college": college_doc["college_name"],
                "collegeId": college_doc["_id"],
                "department": department_doc["department_name"],
                "departmentId": department_doc["_id"],
                "rollNumber": roll_number,
            })

        elif role == "trainer":
            # Trainers do NOT select college/department at registration —
            # the Super Admin assigns these after approval
            # (see PUT /api/admin/trainers/<id>/assign in colleges.py).
            employee_id = (data.get("employeeId") or data.get("idValue") or "").strip()
            if not employee_id:
                return error("Employee ID is required.")
            if users.find_one({"role": "trainer", "employeeId": employee_id}):
                return error("An account with this Employee ID already exists.")
            role_fields["employeeId"] = employee_id

        elif role == "college_admin":
            # College is selected from the same live db.colleges dropdown as
            # everywhere else on the platform (see /api/public/colleges in
            # colleges.py) — never free text, and never hardcoded on the
            # frontend, so newly added colleges show up automatically.
            # (The old "district" free-text field was never actually
            # rendered anywhere in the registration form, which meant this
            # role could never successfully register; it has been dropped
            # in favor of the same collegeId-based flow every other
            # college-scoped role already uses.)
            college_id = (data.get("collegeId") or data.get("college") or "").strip()
            id_value = (data.get("tneaCode") or data.get("idValue") or "").strip()

            if not college_id:
                return error("College is required.")
            if not id_value:
                return error(f"{ID_FIELD_BY_ROLE[role]} is required.")

            college_doc = resolve_active_college(db, college_id)
            if not college_doc:
                return error("Selected college is invalid or inactive.")

            if users.find_one({"role": "college_admin", "tneaCode": id_value}):
                return error("An account with this identifier already exists.")

            role_fields.update({
                "college": college_doc["college_name"],
                "collegeId": college_doc["_id"],
                "tneaCode": id_value,
            })

        otp_err = _consume_otp_token(email, "register", otp_token)
        if otp_err:
            return error(otp_err)

        if find_user_by_email(email):
            return error("An account with this email already exists.")

        password_hash = bcrypt.generate_password_hash(password).decode("utf-8")

        user_doc = {
            "fullName": full_name,
            "email": email,
            "mobile": mobile,
            "role": role,
            "passwordHash": password_hash,
            "approvalStatus": "pending",
            "approvedBy": None,
            "approvedDate": None,
            "googleLogin": False,
            # Soft-delete flag — never set at registration time, only by
            # the Super Admin's "Delete" action (see superadmin.py
            # soft_delete_user). login() below rejects any account where
            # this is true, before it even looks at approvalStatus.
            "isDeleted": False,
            # Students start in Entry Level: cohort is None until BOTH their
            # baseline assessment AND manual interview are scored (see
            # quiz_common.check_and_generate_cohort in the quiz engine).
            "cohort": None,
            "baselineAssessmentScore": None,
            "interviewScore": None,
            "finalEmployabilityScore": None,
            "cohortAssignedAt": None,
            "createdAt": _now(),
            "updatedAt": _now(),
        }
        user_doc.update(role_fields)

        try:
            users.insert_one(user_doc)
        except DuplicateKeyError:
            return error(
                "An account with this email or identifier (Roll Number / "
                "Employee ID / TNEA Code) already exists."
            )

        return ok(
            message="Account created successfully. Your account is waiting for approval.",
            status=201,
        )

    # ==========================================================
    # LOGIN
    # ==========================================================
    @auth_bp.route("/login", methods=["POST"])
    @limiter.limit("15 per minute")
    def login():
        data = request.get_json(silent=True) or {}
        selected_role = normalize_role(data.get("role"))
        email = (data.get("email") or "").strip().lower()
        password = data.get("password") or ""

        if not selected_role or not email or not password:
            return error("Role, email and password are required.")

        # -------- Super Admin: hardcoded, single account, no DB lookup --------
        if selected_role == "super_admin":
            if email == SUPER_ADMIN_EMAIL.lower() and password == SUPER_ADMIN_PASSWORD:
                session_id = _start_new_session("super_admin", "super_admin")
                token = issue_token_response("super_admin", "super_admin", session_id)
                resp = ok(
                    {
                        "token": token,
                        "role": "super_admin",
                        "redirect": ROLE_REDIRECTS["super_admin"],
                    },
                    message="Login successful.",
                )
                set_access_cookies(resp[0], token)
                return resp
            return error("Invalid email or password.", 401)

        # -------- Everyone else: look up in MongoDB --------
        user = find_user_by_email(email)
        if not user:
            return error("Invalid email or password.", 401)

        if not bcrypt.check_password_hash(user["passwordHash"], password):
            return error("Invalid email or password.", 401)

        if normalize_role(user["role"]) != selected_role:
            return error("Wrong role selected.")

        # Soft-deleted accounts are rejected outright, before the
        # approvalStatus checks below — a deleted user should never be
        # able to log in again unless a Super Admin restores the record.
        if user.get("isDeleted"):
            return error("This account has been deleted. Contact the administrator.", 403)

        status = user.get("approvalStatus")
        if status == "pending":
            return error("Your account is awaiting approval.", 403)
        if status == "rejected":
            return error("Your account has been rejected.", 403)
        if status != "approved":
            return error("Your account is not active. Contact the administrator.", 403)

        # Accounts created directly by the Super Admin (see
        # POST /api/admin/trainers in superadmin.py) must complete a
        # first-login verification before they get any usable session:
        # they set a password of their own and confirm it with an OTP.
        # NO token is issued here — login.html shows the verification
        # panel and only then calls /complete-first-login.
        if user.get("firstLoginVerify"):
            return ok(
                {
                    "requiresFirstLoginVerification": True,
                    "role": user["role"],
                    "email": email,
                },
                message="Super Admin created this account. Set your password and verify to continue.",
            )

        session_id = _start_new_session(user["role"], str(user["_id"]))
        token = issue_token_response(user["role"], str(user["_id"]), session_id)
        resp = ok(
            {
                "token": token,
                "role": user["role"],
                "redirect": ROLE_REDIRECTS.get(user["role"], "login.html"),
                "user": public_user(user),
            },
            message="Login successful.",
        )
        set_access_cookies(resp[0], token)
        return resp

    # ==========================================================
    # LOGOUT
    # ==========================================================
    @auth_bp.route("/logout", methods=["POST"])
    @jwt_required(optional=True)
    def logout():
        # optional=True: logout must succeed even if the token here is
        # already stale (e.g. this very tab was the one that just got
        # kicked out by a login elsewhere) — we still want to clear
        # whatever session record remains and unset the cookies.
        claims = get_jwt() or {}
        identity = get_jwt_identity()
        role = claims.get("role")
        if role == "super_admin":
            super_admin_sessions.update_one(
                {"_id": SUPER_ADMIN_SESSION_KEY},
                {"$set": {"sessionId": None}},
                upsert=True,
            )
        elif identity:
            try:
                oid = ObjectId(identity)
            except InvalidId:
                oid = None
            if oid:
                users.update_one({"_id": oid}, {"$set": {"currentSessionId": None}})
        resp = ok(message="Logged out.")
        unset_jwt_cookies(resp[0])
        return resp

    # ==========================================================
    # FORGOT PASSWORD — thin wrapper over send-otp(purpose=forgot)
    # ==========================================================
    @auth_bp.route("/forgot-password", methods=["POST"])
    @limiter.limit("5 per minute")
    def forgot_password():
        data = request.get_json(silent=True) or {}
        email = (data.get("email") or "").strip().lower()

        email_err = validate_email(email)
        if email_err:
            return error(email_err)
        if not find_user_by_email(email):
            return error("No account found with this email.")

        otp_verifications.update_one(
            {"email": email, "purpose": "forgot"},
            {
                "$set": {
                    "email": email,
                    "purpose": "forgot",
                    "otp": OTP_DEV_CODE,
                    "verified": False,
                    "token": None,
                    "createdAt": _now(),
                    "expiresAt": _now() + timedelta(minutes=OTP_EXPIRY_MINUTES),
                }
            },
            upsert=True,
        )
        return ok(message="OTP sent.")

    # ==========================================================
    # RESET PASSWORD
    # ==========================================================
    @auth_bp.route("/reset-password", methods=["POST"])
    @limiter.limit("10 per minute")
    def reset_password():
        data = request.get_json(silent=True) or {}
        email = (data.get("email") or "").strip().lower()
        new_password = data.get("newPassword") or ""
        confirm_password = data.get("confirmPassword") or ""
        otp_token = data.get("otpToken") or ""

        email_err = validate_email(email)
        if email_err:
            return error(email_err)

        password_err = validate_password(new_password)
        if password_err:
            return error(password_err)
        if new_password != confirm_password:
            return error("Passwords do not match.")

        otp_err = _consume_otp_token(email, "forgot", otp_token)
        if otp_err:
            return error(otp_err)

        user = find_user_by_email(email)
        if not user:
            return error("No account found with this email.", 404)

        password_hash = bcrypt.generate_password_hash(new_password).decode("utf-8")
        users.update_one(
            {"_id": user["_id"]},
            {"$set": {"passwordHash": password_hash, "updatedAt": _now()}},
        )
        return ok(message="Password updated successfully.")

    # ==========================================================
    # FIRST-LOGIN VERIFICATION — Super Admin-created accounts
    # (trainers). login() refuses to issue a session for these until
    # the user sets their own password AND proves it with an OTP.
    # Only after that do they get a real token + session + redirect.
    # ==========================================================
    @auth_bp.route("/complete-first-login", methods=["POST"])
    @limiter.limit("10 per minute")
    def complete_first_login():
        data = request.get_json(silent=True) or {}
        email = (data.get("email") or "").strip().lower()
        new_password = data.get("newPassword") or ""
        confirm_password = data.get("confirmPassword") or ""
        otp_token = data.get("otpToken") or ""

        user = find_user_by_email(email)
        if not user:
            return error("No account found with this email.", 404)
        if not user.get("firstLoginVerify"):
            return error("This account does not require first-login verification.", 400)

        password_err = validate_password(new_password)
        if password_err:
            return error(password_err)
        if new_password != confirm_password:
            return error("Passwords do not match.")

        otp_err = _consume_otp_token(email, "first_login", otp_token)
        if otp_err:
            return error(otp_err)

        password_hash = bcrypt.generate_password_hash(new_password).decode("utf-8")
        users.update_one(
            {"_id": user["_id"]},
            {"$set": {
                "passwordHash": password_hash,
                "firstLoginVerify": False,
                "firstLoginVerifiedAt": _now(),
                "updatedAt": _now(),
            }},
        )

        session_id = _start_new_session(user["role"], str(user["_id"]))
        token = issue_token_response(user["role"], str(user["_id"]), session_id)
        resp = ok(
            {
                "token": token,
                "role": user["role"],
                "redirect": ROLE_REDIRECTS.get(user["role"], "login.html"),
                "user": public_user(user),
            },
            message="Verification complete. Welcome to the platform!",
        )
        set_access_cookies(resp[0], token)
        return resp

    # ==========================================================
    # GOOGLE LOGIN (Firebase Authentication — Google Sign-In)
    # ==========================================================
    @auth_bp.route("/google", methods=["POST"])
    @limiter.limit("15 per minute")
    def google_login():
        if not firebase_ready or firebase_auth is None:
            return error(
                "Google Sign-In is not configured on the server yet. "
                "Set FIREBASE_CREDENTIALS_PATH in .env and restart.",
                501,
            )

        data = request.get_json(silent=True) or {}
        id_token = data.get("idToken")
        requested_role = normalize_role(data.get("role")) or DEFAULT_GOOGLE_SIGNUP_ROLE

        if not id_token:
            return error("Missing Google idToken.")

        try:
            decoded = firebase_auth.verify_id_token(id_token)
        except Exception:
            return error("Invalid or expired Google sign-in token.", 401)

        email = (decoded.get("email") or "").strip().lower()
        full_name = decoded.get("name") or email.split("@")[0]
        if not email:
            return error("Google account has no email on file.")

        user = find_user_by_email(email)

        if not user:
            if requested_role not in REGISTERABLE_ROLES:
                requested_role = DEFAULT_GOOGLE_SIGNUP_ROLE
            user_doc = {
                "fullName": full_name,
                "email": email,
                "mobile": None,
                "role": requested_role,
                "college": None,
                "district": None,
                "rollNumber": None,
                "employeeId": None,
                "tneaCode": None,
                "passwordHash": None,
                "approvalStatus": "pending",
                "approvedBy": None,
                "approvedDate": None,
                "googleLogin": True,
                "cohort": None,
                "baselineAssessmentScore": None,
                "interviewScore": None,
                "finalEmployabilityScore": None,
                "cohortAssignedAt": None,
                "createdAt": _now(),
                "updatedAt": _now(),
            }
            users.insert_one(user_doc)
            return ok(
                message="Account created via Google. Your account is waiting for approval.",
                status=201,
            )

        status = user.get("approvalStatus")
        if status == "pending":
            return error("Your account is awaiting approval.", 403)
        if status == "rejected":
            return error("Your account has been rejected.", 403)

        session_id = _start_new_session(user["role"], str(user["_id"]))
        token = issue_token_response(user["role"], str(user["_id"]), session_id)
        resp = ok(
            {
                "token": token,
                "role": user["role"],
                "redirect": ROLE_REDIRECTS.get(user["role"], "login.html"),
                "user": public_user(user),
            },
            message="Login successful.",
        )
        set_access_cookies(resp[0], token)
        return resp

    # ==========================================================
    # CURRENT USER — small protected-route sanity check
    # ==========================================================
    @auth_bp.route("/me", methods=["GET"])
    @jwt_required()
    def me():
        claims = get_jwt()
        identity = get_jwt_identity()
        if claims.get("role") == "super_admin":
            return ok({"user": {"role": "super_admin", "email": SUPER_ADMIN_EMAIL}})
        try:
            user = users.find_one({"_id": ObjectId(identity)})
        except InvalidId:
            user = None
        if not user:
            return error("User not found.", 404)
        return ok({"user": public_user(user)})

    # ==========================================================
    # APPROVAL WORKFLOW
    # College Admin & Trainer -> only Super Admin approves/rejects
    # Student                 -> Super Admin OR Trainer approves/rejects
    # ==========================================================
    def _pending_list(role):
        docs = users.find({"role": role, "approvalStatus": "pending"}).sort("createdAt", 1)
        return [public_user(d) for d in docs]

    @auth_bp.route("/pending/college-admins", methods=["GET"])
    @role_required("super_admin")
    def pending_college_admins():
        return ok({"pending": _pending_list("college_admin")})

    @auth_bp.route("/pending/trainers", methods=["GET"])
    @role_required("super_admin")
    def pending_trainers():
        return ok({"pending": _pending_list("trainer")})

    @auth_bp.route("/pending/students", methods=["GET"])
    @role_required("super_admin", "trainer")
    def pending_students():
        return ok({"pending": _pending_list("student")})

    # ==========================================================
    # APPROVAL STATS — live counts for dashboard cards
    # (Super Admin: full picture. Trainer: student-only count,
    # since that's the only queue a Trainer can act on.)
    # ==========================================================
    @auth_bp.route("/stats", methods=["GET"])
    @role_required("super_admin", "trainer")
    def approval_stats():
        claims = get_jwt()
        role = claims.get("role")

        if role == "trainer":
            return ok({
                "stats": {
                    "pendingStudents": users.count_documents(
                        {"role": "student", "approvalStatus": "pending"}
                    ),
                }
            })

        stats = {
            "pendingCollegeAdmins": users.count_documents(
                {"role": "college_admin", "approvalStatus": "pending"}
            ),
            "pendingTrainers": users.count_documents(
                {"role": "trainer", "approvalStatus": "pending"}
            ),
            "pendingStudents": users.count_documents(
                {"role": "student", "approvalStatus": "pending"}
            ),
            "approvedUsers": users.count_documents({"approvalStatus": "approved"}),
            "rejectedUsers": users.count_documents({"approvalStatus": "rejected"}),
        }
        return ok({"stats": stats})

    def _set_approval(user_id, new_status):
        claims = get_jwt()
        approver_role = claims.get("role")
        approver_identity = get_jwt_identity()

        try:
            oid = ObjectId(user_id)
        except InvalidId:
            return error("Invalid user id.", 400)

        target = users.find_one({"_id": oid})
        if not target:
            return error("User not found.", 404)

        target_role = target.get("role")
        if target_role in ("college_admin", "trainer") and approver_role != "super_admin":
            return error("Only the Super Admin can approve or reject this role.", 403)
        if target_role == "student" and approver_role not in ("super_admin", "trainer"):
            return error("Only a Trainer or Super Admin can approve or reject students.", 403)

        # Only the FIRST approve/reject action is accepted. This matters most
        # for students, who are visible to both Trainer and Super Admin at
        # once — an atomic conditional update (matched only while still
        # "pending") guarantees the second actor can never overwrite the
        # first actor's decision, even under a race.
        result = users.update_one(
            {"_id": oid, "approvalStatus": "pending"},
            {
                "$set": {
                    "approvalStatus": new_status,
                    "approvedBy": approver_identity,
                    "approvedDate": _now(),
                    "updatedAt": _now(),
                }
            },
        )
        if result.matched_count == 0:
            current_status = target.get("approvalStatus")
            return error(
                f"This request has already been {current_status} by another approver.",
                409,
            )
        return ok(message=f"User {new_status}.")

    @auth_bp.route("/approve/<user_id>", methods=["PUT"])
    @role_required("super_admin", "trainer")
    def approve_user(user_id):
        return _set_approval(user_id, "approved")

    @auth_bp.route("/reject/<user_id>", methods=["PUT"])
    @role_required("super_admin", "trainer")
    def reject_user(user_id):
        return _set_approval(user_id, "rejected")

    # ==========================================================
    # SINGLE ACTIVE SESSION — global JWT enforcement
    # ==========================================================
    # These two callbacks are registered on the app's single JWTManager
    # instance, so flask-jwt-extended runs them automatically inside
    # EVERY jwt_required() check across the whole app (this file's own
    # role_required(), quiz_common.role_required() used by superadmin.py /
    # trainer.py / student.py / collegeadmin.py / colleges.py /
    # quiz_module.py / question_bank.py, and the bare @jwt_required()
    # routes like /me). No other file needs to change.
    @jwt.token_in_blocklist_loader
    def _reject_superseded_sessions(_jwt_header, jwt_payload):
        """Return True (token is revoked / rejected) whenever the token's
        embedded session id no longer matches the latest one on record —
        i.e. a newer login has happened since this token was issued."""
        role = jwt_payload.get("role")
        token_sid = jwt_payload.get("sid")

        # Tokens issued before this feature shipped carry no "sid" claim.
        # Treat them as superseded so every account is forced through the
        # new single-session flow instead of silently bypassing it.
        if not token_sid:
            return True

        if role == "super_admin":
            record = super_admin_sessions.find_one({"_id": SUPER_ADMIN_SESSION_KEY}, {"sessionId": 1})
            current_sid = record.get("sessionId") if record else None
            if current_sid == token_sid:
                # Heartbeat — proves this account is actively using the portal
                # right now, which is what the Super Admin "Active Users"
                # dashboard counts (see superadmin.py /dashboard/live-stats).
                super_admin_sessions.update_one(
                    {"_id": SUPER_ADMIN_SESSION_KEY},
                    {"$set": {"lastActiveAt": _now()}},
                )
            return current_sid != token_sid

        identity = jwt_payload.get("sub")
        try:
            oid = ObjectId(identity)
        except (InvalidId, TypeError):
            return True

        user = users.find_one({"_id": oid}, {"currentSessionId": 1, "approvalStatus": 1})
        if not user or user.get("approvalStatus") != "approved":
            return True
        current_sid = user.get("currentSessionId")
        if current_sid == token_sid:
            # Heartbeat — see above; also what makes "Active Users" drop
            # automatically after session expiry / idle timeouts.
            users.update_one({"_id": oid}, {"$set": {"lastActiveAt": _now()}})
            return False
        return True

    @jwt.revoked_token_loader
    def _superseded_session_response(_jwt_header, _jwt_payload):
        return jsonify(
            success=False,
            message="Your account has been logged in from another device.",
        ), 401

    return auth_bp
