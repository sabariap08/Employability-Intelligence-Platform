"""
============================================================
 Employability Intelligence Platform — Flask Application
============================================================
Main entrypoint. Responsible for:
  - Loading configuration from environment (.env)
  - Wiring up MongoDB, Bcrypt, JWT, CORS, and Rate Limiting
  - Initializing Firebase Admin (for Google Sign-In verification)
  - Registering the auth blueprint from login.py
  - Serving the existing static frontend (login.html + role dashboards)

Only two Python modules exist by design: app.py (this file) and
login.py (all authentication logic). Future modules — superadmin.py,
trainer.py, student.py, collegeadmin.py — will each expose their own
Blueprint and be registered here the same way login.py is, without
touching the authentication architecture.

Run with:  python app.py
============================================================
"""

import os
from datetime import timedelta

from flask import Flask, jsonify, send_from_directory
from flask_cors import CORS
from flask_bcrypt import Bcrypt
from flask_jwt_extended import JWTManager
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from pymongo import MongoClient
from pymongo.errors import PyMongoError
from dotenv import load_dotenv

# Firebase Admin is optional at boot time — if no service-account
# credentials are configured, Google Sign-In simply reports itself
# as "not configured" instead of crashing the whole app.
try:
    import firebase_admin
    from firebase_admin import credentials as firebase_credentials
    FIREBASE_SDK_AVAILABLE = True
except ImportError:
    FIREBASE_SDK_AVAILABLE = False

load_dotenv()

# ------------------------------------------------------------
# 1. CONFIGURATION — everything pulled from environment variables.
#    See .env.example for the full list and sensible defaults.
# ------------------------------------------------------------
MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017")
DB_NAME = os.getenv("DB_NAME", "eip_platform")

JWT_SECRET_KEY = os.getenv("JWT_SECRET_KEY")
if not JWT_SECRET_KEY:
    raise RuntimeError(
        "JWT_SECRET_KEY is not set. Create a .env file (see .env.example) "
        "and set a strong random JWT_SECRET_KEY before starting the server."
    )
JWT_ACCESS_HOURS = int(os.getenv("JWT_ACCESS_TOKEN_EXPIRES_HOURS", "12"))
JWT_COOKIE_SECURE = os.getenv("JWT_COOKIE_SECURE", "false").lower() == "true"

FRONTEND_ORIGINS = [
    o.strip() for o in os.getenv("FRONTEND_ORIGIN", "http://localhost:5000").split(",")
    if o.strip()
]

RATE_LIMIT_DEFAULT = os.getenv("RATE_LIMIT_DEFAULT", "200 per day;50 per hour")
RATE_LIMIT_STORAGE_URI = os.getenv("RATE_LIMIT_STORAGE_URI", "memory://")

FIREBASE_CREDENTIALS_PATH = os.getenv("FIREBASE_CREDENTIALS_PATH", "")

FLASK_DEBUG = os.getenv("FLASK_DEBUG", "true").lower() == "true"
FLASK_HOST = os.getenv("FLASK_HOST", "0.0.0.0")
FLASK_PORT = int(os.getenv("FLASK_PORT", "5000"))

FRONTEND_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "frontend")

# ------------------------------------------------------------
# 2. APP + EXTENSIONS
# ------------------------------------------------------------
app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path="")

app.config["JWT_SECRET_KEY"] = JWT_SECRET_KEY
app.config["JWT_TOKEN_LOCATION"] = ["headers", "cookies"]
app.config["JWT_ACCESS_TOKEN_EXPIRES"] = timedelta(hours=JWT_ACCESS_HOURS)
app.config["JWT_COOKIE_SECURE"] = JWT_COOKIE_SECURE
app.config["JWT_COOKIE_SAMESITE"] = "Lax"
# CSRF-on-cookie protection is left off by default for local dev simplicity.
# Turn it on (and implement the matching X-CSRF-TOKEN header on the frontend)
# before shipping this to production.
app.config["JWT_COOKIE_CSRF_PROTECT"] = os.getenv("JWT_COOKIE_CSRF_PROTECT", "false").lower() == "true"

# Single Active Session: every jwt_required() check also runs the
# token_in_blocklist_loader registered in login.py, which rejects a token
# the moment a newer login has superseded it. Only access tokens exist in
# this app (no refresh-token flow), so only that check type is enabled.
app.config["JWT_BLOCKLIST_ENABLED"] = True
app.config["JWT_BLOCKLIST_TOKEN_CHECKS"] = ["access"]

bcrypt = Bcrypt(app)
jwt = JWTManager(app)

CORS(
    app,
    resources={r"/api/*": {"origins": FRONTEND_ORIGINS}},
    supports_credentials=True,
)

limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[l.strip() for l in RATE_LIMIT_DEFAULT.split(";") if l.strip()],
    storage_uri=RATE_LIMIT_STORAGE_URI,
)

# ------------------------------------------------------------
# 3. DATABASE
# ------------------------------------------------------------
try:
    mongo_client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
    mongo_client.admin.command("ping")
    db = mongo_client[DB_NAME]
    # Unique index on email guarantees no duplicate registrations at the
    # database layer even under race conditions, on top of the app-level check.
    db.users.create_index("email", unique=True)
    db.users.create_index("cohort")
    db.users.create_index([("role", 1), ("college", 1)])
    db.otp_verifications.create_index("email")
    db.otp_verifications.create_index("expiresAt", expireAfterSeconds=0)

    # Per-role unique identifiers. Partial indexes so users of OTHER roles
    # (whose corresponding field is None) never collide on "null == null".
    db.users.create_index(
        "rollNumber", unique=True,
        partialFilterExpression={"rollNumber": {"$type": "string"}},
    )
    db.users.create_index(
        "employeeId", unique=True,
        partialFilterExpression={"employeeId": {"$type": "string"}},
    )
    db.users.create_index(
        "tneaCode", unique=True,
        partialFilterExpression={"tneaCode": {"$type": "string"}},
    )

    # College Management collections — single source of truth for every
    # college/department dropdown on the platform (registration + quiz).
    db.colleges.create_index(
        "college_name", unique=True, collation={"locale": "en", "strength": 2}
    )
    db.departments.create_index(
        [("college_id", 1), ("department_name", 1)],
        unique=True, collation={"locale": "en", "strength": 2},
    )
    db.departments.create_index("college_id")


    # Quiz engine indexes
    db.questions.create_index("section")
    db.questions.create_index("active")
    db.assessments.create_index("cohortTarget")
    db.assessments.create_index("status")
    db.assessments.create_index("college")
    db.assessment_attempts.create_index([("studentId", 1), ("assessmentId", 1)])
    db.assessment_attempts.create_index("status")
    db.manual_interviews.create_index("studentId")
    db.manual_interviews.create_index("status")
    db.placement_rules.create_index("active")

    # Create-Quiz feature (quiz_module.py) — dedicated collection, separate
    # from the cohort/placement assessment engine above.
    db.quizzes.create_index("state")
    db.quizzes.create_index("startDateTime")
    db.quizzes.create_index("endDateTime")
    db.quizzes.create_index([("createdBy.id", 1)])
    db.quizzes.create_index([("createdBy.college", 1)])
    db.quiz_sections.create_index("name", unique=True)

    # Question Bank (question_bank.py, Create Quiz Part 5) — permanent
    # store of every validated question Trainer/Super Admin have ever
    # entered, drawn from by Super Admin's Question Bank quiz-creation
    # mode. questionHash is unique so a duplicate (same section + text +
    # options) can never be inserted twice, even under concurrent writes.
    db.question_bank.create_index("questionHash", unique=True)
    db.question_bank.create_index([("section", 1), ("difficulty", 1), ("active", 1)])
    db.question_bank.create_index("uploadedAt")

    # StudentCohort (quiz_common.py, Part 5) — one document per student,
    # kept as an additive mirror of db.users.cohort (see
    # sync_student_cohort_record's docstring for why db.users.cohort
    # itself stays the source every eligibility check reads).
    db.student_cohort.create_index("studentId", unique=True)
    db.student_cohort.create_index("rollNumber")

    # Student Quiz Attempt module (student.py) — one doc per student per
    # quiz; the compound index also enforces "no second attempt" lookups
    # stay O(1) as attempt volume grows.
    db.quiz_attempts.create_index([("studentId", 1), ("quizId", 1)])
    db.quiz_attempts.create_index("status")
    # Marks Management / Interview Verification / Validation Verification —
    # every trainer & super-admin listing filters by college + resultStatus,
    # and student "My Results" filters by studentId; keep both O(log n).
    db.quiz_attempts.create_index([("college", 1), ("resultStatus", 1)])
    db.quiz_attempts.create_index([("studentId", 1), ("status", 1)])

    # Schedule Session (Super Admin > Schedule Session, superadmin.py) —
    # db.workshop_sessions. collegeIds/departmentIds/trainerIds are each
    # arrays of ObjectId, so a multikey index on any one of them also
    # speeds up "does college X have any sessions" style lookups.
    db.workshop_sessions.create_index("status")
    db.workshop_sessions.create_index("collegeIds")
    db.workshop_sessions.create_index("departmentIds")
    db.workshop_sessions.create_index("trainerIds")
    db.workshop_sessions.create_index("date")

    # Attendance (Super Admin > Attendance, superadmin.py) — db.attendance.
    # One document per student per date: the unique compound index both
    # keeps lookups O(log n) and enforces "no duplicate attendance for the
    # same student on the same day" at the database layer (upserts update
    # the existing record rather than inserting a second one).
    db.attendance.create_index(
        [("studentId", 1), ("date", 1)],
        unique=True,
        partialFilterExpression={"studentId": {"$type": "objectId"}},
    )
    db.attendance.create_index([("date", 1), ("college", 1)])
    db.attendance.create_index("markedBy.id")

    # Activity logging — powers Recent Activity on every dashboard
    db.activity_log.create_index([("college", 1), ("createdAt", -1)])
    db.activity_log.create_index([("studentId", 1), ("createdAt", -1)])
    db.activity_log.create_index("actorId")
    print(f"[app] Connected to MongoDB at {MONGO_URI}, database '{DB_NAME}'.")
except PyMongoError as exc:
    raise RuntimeError(
        f"Could not connect to MongoDB at {MONGO_URI}. Is MongoDB running? "
        f"Original error: {exc}"
    )

# ------------------------------------------------------------
# 4. FIREBASE ADMIN (Google Sign-In verification) — optional
# ------------------------------------------------------------
firebase_ready = False
if FIREBASE_SDK_AVAILABLE and FIREBASE_CREDENTIALS_PATH:
    if os.path.exists(FIREBASE_CREDENTIALS_PATH):
        cred = firebase_credentials.Certificate(FIREBASE_CREDENTIALS_PATH)
        firebase_admin.initialize_app(cred)
        firebase_ready = True
        print("[app] Firebase Admin initialized — Google Sign-In is active.")
    else:
        print(
            f"[app] FIREBASE_CREDENTIALS_PATH is set to '{FIREBASE_CREDENTIALS_PATH}' "
            "but the file does not exist. Google Sign-In will report as unavailable."
        )
else:
    print("[app] Firebase Admin not configured. Google Sign-In will report as unavailable "
          "until FIREBASE_CREDENTIALS_PATH is set in .env.")

# ------------------------------------------------------------
# 5. REGISTER BLUEPRINTS
# ------------------------------------------------------------
from login import init_auth  # noqa: E402  (import after app/db/bcrypt exist)
from superadmin import init_superadmin  # noqa: E402
from trainer import init_trainer  # noqa: E402
from student import init_student  # noqa: E402
from colleges import init_colleges  # noqa: E402
from quiz_module import init_quiz  # noqa: E402
from question_bank import init_question_bank  # noqa: E402
from collegeadmin import init_collegeadmin  # noqa: E402

auth_bp = init_auth(
    bcrypt=bcrypt,
    db=db,
    limiter=limiter,
    jwt=jwt,
    firebase_ready=firebase_ready,
)
app.register_blueprint(auth_bp, url_prefix="/api/auth")

# Quiz Management & Dashboard modules — each owns one role's routes and
# all three share quiz_common.py so cohort/placement-rule logic can
# never drift out of sync between dashboards.
app.register_blueprint(init_superadmin(db=db, bcrypt=bcrypt), url_prefix="/api/admin")
app.register_blueprint(init_trainer(db=db, bcrypt=bcrypt), url_prefix="/api/trainer")
app.register_blueprint(init_student(db=db), url_prefix="/api/student")

# Create Quiz feature — dedicated `quizzes` collection, mounted once per
# dashboard so each gets its own scoping (trainer: own college + own
# quizzes; super admin: full platform visibility and control).
app.register_blueprint(init_quiz(db=db, scope="trainer"), url_prefix="/api/trainer")
app.register_blueprint(init_quiz(db=db, scope="super_admin"), url_prefix="/api/admin")

# Question Bank (Part 4/5/6) — read-only summary/list endpoints; writes
# only ever happen internally from quiz_module.py's create/update-quiz
# flow, never directly from the frontend.
app.register_blueprint(init_question_bank(db=db, scope="trainer"), url_prefix="/api/trainer")
app.register_blueprint(init_question_bank(db=db, scope="super_admin"), url_prefix="/api/admin")

# College Management — /api/admin/colleges (Super Admin CRUD) and
# /api/public/colleges (unauthenticated dropdown reads).
app.register_blueprint(init_colleges(db=db, bcrypt=bcrypt), url_prefix="/api")

# College Admin — read-only Assessment Management, scoped to their
# own assigned college. No create/update/delete routes exist here.
app.register_blueprint(init_collegeadmin(db=db), url_prefix="/api/collegeadmin")

# Assessment Question Banks & Templates — serves the Super Admin's
# GET /api/assessments/question-banks and /api/assessments/templates
# (read-only, computed live from db.question_bank / db.quiz_sections).
from assessments import init_assessments  # noqa: E402
app.register_blueprint(init_assessments(db=db), url_prefix="/api/assessments")


# ------------------------------------------------------------
# 6. STATIC FRONTEND — serves the untouched HTML/CSS/JS as-is
# ------------------------------------------------------------
@app.route("/")
def serve_home_page():
    return send_from_directory(FRONTEND_DIR, "Home.html")


@app.route("/<path:filename>")
def serve_frontend_file(filename):
    return send_from_directory(FRONTEND_DIR, filename)


# ------------------------------------------------------------
# 7. ERROR HANDLERS — consistent JSON errors for the /api/* surface
# ------------------------------------------------------------
@app.errorhandler(404)
def not_found(_err):
    return jsonify(success=False, message="Resource not found."), 404


@app.errorhandler(500)
def server_error(_err):
    return jsonify(success=False, message="Internal server error."), 500


@app.errorhandler(429)
def rate_limited(_err):
    return jsonify(success=False, message="Too many requests. Please slow down and try again."), 429


if __name__ == "__main__":
    app.run(host=FLASK_HOST, port=FLASK_PORT, debug=FLASK_DEBUG)
