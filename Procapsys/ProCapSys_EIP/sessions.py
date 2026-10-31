"""
============================================================
 sessions.py — SHARED Workshop Session + Attendance helpers
============================================================
Extracted from superadmin.py so the Trainer dashboard uses the
EXACT same session/attendance rules as Super Admin (statuses,
window timing, cohort labels, audience resolution) and the two
dashboards can never drift.

Backing collections (created/indexed in app.py):
  - db.workshop_sessions — one document per scheduled session
  - db.attendance        — one document per student per date
"""

import re
from datetime import datetime

from quiz_common import (
    now, to_object_id, iso_utc, fmt_ist,
    student_cohort_label, VALID_COHORTS, ENTRY_LEVEL,
)

# ------------------------------------------------------------
# SESSION CONSTANTS — Workshop Session scheduling options
# ------------------------------------------------------------
SESSION_COHORT_OPTIONS = ("All Cohorts", "Cohort A", "Cohort B", "Cohort C")
SESSION_ATTENDANCE_OPTIONS = ("mandatory", "optional")
SESSION_STATUS_OPTIONS = ("scheduled", "completed", "cancelled")
# Mode of Conduct — how a session is delivered. Offline sessions need a
# venue; online sessions carry no venue and never take a meeting link at
# creation time (the link is added later via the Update Meeting Link flow,
# only while the session is live).
SESSION_MODE_OPTIONS = ("online", "offline")
_TIME_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")

# ------------------------------------------------------------
# ATTENDANCE CONSTANTS
# ------------------------------------------------------------
ATTENDANCE_STATUSES = ("present", "absent")
ATTENDANCE_STATUS_LABELS = {"present": "Present", "absent": "Absent", "not_marked": "Not Yet Marked"}
# Cohort labels — derived from the platform's single cohort system
# (VALID_COHORTS + ENTRY_LEVEL), NOT a separate hardcoded list.
ATTENDANCE_COHORT_LABELS = {
    "A": "Cohort A – Placement Ready",
    "B": "Cohort B – Near Ready",
    "C": "Cohort C",
    ENTRY_LEVEL: "Entry Level",
}
_ATTENDANCE_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def parse_optional_id_list(raw):
    """Like parse_id_list but a filter list that may be empty/absent
    (an empty selection means 'all'). Returns (object_ids, error)."""
    if raw is None or raw == "":
        return [], None
    if isinstance(raw, str):
        parts = [p.strip() for p in raw.split(",") if p.strip()]
    elif isinstance(raw, list):
        parts = [str(p) for p in raw]
    else:
        return None, "Invalid id list."
    seen = set()
    oids = []
    for item in parts:
        oid = to_object_id(item)
        if not oid:
            return None, "One of the selected ids is invalid."
        if oid not in seen:
            seen.add(oid)
            oids.append(oid)
    return oids, None


def parse_id_list(raw, field_label):
    """Normalize a request field that must be a non-empty list of id
    strings. Returns (object_ids, error_message) — object_ids is a
    deduped list preserving order; error_message is None on success."""
    if raw is None:
        raw = []
    if not isinstance(raw, list):
        return None, f"{field_label} must be a list of ids."
    if not raw:
        return None, f"Select at least one {field_label.lower()}."
    seen = set()
    oids = []
    for item in raw:
        oid = to_object_id(item)
        if not oid:
            return None, f"One of the selected {field_label.lower()} is invalid."
        if oid not in seen:
            seen.add(oid)
            oids.append(oid)
    return oids, None


def resolve_colleges(db, college_ids):
    """Every id must correspond to an existing, active college. Returns
    (docs_in_request_order, error)."""
    docs = {d["_id"]: d for d in db.colleges.find({"_id": {"$in": college_ids}, "status": "active"})}
    missing = [str(oid) for oid in college_ids if oid not in docs]
    if missing:
        return None, "One or more selected colleges is invalid or inactive."
    return [docs[oid] for oid in college_ids], None


def resolve_departments(db, department_ids):
    """Returns (docs_in_request_order, college_names_by_id, error)."""
    docs = {d["_id"]: d for d in db.departments.find({"_id": {"$in": department_ids}, "status": "active"})}
    missing = [str(oid) for oid in department_ids if oid not in docs]
    if missing:
        return None, None, "One or more selected departments is invalid or inactive."
    college_names = {c["_id"]: c.get("college_name") for c in db.colleges.find({}, {"college_name": 1})}
    ordered = [docs[oid] for oid in department_ids]
    return ordered, college_names, None


def resolve_trainers(db, trainer_ids):
    docs = {
        d["_id"]: d for d in db.users.find(
            {"_id": {"$in": trainer_ids}, "role": "trainer", "approvalStatus": {"$in": ["approved", "suspended"]}}
        )
    }
    missing = [str(oid) for oid in trainer_ids if oid not in docs]
    if missing:
        return None, "One or more selected trainers is invalid."
    return [docs[oid] for oid in trainer_ids], None


def attendance_student_query(db, college_oids, department_oids, cohorts):
    """Mongo filter for students matching the Mark/View filters. An empty
    list on any dimension means 'no restriction' (i.e. all). Cohorts is a
    list of 'A'/'B'/'C'/'entry_level' values; 'all' disables the filter."""
    query = {
        "role": "student",
        "approvalStatus": {"$in": ["approved", "suspended"]},
        "isDeleted": {"$ne": True},
    }
    if college_oids:
        query["collegeId"] = {"$in": college_oids}
    if department_oids:
        query["departmentId"] = {"$in": department_oids}
    if cohorts and "all" not in cohorts:
        branches = []
        if ENTRY_LEVEL in cohorts:
            branches.append({"cohort": {"$nin": list(VALID_COHORTS)}})
        for c in cohorts:
            if c in VALID_COHORTS:
                branches.append({"cohort": c})
        if branches:
            query["$or"] = branches
    return query


def attendance_student_public(doc):
    """Serialize a student user doc for the attendance mark/list views.
    cohortLabel uses the same labels the rest of the platform uses."""
    cohort = student_cohort_label(doc)
    return {
        "id": str(doc["_id"]),
        "name": doc.get("fullName") or doc.get("email") or "—",
        "rollNumber": doc.get("rollNumber") or "—",
        "college": doc.get("college") or "—",
        "collegeId": str(doc["collegeId"]) if doc.get("collegeId") else None,
        "department": doc.get("department") or "—",
        "departmentId": str(doc["departmentId"]) if doc.get("departmentId") else None,
        "cohort": cohort,
        "cohortLabel": ATTENDANCE_COHORT_LABELS.get(cohort, cohort),
    }


def session_target_students(db, doc):
    """The exact audience a workshop session was scheduled for (spec §28).
    If the session stored specific student ids, those are returned. Otherwise
    students are matched by the session's collegeIds + departmentIds + cohort
    — the exact filters chosen at schedule time, never a blanket college
    dump. Used by both Mark Attendance and the Student Workshops feed."""
    specific = doc.get("studentIds")
    if specific:
        oids = [o for o in (to_object_id(s) for s in specific) if o]
        if oids:
            return list(db.users.find(
                {"_id": {"$in": oids}, "role": "student", "isDeleted": {"$ne": True}}
            ).sort("fullName", 1))

    query = {"role": "student", "isDeleted": {"$ne": True}}
    college_oids = [o for o in (to_object_id(x) for x in doc.get("collegeIds", [])) if o]
    if college_oids:
        query["collegeId"] = {"$in": college_oids}
    department_oids = [o for o in (to_object_id(x) for x in doc.get("departmentIds", [])) if o]
    if department_oids:
        query["departmentId"] = {"$in": department_oids}
    cohort = doc.get("cohort")
    if cohort in ("Cohort A", "Cohort B", "Cohort C"):
        query["cohort"] = cohort[-1]
    return list(db.users.find(query).sort("fullName", 1))


def session_start_dt(doc):
    """Naive datetime for a session's start (date + startTime), or None if
    the stored schedule can't be parsed. Server-local time throughout —
    the same convention every other session-time rule in this module uses
    (session_started / attendance_window_state), so all timing decisions
    agree with each other and the browser is never the time source."""
    date_str, start = doc.get("date"), doc.get("startTime")
    if not (date_str and start):
        return None
    try:
        return datetime.strptime(f"{date_str} {start}", "%Y-%m-%d %H:%M")
    except ValueError:
        return None


def session_started(doc):
    """A workshop session counts as started once its start datetime has
    passed. Business rule: Edit / Delete are only allowed before that
    moment — enforced server-side here, and mirrored in the frontend by
    hiding the Edit/Delete buttons on started sessions."""
    start_dt = session_start_dt(doc)
    return start_dt is not None and datetime.now() >= start_dt


def auto_complete_workshop_sessions(db):
    """Auto-complete every scheduled workshop session whose end datetime
    has already passed. Stored date + endTime are plain wall-clock strings
    (the same convention attendance_window_state / session_started use),
    so they are compared to the server's local wall clock — the browser is
    never the time source. The status change is persisted to
    db.workshop_sessions, so every dashboard reading this collection shows
    Completed without a manual status change or a frontend refresh.
    Cancelled sessions are left untouched."""
    now_dt = datetime.now()
    updated = 0
    for doc in db.workshop_sessions.find(
        {"status": "scheduled"}, {"date": 1, "endTime": 1}
    ):
        date_str, end = doc.get("date"), doc.get("endTime")
        if not (date_str and end):
            continue
        try:
            end_dt = datetime.strptime(f"{date_str} {end}", "%Y-%m-%d %H:%M")
        except ValueError:
            continue
        if now_dt >= end_dt:
            db.workshop_sessions.update_one(
                {"_id": doc["_id"]},
                {"$set": {"status": "completed", "updatedAt": now()}},
            )
            updated += 1
    return updated


def attendance_window_open(doc):
    """True only for a Mandatory session, and only while 'now' falls
    within that session's own Session Date + Start/End Time (spec §10).
    Optional sessions never expose an attendance window/button."""
    return attendance_window_state(doc) == "open"


def attendance_window_state(doc):
    """Three-state attendance availability for a session card, computed
    entirely from the stored session schedule against the server clock
    (the client only mirrors it):
      - "not_applicable" — attendance not required (Optional)
      - "upcoming"       — before the session start, attendance locked
      - "open"           — inside the session window (start <= now <= end)
      - "closed"         — after the session end time, attendance locked
    The session date/start/end are stored as plain wall-clock strings the
    admin entered, so they are compared to the server's local wall clock —
    the same convention every other session-time rule in this module uses
    (session_started / meeting_link_window_open), keeping them consistent
    and keeping the browser out of the time decision entirely."""
    if doc.get("attendanceRequirement") != "mandatory":
        return "not_applicable"
    date_str, start, end = doc.get("date"), doc.get("startTime"), doc.get("endTime")
    if not (date_str and start and end):
        return "not_applicable"
    try:
        start_dt = datetime.strptime(f"{date_str} {start}", "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(f"{date_str} {end}", "%Y-%m-%d %H:%M")
    except ValueError:
        return "not_applicable"
    now_dt = datetime.now()
    if now_dt < start_dt:
        return "upcoming"
    if now_dt > end_dt:
        return "closed"
    return "open"


def meeting_link_window_open(doc):
    """Online sessions may take a meeting link only while the session is
    live: Start Time <= now <= End Time. Before the start time, and once
    the session has reached its end time, the meeting link can no longer
    be updated. Computed entirely from the server clock against the stored
    date/start/end wall-clock schedule (same convention as
    attendance_window_state) — the client only mirrors it."""
    if doc.get("mode") != "online":
        return False
    date_str, start, end = doc.get("date"), doc.get("startTime"), doc.get("endTime")
    if not (date_str and start and end):
        return False
    try:
        start_dt = datetime.strptime(f"{date_str} {start}", "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(f"{date_str} {end}", "%Y-%m-%d %H:%M")
    except ValueError:
        return False
    now_dt = datetime.now()
    return start_dt <= now_dt <= end_dt


def workshop_session_public(doc):
    if not doc:
        return None
    colleges = doc.get("collegeNames", [])
    departments = doc.get("departmentNames", [])
    trainers = doc.get("trainerNames", [])
    return {
        "id": str(doc["_id"]),
        "name": doc.get("name"),
        "collegeIds": [str(x) for x in doc.get("collegeIds", [])],
        "colleges": colleges,
        "college": ", ".join(colleges) or "—",
        "departmentIds": [str(x) for x in doc.get("departmentIds", [])],
        "departments": departments,
        "department": ", ".join(departments) or "—",
        "trainerIds": [str(x) for x in doc.get("trainerIds", [])],
        "trainers": trainers,
        "trainer": ", ".join(trainers) or "—",
        "cohort": doc.get("cohort"),
        "date": doc.get("date"),
        "startTime": doc.get("startTime"),
        "endTime": doc.get("endTime"),
        "time": f"{doc.get('startTime', '')} \u2013 {doc.get('endTime', '')}",
        "mode": doc.get("mode") or "offline",
        "venue": doc.get("venue"),
        "meetingLink": doc.get("meetingLink"),
        "meetingLinkUpdatedAt": iso_utc(doc.get("meetingLinkUpdatedAt")),
        "meetingLinkUpdatedAtIST": fmt_ist(doc.get("meetingLinkUpdatedAt")),
        "meetingLinkUpdatedBy": doc.get("meetingLinkUpdatedBy"),
        "meetingLinkWindowOpen": meeting_link_window_open(doc),
        "attendanceRequirement": doc.get("attendanceRequirement"),
        "attendanceWindowOpen": attendance_window_open(doc),
        "attendanceWindowState": attendance_window_state(doc),
        "status": doc.get("status", "scheduled"),
        "approvalStatus": doc.get("approvalStatus", "pending"),
        "createdBy": doc.get("createdBy"),
        "createdAt": iso_utc(doc.get("createdAt")),
        "createdAtIST": fmt_ist(doc.get("createdAt")),
        "updatedAt": iso_utc(doc.get("updatedAt")),
        "updatedAtIST": fmt_ist(doc.get("updatedAt")),
    }
