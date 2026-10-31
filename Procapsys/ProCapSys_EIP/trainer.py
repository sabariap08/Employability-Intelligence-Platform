"""
============================================================
 trainer.py — Trainer Dashboard backend
============================================================
Registered at /api/trainer. Mirrors superadmin.py's assessment +
cohort + quiz-management behaviour so the two dashboards can never
disagree, but every query here is scoped to the trainer's own
college (trainers only manage their own students).

The Entry Level cohort option and assignment rule is IDENTICAL to
Super Admin — both call the same quiz_common helpers.
"""

import hashlib
import io
import re

from datetime import datetime, timedelta

from flask import Blueprint, request, send_file
from flask_jwt_extended import get_jwt_identity
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

from quiz_common import (
    ok, error, role_required, now, to_object_id, serialize, iso_utc, fmt_ist,
    attendance_summary,
    VALID_COHORT_TARGETS, ENTRY_LEVEL,
    cohort_counts, record_interview_score_for_cohort,
    log_activity, get_recent_activity,
    list_quiz_results, set_quiz_interview_marks, validate_quiz_result,
    serialize_quiz_result, RESULT_STATUS_PENDING,
    RESULT_STATUS_INTERVIEW_DONE, RESULT_STATUS_VALIDATED,
    compute_quiz_analytics, list_quiz_responses,
    list_distinct_departments, build_quiz_responses_workbook,
    attempt_section_counts,
)
from colleges import resolve_active_department, _generate_temp_password, department_public
from reporting import excel_bytes, pdf_bytes
import reporting
from login import validate_email, validate_mobile

# ------------------------------------------------------------
# ATTENDANCE MODULE — reuses the exact same query/serialization
# helpers and constants the Super Admin "Mark Attendance" / "View
# Records" screens are built on (superadmin.py), so Trainer and
# Super Admin can never disagree about who counts as present/absent,
# which cohort a student is in, or what a date filter matches. Per
# the platform's role model, Trainers report to Super Admin (not to
# a single college) and therefore get the SAME attendance data scope
# Super Admin has — the routes below are trainer-authenticated
# mirrors of the Super Admin attendance routes, not a restricted
# subset.
# ------------------------------------------------------------
from superadmin import (
    ATTENDANCE_STATUSES, ATTENDANCE_COHORT_LABELS, _ATTENDANCE_DATE_RE,
    _parse_optional_id_list, _attendance_student_query, _attendance_student_public,
    SESSION_STATUS_OPTIONS, _intervention_status_label, _safe_filename,
    _auto_complete_workshop_sessions, _validate_session_payload,
    workshop_session_public, _session_started, _TIME_RE,
)
from quiz_common import VALID_COHORTS, student_cohort_label


# ------------------------------------------------------------
# Dashboard insights — every value in GET /dashboard/insights is
# computed live from db.assessment_attempts (the same source every
# placement-readiness number in this module uses), scoped to the
# trainer's own college. Nothing here is hardcoded, and nothing is
# cached beyond the lifetime of a single request.
# ------------------------------------------------------------
COHORT_RADAR_LABELS = {
    "A": "Cohort A",
    "B": "Cohort B",
    "C": "Cohort C",
    "entry_level": "Entry Level",
}
RADAR_COHORT_ORDER = ["entry_level", "A", "B", "C"]


def _week_buckets(weeks=8):
    """Monday-anchored calendar weeks, oldest first, ending with the week
    containing today. Returns (start, end, label) triples using naive UTC
    datetimes — the same clock stored submittedAt values live on, matching
    the _month_bounds() convention in superadmin.py. Week boundaries are
    midnight-aligned (Monday 00:00), like _month_bounds() month starts."""
    today = datetime.utcnow()
    anchor = (today - timedelta(days=today.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    return [
        (
            anchor - timedelta(weeks=offset),
            anchor - timedelta(weeks=offset - 1),
            (anchor - timedelta(weeks=offset)).strftime("%b %d"),
        )
        for offset in range(weeks - 1, -1, -1)
    ]


def init_trainer(db, bcrypt=None):
    bp = Blueprint("trainer", __name__)

    users = db.users
    questions = db.questions
    assessments = db.assessments
    attempts = db.assessment_attempts
    manual_interviews = db.manual_interviews
    quiz_attempts = db.quiz_attempts

    def _trainer_doc():
        return users.find_one({"_id": to_object_id(get_jwt_identity())})

    def _trainer_college():
        trainer = _trainer_doc()
        return (trainer or {}).get("college")

    # ------------------------------------------------------------
    # Shared student-roster query — backs GET /students AND the
    # /students/export download so a filtered export is always
    # byte-for-byte the same rows the table is showing. Trainers
    # manage the whole platform roster (no trainer↔student binding),
    # so there is deliberately NO college scoping here.
    # ------------------------------------------------------------
    def _student_roster(college="", department="", cohort="", search=""):
        query = {"role": "student"}
        if college:
            query["college"] = {"$in": [c.strip() for c in college.split(",") if c.strip()]}
        if department:
            query["department"] = {"$in": [d.strip() for d in department.split(",") if d.strip()]}
        and_parts = []
        if cohort:
            or_parts = []
            for c in [x.strip() for x in cohort.split(",") if x.strip()]:
                if c == ENTRY_LEVEL:
                    or_parts.append({"cohort": None})
                    or_parts.append({"cohort": {"$exists": False}})
                elif c in {"A", "B", "C"}:
                    or_parts.append({"cohort": c})
            if or_parts:
                and_parts.append({"$or": or_parts})
        if search:
            rx = {"$regex": re.escape(search), "$options": "i"}
            and_parts.append({"$or": [
                {"fullName": rx}, {"rollNumber": rx}, {"email": rx},
            ]})
        if and_parts:
            query["$and"] = and_parts

        docs = list(users.find(query, {"passwordHash": 0}).sort("createdAt", -1))
        ids = [d["_id"] for d in docs]

        att_map = {}
        for r in db.attendance.find(
            {"studentId": {"$in": ids}}, {"studentId": 1, "status": 1}
        ):
            bucket = att_map.setdefault(str(r.get("studentId")), [0, 0])
            if r.get("status") == "present":
                bucket[0] += 1
            elif r.get("status") == "absent":
                bucket[1] += 1

        score_map = {}
        for a in quiz_attempts.find(
            {"status": "submitted", "studentId": {"$in": ids}},
            {"studentId": 1, "overall.percentage": 1},
        ):
            pct = (a.get("overall") or {}).get("percentage")
            if pct is not None:
                score_map.setdefault(str(a.get("studentId")), []).append(pct)

        rows = []
        for d in docs:
            present, absent = att_map.get(str(d["_id"]), [0, 0])
            total = present + absent
            scores = score_map.get(str(d["_id"]), [])
            rows.append({
                "id": str(d["_id"]),
                "fullName": d.get("fullName"),
                "rollNumber": d.get("rollNumber"),
                "email": d.get("email"),
                "mobile": d.get("mobile"),
                "college": d.get("college"),
                "collegeId": str(d["collegeId"]) if d.get("collegeId") else None,
                "department": d.get("department"),
                "departmentId": str(d["departmentId"]) if d.get("departmentId") else None,
                "cohort": d.get("cohort"),
                "attendancePct": round(present / total * 100, 1) if total else None,
                "quizScore": round(sum(scores) / len(scores), 1) if scores else None,
                "status": d.get("approvalStatus") or "approved",
                "createdAt": iso_utc(d.get("createdAt")),
            })
        return rows

    @bp.route("/students", methods=["GET"])
    @role_required("trainer")
    def list_students():
        rows = _student_roster(
            college=(request.args.get("college") or "").strip(),
            department=(request.args.get("department") or "").strip(),
            cohort=(request.args.get("cohort") or "").strip(),
            search=(request.args.get("search") or "").strip(),
        )
        return ok({"students": rows, "total": len(rows)})

    # ==========================================================
    # ASSESSMENT CREATION — same rules as Super Admin, incl. Entry Level
    # ==========================================================
    @bp.route("/assessments", methods=["POST"])
    @role_required("trainer")
    def create_assessment():
        data = request.get_json(silent=True) or {}
        name = (data.get("name") or "").strip()
        assessment_type = (data.get("type") or "custom").strip().lower()
        cohort_target = (data.get("cohortTarget") or "").strip()
        section_counts = data.get("sectionCounts") or {}

        if not name:
            return error("Assessment name is required.")
        if cohort_target not in VALID_COHORT_TARGETS:
            return error(
                "cohortTarget must be one of: A, B, C, all, entry_level."
            )
        if not isinstance(section_counts, dict) or not section_counts:
            return error("sectionCounts is required, e.g. {'Communication': 10}.")
        for section, count in section_counts.items():
            if not isinstance(count, int) or count <= 0:
                return error(f"Question count for section '{section}' must be a positive integer.")

        status = (data.get("status") or "active").strip().lower()
        if status not in {"draft", "active"}:
            return error("status must be 'draft' or 'active'.")

        # Department dropdown, scoped to the trainer's own (assigned) college —
        # reads from the same departments collection as everywhere else.
        department_id = data.get("departmentId")
        department_name = None
        trainer = _trainer_doc()
        trainer_college_id = trainer.get("collegeId") if trainer else None
        if department_id:
            if not trainer_college_id:
                return error("Your account has no college assigned yet. Contact the Super Admin.")
            department_doc = resolve_active_department(db, department_id, str(trainer_college_id))
            if not department_doc:
                return error("Selected department is invalid or inactive for your college.")
            department_name = department_doc["department_name"]

        doc = {
            "name": name,
            "type": assessment_type,
            "cohortTarget": cohort_target,
            "sectionCounts": {k: int(v) for k, v in section_counts.items()},
            "totalQuestions": sum(int(v) for v in section_counts.values()),
            "scheduledAt": data.get("scheduledAt"),
            "availableFrom": data.get("availableFrom"),
            "availableTo": data.get("availableTo"),
            "college": _trainer_college(),
            "department": department_name,
            "createdBy": "trainer",
            "createdByTrainerId": get_jwt_identity(),
            "createdAt": now(),
            "status": status,
        }
        result = assessments.insert_one(doc)
        doc["_id"] = result.inserted_id
        log_activity(
            db, get_jwt_identity(), "trainer", "assessment_created",
            f'Created assessment "{name}"' + (" (published)" if status == "active" else " (draft)"),
            college=doc["college"], meta={"assessmentId": str(result.inserted_id)},
        )
        return ok({"assessment": serialize(doc)}, message="Assessment created.", status=201)

    @bp.route("/assessments", methods=["GET"])
    @role_required("trainer")
    def list_assessments():
        college = _trainer_college()
        query = {"$or": [{"college": college}, {"createdBy": "super_admin"}]} if college else {}
        cursor = assessments.find(query).sort("createdAt", -1)
        return ok({"assessments": [serialize(a) for a in cursor]})

    def _own_trainer_assessment(assessment_id):
        """Look up an assessment, but only if this trainer created it —
        trainers may publish/unpublish/delete their own assessments only.
        Assessments created by Super Admin are read-only here."""
        oid = to_object_id(assessment_id)
        if not oid:
            return None, error("Invalid assessment id.", 404)
        doc = assessments.find_one({"_id": oid})
        if not doc:
            return None, error("Assessment not found.", 404)
        if doc.get("createdByTrainerId") != get_jwt_identity():
            return None, error("You can only manage assessments you created.", 403)
        return doc, None

    @bp.route("/assessments/<assessment_id>/status", methods=["PATCH"])
    @role_required("trainer")
    def update_assessment_status(assessment_id):
        doc, err = _own_trainer_assessment(assessment_id)
        if err:
            return err
        data = request.get_json(silent=True) or {}
        status = (data.get("status") or "").strip().lower()
        if status not in {"draft", "active"}:
            return error("status must be 'draft' (unpublished) or 'active' (published).")
        assessments.update_one({"_id": doc["_id"]}, {"$set": {"status": status, "updatedAt": now()}})
        updated = assessments.find_one({"_id": doc["_id"]})
        action = "assessment_published" if status == "active" else "assessment_unpublished"
        verb = "published" if status == "active" else "moved to draft"
        log_activity(
            db, get_jwt_identity(), "trainer", action,
            f'Assessment "{doc.get("name")}" {verb}',
            college=doc.get("college"), meta={"assessmentId": str(doc["_id"])},
        )
        return ok({"assessment": serialize(updated)}, message=f"Assessment marked as {status}.")

    @bp.route("/assessments/<assessment_id>", methods=["DELETE"])
    @role_required("trainer")
    def delete_assessment(assessment_id):
        doc, err = _own_trainer_assessment(assessment_id)
        if err:
            return err
        assessments.delete_one({"_id": doc["_id"]})
        log_activity(
            db, get_jwt_identity(), "trainer", "assessment_deleted",
            f'Deleted assessment "{doc.get("name")}"',
            college=doc.get("college"), meta={"assessmentId": str(doc["_id"])},
        )
        return ok(message="Assessment deleted.")

    # ==========================================================
    # COHORTS — Entry Level + A/B/C, scoped to trainer's college.
    # Uses the exact same cohort_counts() helper as Super Admin.
    # ==========================================================
    @bp.route("/cohorts/counts", methods=["GET"])
    @role_required("trainer")
    def get_cohort_counts():
        return ok({"cohortCounts": cohort_counts(db, college=_trainer_college())})

    @bp.route("/cohorts/students", methods=["GET"])
    @role_required("trainer")
    def list_cohort_students():
        cohort = request.args.get("cohort", "").strip()
        # Trainers manage the whole roster — no college scoping.
        query = {"role": "student"}
        if cohort == ENTRY_LEVEL:
            query["$or"] = [{"cohort": None}, {"cohort": {"$exists": False}}]
        elif cohort in {"A", "B", "C"}:
            query["cohort"] = cohort
        docs = users.find(query, {"passwordHash": 0}).sort("createdAt", -1)
        return ok({"students": [serialize(d) for d in docs]})

    # ==========================================================
    # QUIZ MANAGEMENT / QUIZ RESPONSES / MANUAL INTERVIEW — live, college-scoped
    # ==========================================================
    @bp.route("/quiz-management/summary", methods=["GET"])
    @role_required("trainer")
    def quiz_management_summary():
        college = _trainer_college()
        student_ids = [s["_id"] for s in users.find({"role": "student", "college": college}, {"_id": 1})]
        total_attempts = attempts.count_documents({"studentId": {"$in": student_ids}})
        submitted_attempts = attempts.count_documents({"studentId": {"$in": student_ids}, "status": "submitted"})
        in_progress = attempts.count_documents({"studentId": {"$in": student_ids}, "status": "in_progress"})
        return ok({
            "totalQuestions": questions.count_documents({"active": True}),
            "totalAssessments": assessments.count_documents({"college": college}),
            "totalAttempts": total_attempts,
            "submittedAttempts": submitted_attempts,
            "inProgressAttempts": in_progress,
            "cohortCounts": cohort_counts(db, college=college),
        })

    @bp.route("/quiz-responses", methods=["GET"])
    @role_required("trainer")
    def quiz_responses():
        """Raw per-attempt rows for the Trainer's Quiz Responses table —
        scoped to the trainer's own college. Root-cause fix: this used to
        query db.assessment_attempts/db.assessments (the older, largely
        unused cohort-placement engine), which is why this page always
        appeared empty — students actually submit into db.quiz_attempts
        via Quiz Management. Now shares the exact same query Super
        Admin's working Quiz Responses table uses (list_quiz_responses),
        just scoped to this trainer's college, so the two can never
        disagree.
        """
        college = _trainer_college()
        cohort = request.args.get("cohort") or "all"
        quiz_id = request.args.get("quizId") or "all"
        department = request.args.get("department") or "all"
        search = request.args.get("search") or ""
        return ok({"responses": list_quiz_responses(
            db, college=college, cohort=cohort, quiz_id=quiz_id,
            department=department, search=search,
        )})

    @bp.route("/quiz-responses/filters", methods=["GET"])
    @role_required("trainer")
    def quiz_responses_filters():
        """Course/Department dropdown options for the Quiz Responses
        filter bar — distinct values actually in use by this trainer's
        own students, read live from the database, never hardcoded."""
        return ok({"departments": list_distinct_departments(db, college=_trainer_college())})

    @bp.route("/quiz-responses/export", methods=["GET"])
    @role_required("trainer")
    def quiz_responses_export():
        """Backend-generated .xlsx of exactly what the Quiz Responses
        table is currently showing — same filters, same query, same
        rows; nothing is regenerated or recomputed differently here."""
        college = _trainer_college()
        cohort = request.args.get("cohort") or "all"
        quiz_id = request.args.get("quizId") or "all"
        department = request.args.get("department") or "all"
        search = request.args.get("search") or ""
        rows = list_quiz_responses(
            db, college=college, cohort=cohort, quiz_id=quiz_id,
            department=department, search=search, limit=100000,
        )
        buf = build_quiz_responses_workbook(rows)
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="quiz_responses.xlsx",
        )

    @bp.route("/quiz-analytics", methods=["GET"])
    @role_required("trainer")
    def quiz_analytics():
        """Student Participation, Performance Metrics, Leaderboard and
        Section-wise Performance for this trainer's own college — the
        exact same computation Super Admin's Quiz Responses page uses
        (quiz_common.compute_quiz_analytics), just pre-scoped to the
        trainer's college rather than accepting a college filter (a
        trainer only ever sees their own college's data).
        """
        cohort = request.args.get("cohort") or "all"
        quiz_id = request.args.get("quizId") or "all"
        department = request.args.get("department") or "all"
        search = request.args.get("search") or ""
        return ok(compute_quiz_analytics(
            db, college=_trainer_college(), cohort=cohort, quiz_id=quiz_id,
            department=department, search=search,
        ))

    @bp.route("/manual-interview", methods=["GET"])
    @role_required("trainer")
    def manual_interview_list():
        college = _trainer_college()
        student_ids = [s["_id"] for s in users.find({"role": "student", "college": college}, {"_id": 1})]
        docs = manual_interviews.find({"studentId": {"$in": student_ids}}).sort("createdAt", -1)
        rows = []
        for d in docs:
            student = users.find_one({"_id": d["studentId"]}, {"passwordHash": 0})
            rows.append({
                "interviewId": str(d["_id"]),
                "studentId": str(d["studentId"]),
                "studentName": (student or {}).get("fullName"),
                "cohort": (student or {}).get("cohort") or ENTRY_LEVEL,
                "status": d.get("status"),
                "scheduledAt": d.get("scheduledAt"),
                "notes": d.get("notes"),
                "score": d.get("score"),
            })
        return ok({"manualInterviews": rows})

    @bp.route("/manual-interview/<student_id>/schedule", methods=["POST"])
    @role_required("trainer")
    def schedule_manual_interview(student_id):
        """Step 1: schedule (or reschedule) a manual interview for a student.
        This does NOT score the interview or touch cohort generation."""
        data = request.get_json(silent=True) or {}
        oid = to_object_id(student_id)
        if not oid or not users.find_one({"_id": oid, "role": "student"}):
            return error("Invalid student id.", 404)

        doc = {
            "studentId": oid,
            "status": "scheduled",
            "scheduledAt": data.get("scheduledAt"),
            "notes": data.get("notes", ""),
            "score": None,
            "scheduledBy": get_jwt_identity(),
            "createdAt": now(),
            "updatedAt": now(),
        }
        existing = manual_interviews.find_one({"studentId": oid, "status": {"$ne": "completed"}})
        if existing:
            manual_interviews.update_one({"_id": existing["_id"]}, {"$set": {
                "scheduledAt": doc["scheduledAt"], "notes": doc["notes"], "updatedAt": now(),
            }})
            interview_id = existing["_id"]
        else:
            interview_id = manual_interviews.insert_one(doc).inserted_id

        student = users.find_one({"_id": oid})
        log_activity(
            db, get_jwt_identity(), "trainer", "manual_interview_scheduled",
            f'Scheduled manual interview for {(student or {}).get("fullName", "a student")}',
            college=_trainer_college(), student_id=oid,
            meta={"interviewId": str(interview_id)},
        )
        return ok({"interviewId": str(interview_id)}, message="Manual interview scheduled.")

    @bp.route("/manual-interview/<interview_id>/score", methods=["POST"])
    @role_required("trainer")
    def score_manual_interview(interview_id):
        """
        Step 2: record the interview score (0-100). This feeds the
        two-stage cohort generation — once BOTH the assessment score
        and this interview score exist, a cohort (A/B/C) is generated
        automatically using the Super Admin's current Placement Rules.
        """
        data = request.get_json(silent=True) or {}
        oid = to_object_id(interview_id)
        if not oid:
            return error("Invalid interview id.", 404)
        interview = manual_interviews.find_one({"_id": oid})
        if not interview:
            return error("Manual interview not found.", 404)

        score = data.get("score")
        if score is None or not isinstance(score, (int, float)) or not (0 <= score <= 100):
            return error("score is required and must be a number between 0 and 100.")

        manual_interviews.update_one(
            {"_id": oid},
            {"$set": {
                "status": "completed",
                "score": score,
                "notes": data.get("notes", interview.get("notes", "")),
                "conductedBy": get_jwt_identity(),
                "completedAt": now(),
                "updatedAt": now(),
            }},
        )

        new_cohort = record_interview_score_for_cohort(db, interview["studentId"], score)
        student = users.find_one({"_id": interview["studentId"]})
        log_activity(
            db, get_jwt_identity(), "trainer", "manual_interview_scored",
            f'Scored manual interview for {(student or {}).get("fullName", "a student")} — {score}%',
            college=_trainer_college(), student_id=interview["studentId"],
            meta={"interviewId": str(oid)},
        )
        response = {"interviewId": str(oid), "score": score}
        if new_cohort:
            response["cohortGenerated"] = new_cohort
        return ok(response, message="Manual interview scored.")

    # ==========================================================
    # MARKS MANAGEMENT — Create-Quiz results, Interview Verification,
    # and Validation Verification. Deliberately separate from the
    # /quiz-responses + /manual-interview routes above, which power the
    # OLDER random-baseline assessment engine (db.assessments). This
    # block is the workflow for manually authored quizzes (db.quizzes /
    # db.quiz_attempts): a student appears here the instant — and ONLY
    # the instant — they submit a quiz; never before.
    # ==========================================================
    @bp.route("/quiz-results", methods=["GET"])
    @role_required("trainer")
    def quiz_results():
        """Assessment Responses: every student who has actually
        submitted a Create-Quiz quiz, scoped to this trainer's college.
        `search` matches Student Name / Roll Number / Register Number,
        resolved entirely by the database (see quiz_common._student_search_query)."""
        search = request.args.get("search") or ""
        return ok({"results": list_quiz_results(db, college=_trainer_college(), search=search)})

    @bp.route("/quiz-interview-verification", methods=["GET"])
    @role_required("trainer")
    def quiz_interview_verification():
        """Students eligible for interview verification: anyone who has
        submitted at least one quiz. Includes rows already scored (so the
        trainer can see/update them) — the frontend groups by `status`.
        `search` matches Assessment Name / Student Name / Roll Number /
        College / Department, resolved entirely by the database (see
        quiz_common._verification_search_query)."""
        search = request.args.get("search") or ""
        return ok({"results": list_quiz_results(db, college=_trainer_college(), search=search, broad_search=True)})

    @bp.route("/quiz-interview-verification/<attempt_id>/marks", methods=["POST"])
    @role_required("trainer")
    def enter_quiz_interview_marks(attempt_id):
        oid = to_object_id(attempt_id)
        if not oid:
            return error("Invalid quiz result id.", 404)
        data = request.get_json(silent=True) or {}
        updated, err = set_quiz_interview_marks(
            db, oid, data.get("marks"), get_jwt_identity(), college=_trainer_college(),
        )
        if err:
            return error(err)
        log_activity(
            db, get_jwt_identity(), "trainer", "quiz_interview_scored",
            f'Entered interview marks for {updated.get("studentName", "a student")} '
            f'({updated.get("quizTitle", "a quiz")}) — {updated.get("interviewMarks")}%',
            college=_trainer_college(), student_id=updated.get("studentId"),
            meta={"attemptId": attempt_id, "finalAverage": updated.get("finalAverage"),
                  "cohort": updated.get("assignedCohort")},
        )
        return ok({"result": serialize_quiz_result(db, updated)}, message="Interview marks saved.")

    @bp.route("/quiz-validation-verification", methods=["GET"])
    @role_required("trainer")
    def quiz_validation_verification():
        """Only results that have a Final Average + assigned Cohort —
        i.e. interview marks have already been entered. `search` matches
        Assessment Name / Student Name / Roll Number / College /
        Department, same as Interview Verification above."""
        search = request.args.get("search") or ""
        return ok({"results": list_quiz_results(
            db, college=_trainer_college(),
            statuses=[RESULT_STATUS_INTERVIEW_DONE, RESULT_STATUS_VALIDATED],
            search=search, broad_search=True,
        )})

    @bp.route("/quiz-validation-verification/<attempt_id>/validate", methods=["POST"])
    @role_required("trainer")
    def validate_quiz_result_route(attempt_id):
        oid = to_object_id(attempt_id)
        if not oid:
            return error("Invalid quiz result id.", 404)
        updated, err = validate_quiz_result(db, oid, get_jwt_identity(), college=_trainer_college())
        if err:
            return error(err)
        log_activity(
            db, get_jwt_identity(), "trainer", "quiz_result_validated",
            f'Validated result for {updated.get("studentName", "a student")} '
            f'({updated.get("quizTitle", "a quiz")}) — Cohort {updated.get("assignedCohort")}',
            college=_trainer_college(), student_id=updated.get("studentId"),
            meta={"attemptId": attempt_id},
        )
        return ok({"result": serialize_quiz_result(db, updated)}, message="Result validated.")

    # ==========================================================
    # RECENT ACTIVITY — replaces the hardcoded activity-list on the
    # Trainer dashboard. Every entry here comes from log_activity()
    # calls above (and, over time, other trainer/student actions).
    # ==========================================================
    @bp.route("/activity", methods=["GET"])
    @role_required("trainer")
    def recent_activity():
        trainer = _trainer_doc()
        actor_id = str(trainer["_id"]) if trainer and trainer.get("_id") else None
        limit = max(1, min(int(request.args.get("limit", 5)), 50))
        query = {"actorId": actor_id} if actor_id else {}
        rows = get_recent_activity(db, query, limit=limit)
        return ok({"activity": rows})

    # ==========================================================
    # DASHBOARD SUMMARY — real numbers for the hero stat pills.
    # Sessions Today and Attendance % are now included: db.workshop_sessions
    # and db.attendance both exist and are actively used elsewhere in this
    # file (Upcoming Sessions, Mark Attendance), so hiding these two pills
    # is no longer necessary — doing so would itself be withholding real,
    # available data from the dashboard.
    # ==========================================================
    @bp.route("/dashboard/summary", methods=["GET"])
    @role_required("trainer")
    def dashboard_summary():
        # Trainers manage the whole roster (no trainer↔student binding), so
        # every hero value is platform-wide. Average Score reads the same
        # db.quiz_attempts collection students actually submit into via Quiz
        # Management (db.assessment_attempts is the older, unused engine and
        # stays empty — see quiz-responses notes).
        total_students = users.count_documents({"role": "student"})
        scores = [
            a["overall"]["percentage"]
            for a in quiz_attempts.find(
                {"status": "submitted"},
                {"overall.percentage": 1},
            )
            if a.get("overall", {}).get("percentage") is not None
        ]
        avg_quiz_score = round(sum(scores) / len(scores), 1) if scores else None

        # Sessions Today — this trainer's own workshop sessions scheduled
        # for today's date (server-local wall clock, same convention as
        # dashboard_insights' Upcoming Sessions and the notifications feed).
        sessions_today = None
        trainer = _trainer_doc()
        if trainer is not None and trainer.get("_id") is not None:
            today_str = datetime.now().strftime("%Y-%m-%d")
            sessions_today = db.workshop_sessions.count_documents({
                "trainerIds": trainer["_id"],
                "date": today_str,
                "status": {"$in": ["scheduled", "completed"]},
            })

        # Attendance % — share of today's marked attendance records that
        # are "present", across the whole roster (trainers have no
        # college↔student binding, same scope as every other hero value).
        # If nothing has been marked yet today, this is genuinely "no data"
        # rather than a fabricated 0%.
        attendance_pct = None
        today_str = datetime.now().strftime("%Y-%m-%d")
        present = db.attendance.count_documents({"date": today_str, "status": "present"})
        absent = db.attendance.count_documents({"date": today_str, "status": "absent"})
        if present + absent > 0:
            attendance_pct = round(present / (present + absent) * 100, 1)

        return ok({
            "studentsTrained": total_students,
            "avgQuizScore": avg_quiz_score,
            "sessionsToday": sessions_today,
            "attendancePct": attendance_pct,
        })

    # ==========================================================
    # DYNAMIC CHARTS — same shape as Super Admin, scoped to college
    # ==========================================================
    @bp.route("/dashboard/charts", methods=["GET"])
    @role_required("trainer")
    def dashboard_charts():
        college = _trainer_college()
        student_ids = [s["_id"] for s in users.find({"role": "student", "college": college}, {"_id": 1})]
        # Real submissions live in db.quiz_attempts, NOT the legacy/unused
        # db.assessment_attempts engine — see dashboard_summary's note above.
        section_agg = {}
        overall_scores = []
        trend_points = []
        quizzes_cache = {}
        for att in quiz_attempts.find(
            {"status": "submitted", "studentId": {"$in": student_ids}}
        ).sort("submittedAt", 1):
            overall = att.get("overall", {})
            if "percentage" in overall:
                overall_scores.append(overall["percentage"])
                trend_points.append({
                    "date": iso_utc(att.get("submittedAt")),
                    "percentage": overall["percentage"],
                })
            for section, (correct, total) in attempt_section_counts(db, att, quizzes_cache).items():
                bucket = section_agg.setdefault(section, {"correct": 0, "total": 0})
                bucket["correct"] += correct
                bucket["total"] += total

        skill_radar = {
            section: round(b["correct"] / b["total"] * 100, 2) if b["total"] else 0
            for section, b in section_agg.items()
        }
        overall_avg = round(sum(overall_scores) / len(overall_scores), 2) if overall_scores else 0

        return ok({
            "skillRadar": skill_radar,
            "scoreByCategory": skill_radar,
            "overallScore": overall_avg,
            "categoryPercentage": skill_radar,
            "performanceTrend": trend_points,
        })

    # ==========================================================
    # SHARED AGGREGATION PASS — one scan over every submitted
    # quiz_attempts document, reused by /dashboard/insights AND
    # /dashboard/analytics so Readiness Trend, Department Ranking,
    # Skill Radar, Avg Improvement and Assessment Completion/
    # Performance can never drift out of sync with each other or
    # with one another's numbers.
    # ==========================================================
    def _trainer_attempt_pass(student_ids):
        buckets = _week_buckets()
        week_agg = {}             # week index -> [sum, n]
        dept_best = {}            # (department, college) -> {studentId: best_pct}
        cohort_best = {}          # cohort slug -> {studentId: best_pct}
        student_results = {}      # studentId -> {quizId: (submittedAt, pct)}
        attempted_ids = set()     # every student with >=1 submitted attempt

        for att in quiz_attempts.find(
            {"status": "submitted", "studentId": {"$in": student_ids}}
        ).sort("submittedAt", 1):
            pct = (att.get("overall") or {}).get("percentage")
            if pct is None:
                continue
            sid = att.get("studentId")
            submitted_at = att.get("submittedAt")
            dept = att.get("department")
            if sid is not None:
                attempted_ids.add(sid)

            if submitted_at is not None:
                for idx, (ws, we, _label) in enumerate(buckets):
                    if ws <= submitted_at < we:
                        bucket = week_agg.setdefault(idx, [0.0, 0])
                        bucket[0] += pct
                        bucket[1] += 1
                        break

            if dept and sid is not None:
                best = dept_best.setdefault((dept, att.get("college")), {})
                if sid not in best or pct > best[sid]:
                    best[sid] = pct

            if sid is not None:
                slug = att.get("cohort")
                slug = slug if slug in COHORT_RADAR_LABELS else ENTRY_LEVEL
                best = cohort_best.setdefault(slug, {})
                if sid not in best or pct > best[sid]:
                    best[sid] = pct

            if sid is not None and submitted_at is not None:
                quiz_key = str(att.get("quizId") or att.get("quizTitle") or "")
                per_student = student_results.setdefault(sid, {})
                prev = per_student.get(quiz_key)
                if prev is None or submitted_at >= prev[0]:
                    per_student[quiz_key] = (submitted_at, pct)

        return {
            "buckets": buckets,
            "week_agg": week_agg,
            "dept_best": dept_best,
            "cohort_best": cohort_best,
            "student_results": student_results,
            "attempted_ids": attempted_ids,
        }

    def _avg_improvement_deltas(student_results):
        """Mean of every per-student delta between two consecutive,
        distinct quiz results (previous -> subsequent), in the
        chronological order of actual submissions — never a hardcoded
        'Baseline -> Mid' style label. Shared by /dashboard/insights'
        Improvement Ring and /dashboard/analytics' Avg Improvement stat
        so the two can never disagree."""
        deltas = []
        for per_student in student_results.values():
            ordered = sorted(per_student.values(), key=lambda item: item[0])
            for i in range(1, len(ordered)):
                deltas.append(ordered[i][1] - ordered[i - 1][1])
        return deltas

    # ==========================================================
    # DASHBOARD INSIGHTS — Readiness Trend, Department Ranking,
    # Skill Radar, Avg Improvement ring and Upcoming Sessions, all
    # computed live from db.quiz_attempts (the collection real student
    # quiz submissions actually land in — see dashboard_summary's note;
    # db.assessment_attempts is the legacy, unused engine and always
    # stays empty) / db.workshop_sessions, scoped to the trainer's own
    # college. The frontend renders these five widgets entirely from
    # this one response.
    # ==========================================================
    @bp.route("/dashboard/insights", methods=["GET"])
    @role_required("trainer")
    def dashboard_insights():
        trainer = _trainer_doc()
        # Trainers manage the whole roster — insights aggregate every
        # submitted quiz attempt platform-wide.
        student_ids = [s["_id"] for s in users.find({"role": "student"}, {"_id": 1})]

        empty = {
            "readinessTrend": [],
            "departmentRanking": [],
            "skillRadar": [],
            "improvementRing": None,
            "upcomingSessions": [],
        }
        if not student_ids:
            return ok(empty)

        agg = _trainer_attempt_pass(student_ids)
        buckets = agg["buckets"]
        week_agg = agg["week_agg"]
        dept_best = agg["dept_best"]
        cohort_best = agg["cohort_best"]
        student_results = agg["student_results"]

        readiness_trend = [
            {"label": buckets[idx][2], "value": round(week_agg[idx][0] / week_agg[idx][1], 1)}
            for idx in sorted(week_agg)
        ]

        department_ranking = []
        for (dept, college), best_map in dept_best.items():
            if not best_map:
                continue
            values = list(best_map.values())
            department_ranking.append({
                "name": dept,
                "college": college,
                "value": round(sum(values) / len(values), 1),
            })
        department_ranking.sort(key=lambda r: r["value"], reverse=True)
        department_ranking = department_ranking[:5]

        skill_radar = []
        for slug in RADAR_COHORT_ORDER:
            best_map = cohort_best.get(slug)
            if not best_map:
                continue
            skill_radar.append({
                "label": COHORT_RADAR_LABELS[slug],
                "value": round(sum(best_map.values()) / len(best_map), 1),
            })

        deltas = _avg_improvement_deltas(student_results)

        improvement_ring = None
        if deltas:
            avg_delta = sum(deltas) / len(deltas)
            improvement_ring = {
                "pct": round(max(0.0, min(100.0, avg_delta)), 1),
                "improvementLabel": "{}{}%".format(
                    "+" if avg_delta > 0 else "", round(avg_delta, 1)
                ),
                "workshopsCount": None,
                "avgRating": None,
            }

        # 5. Upcoming Sessions — this trainer's assigned workshop sessions
        # whose start datetime hasn't passed yet. Dates are stored as
        # server-local wall clock (YYYY-MM-DD + HH:MM), the same convention
        # superadmin.py's session rules use.
        upcoming_sessions = []
        if trainer is not None and trainer.get("_id") is not None:
            trainer_oid = trainer["_id"]
            today = datetime.now()
            if improvement_ring is not None:
                improvement_ring["workshopsCount"] = db.workshop_sessions.count_documents(
                    {"trainerIds": trainer_oid}
                )
            for doc in db.workshop_sessions.find({
                "trainerIds": trainer_oid,
                "status": "scheduled",
                "date": {"$gte": today.strftime("%Y-%m-%d")},
            }):
                date_str = doc.get("date")
                start_time = doc.get("startTime")
                try:
                    start_dt = datetime.strptime(f"{date_str} {start_time}", "%Y-%m-%d %H:%M")
                except (ValueError, TypeError):
                    continue
                if start_dt < today:
                    continue
                end_time = doc.get("endTime")
                parsed_date = datetime.strptime(date_str, "%Y-%m-%d")
                time_str = f"{start_time}\u2013{end_time}" if end_time else start_time
                upcoming_sessions.append({
                    "name": doc.get("name"),
                    "meta": f"{parsed_date.strftime('%d %b %Y')} \u2022 {time_str}",
                    "isToday": date_str == today.strftime("%Y-%m-%d"),
                    "_sort_dt": start_dt,
                })
            upcoming_sessions.sort(key=lambda s: s["_sort_dt"])
            for s in upcoming_sessions:
                s.pop("_sort_dt", None)
            upcoming_sessions = upcoming_sessions[:5]

        return ok({
            "readinessTrend": readiness_trend,
            "departmentRanking": department_ranking,
            "skillRadar": skill_radar,
            "improvementRing": improvement_ring,
            "upcomingSessions": upcoming_sessions,
        })

    # ==========================================================
    # ANALYTICS PAGE — was previously calling a route that did not
    # exist (404 on every visit to the Analytics tab). Reuses the exact
    # same shared aggregation pass as /dashboard/insights so Avg
    # Improvement here can never disagree with the Improvement Ring on
    # the dashboard home, and never invents a second scoring formula.
    #
    # Workshop Effectiveness has NO backing data anywhere in the schema
    # (db.workshop_sessions has no rating/feedback field) — rather than
    # fabricate a number, that stat and its bars are returned as
    # genuine no-data states, same principle as improvementRing's
    # avgRating already uses.
    # ==========================================================
    @bp.route("/dashboard/analytics", methods=["GET"])
    @role_required("trainer")
    def dashboard_analytics():
        student_ids = [s["_id"] for s in users.find({"role": "student"}, {"_id": 1})]
        total_students = len(student_ids)

        empty = {
            "assessmentCompletion": {"value": None, "delta": None},
            "workshopEffectiveness": {"value": None, "delta": None},
            "avgImprovement": {"value": None, "delta": None},
            "assessmentPerformance": [],
            "workshopEffectivenessBars": [],
        }
        if not student_ids:
            return ok(empty)

        agg = _trainer_attempt_pass(student_ids)
        buckets = agg["buckets"]
        week_agg = agg["week_agg"]
        student_results = agg["student_results"]
        attempted_ids = agg["attempted_ids"]

        # Assessment Completion — % of the roster with >=1 submitted quiz
        # attempt, as of now vs. as of 7 days ago (same denominator both
        # times, isolating the change in submission behaviour).
        completion_now = round(len(attempted_ids) / total_students * 100, 1) if total_students else None
        week_ago = datetime.utcnow() - timedelta(days=7)
        attempted_week_ago = {
            a["studentId"] for a in quiz_attempts.find(
                {"status": "submitted", "studentId": {"$in": student_ids}, "submittedAt": {"$lte": week_ago}},
                {"studentId": 1},
            )
        }
        completion_week_ago = round(len(attempted_week_ago) / total_students * 100, 1) if total_students else None
        completion_delta = (
            round(completion_now - completion_week_ago, 1)
            if completion_now is not None and completion_week_ago is not None else None
        )

        # Avg Improvement — identical calculation to the dashboard home's
        # Improvement Ring (previous -> next submitted quiz, chronological
        # order), so the two never disagree.
        deltas = _avg_improvement_deltas(student_results)
        avg_improvement = round(sum(deltas) / len(deltas), 1) if deltas else None

        # Assessment Performance Over Time — the same weekly readiness
        # buckets as the dashboard home's Readiness Trend, reshaped as bars.
        assessment_performance = [
            {"label": buckets[idx][2], "pct": round(week_agg[idx][0] / week_agg[idx][1], 1)}
            for idx in sorted(week_agg)
        ]

        # This stat band's frontend renders value/delta as raw text with no
        # formatting applied client-side (unlike readiness trend/department
        # ranking, which append "%" in their own templates) — so values and
        # deltas are formatted here as ready-to-display strings.
        def _fmt_pct(value):
            return f"{value}%" if value is not None else None

        def _fmt_delta(delta):
            if delta is None:
                return None
            return "{}{}%".format("+" if delta > 0 else "", delta)

        return ok({
            "assessmentCompletion": {"value": _fmt_pct(completion_now), "delta": _fmt_delta(completion_delta)},
            "workshopEffectiveness": {"value": None, "delta": None},
            "avgImprovement": {"value": _fmt_delta(avg_improvement) if avg_improvement is not None else None, "delta": None},
            "assessmentPerformance": assessment_performance,
            "workshopEffectivenessBars": [],
        })

    # ==========================================================
    # SKILL RADAR (per cohort) — powers the Skill Radar widget's
    # Entry Level / Cohort A / Cohort B / Cohort C selector. Unlike
    # dashboard_insights' skillRadar (one point per cohort, showing
    # how cohorts compare to each other), this returns the SELECTED
    # cohort's own per-section average, so a trainer can drill into
    # what that cohort is actually strong/weak at. Computed live from
    # db.quiz_attempts — each attempt is scored against its own quiz's
    # question pool via attempt_section_counts() (quiz_attempts never
    # stores a precomputed sectionScores field the way the legacy
    # db.assessment_attempts engine did) — never hardcoded, and a
    # cohort with no submitted attempts yet returns an empty list so
    # the frontend can show a proper no-data state instead of zeros.
    # ==========================================================
    @bp.route("/dashboard/skill-radar", methods=["GET"])
    @role_required("trainer")
    def dashboard_skill_radar():
        cohort = (request.args.get("cohort") or ENTRY_LEVEL).strip()
        if cohort not in RADAR_COHORT_ORDER:
            return error("Invalid cohort. Use entry_level, A, B or C.", status=400)

        student_ids = [s["_id"] for s in users.find({"role": "student"}, {"_id": 1})]
        section_agg = {}
        quizzes_cache = {}
        if student_ids:
            for att in quiz_attempts.find({"status": "submitted", "studentId": {"$in": student_ids}}):
                slug = att.get("cohort")
                slug = slug if slug in COHORT_RADAR_LABELS else ENTRY_LEVEL
                if slug != cohort:
                    continue
                for section, (correct, total) in attempt_section_counts(db, att, quizzes_cache).items():
                    bucket = section_agg.setdefault(section, {"correct": 0, "total": 0})
                    bucket["correct"] += correct
                    bucket["total"] += total

        skill_radar = [
            {"label": section, "value": round(b["correct"] / b["total"] * 100, 1)}
            for section, b in section_agg.items() if b["total"]
        ]
        skill_radar.sort(key=lambda r: r["label"])
        return ok({"cohort": cohort, "cohortLabel": COHORT_RADAR_LABELS[cohort], "skillRadar": skill_radar})

    # ==========================================================
    # NOTIFICATIONS — the Trainer dashboard bell. Aggregates every
    # trainer-facing pending item into one live feed; nothing here is
    # ever stored, each item is recomputed from its source collection
    # on every call (student approvals, quiz verification queues and
    # the trainer's own upcoming sessions). Only read/unread state
    # persists, in the same db.notification_reads collection the Super
    # Admin feed uses — each id is scoped with the trainer's user id so
    # the two dashboards can never share read-state by accident.
    # ==========================================================
    @bp.route("/dashboard/notifications", methods=["GET"])
    @role_required("trainer")
    def trainer_dashboard_notifications():
        trainer = _trainer_doc()
        trainer_id = str(trainer["_id"]) if trainer and trainer.get("_id") else "anonymous"
        college = _trainer_college()

        def _notification_id(kind, source_id):
            """Deterministic, trainer-scoped id so read/unread state in
            db.notification_reads survives between requests and never
            collides with the Super Admin's identically-named kinds."""
            raw = f"trainer:{kind}:{trainer_id}:{source_id}"
            return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:24]

        now_dt = datetime.utcnow()
        recent_cutoff = now_dt - timedelta(days=2)
        items = []

        # 1. Pending Approvals — student registrations awaiting
        # verification. Same shared queue as the Student Verification
        # page (GET /api/auth/pending/students), so the count here
        # always matches the badge on that page.
        pending_students = users.count_documents({"role": "student", "approvalStatus": "pending"})
        if pending_students:
            items.append({
                "type": "student_verification",
                "action": "verification",
                "title": "Student approvals pending",
                "description": f"{pending_students} student registration(s) awaiting verification.",
                "createdAt": now_dt,
                "sourceId": "student-pending-bucket",
            })

        # 2. Pending Approvals — quiz results waiting for interview
        # marks (the Interview Verification tab), scoped to the same
        # college that tab queries.
        interview_query = {"status": "submitted", "resultStatus": RESULT_STATUS_PENDING}
        if college:
            interview_query["college"] = college
        interview_pending = db.quiz_attempts.count_documents(interview_query)
        if interview_pending:
            items.append({
                "type": "interview_verification",
                "action": "interview-verification",
                "title": "Interview verification pending",
                "description": f"{interview_pending} quiz result(s) are waiting for interview marks.",
                "createdAt": now_dt,
                "sourceId": "interview-pending-bucket",
            })

        # 3. Pending Approvals — quiz results whose interview marks are
        # in but which have not been validated (Validation Verification).
        validation_query = {"status": "submitted", "resultStatus": RESULT_STATUS_INTERVIEW_DONE}
        if college:
            validation_query["college"] = college
        validation_pending = db.quiz_attempts.count_documents(validation_query)
        if validation_pending:
            items.append({
                "type": "validation_verification",
                "action": "validation-verification",
                "title": "Validation verification pending",
                "description": f"{validation_pending} quiz result(s) are waiting for validation.",
                "createdAt": now_dt,
                "sourceId": "validation-pending-bucket",
            })

        # 4. Other pending items — quiz submissions in the last 48h
        # that have not been reviewed in the Quiz Responses tab.
        submissions_query = {"status": "submitted", "submittedAt": {"$gte": recent_cutoff}}
        if college:
            submissions_query["college"] = college
        new_submissions = db.quiz_attempts.count_documents(submissions_query)
        if new_submissions:
            items.append({
                "type": "new_submissions",
                "action": "responses",
                "title": "New quiz submissions to review",
                "description": f"{new_submissions} quiz submission(s) landed in the last 48 hours.",
                "createdAt": now_dt,
                "sourceId": "new-submissions-bucket",
            })

        # 5. Upcoming Session Reminder — only generated when a session
        # assigned to THIS trainer starts within the next 2 hours. Uses
        # the same server-local wall-clock (YYYY-MM-DD HH:MM) convention
        # as dashboard_insights' upcoming-sessions query, so a session
        # that already started (or is further out than 2 hours) produces
        # no reminder at all.
        if trainer is not None and trainer.get("_id") is not None:
            trainer_oid = trainer["_id"]
            local_now = datetime.now()
            reminder_cutoff = local_now + timedelta(hours=2)
            for doc in db.workshop_sessions.find(
                {"trainerIds": trainer_oid, "status": "scheduled"},
                {"name": 1, "date": 1, "startTime": 1, "endTime": 1, "mode": 1, "venue": 1},
            ):
                try:
                    start_dt = datetime.strptime(
                        f"{doc.get('date')} {doc.get('startTime')}", "%Y-%m-%d %H:%M"
                    )
                except (ValueError, TypeError):
                    continue
                if start_dt <= local_now or start_dt > reminder_cutoff:
                    continue
                end_time = doc.get("endTime")
                time_str = f"{doc.get('startTime')}\u2013{end_time}" if end_time else doc.get("startTime")
                when = f"{doc.get('date')} {time_str}"
                where = f" at {doc.get('venue')}" if doc.get("venue") else ""
                items.append({
                    "type": "session_reminder",
                    "action": "workshops",
                    "title": "Upcoming session reminder",
                    "description": f'"{doc.get("name", "A session")}" starts soon — {when}{where}.',
                    "createdAt": now_dt,
                    "sourceId": str(doc["_id"]),
                })

        for it in items:
            it["id"] = _notification_id(it["type"], it["sourceId"])
            it["createdAt"] = iso_utc(it["createdAt"]) if it["createdAt"] else None

        ids = [it["id"] for it in items]
        read_ids = {r["_id"] for r in db.notification_reads.find({"_id": {"$in": ids}}, {"_id": 1})}
        for it in items:
            it["unread"] = it["id"] not in read_ids

        items.sort(key=lambda x: x["createdAt"] or "", reverse=True)
        return ok({"notifications": items, "unreadCount": sum(1 for it in items if it["unread"])})

    @bp.route("/dashboard/notifications/<notification_id>/read", methods=["PATCH"])
    @role_required("trainer")
    def mark_trainer_notification_read(notification_id):
        db.notification_reads.update_one(
            {"_id": notification_id}, {"$set": {"readAt": now()}}, upsert=True,
        )
        return ok(message="Marked as read.")

    @bp.route("/dashboard/notifications/read-all", methods=["POST"])
    @role_required("trainer")
    def mark_all_trainer_notifications_read():
        data = request.get_json(silent=True) or {}
        ids = data.get("ids") or []
        if not isinstance(ids, list) or not ids:
            return error("ids (non-empty list) is required.")
        ops = [
            {"filter": {"_id": nid}, "update": {"$set": {"readAt": now()}}, "upsert": True}
            for nid in ids
        ]
        for op in ops:
            db.notification_reads.update_one(op["filter"], op["update"], upsert=op["upsert"])
        return ok(message="All notifications marked as read.")

    # ==========================================================
    # STUDENT ROSTER — All Students list, single add, bulk import
    # and PDF/Excel export. Trainers manage the whole platform
    # roster (no trainer↔student binding), so the trainer picks the
    # college on the add form / per bulk row and every query here
    # is platform-wide.
    #
    # New students are created APPROVED (no approval queue) with a
    # server-generated temporary password that is returned EXACTLY
    # once so the trainer can hand it to the student securely.
    # firstLoginVerify forces the student to set their own password
    # + OTP-verify on first login (existing login.py flow). Cohort
    # starts at Entry Level (None) — the identical rule to
    # self-registration (login.py) and Super Admin bulk import.
    # ==========================================================
    @bp.route("/students/departments", methods=["GET"])
    @role_required("trainer")
    def trainer_student_departments():
        """Every active college with its active departments — feeds the
        Add Student college + dynamic department dropdowns."""
        depts = list(db.departments.find({"status": "active"}).sort("department_name", 1))
        by_college = {}
        for d in depts:
            by_college.setdefault(str(d.get("college_id")), []).append(department_public(d))
        colleges = []
        for c in db.colleges.find({"status": "active"}).sort("college_name", 1):
            colleges.append({
                "id": str(c["_id"]),
                "name": c.get("college_name"),
                "departments": by_college.get(str(c["_id"]), []),
            })
        return ok({"colleges": colleges})

    @bp.route("/students", methods=["POST"])
    @role_required("trainer")
    def create_student():
        data = request.get_json(silent=True) or {}
        full_name = (data.get("fullName") or "").strip()
        email = (data.get("email") or "").strip().lower()
        mobile = (data.get("mobile") or "").strip()
        roll_number = (data.get("rollNumber") or "").strip()
        department_id = data.get("departmentId")
        college_id = data.get("collegeId")

        if not full_name:
            return error("Full name is required.")
        if not roll_number:
            return error("Roll Number is required.")
        if not department_id:
            return error("Department is required.")
        if not college_id:
            return error("Select a college.")

        email_err = validate_email(email)
        if email_err:
            return error(email_err)
        mobile_err = validate_mobile(mobile)
        if mobile_err:
            return error(mobile_err)

        if users.find_one({"email": email}):
            return error("An account with this email already exists.")
        if users.find_one({"role": "student", "rollNumber": roll_number}):
            return error("A student with this Roll Number already exists.")

        # Exact college/department relationship check: the department must
        # actually belong to the chosen (active) college.
        college_doc = db.colleges.find_one({"_id": to_object_id(college_id), "status": "active"})
        if not college_doc:
            return error("Selected college is invalid or inactive.")
        department_doc = resolve_active_department(db, department_id, str(college_id))
        if not department_doc:
            return error("Selected department is invalid or inactive for the chosen college.")

        # Password: the backend always generates a random temporary
        # password for the student — the trainer never enters one. The
        # student must change it at first login (firstLoginVerify).
        temp_password = _generate_temp_password()

        if bcrypt is None:
            return error("Student creation is not available right now.", 503)

        password_hash = bcrypt.generate_password_hash(temp_password).decode("utf-8")
        user_doc = {
            "fullName": full_name,
            "email": email,
            "mobile": mobile,
            "role": "student",
            "passwordHash": password_hash,
            "approvalStatus": "approved",
            "approvedBy": f"trainer:{get_jwt_identity()}",
            "approvedDate": now(),
            "googleLogin": False,
            "isDeleted": False,
            "firstLoginVerify": True,
            "firstLoginVerifiedAt": None,
            "cohort": None,
            "baselineAssessmentScore": None,
            "interviewScore": None,
            "finalEmployabilityScore": None,
            "cohortAssignedAt": None,
            "rollNumber": roll_number,
            "tneaCode": None,
            "college": college_doc["college_name"],
            "collegeId": college_doc["_id"],
            "department": department_doc["department_name"],
            "departmentId": department_doc["_id"],
            "district": None,
            "employeeId": None,
            "createdAt": now(),
            "updatedAt": now(),
        }
        result = users.insert_one(user_doc)
        log_activity(
            db, get_jwt_identity(), "trainer", "student_added",
            f'Added student {full_name} ({roll_number}) to {department_doc["department_name"]}, {college_doc["college_name"]}',
            college=college_doc["college_name"], student_id=result.inserted_id,
            meta={"department": department_doc["department_name"], "college": college_doc["college_name"]},
        )
        return ok({
            "student": {
                "id": str(result.inserted_id),
                "name": full_name,
                "email": email,
                "rollNumber": roll_number,
                "college": college_doc["college_name"],
                "department": department_doc["department_name"],
            },
            "temporaryPassword": temp_password,
            "passwordSet": False,
        }, message="Student created. They must set their own password at first login.", status=201)

    @bp.route("/students/bulk-import", methods=["POST"])
    @role_required("trainer")
    def bulk_import_students():
        data = request.get_json(silent=True) or {}
        rows = data.get("students") or []
        if not isinstance(rows, list) or not rows:
            return error("No student rows provided.")
        if len(rows) > 1000:
            return error("Maximum 1000 rows per import.")
        if bcrypt is None:
            return error("Student import is not available right now.", 503)

        colleges_by_id = {}
        colleges_by_name = {}
        for c in db.colleges.find({}):
            colleges_by_id[str(c["_id"])] = c
            colleges_by_name[str(c.get("college_name") or "").strip().lower()] = c

        existing_emails = {u["email"] for u in users.find(
            {"role": "student", "email": {"$ne": None}}, {"email": 1})}
        existing_rolls = {str(u["rollNumber"]).lower() for u in users.find(
            {"role": "student", "rollNumber": {"$ne": None}}, {"rollNumber": 1})}

        imported = 0
        rejected = []
        created = []
        batch_emails = set()
        batch_rolls = set()

        for i, r in enumerate(rows):
            if not isinstance(r, dict):
                rejected.append({"index": i, "name": "", "reason": "Row is not an object."})
                continue
            name = (r.get("name") or "").strip()
            roll = (r.get("roll") or "").strip()
            email = (r.get("email") or "").strip().lower()
            phone = (r.get("phone") or "").strip()
            dept = (r.get("dept") or "").strip()
            college_ref = (r.get("college") or "").strip()

            reasons = []
            if not name:
                reasons.append("Missing name")
            if not roll:
                reasons.append("Missing roll number")
            email_err = validate_email(email) if email else "Missing email"
            if email_err:
                reasons.append(email_err)
            mobile_err = validate_mobile(phone) if phone else "Missing phone"
            if mobile_err:
                reasons.append(mobile_err)
            if not dept:
                reasons.append("Missing department")
            if not college_ref:
                reasons.append("Missing college")

            college_doc = None
            if college_ref:
                college_doc = colleges_by_id.get(college_ref) or colleges_by_name.get(college_ref.lower())
                if not college_doc:
                    reasons.append(f"Unknown college: {college_ref}")

            dept_doc = None
            if college_doc and dept:
                dept_doc = db.departments.find_one({
                    "college_id": college_doc["_id"],
                    "department_name": {"$regex": f"^{re.escape(dept)}$", "$options": "i"},
                    "status": "active",
                })
                if not dept_doc:
                    reasons.append(f"Department '{dept}' does not exist in {college_doc.get('college_name')}")

            if email:
                if email in existing_emails or email in batch_emails:
                    reasons.append("Email already registered")
            if roll:
                if roll.lower() in existing_rolls or roll.lower() in batch_rolls:
                    reasons.append("Roll number already exists")

            if reasons:
                rejected.append({"index": i, "name": name, "reason": "; ".join(reasons)})
                continue

            temp_password = _generate_temp_password()
            password_hash = bcrypt.generate_password_hash(temp_password).decode("utf-8")
            user_doc = {
                "fullName": name,
                "email": email,
                "mobile": phone,
                "role": "student",
                "passwordHash": password_hash,
                "approvalStatus": "approved",
                "approvedBy": f"trainer:{get_jwt_identity()}",
                "approvedDate": now(),
                "googleLogin": False,
                "isDeleted": False,
                "firstLoginVerify": True,
                "firstLoginVerifiedAt": None,
                "cohort": None,
                "baselineAssessmentScore": None,
                "interviewScore": None,
                "finalEmployabilityScore": None,
                "cohortAssignedAt": None,
                "rollNumber": roll,
                "tneaCode": None,
                "college": college_doc["college_name"],
                "collegeId": college_doc["_id"],
                "department": dept_doc["department_name"],
                "departmentId": dept_doc["_id"],
                "district": None,
                "employeeId": None,
                "createdAt": now(),
                "updatedAt": now(),
            }
            res = users.insert_one(user_doc)
            existing_emails.add(email)
            batch_emails.add(email)
            existing_rolls.add(roll.lower())
            batch_rolls.add(roll.lower())
            imported += 1
            created.append({
                "id": str(res.inserted_id), "name": name, "email": email,
                "temporaryPassword": temp_password,
            })

        if imported:
            log_activity(
                db, get_jwt_identity(), "trainer", "student_bulk_imported",
                f"Bulk imported {imported} student{'s' if imported != 1 else ''} "
                f"({len(rejected)} rejected)",
                meta={"imported": imported, "rejected": len(rejected)},
            )
        return ok({
            "imported": imported,
            "rejected": rejected,
            "students": created,
        }, message=f"{imported} student(s) imported.")

    @bp.route("/students/export", methods=["GET"])
    @role_required("trainer")
    def students_export():
        """Server-side PDF / Excel export of exactly what the All Students
        table is currently showing (same filters, same query)."""
        rows = _student_roster(
            college=(request.args.get("college") or "").strip(),
            department=(request.args.get("department") or "").strip(),
            cohort=(request.args.get("cohort") or "").strip(),
            search=(request.args.get("search") or "").strip(),
        )
        columns = ["Student Name", "Roll Number", "Email", "Mobile", "College",
                   "Department", "Cohort", "Attendance %", "Avg Quiz Score", "Status"]

        def _cohort_label(c):
            if c in {"A", "B", "C"}:
                return f"Cohort {c}"
            return c or "Entry Level"

        data_rows = [[
            r["fullName"], r["rollNumber"], r["email"], r["mobile"], r["college"],
            r["department"], _cohort_label(r["cohort"]),
            r["attendancePct"] if r["attendancePct"] is not None else "—",
            r["quizScore"] if r["quizScore"] is not None else "—",
            r["status"],
        ] for r in rows]
        fmt = (request.args.get("format") or "xlsx").lower()
        if fmt == "pdf":
            buf = pdf_bytes(
                "Student Directory",
                f"Generated {fmt_ist(now())} IST · {len(data_rows)} students",
                columns, data_rows,
            )
            return send_file(buf, mimetype="application/pdf", as_attachment=True,
                             download_name="student_directory.pdf")
        buf = excel_bytes(columns, data_rows, "Students")
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="student_directory.xlsx")

    # ==========================================================
    # TRAINER — CALENDAR EVENTS (db.calendar_events).
    # The trainer's own calendar. Event creation is deliberately
    # minimal — Title, Date, From Time, To Time, Description only,
    # no lookup/type fields. Events live in their own collection
    # (separate from the student-facing db.meetings feed) and are
    # surfaced back in the calendar grid for the requested month.
    # ==========================================================
    @bp.route("/calendar/events", methods=["GET"])
    @role_required("trainer")
    def trainer_calendar_events():
        try:
            month = int(request.args.get("month") or datetime.now().month)
            year = int(request.args.get("year") or datetime.now().year)
        except ValueError:
            return error("month and year must be integers.")
        if not 1 <= month <= 12:
            return error("month must be between 1 and 12.")
        prefix = f"{year:04d}-{month:02d}"
        docs = list(db.calendar_events.find(
            {"date": {"$regex": f"^{re.escape(prefix)}"}}
        ).sort([("date", 1), ("fromTime", 1)]))
        return ok({"events": [{
            "id": str(d["_id"]),
            "title": d.get("title"),
            "date": d.get("date"),
            "fromTime": d.get("fromTime"),
            "toTime": d.get("toTime"),
            "description": d.get("description") or "",
            "day": int(d["date"].split("-")[2]),
            "type": "event",
            "label": d.get("title") or "Event",
        } for d in docs]})

    @bp.route("/calendar/events", methods=["POST"])
    @role_required("trainer")
    def trainer_calendar_event_create():
        data = request.get_json(silent=True) or {}
        title = (data.get("title") or "").strip()
        date_str = (data.get("date") or "").strip()
        from_time = (data.get("fromTime") or "").strip()
        to_time = (data.get("toTime") or "").strip()
        description = (data.get("description") or "").strip()

        if not title:
            return error("Event title is required.")
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", date_str):
            return error("A valid event Date (YYYY-MM-DD) is required.")
        if not (_TIME_RE.match(from_time) and _TIME_RE.match(to_time)):
            return error("A valid From Time and To Time are required.")
        if to_time <= from_time:
            return error("To Time must be after From Time.")

        trainer = _trainer_doc()
        doc = {
            "title": title,
            "date": date_str,
            "fromTime": from_time,
            "toTime": to_time,
            "description": description,
            "createdBy": {"id": str(trainer["_id"]), "name": trainer.get("fullName") or "Trainer", "role": "trainer"},
            "createdAt": now(),
            "updatedAt": now(),
        }
        result = db.calendar_events.insert_one(doc)
        doc["_id"] = result.inserted_id
        log_activity(
            db, get_jwt_identity(), "trainer", "calendar_event_created",
            f'Created calendar event "{title}" on {date_str}',
            college=_trainer_college(),
            meta={"eventId": str(result.inserted_id), "date": date_str},
        )
        return ok({
            "event": {
                "id": str(result.inserted_id),
                "title": title,
                "date": date_str,
                "fromTime": from_time,
                "toTime": to_time,
                "description": description,
                "day": int(date_str.split("-")[2]),
                "type": "event",
                "label": title,
            },
        }, message="Event added to the calendar.", status=201)

    # ==========================================================
    # TRAINER — INTERVENTIONS (db.workshop_sessions).
    # The trainer submits intervention sessions for approval; the
    # Super Admin approves/rejects them (superadmin.py). This
    # reuses the exact same validation, serialization and
    # auto-completion helpers as the Super Admin workshop module
    # so the two dashboards can never disagree. A trainer may only
    # manage sessions they created, and only before they start.
    # ==========================================================
    def _trainer_owns_session(doc):
        trainer_oid = _trainer_doc().get("_id")
        return bool(doc) and trainer_oid is not None and (
            doc.get("createdBy", {}).get("id") == str(trainer_oid)
        )

    @bp.route("/interventions/options", methods=["GET"])
    @role_required("trainer")
    def trainer_intervention_options():
        colleges = [
            {"id": str(c["_id"]), "name": c.get("college_name") or "—"}
            for c in db.colleges.find({"status": "active"}).sort("college_name", 1)
        ]
        college_names = {str(c["_id"]): c["name"] for c in colleges}
        departments = [
            {"id": str(d["_id"]), "name": d.get("department_name") or "—",
             "collegeId": str(d.get("college_id")) if d.get("college_id") else None,
             "collegeName": college_names.get(str(d.get("college_id"))) or "—"}
            for d in db.departments.find({"status": "active"}).sort("department_name", 1)
        ]
        trainers = [
            {"id": str(t["_id"]), "name": t.get("fullName") or "—"}
            for t in users.find(
                {"role": "trainer", "approvalStatus": {"$in": ["approved", "suspended"]}},
                {"fullName": 1}).sort("fullName", 1)
        ]
        return ok({
            "colleges": colleges,
            "departments": departments,
            "trainers": trainers,
        })

    @bp.route("/interventions", methods=["GET"])
    @role_required("trainer")
    def trainer_list_interventions():
        _auto_complete_workshop_sessions(db)
        trainer_oid = _trainer_doc().get("_id")
        if trainer_oid is None:
            return ok({"sessions": []})
        docs = list(db.workshop_sessions.find({
            "$or": [
                {"createdBy.id": str(trainer_oid)},
                {"trainerIds": trainer_oid},
            ],
        }).sort("createdAt", -1))
        sessions = []
        for d in docs:
            pub = workshop_session_public(d)
            pub["statusLabel"] = _intervention_status_label(d)
            sessions.append(pub)
        return ok({"sessions": sessions})

    @bp.route("/interventions/schedule", methods=["POST"])
    @role_required("trainer")
    def trainer_schedule_intervention():
        data = request.get_json(silent=True) or {}
        fields, err = _validate_session_payload(data)
        if err:
            return error(err)

        trainer = _trainer_doc()
        doc = dict(fields)
        doc.update({
            "status": "scheduled",
            "approvalStatus": "pending",
            "createdBy": {"id": str(trainer["_id"]), "name": trainer.get("fullName") or "Trainer", "role": "trainer"},
            "createdByRole": "trainer",
            "createdAt": now(),
            "updatedAt": now(),
        })
        result = db.workshop_sessions.insert_one(doc)
        doc["_id"] = result.inserted_id
        log_activity(
            db, get_jwt_identity(), "trainer", "workshop_session_scheduled",
            f'Submitted intervention "{fields["name"]}" for approval',
            college=_trainer_college(),
            meta={"sessionId": str(doc["_id"])},
        )
        return ok({"session": workshop_session_public(doc)},
                  message="Intervention submitted for approval.", status=201)

    @bp.route("/interventions/sessions/<session_id>", methods=["PATCH"])
    @role_required("trainer")
    def trainer_update_intervention(session_id):
        oid = to_object_id(session_id)
        if not oid:
            return error("Invalid session id.", 404)
        existing = db.workshop_sessions.find_one({"_id": oid})
        if not existing:
            return error("Session not found.", 404)
        if not _trainer_owns_session(existing):
            return error("You can only manage interventions you created.", 403)

        data = request.get_json(silent=True) or {}
        if "status" in data and len(data) == 1:
            status = (data.get("status") or "").strip().lower()
            if status != "cancelled":
                return error("Trainers can only cancel their own interventions.")
            if _session_started(existing):
                return error("Sessions can only be cancelled before they start.", 409)
            db.workshop_sessions.update_one(
                {"_id": oid}, {"$set": {"status": "cancelled", "updatedAt": now()}})
            updated = db.workshop_sessions.find_one({"_id": oid})
            log_activity(
                db, get_jwt_identity(), "trainer", "workshop_session_cancelled",
                f'Cancelled intervention "{existing.get("name")}"',
                college=_trainer_college(),
                meta={"sessionId": str(oid)},
            )
            return ok({"session": workshop_session_public(updated)}, message="Intervention cancelled.")

        if _session_started(existing):
            return error("Sessions can only be edited before they start.", 409)
        fields, err = _validate_session_payload(data)
        if err:
            return error(err)
        fields["updatedAt"] = now()
        db.workshop_sessions.update_one({"_id": oid}, {"$set": fields})
        updated = db.workshop_sessions.find_one({"_id": oid})
        log_activity(
            db, get_jwt_identity(), "trainer", "workshop_session_updated",
            f'Updated intervention "{fields["name"]}"',
            college=_trainer_college(),
            meta={"sessionId": str(oid)},
        )
        return ok({"session": workshop_session_public(updated)}, message="Intervention updated.")

    @bp.route("/interventions/sessions/<session_id>", methods=["DELETE"])
    @role_required("trainer")
    def trainer_delete_intervention(session_id):
        oid = to_object_id(session_id)
        if not oid:
            return error("Invalid session id.", 404)
        existing = db.workshop_sessions.find_one({"_id": oid})
        if not existing:
            return error("Session not found.", 404)
        if not _trainer_owns_session(existing):
            return error("You can only manage interventions you created.", 403)
        if _session_started(existing):
            return error("Sessions can only be deleted before they start.", 409)
        db.workshop_sessions.delete_one({"_id": oid})
        log_activity(
            db, get_jwt_identity(), "trainer", "workshop_session_deleted",
            f'Deleted intervention "{existing.get("name")}"',
            college=_trainer_college(),
            meta={"sessionId": str(oid)},
        )
        return ok(message="Intervention deleted.")

    # ==========================================================
    # TRAINER — ATTENDANCE MODULE (db.attendance).
    # Trainer-authenticated mirror of the Super Admin "Mark Attendance" /
    # "View Records" routes in superadmin.py. Same collection, same
    # query builder (_attendance_student_query), same serialization
    # (_attendance_student_public) and the same unique (studentId, date)
    # upsert behaviour, so re-marking a student on an existing date
    # always updates that one record instead of creating a duplicate.
    # Trainers are not bound to a single college (they report to Super
    # Admin), so — exactly like Super Admin — no college/department is
    # ever hardcoded here; every filter option is resolved from the
    # database and an empty filter means "no restriction".
    # ==========================================================
    @bp.route("/attendance/filters", methods=["GET"])
    @role_required("trainer")
    def trainer_attendance_filters():
        college_docs = list(db.colleges.find({"status": "active"}).sort("college_name", 1))
        dept_docs = list(db.departments.find({"status": "active"}).sort("department_name", 1))
        trainer_docs = list(users.find(
            {"role": "trainer", "approvalStatus": {"$in": ["approved", "suspended"]}}
        ).sort("fullName", 1))
        return ok({
            "colleges": [{"id": str(c["_id"]), "name": c.get("college_name")} for c in college_docs],
            "departments": [
                {"id": str(d["_id"]), "collegeId": str(d["college_id"]), "name": d.get("department_name")}
                for d in dept_docs
            ],
            "cohorts": [{"value": "all", "label": "All Cohorts"}] + [
                {"value": c, "label": ATTENDANCE_COHORT_LABELS[c]} for c in sorted(VALID_COHORTS)
            ] + [{"value": ENTRY_LEVEL, "label": ATTENDANCE_COHORT_LABELS[ENTRY_LEVEL]}],
            "trainers": [{"id": str(t["_id"]), "name": t.get("fullName") or "—"} for t in trainer_docs],
        })

    @bp.route("/attendance/students", methods=["GET"])
    @role_required("trainer")
    def trainer_list_attendance_students():
        college_ids, err = _parse_optional_id_list(request.args.get("collegeIds"))
        if err:
            return error(err)
        department_ids, err = _parse_optional_id_list(request.args.get("departmentIds"))
        if err:
            return error(err)
        cohort_param = (request.args.get("cohorts") or "").strip()
        cohorts = [c for c in cohort_param.split(",") if c.strip()] if cohort_param else []
        allowed = set(VALID_COHORTS) | {ENTRY_LEVEL, "all"}
        if any(c not in allowed for c in cohorts):
            return error("Invalid cohort filter.")
        query = _attendance_student_query(db, college_ids, department_ids, cohorts)
        students = list(users.find(query).sort("fullName", 1))
        return ok({"students": [_attendance_student_public(d) for d in students], "total": len(students)})

    @bp.route("/attendance/mark", methods=["POST"])
    @role_required("trainer")
    def trainer_mark_attendance():
        data = request.get_json(silent=True) or {}
        date_str = (data.get("date") or "").strip()
        if not _ATTENDANCE_DATE_RE.match(date_str):
            return error("A valid date (YYYY-MM-DD) is required.")

        marks = data.get("marks")
        if not isinstance(marks, list) or not marks:
            return error("No attendance marks provided.")
        if len(marks) > 2000:
            return error("Too many records in one submission.")

        student_ids = []
        status_by_oid = {}
        for m in marks:
            sid = m.get("studentId")
            status = (m.get("status") or "").strip().lower()
            if status not in ATTENDANCE_STATUSES:
                return error("Attendance status must be 'present' or 'absent'.")
            oid = to_object_id(sid)
            if not oid:
                return error("One of the student ids is invalid.")
            if oid not in status_by_oid:
                status_by_oid[oid] = status
                student_ids.append(oid)

        students_by_id = {d["_id"]: d for d in users.find(
            {"_id": {"$in": student_ids}, "role": "student"})}
        missing = [str(oid) for oid in student_ids if oid not in students_by_id]
        if missing:
            return error("One or more selected students no longer exist.")

        actor_id = get_jwt_identity()
        trainer = _trainer_doc() or {}
        actor_name = trainer.get("fullName") or "Trainer"
        now_ts = now()
        inserted = updated = 0
        for oid, status in status_by_oid.items():
            student = students_by_id[oid]
            cohort = student_cohort_label(student)
            payload = {
                "studentId": oid,
                "date": date_str,
                "studentName": student.get("fullName") or "—",
                "rollNumber": student.get("rollNumber") or "—",
                "collegeId": student.get("collegeId"),
                "college": student.get("college") or "—",
                "departmentId": student.get("departmentId"),
                "department": student.get("department") or "—",
                "cohort": cohort,
                "status": status,
                "markedBy": {"id": str(actor_id), "name": actor_name, "role": "trainer"},
                "markedAt": now_ts,
                "updatedAt": now_ts,
            }
            result = db.attendance.update_one(
                {"studentId": oid, "date": date_str},
                {"$set": payload},
                upsert=True,
            )
            if result.upserted_id:
                inserted += 1
            else:
                updated += 1

        present = sum(1 for s in status_by_oid.values() if s == "present")
        absent = len(status_by_oid) - present
        log_activity(
            db, actor_id, "trainer", "attendance_marked",
            f"Marked attendance for {len(status_by_oid)} students on {date_str}",
            meta={"date": date_str, "present": present, "absent": absent},
        )
        return ok({
            "saved": len(status_by_oid),
            "inserted": inserted,
            "updated": updated,
            "present": present,
            "absent": absent,
        }, message="Attendance saved.")

    @bp.route("/attendance/records", methods=["GET"])
    @role_required("trainer")
    def trainer_list_attendance_records():
        college_ids, err = _parse_optional_id_list(request.args.get("collegeIds"))
        if err:
            return error(err)
        department_ids, err = _parse_optional_id_list(request.args.get("departmentIds"))
        if err:
            return error(err)
        trainer_ids, err = _parse_optional_id_list(request.args.get("trainerIds"))
        if err:
            return error(err)
        date_str = (request.args.get("date") or "").strip()
        if not _ATTENDANCE_DATE_RE.match(date_str):
            return error("A valid date (YYYY-MM-DD) is required.")

        student_query = _attendance_student_query(db, college_ids, department_ids, [])
        students = list(users.find(student_query).sort("fullName", 1))
        student_ids = [d["_id"] for d in students]

        record_query = {"date": date_str, "studentId": {"$in": student_ids}}
        if trainer_ids:
            record_query["markedBy.id"] = {"$in": [str(t) for t in trainer_ids]}
        records = list(db.attendance.find(record_query))
        record_by_student = {r["studentId"]: r for r in records}

        rows = []
        for student in students:
            marked = record_by_student.get(student["_id"])
            cohort = student_cohort_label(student)
            rows.append({
                "studentId": str(student["_id"]),
                "studentName": student.get("fullName") or "—",
                "rollNumber": student.get("rollNumber") or "—",
                "college": student.get("college") or "—",
                "department": student.get("department") or "—",
                "cohort": cohort,
                "cohortLabel": ATTENDANCE_COHORT_LABELS.get(cohort, cohort),
                "status": marked["status"] if marked else "not_marked",
                "markedBy": marked["markedBy"] if marked else None,
                "markedAt": marked["markedAt"] if marked else None,
            })

        present = sum(1 for r in rows if r["status"] == "present")
        absent = sum(1 for r in rows if r["status"] == "absent")
        not_marked = sum(1 for r in rows if r["status"] == "not_marked")
        return ok({
            "records": rows,
            "date": date_str,
            "present": present,
            "absent": absent,
            "notMarked": not_marked,
            "total": len(rows),
        })

    # ==========================================================
    # TRAINER — REPORTS CENTER.
    # Platform-wide mirror of Super Admin's Report Center
    # (superadmin.py §"REPORT EXPORTS"). Same reports, same shared
    # reporting.py helpers, same PDF/Excel generation. Every filter
    # dropdown is DB-driven (all colleges, departments, students,
    # trainers, sessions) and the College Report supports one or many
    # colleges in a single file.
    # ==========================================================
    _fmt = reporting.fmt_value

    def _trainer_college_doc():
        """The db.colleges document matching this trainer's assigned
        college name, or None when no college is assigned / found."""
        name = _trainer_college()
        if not name:
            return None
        return db.colleges.find_one({"college_name": name})

    def _trainer_overall_report_rows(student_oid=None, quiz_ids=None, department=None):
        """Same shared row builder Super Admin uses — platform-wide
        (reporting.overall_report_rows), matching the trainers' access to
        the whole roster."""
        return reporting.overall_report_rows(
            db, users, student_oid=student_oid, quiz_ids=quiz_ids,
            college=None, department=department,
        )

    def _college_report_payload(college_doc):
        """Shared builder for the College Report — same real-data
        implementation as Super Admin's: the college's own record, its
        department list, and its students with live quiz/interview/overall
        averages computed from every submitted db.quiz_attempts record.
        Students with no submitted attempts show a dash (—) for each score
        instead of a fabricated 0."""
        oid = college_doc["_id"]
        college_name = college_doc.get("college_name") or "College"
        departments = list(db.departments.find(
            {"college_id": oid, "status": "active"}).sort("department_name", 1))
        dept_names = [d.get("department_name") for d in departments if d.get("department_name")]
        dept_line = f"Departments ({len(departments)}):  {', '.join(dept_names) or '—'}"
        college_columns = ["Student", "Register No", "Department",
                           "Avg Quiz %", "Avg Interview %", "Avg Overall %", "Cohort"]

        stats = {}
        for a in db.quiz_attempts.find(
            {"status": "submitted", "college": college_name, "studentId": {"$exists": True}},
            {"studentId": 1, "overall.percentage": 1, "interviewMarks": 1, "finalAverage": 1},
        ):
            sid = a.get("studentId")
            if sid is None:
                continue
            bucket = stats.setdefault(sid, {"quiz": [], "interview": [], "overall": []})
            pct = (a.get("overall") or {}).get("percentage")
            if isinstance(pct, (int, float)) and not isinstance(pct, bool):
                bucket["quiz"].append(float(pct))
            interview = a.get("interviewMarks")
            if isinstance(interview, (int, float)) and not isinstance(interview, bool):
                bucket["interview"].append(float(interview))
            final_average = a.get("finalAverage")
            if isinstance(final_average, (int, float)) and not isinstance(final_average, bool):
                bucket["overall"].append(float(final_average))

        def _avg(values):
            return round(sum(values) / len(values), 2) if values else None

        students = list(users.find(
            {"role": "student", "college": college_name, "isDeleted": {"$ne": True}},
            {"fullName": 1, "rollNumber": 1, "tneaCode": 1, "department": 1, "cohort": 1},
        ).sort("fullName", 1))
        student_rows = []
        for s in students:
            bucket = stats.get(s["_id"], {})
            student_rows.append([
                s.get("fullName") or "—",
                s.get("rollNumber") or s.get("tneaCode") or "—",
                s.get("department") or "—",
                _fmt(_avg(bucket.get("quiz"))) if bucket.get("quiz") else "—",
                _fmt(_avg(bucket.get("interview"))) if bucket.get("interview") else "—",
                _fmt(_avg(bucket.get("overall"))) if bucket.get("overall") else "—",
                student_cohort_label(s),
            ])

        subtitle = (f"{college_name} · {len(departments)} department(s) · "
                    f"{len(students)} student(s) · generated {fmt_ist(now())}")
        return {
            "name": college_name,
            "filename": _safe_filename(college_name),
            "dept_line": dept_line,
            "num_departments": len(departments),
            "num_students": len(students),
            "columns": college_columns,
            "rows": student_rows,
            "subtitle": subtitle,
        }

    def _student_report_payload(student):
        """Shared builder for the Individual Student Report — same
        live-data implementation as Super Admin's. Returns the summary
        lines and the Performance / Interventions data sections."""
        oid = student["_id"]
        name = student.get("fullName") or student.get("name") or "Student"
        reg = student.get("rollNumber") or student.get("tneaCode") or "—"
        college = student.get("college") or "—"
        department = student.get("department") or "—"
        cohort = student_cohort_label(student)
        att = attendance_summary(db, oid)
        att_pct = att.get("percentage") if att and isinstance(att, dict) else None

        def pct_of(v):
            return _fmt(v) if v is not None else "—"

        perf_columns = ["Assessment", "Quiz %", "Interview %", "Average %", "Final Overall %", "Submitted"]
        perf_data = [[r["assessmentName"] or "—", pct_of(r["quizMarks"]), pct_of(r["interviewMarks"]),
                      pct_of(r["averageMarks"]), pct_of(r["finalOverallMarks"]),
                      (r["submittedAt"] or "")[:16].replace("T", " ")]
                     for r in _trainer_overall_report_rows(student_oid=oid)]

        sess_query = {"collegeNames": {"$in": [college]}} if college and college != "—" else {}
        session_columns = ["Session", "Date", "Time", "Trainer", "Attendance Req", "Status", "Student Status"]
        session_data = []
        for sc in db.workshop_sessions.find(sess_query).sort("date", -1).limit(10):
            date_str = sc.get("date")
            marked = None
            if date_str:
                marked = db.attendance.find_one({
                    "studentId": oid,
                    "date": date_str,
                    "status": {"$in": ["present", "absent"]},
                })
            if marked:
                st = marked.get("status")
                student_status = "Present" if st == "present" else ("Absent" if st == "absent" else "—")
            else:
                student_status = "Attendance Not Marked"
            session_data.append([
                sc.get("name") or "—",
                date_str or "—",
                f"{sc.get('startTime','')} – {sc.get('endTime','')}" if sc.get("startTime") else "—",
                ", ".join(sc.get("trainerNames", [])) or "—",
                (sc.get("attendanceRequirement") or "—").title(),
                (sc.get("status") or "—").title(),
                student_status,
            ])

        summary = [
            f"Student: {name}  ({reg})",
            f"College: {college} · Department: {department} · Cohort: {cohort}",
            (f"Overall Employability: {pct_of(student.get('finalEmployabilityScore'))}% · "
             f"Interview: {pct_of(student.get('interviewScore'))}% · Attendance: {_fmt(att_pct)}%"),
        ]
        return {
            "name": name,
            "filename": _safe_filename(name),
            "summary": summary,
            "subtitle": " · ".join(summary),
            "sections": [
                ("Assessment Performance", perf_columns, perf_data),
                ("Interventions", session_columns, session_data),
            ],
            "excel_sections": [
                ("Performance", perf_columns, perf_data),
                ("Interventions", session_columns, session_data),
            ],
        }

    @bp.route("/reports/filters", methods=["GET"])
    @role_required("trainer")
    def reports_filters():
        """Picker options for the Report Center export modals —
        platform-wide: every active college, its departments, all
        students, workshop sessions and trainers."""
        statuses = [{"id": s, "label": s.title()} for s in SESSION_STATUS_OPTIONS]

        colleges = [
            {"id": str(c["_id"]), "name": c.get("college_name") or "—"}
            for c in db.colleges.find({"status": "active"}).sort("college_name", 1)
        ]
        departments = [
            {"id": str(d["_id"]), "name": d.get("department_name") or "—",
             "collegeId": str(d.get("college_id")) if d.get("college_id") else None}
            for d in db.departments.find({"status": "active"}).sort("department_name", 1)
        ]

        sessions = []
        trainers = []
        seen_trainers = set()
        for sc in db.workshop_sessions.find(
            {}, {"name": 1, "date": 1, "trainerNames": 1}).sort("date", -1).limit(500):
            date_str = sc.get("date") or ""
            sessions.append({
                "id": str(sc["_id"]),
                "label": (sc.get("name") or "—") + (f" · {date_str}" if date_str else ""),
            })
            for t in (sc.get("trainerNames") or []):
                if t and t not in seen_trainers:
                    seen_trainers.add(t)
                    trainers.append({"id": t, "label": t})

        if not trainers:
            trainers = [{"id": u.get("fullName") or str(u["_id"]),
                         "label": u.get("fullName") or "—"}
                        for u in users.find(
                            {"role": "trainer", "approvalStatus": {"$in": ["approved", "suspended"]}},
                            {"fullName": 1}).sort("fullName", 1)]

        students = [{
            "id": str(s["_id"]),
            "label": f"{s.get('fullName') or '—'} · {s.get('rollNumber') or s.get('tneaCode') or '—'}",
        } for s in users.find({"role": "student", "isDeleted": {"$ne": True}},
                              {"fullName": 1, "rollNumber": 1, "tneaCode": 1}).sort("fullName", 1)]

        return ok({
            "colleges": colleges,
            "departments": departments,
            "students": students,
            "sessions": sessions,
            "trainers": trainers,
            "statuses": statuses,
        })

    @bp.route("/reports/overall", methods=["GET"])
    @role_required("trainer")
    def reports_overall():
        """The Overall Report for the trainer — every submitted quiz
        attempt platform-wide, or for one student when studentId is passed."""
        student_id = (request.args.get("studentId") or "").strip()
        oid = to_object_id(student_id) if student_id else None
        rows = _trainer_overall_report_rows(student_oid=oid)
        return ok({"report": rows, "college": _trainer_college(), "studentId": student_id})

    @bp.route("/reports/overall/export", methods=["GET"])
    @role_required("trainer")
    def reports_overall_export():
        """Overall Platform Report export (platform-wide). Mirrors the
        Overall Report screen filters: optional single department plus a
        free-text q search across student/roll/college/department/assessment."""
        fmt = (request.args.get("format") or "pdf").lower()
        department = (request.args.get("department") or "").strip()
        search = (request.args.get("q") or "").strip().lower()
        rows = _trainer_overall_report_rows(
            department=department if department and department != "all" else None)
        if search:
            def _matches(r):
                hay = " ".join(str(r.get(k) or "") for k in
                                ("studentName", "rollNumber", "college", "department", "assessmentName")).lower()
                return search in hay
            rows = [r for r in rows if _matches(r)]
        columns = ["Student", "Register No", "Department", "College", "Assessment",
                   "Quiz %", "Interview %", "Average %", "Final Overall %", "Submitted"]
        data = [[r["studentName"] or "—", r["rollNumber"] or "—", r["department"] or "—",
                 r["college"] or "—", r["assessmentName"] or "—", _fmt(r["quizMarks"]),
                 _fmt(r["interviewMarks"]), _fmt(r["averageMarks"]), _fmt(r["finalOverallMarks"]),
                 (r["submittedAt"] or "")[:16].replace("T", " ")] for r in rows]
        filter_bits = []
        if department and department != "all":
            filter_bits.append(f"department: {department}")
        if search:
            filter_bits.append(f"search: \"{search}\"")
        filter_note = f" · filtered by {', '.join(filter_bits)}" if filter_bits else ""
        subtitle = (f"{len(rows)} submitted attempt(s){filter_note} · generated {fmt_ist(now())}")
        if fmt == "excel":
            return send_file(
                excel_bytes(columns, data, "Overall Report"),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True, download_name="overall_platform_report.xlsx",
            )
        return send_file(
            pdf_bytes("Overall Platform Report", subtitle, columns, data),
            mimetype="application/pdf", as_attachment=True,
            download_name="overall_platform_report.pdf",
        )

    @bp.route("/reports/college/export", methods=["GET"])
    @role_required("trainer")
    def reports_college_export():
        """College Report (PDF/Excel) — same report the Super Admin
        generates for any college. Supports a `colleges` multi-select so a
        single file can combine one or many colleges; falls back to the
        trainer's assigned college, then all active colleges."""
        fmt = (request.args.get("format") or "pdf").lower()

        college_oids, err = _parse_optional_id_list(request.args.get("colleges"))
        if err:
            return error(err)
        if college_oids:
            college_docs = list(db.colleges.find({"_id": {"$in": college_oids}, "status": "active"}))
            if not college_docs:
                return error("Selected college(s) not found.")
        else:
            assigned = _trainer_college_doc()
            if assigned:
                college_docs = [assigned]
            else:
                college_docs = list(db.colleges.find({"status": "active"}).sort("college_name", 1))
        if not college_docs:
            return error("No active colleges found.")

        payloads = [_college_report_payload(cd) for cd in college_docs]

        if fmt == "excel":
            wb = Workbook()
            ws = wb.active
            ws.title = "Overview"
            ws.append(["College Report", len(payloads)])
            ws.append(["Generated", fmt_ist(now())])
            ws.append([])
            ws.append(["College", "Departments", "Students"])
            for p in payloads:
                ws.append([p["name"], p["num_departments"], p["num_students"]])
            for c in range(1, 4):
                cell = ws.cell(row=4, column=c)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F9E93")
            for p in payloads:
                sheet = wb.create_sheet(_safe_filename(p["name"])[:31])
                sheet.append(["College Report", p["name"]])
                sheet.append(["Departments", p["dept_line"]])
                sheet.append(["Students", p["num_students"]])
                sheet.append([])
                sheet.append(p["columns"])
                for r in p["rows"]:
                    sheet.append(r)
                for c in range(1, len(p["columns"]) + 1):
                    cell = sheet.cell(row=5, column=c)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1F9E93")
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return send_file(
                buf,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="college_report.xlsx",
            )

        if len(payloads) == 1:
            p = payloads[0]
            return send_file(
                pdf_bytes(f"College Report — {p['name']}", p["subtitle"], p["columns"], p["rows"]),
                mimetype="application/pdf", as_attachment=True,
                download_name=f"college_report_{p['filename']}.pdf",
            )

        sections = [
            (p["name"], [p["dept_line"]], [[f"{p['num_students']} student(s)"]])
            for p in payloads
        ]
        subtitle = (f"{len(payloads)} college(s) · {sum(p['num_students'] for p in payloads)} student(s)"
                    f" · generated {fmt_ist(now())}")
        return send_file(
            reporting.pdf_sections("College Report", subtitle, sections),
            mimetype="application/pdf", as_attachment=True,
            download_name="college_report.pdf",
        )

    @bp.route("/reports/students/export", methods=["GET"])
    @role_required("trainer")
    def reports_students_export():
        """Multi-student report — the Individual Student Report for a list
        of students picked in the Report Center (optionally scoped by
        department ids), across the whole platform. One summary block per
        student in the PDF; one sheet per student in Excel."""
        fmt = (request.args.get("format") or "pdf").lower()
        student_oids, err = _parse_optional_id_list(request.args.get("students"))
        if err:
            return error(err)
        department_oids, err = _parse_optional_id_list(request.args.get("departments"))
        if err:
            return error(err)

        query = {"role": "student", "isDeleted": {"$ne": True}}
        if student_oids:
            query["_id"] = {"$in": student_oids}
        if department_oids:
            query["departmentId"] = {"$in": department_oids}
        students = list(users.find(query).sort("fullName", 1))
        if not students:
            return error("No students match the selected filters.")

        payloads = [_student_report_payload(s) for s in students]
        total_sessions = sum(len(p["sections"][1][2]) for p in payloads)
        subtitle = (f"{len(students)} student report(s) · {total_sessions} intervention record(s)"
                    f" · generated {fmt_ist(now())}")

        if fmt == "excel":
            wb = Workbook()
            ws = wb.active
            ws.title = "Overview"
            ws.append(["Student Reports", len(students)])
            ws.append(["Generated", fmt_ist(now())])
            ws.append([])
            ws.append(["Student", "Register No", "College", "Department", "Cohort"])
            for p, s in zip(payloads, students):
                ws.append([
                    p["name"], s.get("rollNumber") or s.get("tneaCode") or "—",
                    s.get("college") or "—", s.get("department") or "—",
                    student_cohort_label(s),
                ])
            for p in payloads:
                sheet = wb.create_sheet(_safe_filename(p["name"])[:31])
                for line in p["summary"]:
                    sheet.append([line])
                sheet.append([])
                for sheet_name, cols, data in p["excel_sections"]:
                    sheet.append([f"— {sheet_name} —"])
                    sheet.append(cols)
                    for c in range(1, len(cols) + 1):
                        cell = sheet.cell(row=sheet.max_row - len(data), column=c)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill("solid", fgColor="1F9E93")
                    for r in data:
                        sheet.append(r)
                    sheet.append([])
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return send_file(
                buf,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True, download_name="students_report.xlsx",
            )

        sections = []
        for p in payloads:
            sections.append((p["name"], ["Detail"], [[line] for line in p["summary"]]))
            sections.extend(p["sections"])
        return send_file(
            reporting.pdf_sections("Individual Student Report", subtitle, sections),
            mimetype="application/pdf", as_attachment=True,
            download_name="students_report.pdf",
        )

    @bp.route("/reports/assessments/export", methods=["GET"])
    @role_required("trainer")
    def reports_assessments_export():
        """Individual Assessment Report — one or more assessments selected
        by the trainer, downloaded as PDF/Excel, across the whole platform.
        Same structure as Super Admin's: an Assessment Performance summary
        section plus, for each selected assessment, every student who
        submitted it."""
        fmt = (request.args.get("format") or "pdf").lower()
        raw_ids = [x.strip() for x in (request.args.get("assessments") or "").split(",") if x.strip()]
        quiz_ids = []
        for qid in raw_ids:
            oid = to_object_id(qid)
            if oid:
                quiz_ids.append(oid)
        if not quiz_ids:
            return error("Select at least one assessment to export.")
        names = {str(d["_id"]): (d.get("title") or "Quiz")
                 for d in db.quizzes.find({"_id": {"$in": quiz_ids}}, {"title": 1})}

        rows = _trainer_overall_report_rows(quiz_ids=quiz_ids)
        str_ids = [str(qid) for qid in quiz_ids]
        by_assessment = {sid: [] for sid in str_ids}
        for r in rows:
            aid = r.get("assessmentId")
            if aid in by_assessment:
                by_assessment[aid].append(r)

        ordered = [qid for qid in quiz_ids if by_assessment[str(qid)]]

        def marks_of(r):
            got = r.get("quizMarksObtained")
            tot = r.get("quizTotalMarks")
            if got is None or tot is None:
                return "—"
            return f"{got}/{tot}"

        def result_of(r):
            pf = r.get("passFail")
            if pf in ("Pass", "Fail"):
                return pf
            pct = r.get("quizMarks")
            if pct is None:
                return "—"
            return "Pass" if float(pct) >= 40 else "Fail"

        def _avg_of(vals):
            vals = [float(v) for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
            return round(sum(vals) / len(vals), 2) if vals else None

        perf_columns = ["Assessment", "Avg Quiz %", "Avg Interview %",
                         "Avg Current Assessment %", "Avg Overall Average %", "Cohort(s)", "Submitted"]
        perf_rows = []
        for qid in quiz_ids:
            recs = by_assessment.get(str(qid), [])
            cohorts_seen = sorted({r["cohort"] for r in recs if r.get("cohort")})
            perf_rows.append([
                names.get(str(qid), "Assessment"),
                _fmt(_avg_of([r["quizMarks"] for r in recs])) if recs else "—",
                _fmt(_avg_of([r["interviewMarks"] for r in recs])) if recs else "—",
                _fmt(_avg_of([r["averageMarks"] for r in recs])) if recs else "—",
                _fmt(_avg_of([r["finalOverallMarks"] for r in recs])) if recs else "—",
                ", ".join(cohorts_seen) or "—",
                len(recs),
            ])

        columns = ["Assessment", "Student", "Register No", "Department", "College",
                   "Quiz %", "Marks", "Result", "Interview %", "Average %", "Final Overall %", "Submitted"]
        def row_of(r):
            return [names.get(r.get("assessmentId"), r["assessmentName"] or "—"),
                    r["studentName"] or "—", r["rollNumber"] or "—", r["department"] or "—",
                    r["college"] or "—", _fmt(r["quizMarks"]), marks_of(r), result_of(r),
                    _fmt(r["interviewMarks"]), _fmt(r["averageMarks"]), _fmt(r["finalOverallMarks"]),
                    (r["submittedAt"] or "")[:16].replace("T", " ")]
        data = [row_of(r) for r in rows]

        subtitle = (f"{len(rows)} submitted attempt(s) across {len(ordered)} selected assessment(s)"
                    f" · generated {fmt_ist(now())}")
        if fmt == "excel":
            wb = Workbook()
            perf_ws = wb.active
            perf_ws.title = "Assessment Performance"
            perf_ws.append(perf_columns)
            for c in range(1, len(perf_columns) + 1):
                cell = perf_ws.cell(row=1, column=c)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F9E93")
            for r in perf_rows:
                perf_ws.append(r)
            ws = wb.create_sheet("All Assessments")
            ws.append(columns)
            for c in range(1, len(columns) + 1):
                cell = ws.cell(row=1, column=c)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F9E93")
            for r in data:
                ws.append(r)
            for qid in ordered:
                name = names.get(str(qid), "Assessment")
                sheet = wb.create_sheet(_safe_filename(name)[:31])
                per_columns = columns[1:]
                sheet.append(per_columns)
                for c in range(1, len(per_columns) + 1):
                    cell = sheet.cell(row=1, column=c)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1F9E93")
                for r in by_assessment[str(qid)]:
                    sheet.append([r["studentName"] or "—", r["rollNumber"] or "—",
                                  r["department"] or "—", r["college"] or "—",
                                  _fmt(r["quizMarks"]), marks_of(r), result_of(r),
                                  _fmt(r["interviewMarks"]), _fmt(r["averageMarks"]),
                                  _fmt(r["finalOverallMarks"]),
                                  (r["submittedAt"] or "")[:16].replace("T", " ")])
            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            return send_file(
                buf,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="assessment_report.xlsx",
            )
        sections = [("Assessment Performance", perf_columns, perf_rows)]
        sections += [(names.get(str(qid), "Assessment"), columns[1:],
                     [[r["studentName"] or "—", r["rollNumber"] or "—", r["department"] or "—",
                       r["college"] or "—", _fmt(r["quizMarks"]), marks_of(r), result_of(r),
                       _fmt(r["interviewMarks"]), _fmt(r["averageMarks"]),
                       _fmt(r["finalOverallMarks"]),
                       (r["submittedAt"] or "")[:16].replace("T", " ")]
                      for r in by_assessment[str(qid)]])
                    for qid in ordered]
        if not ordered:
            return error("No submitted attempts found for the selected assessment(s).")
        return send_file(
            reporting.pdf_sections("Individual Assessment Report", subtitle, sections),
            mimetype="application/pdf", as_attachment=True,
            download_name="assessment_report.pdf",
        )

    @bp.route("/reports/interventions/export", methods=["GET"])
    @role_required("trainer")
    def reports_interventions_export():
        """Intervention Report — every workshop session on the platform,
        with its target colleges/departments/trainers, cohort, schedule,
        venue, attendance requirement, a single Status column (folding
        workflow + approval state via _intervention_status_label) and real
        attendance participation."""
        fmt = (request.args.get("format") or "pdf").lower()
        statuses = [s for s in (request.args.get("statuses") or "").split(",") if s.strip() in SESSION_STATUS_OPTIONS]
        department_oids, err = _parse_optional_id_list(request.args.get("departments"))
        if err:
            return error(err)
        trainer_oids, err = _parse_optional_id_list(request.args.get("trainers"))
        if err:
            return error(err)
        session_oids, err = _parse_optional_id_list(request.args.get("sessions"))
        if err:
            return error(err)

        _auto_complete_workshop_sessions(db)

        query = {}
        if statuses:
            query["status"] = {"$in": statuses}
        if session_oids:
            query["_id"] = {"$in": session_oids}
        docs = list(db.workshop_sessions.find(query).sort("date", -1))
        if department_oids:
            docs = [d for d in docs if d.get("departmentIds") and any(
                d_ in department_oids for d_ in d["departmentIds"])]
        if trainer_oids:
            docs = [d for d in docs if d.get("trainerIds") and any(
                t in trainer_oids for t in d["trainerIds"])]

        columns = ["Session", "Colleges", "Departments", "Trainers", "Cohort", "Date",
                   "Time", "Venue", "Attendance Req", "Status",
                   "Marked", "Present", "Absent", "Created"]
        data = []
        for d in docs:
            cnt = reporting.attendance_counts_for_date(db, d.get("date"))
            data.append([
                d.get("name") or "—",
                ", ".join(d.get("collegeNames", [])) or "—",
                ", ".join(d.get("departmentNames", [])) or "—",
                ", ".join(d.get("trainerNames", [])) or "—",
                d.get("cohort") or "—",
                d.get("date") or "—",
                f"{d.get('startTime','')} – {d.get('endTime','')}" if d.get("startTime") else "—",
                d.get("venue") or "—",
                (d.get("attendanceRequirement") or "—").title(),
                _intervention_status_label(d),
                cnt["total"], cnt["present"], cnt["absent"],
                iso_utc(d.get("createdAt"))[:16].replace("T", " "),
            ])
        subtitle = (f"{len(docs)} intervention session(s)"
                    f"{(' · status: ' + ', '.join(statuses)) if statuses else ''}"
                    f" · generated {fmt_ist(now())}")
        if fmt == "excel":
            return send_file(
                excel_bytes(columns, data, "Intervention Report"),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="intervention_report.xlsx",
            )
        return send_file(
            pdf_bytes("Intervention Report", subtitle, columns, data),
            mimetype="application/pdf", as_attachment=True,
            download_name="intervention_report.pdf",
        )

    @bp.route("/reports/performance/export", methods=["GET"])
    @role_required("trainer")
    def reports_performance_export():
        """Student Performance Report — a per-student roster (whole
        platform) with Assessment Quiz Score, Interview %, overall
        employability, cohort and attendance — computed live from db.users
        + db.quiz_attempts + db.attendance. Optional department / student /
        cohort filters."""
        fmt = (request.args.get("format") or "pdf").lower()
        department_oids, err = _parse_optional_id_list(request.args.get("departments"))
        if err:
            return error(err)
        student_oids, err = _parse_optional_id_list(request.args.get("students"))
        if err:
            return error(err)
        cohort_param = (request.args.get("cohorts") or "").strip()
        cohorts = [c for c in cohort_param.split(",") if c.strip()] if cohort_param else []

        query = {"role": "student", "isDeleted": {"$ne": True}}
        if department_oids:
            query["departmentId"] = {"$in": department_oids}
        if student_oids:
            query["_id"] = {"$in": student_oids}
        if cohorts:
            allowed = set(VALID_COHORTS) | {ENTRY_LEVEL}
            if any(c not in allowed for c in cohorts):
                return error("Invalid cohort filter.")
            query["$or"] = []
            if ENTRY_LEVEL in cohorts:
                query["$or"].append({"cohort": None})
                query["$or"].append({"cohort": {"$exists": False}})
            for c in cohorts:
                if c in VALID_COHORTS:
                    query["$or"].append({"cohort": c})

        docs = list(users.find(
            query,
            {"fullName": 1, "rollNumber": 1, "tneaCode": 1, "college": 1,
             "department": 1, "cohort": 1, "finalEmployabilityScore": 1},
        ).sort("fullName", 1))

        oids = [d["_id"] for d in docs]
        att_stats = {}
        if oids:
            for row in db.attendance.aggregate([
                {"$match": {"studentId": {"$in": oids}}},
                {"$group": {
                    "_id": "$studentId",
                    "present": {"$sum": {"$cond": [{"$eq": ["$status", "present"]}, 1, 0]}},
                    "absent": {"$sum": {"$cond": [{"$eq": ["$status", "absent"]}, 1, 0]}},
                }},
            ]):
                att_stats[row["_id"]] = row

        quiz_stats = {}
        if oids:
            for row in db.quiz_attempts.aggregate([
                {"$match": {"status": "submitted", "studentId": {"$in": oids}}},
                {"$group": {
                    "_id": "$studentId",
                    "avgQuiz": {"$avg": "$overall.percentage"},
                    "avgInterview": {"$avg": "$interviewMarks"},
                }},
            ]):
                quiz_stats[row["_id"]] = row

        columns = ["Student", "Register No", "College", "Department", "Cohort",
                   "Assessment Quiz Score %", "Interview %", "Overall Employability %",
                   "Attendance %", "Assessment Status"]
        data = []
        for s in docs:
            a = att_stats.get(s["_id"], {})
            present = a.get("present", 0)
            absent = a.get("absent", 0)
            total = present + absent
            att_pct = round(present / total * 100, 1) if total else 0
            qs = quiz_stats.get(s["_id"], {})
            quiz_score = qs.get("avgQuiz")
            interview = qs.get("avgInterview")
            status = "Completed" if quiz_score is not None and interview is not None else "Pending"
            cohort_label = student_cohort_label(s)
            data.append([
                s.get("fullName") or "—",
                s.get("rollNumber") or s.get("tneaCode") or "—",
                s.get("college") or "—",
                s.get("department") or "—",
                f"Cohort {cohort_label}" if cohort_label in ("A", "B", "C") else "Entry Level",
                _fmt(round(quiz_score, 2)) if quiz_score is not None else "—",
                _fmt(round(interview, 2)) if interview is not None else "—",
                _fmt(s.get("finalEmployabilityScore")) if s.get("finalEmployabilityScore") is not None else "—",
                att_pct, status,
            ])
        subtitle = (f"{len(docs)} student(s)"
                    f"{(' · cohort: ' + ', '.join(cohorts)) if cohorts else ''}"
                    f" · generated {fmt_ist(now())}")
        if fmt == "excel":
            return send_file(
                excel_bytes(columns, data, "Student Performance"),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="student_performance_report.xlsx",
            )
        return send_file(
            pdf_bytes("Student Performance Report", subtitle, columns, data),
            mimetype="application/pdf", as_attachment=True,
            download_name="student_performance_report.pdf",
        )

    @bp.route("/reports/assessment-report/export", methods=["GET"])
    @role_required("trainer")
    def reports_assessment_report_export():
        """Assessment Report card — Department -> Student multi-select,
        platform-wide. PDF: one page per student (reporting.pdf_per_student).
        Excel: one flat table of every selected student's every assessment
        record."""
        fmt = (request.args.get("format") or "pdf").lower()
        department_oids, err = _parse_optional_id_list(request.args.get("departments"))
        if err:
            return error(err)
        student_oids, err = _parse_optional_id_list(request.args.get("students"))
        if err:
            return error(err)

        query = {"role": "student", "isDeleted": {"$ne": True}}
        if department_oids:
            query["departmentId"] = {"$in": department_oids}
        if student_oids:
            query["_id"] = {"$in": student_oids}
        if not (department_oids or student_oids):
            return error("Select at least one department or student to export.")

        students = list(users.find(query).sort("fullName", 1))
        if not students:
            return error("No students match the selected filters.")

        columns = ["Assessment", "Date", "Quiz %", "Interview %", "Overall %", "Cohort", "Status"]

        if fmt == "excel":
            excel_columns = ["Student", "Register No", "College", "Department"] + columns
            excel_rows = []
            for s in students:
                base = [s.get("fullName") or "—", s.get("rollNumber") or s.get("tneaCode") or "—",
                        s.get("college") or "—", s.get("department") or "—"]
                recs = reporting.assessment_report_rows_for_student(db, s["_id"])
                if not recs:
                    excel_rows.append(base + ["—", "—", "—", "—", "—", "—", "No records"])
                    continue
                for r in recs:
                    excel_rows.append(base + [
                        r["assessmentName"], r["date"], _fmt(r["quizMarks"]) if r["quizMarks"] is not None else "—",
                        _fmt(r["interviewMarks"]) if r["interviewMarks"] is not None else "—",
                        _fmt(r["overallMarks"]) if r["overallMarks"] is not None else "—",
                        r["cohort"], r["status"],
                    ])
            return send_file(
                excel_bytes(excel_columns, excel_rows, "Assessment Report"),
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name="assessment_report_card.xlsx",
            )

        pages = []
        for s in students:
            recs = reporting.assessment_report_rows_for_student(db, s["_id"])
            rows = [[r["assessmentName"], r["date"],
                     _fmt(r["quizMarks"]) if r["quizMarks"] is not None else "—",
                     _fmt(r["interviewMarks"]) if r["interviewMarks"] is not None else "—",
                     _fmt(r["overallMarks"]) if r["overallMarks"] is not None else "—",
                     r["cohort"], r["status"]] for r in recs]
            name = s.get("fullName") or s.get("name") or "Student"
            reg = s.get("rollNumber") or s.get("tneaCode") or "—"
            pages.append({
                "heading": f"{name} ({reg})",
                "meta": [f"College: {s.get('college') or '—'}  ·  "
                         f"Department: {s.get('department') or '—'}  ·  "
                         f"Cohort: {student_cohort_label(s)}"],
                "columns": columns,
                "rows": rows,
            })
        subtitle = f"{len(students)} student(s) · generated {fmt_ist(now())}"
        return send_file(
            reporting.pdf_per_student("Assessment Report", subtitle, pages),
            mimetype="application/pdf", as_attachment=True,
            download_name="assessment_report_card.pdf",
        )

    return bp
