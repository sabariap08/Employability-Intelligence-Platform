"""
============================================================
 student.py — Student Dashboard backend
============================================================
Registered at /api/student. Implements the full quiz flow:

  Login -> fetch available assessments -> start assessment ->
  random questions served -> submit answers -> store responses ->
  calculate score -> store result -> display result

Also implements:
  - My Assessments (Upcoming / Completed / Attempted / Pending),
    computed from schedule + cohort (incl. Entry Level) + availability.
  - Assessment Results page: Baseline / Mid / Final + the
    "Assessment Name" dropdown of every assessment the student has
    taken, all from the database.
  - Dynamic graphs (skill radar, category %, overall, trend) for
    this one student.
  - Download Report — generated from real DB values.
"""

from datetime import datetime, timezone, timedelta
from io import BytesIO

from flask import Blueprint, request, send_file
from flask_jwt_extended import get_jwt_identity

from quiz_common import (
    ok, error, role_required, now, to_object_id, serialize,
    select_random_questions_v2, strip_answers, calculate_score,
    record_assessment_score_for_cohort, student_cohort_label,
    cohort_query_for_target, student_matches_cohort_target,
    log_activity, init_quiz_result_fields, serialize_quiz_result,
    list_quiz_results, compute_overall_performance, top_cohort_label,
    iso_utc, attendance_summary,
)

# Manually-authored quizzes (Trainer/Super Admin "Create Quiz" wizard,
# db.quizzes) are a deliberately separate feature from the cohort/placement
# assessment engine above (db.assessments) — see the header comment in
# quiz_module.py. select_random_questions/compute_status here are that
# module's versions (different signature/shape than quiz_common's), so
# they're imported under distinct names to avoid any confusion with the
# assessment-engine helpers already imported above.
from quiz_module import (
    select_random_questions as quiz_draw_questions,
    compute_status as quiz_compute_status,
    normalize_quiz_college_names,
)


def _parse_dt(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def init_student(db):
    bp = Blueprint("student", __name__)

    users = db.users
    questions = db.questions
    assessments = db.assessments
    attempts = db.assessment_attempts

    # Manually-authored Quizzes (separate feature — see import comment above)
    quizzes = db.quizzes
    quiz_attempts = db.quiz_attempts

    def _current_student():
        return users.find_one({"_id": to_object_id(get_jwt_identity())})

    def _aware_dt(dt):
        if not dt:
            return None
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)

    def _assessment_eligible(assessment, student):
        """Full eligibility for an assessment: cohort AND college AND
        department (when the assessment scopes to them).

        Root-cause note: assessment docs already store an authoritative
        college/department name (resolved server-side at creation time
        via resolve_active_college/resolve_active_department in
        trainer.py / superadmin.py — never trusted from raw client IDs,
        so there's no id-vs-name mismatch here, unlike the quiz bug).
        The actual gap was that nothing on the student-visibility side
        ever read those fields — only cohortTarget was checked — so an
        assessment scoped to one college was visible/attemptable by
        students at every college, as long as cohort matched. A None/
        empty college on the assessment means "all colleges" (this is
        the documented behavior in superadmin.create_assessment); a
        college with no department means "all departments in that
        college". Department is only meaningful once college is scoped,
        matching the creation form's own nested logic.

        Null-safety (spec: "Validate Existing Database... handle null
        safely"): an assessment created before cohortTarget existed, or
        with it unset for any other reason, is treated as "all" — the
        same default quiz_module's eligibility check already used — so a
        missing field hides the assessment from nobody instead of hiding
        it from every Cohort A/B/C student (only Entry Level students,
        who have no `cohort` value, would previously have been able to
        see it — an accidental near-total-hide rather than a deliberate
        "all cohorts" default).
        """
        if not student_matches_cohort_target(student, assessment.get("cohortTarget") or "all"):
            return False
        assess_college = assessment.get("college")
        if assess_college:
            student_college = str(student.get("college") or "").strip().lower()
            if not student_college or student_college != str(assess_college).strip().lower():
                return False
            assess_department = assessment.get("department")
            if assess_department:
                student_department = str(student.get("department") or "").strip().lower()
                if not student_department or student_department != str(assess_department).strip().lower():
                    return False
        return True

    def _assessment_deadline(assessment, started_at):
        """Backend cutoff for an in-progress attempt. The assessment
        schema has no per-attempt duration field (only an availableTo
        window close), so the deadline is whichever is earlier: the
        existing 30-minute default the exam UI already counts down from,
        or the assessment's configured window close — availableTo is the
        authoritative "cannot continue after end time" cutoff either way."""
        window_close = _aware_dt(assessment.get("availableTo"))
        default_deadline = started_at + timedelta(seconds=30 * 60)
        if window_close and window_close < default_deadline:
            return window_close
        return default_deadline

    def _pass_fail(assessment, overall):
        """Dormant until the Create Assessment wizard defines a passing
        threshold (no passingMarks/passingPercentage field exists there
        today — only sectionCounts). Reads either key defensively so
        Pass/Fail activates automatically, with zero code changes here,
        the moment such a field is added — matching the same pattern
        already used for quizzes. Never invents a hardcoded default
        threshold, per "No hardcoded values."
        """
        passing_pct = assessment.get("passingPercentage")
        if passing_pct is not None:
            return "Pass" if overall["percentage"] >= float(passing_pct) else "Fail"
        passing_marks = assessment.get("passingMarks")
        if passing_marks is not None:
            return "Pass" if overall["correct"] >= float(passing_marks) else "Fail"
        return None

    def _finalize_assessment_attempt(assessment, attempt, answers=None):
        if answers is None:
            answers = attempt.get("answers") or {}
        student = users.find_one({"_id": attempt["studentId"]}) or {}
        scoring = calculate_score(db, attempt["questionIds"], answers)
        submitted_at = now()
        started_at = attempt.get("startedAt") or submitted_at
        time_taken_seconds = int((submitted_at - _aware_dt(started_at)).total_seconds())
        overall = dict(scoring["overall"])
        # Task 7 naming: same values as correct/total, explicit aliases so
        # "Obtained Marks / Total Marks" is stored under those exact names
        # too (each question is worth 1 mark in the current question bank).
        overall["obtainedMarks"] = overall["correct"]
        overall["totalMarks"] = overall["total"]
        overall["passFail"] = _pass_fail(assessment or {}, overall)

        # Full snapshot required by the Results workflow — captured at
        # submission time so a later profile edit (name change, college
        # transfer, etc.) never rewrites history for a result already on
        # record.
        record_fields = {
            "answers": answers,
            "status": "submitted",
            "submittedAt": submitted_at,
            "sectionScores": scoring["sectionScores"],
            "overall": overall,
            "attemptNumber": attempt.get("attemptNumber", 1),
            "timeTakenSeconds": time_taken_seconds,
            "studentRollNumber": student.get("rollNumber"),
            "studentName": student.get("fullName"),
            "college": student.get("college"),
            "department": student.get("department"),
            "cohort": student_cohort_label(student),
            "assessmentName": (assessment or {}).get("name") or (assessment or {}).get("title"),
            "completionStatus": "completed",
        }
        attempts.update_one({"_id": attempt["_id"]}, {"$set": record_fields})
        attempt.update(record_fields)
        record_assessment_score_for_cohort(
            db, attempt["studentId"],
            assessment_type=(assessment or {}).get("type"),
            percentage=overall["percentage"],
        )
        return attempt

    def _auto_expire_assessment_if_needed(assessment, attempt):
        """Server-side enforcement of "cannot continue after end time"
        (Task 9): if an in-progress attempt's deadline has passed —
        whether the student's browser is even open or not — force-submit
        it with whatever answers were saved, the next time anything
        touches this attempt (listing, resume, or submit)."""
        if attempt.get("status") != "in_progress":
            return False
        deadline = _aware_dt(attempt.get("deadline"))
        if not deadline or now() <= deadline:
            return False
        _finalize_assessment_attempt(assessment, attempt)
        return True

    # ==========================================================
    # 1. FETCH AVAILABLE ASSESSMENTS — My Assessments tab
    # ==========================================================
    @bp.route("/assessments", methods=["GET"])
    @role_required("student")
    def my_assessments():
        student = _current_student()
        if not student:
            return error("Student not found.", 404)

        current_time = now()
        result = {"upcoming": [], "completed": [], "attempted": [], "pending": []}

        # Task 8: the student's status must reflect reality on every load,
        # not just whatever it was when the attempt started — so any
        # in-progress attempt whose deadline has already passed gets
        # auto-submitted right here before we bucket anything below.
        stale = list(attempts.find({"studentId": student["_id"], "status": "in_progress"}))
        if stale:
            assessments_by_id = {a["_id"]: a for a in assessments.find(
                {"_id": {"$in": [s["assessmentId"] for s in stale]}}
            )}
            for s in stale:
                a = assessments_by_id.get(s["assessmentId"])
                if a:
                    _auto_expire_assessment_if_needed(a, s)

        completed_ids = {
            str(a["assessmentId"])
            for a in attempts.find({"studentId": student["_id"], "status": "submitted"}, {"assessmentId": 1})
        }
        in_progress_ids = {
            str(a["assessmentId"])
            for a in attempts.find({"studentId": student["_id"], "status": "in_progress"}, {"assessmentId": 1})
        }

        for a in assessments.find({"status": "active"}).sort("createdAt", -1):
            if not _assessment_eligible(a, student):
                continue
            aid = str(a["_id"])
            available_from = _parse_dt(a.get("availableFrom"))
            available_to = _parse_dt(a.get("availableTo"))

            card = serialize(a)

            if aid in completed_ids:
                result["completed"].append(card)
            elif aid in in_progress_ids:
                result["attempted"].append(card)
            elif available_from and available_from > current_time:
                result["upcoming"].append(card)
            elif available_to and available_to < current_time:
                # Window closed and the student never attempted it — still
                # "Completed" per spec (end time passed OR submitted), not
                # hidden, so it doesn't just vanish from the dashboard.
                result["completed"].append(card)
            else:
                result["pending"].append(card)

        return ok(result)

    # ==========================================================
    # 2. START ASSESSMENT — random questions served fresh every time
    # ==========================================================
    @bp.route("/assessments/<assessment_id>/start", methods=["POST"])
    @role_required("student")
    def start_assessment(assessment_id):
        student = _current_student()
        aoid = to_object_id(assessment_id)
        if not aoid:
            return error("Invalid assessment id.", 404)
        assessment = assessments.find_one({"_id": aoid})
        if not assessment:
            return error("Assessment not found.", 404)
        if not _assessment_eligible(assessment, student):
            return error("This assessment is not available to you.", 403)

        existing = attempts.find_one({
            "studentId": student["_id"], "assessmentId": aoid, "status": "in_progress",
        })
        if existing:
            if _auto_expire_assessment_if_needed(assessment, existing):
                return error("Time is up — this attempt was auto-submitted.", 409)
            qdocs = list(questions.find({"_id": {"$in": existing["questionIds"]}}))
            existing_deadline = _aware_dt(existing.get("deadline"))
            return ok({
                "attemptId": str(existing["_id"]),
                "assessment": serialize(assessment),
                "questions": strip_answers(qdocs),
                "remainingSeconds": max(0, int((existing_deadline - now()).total_seconds())) if existing_deadline else None,
            })

        already_submitted = attempts.find_one({
            "studentId": student["_id"], "assessmentId": aoid, "status": "submitted",
        })
        if already_submitted:
            return error("You have already completed this assessment.", 409)

        # Task 9: backend must enforce the scheduled window itself — never
        # trust the frontend to hide the Start button before/after it.
        current_time = now()
        window_from = _aware_dt(_parse_dt(assessment.get("availableFrom")))
        window_to = _aware_dt(_parse_dt(assessment.get("availableTo")))
        if window_from and current_time < window_from:
            return error("This assessment has not started yet.", 409)
        if window_to and current_time > window_to:
            return error("This assessment's window has already closed.", 409)

        selected, shortage = select_random_questions_v2(db, assessment)
        if shortage:
            return error(
                "The question bank does not have enough questions yet for: "
                + ", ".join(f"{key} (short by {n})" for key, n in shortage.items()),
                409,
            )

        started_at = now()
        attempt_doc = {
            "studentId": student["_id"],
            "assessmentId": aoid,
            "questionIds": [q["_id"] for q in selected],
            "answers": {},
            "status": "in_progress",
            "attemptNumber": 1,
            "startedAt": started_at,
            "deadline": _assessment_deadline(assessment, started_at),
            "submittedAt": None,
        }
        result = attempts.insert_one(attempt_doc)
        log_activity(
            db, student["_id"], "student", "assessment_started",
            f'Started assessment "{assessment.get("name") or assessment.get("title") or ""}"',
            student_id=student["_id"], meta={"assessmentId": str(aoid)},
        )

        return ok({
            "attemptId": str(result.inserted_id),
            "assessment": serialize(assessment),
            "questions": strip_answers(selected),
            "remainingSeconds": int((attempt_doc["deadline"] - started_at).total_seconds()),
        }, status=201)

    # ==========================================================
    # 3. SUBMIT ANSWERS — store responses, calculate + store score
    # ==========================================================
    @bp.route("/attempts/<attempt_id>/submit", methods=["POST"])
    @role_required("student")
    def submit_attempt(attempt_id):
        student = _current_student()
        oid = to_object_id(attempt_id)
        if not oid:
            return error("Invalid attempt id.", 404)
        attempt = attempts.find_one({"_id": oid, "studentId": student["_id"]})
        if not attempt:
            return error("Attempt not found.", 404)
        if attempt["status"] == "submitted":
            return error("This attempt has already been submitted.", 409)

        data = request.get_json(silent=True) or {}
        answers = data.get("answers") or {}
        if not isinstance(answers, dict):
            return error("answers must be an object of {questionId: selectedOption}.")

        assessment = assessments.find_one({"_id": attempt["assessmentId"]})
        attempt = _finalize_assessment_attempt(assessment, attempt, answers)
        log_activity(
            db, student["_id"], "student", "assessment_submitted",
            f'Submitted assessment "{(assessment or {}).get("name") or (assessment or {}).get("title") or ""}"',
            student_id=student["_id"],
            meta={"assessmentId": str(attempt["assessmentId"]), "attemptId": str(oid),
                  "percentage": attempt["overall"]["percentage"]},
        )

        response = {
            "attemptId": str(oid),
            "sectionScores": attempt["sectionScores"],
            "overall": attempt["overall"],
            "timeTakenSeconds": attempt["timeTakenSeconds"],
        }
        return ok(response, message="Assessment submitted.")

    # ==========================================================
    # 3b. COHORT STATUS — shows progress toward cohort generation
    #     (assessment score + interview score -> Final Employability Score)
    # ==========================================================
    @bp.route("/cohort-status", methods=["GET"])
    @role_required("student")
    def cohort_status():
        student = _current_student()
        # Task 3/4/5/8: Overall Score, Overall Cohort, average Interview
        # Score, Assessments Completed and Placement Readiness are ALL
        # computed once, server-side, in compute_overall_performance() —
        # this route (used by Dashboard Home and My Cohort) and /readiness
        # below both read the exact same computed dict, so these numbers
        # can never drift apart between pages.
        overall = compute_overall_performance(db, student["_id"]) or {}
        # Part 5 — StudentCohort collection: db.users.cohort (via
        # student_cohort_label above) stays the authoritative value this
        # route has always returned, so nothing that reads `cohort` here
        # changes behavior; `cohortLastUpdated`/`cohortSource` are new,
        # purely additive fields sourced from the dedicated collection.
        cohort_record = db.student_cohort.find_one({"studentId": student["_id"]})
        return ok({
            "cohort": student_cohort_label(student),
            "baselineAssessmentScore": student.get("baselineAssessmentScore"),
            "interviewScore": student.get("interviewScore"),
            "finalEmployabilityScore": student.get("finalEmployabilityScore"),
            "cohortAssignedAt": iso_utc(student.get("cohortAssignedAt")),
            "cohortLastUpdated": iso_utc(cohort_record.get("lastUpdated")) if cohort_record else None,
            "cohortSource": cohort_record.get("source") if cohort_record else None,
            # New, backend-computed "Overall Cohort Calculation" fields:
            "overallScore": overall.get("overallScore"),
            "overallCohort": overall.get("overallCohort"),
            "overallCohortLabel": overall.get("overallCohortLabel"),
            "averageInterviewScore": overall.get("averageInterviewScore"),
            "assessmentsCompleted": overall.get("assessmentsCompleted"),
            "placementReadiness": overall.get("placementReadiness"),
        })

    # ==========================================================
    # 3c. PLACEMENT READINESS — score + Ready/Not Ready, computed from
    #     compute_overall_performance() (same numbers as /cohort-status
    #     and Dashboard Home), plus a dynamic missing-skills /
    #     recommendations / roadmap breakdown for the Placement
    #     Readiness page. Nothing here is hardcoded content — every
    #     line is derived from this student's own stored results.
    # ==========================================================
    @bp.route("/readiness", methods=["GET"])
    @role_required("student")
    def placement_readiness_route():
        student = _current_student()
        if not student:
            return error("Student not found.", 404)

        overall = compute_overall_performance(db, student["_id"]) or {}
        readiness = overall.get("placementReadiness") or {}

        # Section-wise weak areas, from this student's own submitted
        # baseline-assessment section scores (same aggregation used by
        # /dashboard/charts and the existing download_report() 50%
        # "needs focus" convention above — never a newly-invented rule).
        section_agg = {}
        for att in attempts.find({"studentId": student["_id"], "status": "submitted"}):
            for section, s in att.get("sectionScores", {}).items():
                bucket = section_agg.setdefault(section, {"sum": 0.0, "n": 0})
                bucket["sum"] += s.get("percentage", 0)
                bucket["n"] += 1
        skill_radar = {section: round(b["sum"] / b["n"], 2) if b["n"] else 0 for section, b in section_agg.items()}
        weak_sections = sorted(
            [(section, pct) for section, pct in skill_radar.items() if pct < 50],
            key=lambda pair: pair[1],
        )
        missing_skills = [f"{section} ({pct}%)" for section, pct in weak_sections]
        recommendations = [f"Focus on improving {section} — currently at {pct}%." for section, pct in weak_sections]

        top_cohort = top_cohort_label(db)
        roadmap = [
            {
                "title": "Complete a validated assessment",
                "status": "done" if overall.get("assessmentsCompleted", 0) > 0 else "pending",
                "detail": f"{overall.get('assessmentsCompleted', 0)} assessment(s) validated so far.",
            },
            {
                "title": "Get your manual interview scored",
                "status": "done" if overall.get("averageInterviewScore") is not None else "pending",
                "detail": "Interview marks are entered by your trainer or Super Admin after each assessment.",
            },
            {
                "title": f"Reach Cohort {top_cohort}",
                "status": "done" if readiness.get("ready") else ("progress" if overall.get("overallCohort") else "pending"),
                "detail": readiness.get("summary", ""),
            },
        ]

        return ok({
            "score": readiness.get("score", 0),
            "ready": readiness.get("ready", False),
            "statusLabel": readiness.get("statusLabel", "Not Ready"),
            "summary": readiness.get("summary", ""),
            "overallCohort": overall.get("overallCohort"),
            "overallCohortLabel": overall.get("overallCohortLabel"),
            "assessmentsCompleted": overall.get("assessmentsCompleted"),
            "missingSkills": missing_skills,
            "recommendations": recommendations,
            "roadmap": roadmap,
        })

    # ==========================================================
    # 4. RESULTS PAGE — Assessment Name dropdown + Baseline/Mid/Final
    # ==========================================================
    @bp.route("/results", methods=["GET"])
    @role_required("student")
    def results_list():
        """Every assessment this student has completed — feeds the
        Assessment Name dropdown automatically."""
        student = _current_student()
        interview_score = student.get("interviewScore")
        pipeline = [
            {"$match": {"studentId": student["_id"], "status": "submitted"}},
            {"$sort": {"submittedAt": -1}},
            {"$lookup": {"from": "assessments", "localField": "assessmentId", "foreignField": "_id", "as": "assessment"}},
        ]
        rows = []
        for a in attempts.aggregate(pipeline):
            assessment = a["assessment"][0] if a.get("assessment") else {}
            quiz_score = a.get("overall", {}).get("percentage")
            overall_score = (
                round((quiz_score + interview_score) / 2, 2)
                if quiz_score is not None and interview_score is not None
                else quiz_score
            )
            rows.append({
                "attemptId": str(a["_id"]),
                "assessmentId": str(a["assessmentId"]),
                "assessmentName": a.get("assessmentName") or assessment.get("name"),
                "assessmentType": assessment.get("type"),
                "overallPercentage": quiz_score,
                "quizScore": quiz_score,
                "interviewScore": interview_score,
                "overallScore": overall_score,
                "pendingComponent": None if interview_score is not None else "interview",
                "passFail": a.get("overall", {}).get("passFail"),
                "marksObtained": a.get("overall", {}).get("obtainedMarks"),
                "totalMarks": a.get("overall", {}).get("totalMarks"),
                "status": "Completed",
                "startedAt": iso_utc(a.get("startedAt")),
                "submittedAt": iso_utc(a.get("submittedAt")),
            })
        return ok({"results": rows})

    @bp.route("/results/<assessment_id>", methods=["GET"])
    @role_required("student")
    def result_detail(assessment_id):
        student = _current_student()
        aoid = to_object_id(assessment_id)
        if not aoid:
            return error("Invalid assessment id.", 404)
        attempt = attempts.find_one({
            "studentId": student["_id"], "assessmentId": aoid, "status": "submitted",
        })
        if not attempt:
            return error("No result found for this assessment yet.", 404)
        assessment = assessments.find_one({"_id": aoid})

        # Task 5: Overall Score = Average(Quiz Score, Manual Interview Score).
        # The interview score lives on the student record (one interview
        # covers the student's whole placement readiness, not per-assessment
        # — see quiz_common.record_interview_score_for_cohort), so it's the
        # same value across every assessment's result page. If it hasn't
        # been scored yet, show the quiz score alone and say so explicitly
        # rather than silently averaging against a missing value.
        quiz_score = (attempt.get("overall") or {}).get("percentage")
        interview_score = student.get("interviewScore")
        if quiz_score is not None and interview_score is not None:
            overall_score = round((quiz_score + interview_score) / 2, 2)
            pending_component = None
        else:
            overall_score = quiz_score
            pending_component = "interview" if interview_score is None else "quiz"

        return ok({
            "assessment": serialize(assessment),
            "sectionScores": attempt.get("sectionScores", {}),
            "overall": attempt.get("overall", {}),
            "quizScore": quiz_score,
            "interviewScore": interview_score,
            "overallScore": overall_score,
            "pendingComponent": pending_component,
            "passFail": (attempt.get("overall") or {}).get("passFail"),
            "questionsAttempted": (attempt.get("overall") or {}).get("attempted"),
            "correctAnswers": (attempt.get("overall") or {}).get("correct"),
            "wrongAnswers": (attempt.get("overall") or {}).get("wrong"),
            "unansweredQuestions": (attempt.get("overall") or {}).get("unanswered"),
            "totalQuestions": (attempt.get("overall") or {}).get("total"),
            "marksObtained": (attempt.get("overall") or {}).get("obtainedMarks"),
            "totalMarks": (attempt.get("overall") or {}).get("totalMarks"),
            "startedAt": iso_utc(attempt.get("startedAt")),
            "submittedAt": iso_utc(attempt.get("submittedAt")),
            "timeTakenSeconds": attempt.get("timeTakenSeconds"),
            "studentRollNumber": attempt.get("studentRollNumber") or student.get("rollNumber"),
            "studentName": attempt.get("studentName") or student.get("fullName"),
            "college": attempt.get("college") or student.get("college"),
            "department": attempt.get("department") or student.get("department"),
            "cohort": attempt.get("cohort") or student_cohort_label(student),
        })

    # ==========================================================
    # 5. DYNAMIC GRAPHS for this student
    # ==========================================================
    @bp.route("/dashboard/charts", methods=["GET"])
    @role_required("student")
    def dashboard_charts():
        student = _current_student()
        section_agg = {}
        overall_scores = []
        trend_points = []
        for att in attempts.find({"studentId": student["_id"], "status": "submitted"}).sort("submittedAt", 1):
            overall = att.get("overall", {})
            if "percentage" in overall:
                overall_scores.append(overall["percentage"])
                trend_points.append({
                    "date": iso_utc(att.get("submittedAt")),
                    "percentage": overall["percentage"],
                })
            for section, s in att.get("sectionScores", {}).items():
                bucket = section_agg.setdefault(section, {"sum": 0.0, "n": 0})
                bucket["sum"] += s.get("percentage", 0)
                bucket["n"] += 1

        skill_radar = {section: round(b["sum"] / b["n"], 2) if b["n"] else 0 for section, b in section_agg.items()}
        overall_avg = round(sum(overall_scores) / len(overall_scores), 2) if overall_scores else 0

        return ok({
            "skillRadar": skill_radar,
            "categoryPercentage": skill_radar,
            "overallScore": overall_avg,
            "performanceTrend": trend_points,
        })

    # ==========================================================
    # 6. DOWNLOAD REPORT — generated from real DB values
    # ==========================================================
    @bp.route("/results/<assessment_id>/report", methods=["GET"])
    @role_required("student")
    def download_report(assessment_id):
        student = _current_student()
        aoid = to_object_id(assessment_id)
        if not aoid:
            return error("Invalid assessment id.", 404)
        attempt = attempts.find_one({
            "studentId": student["_id"], "assessmentId": aoid, "status": "submitted",
        })
        if not attempt:
            return error("No result found for this assessment yet.", 404)
        assessment = assessments.find_one({"_id": aoid}) or {}

        overall = attempt.get("overall", {})
        section_scores = attempt.get("sectionScores", {})
        cohort = student_cohort_label(student)

        recommendations = []
        for section, s in section_scores.items():
            if s.get("percentage", 0) < 50:
                recommendations.append(f"Focus on improving {section} — currently at {s['percentage']}%.")
        if not recommendations:
            recommendations.append("Strong performance across all sections — keep it up.")

        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas

            buf = BytesIO()
            c = canvas.Canvas(buf, pagesize=A4)
            width, height = A4
            y = height - 50

            def line(text, size=11, gap=18, bold=False):
                nonlocal y
                c.setFont("Helvetica-Bold" if bold else "Helvetica", size)
                c.drawString(50, y, text)
                y -= gap

            line("Employability Intelligence Platform — Assessment Report", 15, 26, bold=True)
            line(f"Student: {student.get('fullName', '')}")
            line(f"Email: {student.get('email', '')}")
            line(f"College: {student.get('college', '')}")
            line(f"Cohort: {cohort}")
            line(f"Assessment: {assessment.get('name', '')}")
            line(f"Date: {attempt['submittedAt'].strftime('%Y-%m-%d %H:%M UTC') if attempt.get('submittedAt') else ''}")
            y -= 8
            line("Category Scores:", 13, 20, bold=True)
            for section, s in section_scores.items():
                line(f"  {section}: {s['correct']}/{s['total']}  ({s['percentage']}%)")
            y -= 8
            line(f"Overall Score: {overall.get('correct', 0)}/{overall.get('total', 0)}", 13, 20, bold=True)
            line(f"Overall Percentage: {overall.get('percentage', 0)}%", 13, 20, bold=True)
            y -= 8
            line("Recommendations:", 13, 20, bold=True)
            for rec in recommendations:
                line(f"  - {rec}")

            c.showPage()
            c.save()
            buf.seek(0)
            filename = f"{(assessment.get('name') or 'assessment').replace(' ', '_')}_report.pdf"
            return send_file(buf, mimetype="application/pdf", as_attachment=True, download_name=filename)
        except ImportError:
            # reportlab not installed — fall back to a structured JSON report
            # rather than any hardcoded/sample content.
            return ok({
                "studentDetails": {
                    "name": student.get("fullName"),
                    "email": student.get("email"),
                    "college": student.get("college"),
                    "cohort": cohort,
                },
                "assessmentName": assessment.get("name"),
                "date": iso_utc(attempt.get("submittedAt")),
                "categoryScores": section_scores,
                "overallScore": overall,
                "recommendations": recommendations,
            })

    # ==========================================================
    # STUDENT QUIZ MODULE — manually-authored quizzes created via the
    # Trainer/Super Admin "Create Quiz" wizard (db.quizzes, see
    # quiz_module.py). Fully separate collection/flow from the cohort
    # assessment engine above; nothing here touches db.assessments,
    # db.assessment_attempts, or cohort generation. This section only
    # adds new routes on the SAME /api/student blueprint — no existing
    # route above is modified.
    #
    # Known, deliberate scope limits (flagged rather than silently
    # "fixed", since fixing them means changing the Create-Quiz wizard,
    # which this task was not to touch):
    #   - The quiz document has no "department" field today (only
    #     cohortTarget + colleges/collegeNames), so department-level
    #     eligibility filtering isn't possible until the wizard adds one.
    #   - "Passing Marks" and "Negative Marking" were explicitly removed
    #     from the quiz schema/validation in an earlier pass (see the
    #     NOTE in quiz_module.validate_and_normalize). The grading logic
    #     below still checks for a `passingMarks` field defensively so it
    #     activates automatically with zero code changes if that field is
    #     ever reintroduced, but today it will always be null.
    # ==========================================================

    def _quiz_eligible(quiz, student):
        """A student is eligible for a quiz when both their cohort/Entry
        Level status AND their college match what the quiz was targeted
        at. Reuses quiz_common.student_matches_cohort_target so cohort
        rules can never drift out of sync with the assessment engine."""
        if not student_matches_cohort_target(student, quiz.get("cohortTarget") or "all"):
            return False
        allowed_colleges = quiz.get("collegeNames") or quiz.get("colleges") or []
        if allowed_colleges:
            student_college = str(student.get("college") or "").strip().lower()
            allowed = {str(c).strip().lower() for c in allowed_colleges}
            if not student_college or student_college not in allowed:
                return False
        return True

    def _quiz_open_query():
        return {"state": "published", "cancelled": {"$ne": True}, "archived": {"$ne": True}}

    def _quiz_student_status(quiz, attempt):
        """Student-facing status: scheduled / live / attempted / completed /
        cancelled / archived. 'attempted' (already submitted) always wins
        over the time-based status, per spec ("Attempted" quizzes stay
        visible/labelled even after the window closes)."""
        if attempt and attempt.get("status") == "submitted":
            return "attempted"
        base = quiz_compute_status(quiz)
        return "live" if base == "active" else base

    def _aware(dt):
        if not dt:
            return None
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)

    def _quiz_total_marks(quiz):
        return round(sum(float(q.get("marks") or 0) for q in (quiz.get("questions") or [])), 2)

    def _serialize_quiz_card(quiz, attempt):
        status = _quiz_student_status(quiz, attempt)
        has_submitted = bool(attempt and attempt.get("status") == "submitted")
        has_in_progress = bool(attempt and attempt.get("status") == "in_progress")
        end = _aware(quiz.get("endDateTime"))
        remaining_seconds = max(0, int((end - now()).total_seconds())) if (status == "live" and end) else None
        return {
            "id": str(quiz["_id"]),
            "title": quiz.get("title"),
            "category": quiz.get("category"),
            "description": quiz.get("description"),
            "cohortTarget": quiz.get("cohortTarget"),
            "durationMinutes": quiz.get("durationMinutes"),
            "questionsAvailable": quiz.get("questionsAvailable"),
            "questionsDisplayed": quiz.get("questionsDisplayed"),
            "totalMarks": _quiz_total_marks(quiz),
            "passingMarks": quiz.get("passingMarks"),
            "startDateTime": _aware(quiz.get("startDateTime")).isoformat() if quiz.get("startDateTime") else None,
            "endDateTime": _aware(quiz.get("endDateTime")).isoformat() if quiz.get("endDateTime") else None,
            "status": status,
            "remainingSeconds": remaining_seconds,
            "canStart": status == "live" and not has_submitted,
            "action": "resume" if has_in_progress else "start",
            "attemptId": str(attempt["_id"]) if attempt else None,
            "result": attempt.get("overall") if has_submitted else None,
            # Part 4/9 — "Completed in XX Minutes": timeTakenSeconds is
            # already computed and stored on the attempt at submit time
            # (see _finalize_attempt: submittedAt - startedAt), it just
            # wasn't being surfaced on the card payload before.
            "timeTakenSeconds": attempt.get("timeTakenSeconds") if has_submitted else None,
        }

    def _attempt_payload(quiz, attempt, drawn=None):
        """Student-facing attempt payload — correct answers are NEVER
        included (only text/options/type/marks + this student's own
        currently-saved selection for each question)."""
        if drawn is None:
            pool = quiz.get("questions") or []
            drawn = []
            for pq in (attempt.get("pooledQuestions") or []):
                idx = pq.get("poolIndex")
                if idx is not None and 0 <= idx < len(pool):
                    q = dict(pool[idx])
                    q["_poolIndex"] = idx
                    drawn.append(q)
        deadline = _aware(attempt.get("deadline"))
        remaining_seconds = max(0, int((deadline - now()).total_seconds())) if deadline else None
        saved_answers = attempt.get("answers") or {}
        questions_out = []
        for q in drawn:
            idx = q["_poolIndex"]
            questions_out.append({
                "poolIndex": idx,
                "text": q.get("text"),
                "options": q.get("options") or [],
                "type": q.get("type"),
                "section": q.get("section"),
                "difficulty": q.get("difficulty"),
                "marks": q.get("marks") or 0,
                "selected": saved_answers.get(str(idx), []),
            })
        return {
            "attemptId": str(attempt["_id"]),
            "status": attempt.get("status"),
            "quiz": {
                "id": str(quiz["_id"]),
                "title": quiz.get("title"),
                "description": quiz.get("description"),
                "durationMinutes": quiz.get("durationMinutes"),
            },
            "questions": questions_out,
            "remainingSeconds": remaining_seconds,
        }

    def _grade_answers(quiz, served_pool_indices, answers):
        pool = quiz.get("questions") or []
        correct_n = wrong_n = unanswered_n = 0
        marks_obtained = 0.0
        total_marks = 0.0
        for idx in served_pool_indices:
            if idx is None or idx >= len(pool):
                continue
            q = pool[idx]
            marks = float(q.get("marks") or 0)
            total_marks += marks
            raw = answers.get(str(idx))
            submitted = raw if isinstance(raw, list) else ([raw] if raw not in (None, "") else [])
            if not submitted:
                unanswered_n += 1
                continue
            try:
                submitted_idx = sorted({int(x) for x in submitted})
            except (TypeError, ValueError):
                submitted_idx = []
            correct_idx = sorted(set(q.get("correct") or []))
            if submitted_idx and submitted_idx == correct_idx:
                correct_n += 1
                marks_obtained += marks
            else:
                wrong_n += 1
        percentage = round((marks_obtained / total_marks) * 100, 2) if total_marks else 0.0
        passing_marks = quiz.get("passingMarks")
        passed = (marks_obtained >= float(passing_marks)) if passing_marks is not None else None
        return {
            "correct": correct_n,
            "wrong": wrong_n,
            "unanswered": unanswered_n,
            "totalMarks": round(total_marks, 2),
            "marksObtained": round(marks_obtained, 2),
            "percentage": percentage,
            "passed": passed,
        }

    def _finalize_attempt(quiz, attempt, answers):
        served = [pq.get("poolIndex") for pq in (attempt.get("pooledQuestions") or [])]
        overall = _grade_answers(quiz, served, answers)
        submitted_at = now()
        started_at = attempt.get("startedAt") or submitted_at
        time_taken_seconds = int((submitted_at - _aware(started_at)).total_seconds())
        student = users.find_one({"_id": attempt["studentId"]}) or {}
        record_fields = {
            "answers": answers,
            "status": "submitted",
            "submittedAt": submitted_at,
            "overall": overall,
            "timeTakenSeconds": time_taken_seconds,
            "studentRollNumber": student.get("rollNumber"),
            "studentName": student.get("fullName"),
            "college": student.get("college"),
            "department": student.get("department"),
            "cohort": student_cohort_label(student),
            "completionStatus": "completed",
        }
        # Marks Management / Interview Verification: only set these the
        # FIRST time an attempt is finalized (a re-finalize can't happen —
        # submit is one-way — but this guards against ever clobbering an
        # interview score if this function is ever called twice).
        if attempt.get("resultStatus") is None:
            record_fields.update(init_quiz_result_fields())
        quiz_attempts.update_one({"_id": attempt["_id"]}, {"$set": record_fields})
        attempt.update(record_fields)
        log_activity(
            db, attempt["studentId"], "student", "quiz_submitted",
            f'Submitted quiz "{quiz.get("title")}"',
            student_id=attempt["studentId"],
            meta={"quizId": str(quiz["_id"]), "attemptId": str(attempt["_id"]), "percentage": overall["percentage"]},
        )
        return attempt

    def _auto_expire_if_needed(quiz, attempt):
        """Server-side enforcement of the timer/Auto-Submit rule — called
        at the top of every quiz-attempt route so a client that never
        calls submit (closed tab, crashed browser, tampered timer, etc.)
        still gets force-submitted the moment its deadline has passed."""
        if attempt.get("status") != "in_progress":
            return False
        deadline = _aware(attempt.get("deadline"))
        if not deadline or now() <= deadline:
            return False
        _finalize_attempt(quiz, attempt, attempt.get("answers") or {})
        return True

    # ----------------------------------------------------
    # LIST — quizzes eligible for the logged-in student, bucketed by status.
    # Never returns quizzes for another college/cohort, or ones that are
    # draft/cancelled/archived.
    # ----------------------------------------------------
    @bp.route("/quizzes", methods=["GET"])
    @role_required("student")
    def my_quizzes():
        student = _current_student()
        if not student:
            return error("Student not found.", 404)
        docs = [normalize_quiz_college_names(db, q) for q in quizzes.find(_quiz_open_query()).sort("startDateTime", 1)]
        eligible = [q for q in docs if _quiz_eligible(q, student)]
        attempt_docs = quiz_attempts.find({
            "studentId": student["_id"], "quizId": {"$in": [q["_id"] for q in eligible]},
        })
        attempt_by_quiz = {a["quizId"]: a for a in attempt_docs}
        cards = [_serialize_quiz_card(q, attempt_by_quiz.get(q["_id"])) for q in eligible]
        buckets = {"scheduled": [], "live": [], "attempted": [], "completed": []}
        for c in cards:
            buckets.setdefault(c["status"], []).append(c)
        return ok({"quizzes": cards, **buckets})

    # ----------------------------------------------------
    # DETAIL — a single quiz's full card info. Manual URL access to a
    # quiz the student isn't eligible for is blocked with 403/404, never
    # silently allowed.
    # ----------------------------------------------------
    @bp.route("/quizzes/<quiz_id>", methods=["GET"])
    @role_required("student")
    def quiz_detail(quiz_id):
        student = _current_student()
        oid = to_object_id(quiz_id)
        if not oid:
            return error("Invalid quiz id.", 404)
        quiz = quizzes.find_one({**_quiz_open_query(), "_id": oid})
        if not quiz:
            return error("Quiz not found.", 404)
        quiz = normalize_quiz_college_names(db, quiz)
        if not _quiz_eligible(quiz, student):
            return error("This quiz is not available to you.", 403)
        attempt = quiz_attempts.find_one({"studentId": student["_id"], "quizId": oid})
        return ok({"quiz": _serialize_quiz_card(quiz, attempt)})

    # ----------------------------------------------------
    # START / RESUME — generates the per-student random question set
    # (via quiz_module.select_random_questions, respecting Section/
    # Difficulty Distribution exactly as configured) the first time, or
    # returns the already-in-progress attempt unchanged on any later call
    # so refreshing the page never re-randomizes or loses answers.
    # ----------------------------------------------------
    @bp.route("/quizzes/<quiz_id>/start", methods=["POST"])
    @role_required("student")
    def start_quiz(quiz_id):
        student = _current_student()
        oid = to_object_id(quiz_id)
        if not oid:
            return error("Invalid quiz id.", 404)
        quiz = quizzes.find_one({**_quiz_open_query(), "_id": oid})
        if not quiz:
            return error("Quiz not found.", 404)
        quiz = normalize_quiz_college_names(db, quiz)
        if not _quiz_eligible(quiz, student):
            return error("This quiz is not available to you.", 403)

        status = quiz_compute_status(quiz)
        if status == "scheduled":
            return error("This quiz has not started yet.", 409)
        if status in ("completed", "cancelled", "archived"):
            return error("This quiz is no longer available.", 409)

        if quiz_attempts.find_one({"studentId": student["_id"], "quizId": oid, "status": "submitted"}):
            return error("You have already attempted this quiz.", 409)

        existing = quiz_attempts.find_one({"studentId": student["_id"], "quizId": oid, "status": "in_progress"})
        if existing:
            if _auto_expire_if_needed(quiz, existing):
                return error("Time is up — this attempt was auto-submitted.", 409)
            return ok(_attempt_payload(quiz, existing))

        drawn = quiz_draw_questions(quiz)
        if not drawn:
            return error("This quiz has no questions configured yet.", 409)

        started_at = now()
        duration_seconds = int(quiz.get("durationMinutes") or 0) * 60
        deadline = started_at + timedelta(seconds=duration_seconds) if duration_seconds else None
        quiz_end = _aware(quiz.get("endDateTime"))
        if quiz_end and (not deadline or quiz_end < deadline):
            deadline = quiz_end

        attempt_doc = {
            "studentId": student["_id"],
            "quizId": oid,
            "quizTitle": quiz.get("title"),
            "pooledQuestions": [
                {"poolIndex": q["_poolIndex"], "type": q.get("type"), "marks": q.get("marks") or 0}
                for q in drawn
            ],
            "answers": {},
            "status": "in_progress",
            "startedAt": started_at,
            "deadline": deadline,
            "submittedAt": None,
            "overall": None,
        }
        result = quiz_attempts.insert_one(attempt_doc)
        attempt_doc["_id"] = result.inserted_id
        log_activity(
            db, student["_id"], "student", "quiz_started",
            f'Started quiz "{quiz.get("title")}"', student_id=student["_id"],
            meta={"quizId": str(oid)},
        )
        return ok(_attempt_payload(quiz, attempt_doc, drawn=drawn), status=201)

    # ----------------------------------------------------
    # RESUME/POLL — fetch the current state of an attempt (used on page
    # reload so the timer/answers/question-set survive a refresh).
    # ----------------------------------------------------
    @bp.route("/quiz-attempts/<attempt_id>", methods=["GET"])
    @role_required("student")
    def get_quiz_attempt(attempt_id):
        student = _current_student()
        oid = to_object_id(attempt_id)
        if not oid:
            return error("Invalid attempt id.", 404)
        attempt = quiz_attempts.find_one({"_id": oid, "studentId": student["_id"]})
        if not attempt:
            return error("Attempt not found.", 404)
        quiz = quizzes.find_one({"_id": attempt["quizId"]})
        if not quiz:
            return error("Quiz not found.", 404)
        if attempt["status"] == "in_progress" and _auto_expire_if_needed(quiz, attempt):
            attempt = quiz_attempts.find_one({"_id": oid})
        if attempt["status"] == "submitted":
            return ok({"attemptId": str(oid), "status": "submitted", "result": attempt.get("overall")})
        return ok(_attempt_payload(quiz, attempt))

    # ----------------------------------------------------
    # SAVE ANSWER — persists one question's selection during navigation
    # (Next/Previous/answer updates). Every answer selected is saved
    # immediately, not batched until final submit, so nothing is lost.
    # ----------------------------------------------------
    @bp.route("/quiz-attempts/<attempt_id>/save", methods=["POST"])
    @role_required("student")
    def save_quiz_answer(attempt_id):
        student = _current_student()
        oid = to_object_id(attempt_id)
        if not oid:
            return error("Invalid attempt id.", 404)
        attempt = quiz_attempts.find_one({"_id": oid, "studentId": student["_id"]})
        if not attempt:
            return error("Attempt not found.", 404)
        if attempt["status"] != "in_progress":
            return error("This attempt is no longer in progress.", 409)
        quiz = quizzes.find_one({"_id": attempt["quizId"]})
        if not quiz:
            return error("Quiz not found.", 404)
        if _auto_expire_if_needed(quiz, attempt):
            return error("Time is up — this attempt was auto-submitted.", 409)

        data = request.get_json(silent=True) or {}
        pool_index = data.get("poolIndex")
        if pool_index is None:
            return error("poolIndex is required.")
        selected = data.get("selected", [])
        if not isinstance(selected, list):
            selected = [selected] if selected not in (None, "") else []
        answers = attempt.get("answers") or {}
        answers[str(pool_index)] = selected
        quiz_attempts.update_one({"_id": oid}, {"$set": {"answers": answers}})
        return ok({"saved": True})

    # ----------------------------------------------------
    # SUBMIT — locks the attempt permanently (backend-enforced: once
    # status is "submitted" no route above will ever accept further
    # edits or a second submission), grades it, and stores the result.
    # ----------------------------------------------------
    @bp.route("/quiz-attempts/<attempt_id>/submit", methods=["POST"])
    @role_required("student")
    def submit_quiz_attempt(attempt_id):
        student = _current_student()
        oid = to_object_id(attempt_id)
        if not oid:
            return error("Invalid attempt id.", 404)
        attempt = quiz_attempts.find_one({"_id": oid, "studentId": student["_id"]})
        if not attempt:
            return error("Attempt not found.", 404)
        if attempt["status"] == "submitted":
            return error("This attempt has already been submitted.", 409)
        quiz = quizzes.find_one({"_id": attempt["quizId"]})
        if not quiz:
            return error("Quiz not found.", 404)

        data = request.get_json(silent=True) or {}
        incoming = data.get("answers")
        answers = attempt.get("answers") or {}
        if isinstance(incoming, dict):
            for k, v in incoming.items():
                answers[str(k)] = v if isinstance(v, list) else ([v] if v not in (None, "") else [])

        attempt = _finalize_attempt(quiz, attempt, answers)
        return ok({"attemptId": str(oid), "overall": attempt["overall"]}, message="Quiz submitted successfully.")

    # ----------------------------------------------------
    # QUIZ HISTORY — powers the existing (previously unwired) "Quiz
    # History" page in student.html, whose JS already calls
    # GET /api/student/quiz-history and expects exactly this shape.
    # ----------------------------------------------------
    @bp.route("/quiz-history", methods=["GET"])
    @role_required("student")
    def quiz_history():
        student = _current_student()

        # Total Quizzes = every published quiz currently eligible for this
        # student (same cohort/college eligibility rules as the /quizzes
        # list above) — NOT the number of attempts the student happens to
        # have. This is recomputed from the database on every call, so it
        # always reflects the live count the moment a trainer creates a
        # quiz or one becomes available to this student.
        eligible_quizzes = [
            q for q in (normalize_quiz_college_names(db, q) for q in quizzes.find(_quiz_open_query()))
            if _quiz_eligible(q, student)
        ]
        total_quizzes = len(eligible_quizzes)

        docs = list(quiz_attempts.find({"studentId": student["_id"]}).sort("startedAt", -1))
        submitted = [d for d in docs if d.get("status") == "submitted"]
        # Attempted = every quiz this student has started (in progress or
        # submitted) — one attempt document per quiz, so this is just the
        # count of attempt documents.
        attempted = len(docs)
        completed = len(submitted)
        # Pending is always Total - Completed, per spec — never derived
        # from "not yet attempted" alone, so a quiz still shows as pending
        # even before the student has started it.
        pending = max(total_quizzes - completed, 0)
        percentages = [(d.get("overall") or {}).get("percentage") for d in submitted if (d.get("overall") or {}).get("percentage") is not None]
        avg_pct = round(sum(percentages) / len(percentages), 1) if percentages else None
        highest_pct = round(max(percentages), 1) if percentages else None
        lowest_pct = round(min(percentages), 1) if percentages else None
        rows = []
        for d in docs:
            when = d.get("submittedAt") or d.get("startedAt")
            overall = d.get("overall") or {}
            rows.append({
                "attemptId": str(d["_id"]),
                "quizId": str(d["quizId"]),
                "name": d.get("quizTitle") or "Quiz",
                "date": iso_utc(when),
                "score": overall.get("percentage") if d.get("status") == "submitted" else None,
                "marksObtained": overall.get("marksObtained") if d.get("status") == "submitted" else None,
                "totalMarks": overall.get("totalMarks") if d.get("status") == "submitted" else None,
                "passed": overall.get("passed") if d.get("status") == "submitted" else None,
                "timeTakenSeconds": d.get("timeTakenSeconds") if d.get("status") == "submitted" else None,
                "status": "Completed" if d.get("status") == "submitted" else "In Progress",
            })
        return ok({
            "stats": {
                "total": total_quizzes,
                "attempted": attempted,
                "completed": completed,
                "pending": pending,
                "average": avg_pct,
                "averageScore": avg_pct,
                "highest": highest_pct,
                "lowest": lowest_pct,
            },
            "quizzes": rows,
        })

    # ----------------------------------------------------
    # MY RESULTS — Marks Management: every submitted quiz for this
    # student, WITH interview marks / final average / assigned cohort /
    # status the moment a trainer or super admin enters them. Nothing
    # here is hardcoded; a student only ever sees their own rows.
    # ----------------------------------------------------
    @bp.route("/quiz-results", methods=["GET"])
    @role_required("student")
    def my_quiz_results():
        student = _current_student()
        if not student:
            return error("Student not found.", 404)
        results = list_quiz_results(db, student_id=student["_id"])
        return ok({"results": results})

    # ----------------------------------------------------
    # ATTENDANCE — live view for the logged-in student.
    # Attendance is written by the Super Admin (Mark/View Attendance
    # screens -> db.attendance, one doc per student per class date).
    # The student page derives percentages on the fly from the shared
    # attendance_summary() helper (single source of truth) instead of
    # storing or trusting any pre-aggregated number.
    # ----------------------------------------------------
    @bp.route("/attendance", methods=["GET"])
    @role_required("student")
    def my_attendance():
        student = _current_student()
        if not student:
            return error("Student not found.", 404)
        return ok({"attendance": attendance_summary(db, student["_id"])})

    def _session_start_dt_local(doc):
        date_str, start = doc.get("date"), doc.get("startTime")
        if not (date_str and start):
            return None
        try:
            return datetime.strptime(f"{date_str} {start}", "%Y-%m-%d %H:%M")
        except ValueError:
            return None

    def _student_workshop_public(doc, student):
        """Student-facing view of a workshop session — the same document
        superadmin.py serves admins, reduced to what a student needs plus
        the student's own attendance mark and join state."""
        date_str = doc.get("date")
        start_dt = _session_start_dt_local(doc)
        now_dt = datetime.now()
        has_started = start_dt is not None and now_dt >= start_dt
        window_open = (
            doc.get("mode") == "online"
            and start_dt is not None
            and now_dt >= (start_dt - timedelta(hours=5))
        )
        attendance = None
        if date_str:
            attendance = db.attendance.find_one(
                {"studentId": student["_id"], "date": date_str})
        marked = None
        if attendance:
            st = attendance.get("status")
            marked = "Present" if st == "present" else ("Absent" if st == "absent" else None)
        if not has_started:
            label = "Upcoming"
        elif marked:
            label = "Completed"
        else:
            label = "Session Live" if window_open or (start_dt and now_dt <= start_dt + timedelta(hours=5)) else "Completed"
        return {
            "id": str(doc["_id"]),
            "name": doc.get("name"),
            "description": doc.get("description"),
            "trainer": ", ".join(doc.get("trainerNames", [])) or None,
            "venue": doc.get("venue") or "—",
            "mode": doc.get("mode") or "offline",
            "date": date_str,
            "dateLabel": f"{date_str or '—'} · {doc.get('startTime','')} – {doc.get('endTime','')}" if date_str else "—",
            "duration": f"{doc.get('startTime','')} – {doc.get('endTime','')}" if doc.get("startTime") else "—",
            "meetingLink": doc.get("meetingLink"),
            "meetingLinkWindowOpen": bool(window_open),
            "attendanceRequirement": doc.get("attendanceRequirement"),
            "attendance": marked or "Not marked",
            "status": label,
        }

    @bp.route("/workshops", methods=["GET"])
    @role_required("student")
    def my_workshops():
        """Workshops & Interventions page — every scheduled workshop session
        targeted at this student's college, split into upcoming and
        completed/completed-past. Real data from db.workshop_sessions, with
        the student's own attendance mark read live from db.attendance."""
        student = _current_student()
        if not student:
            return error("Student not found.", 404)
        college = student.get("college")
        query = {"status": "scheduled"}
        if college:
            query["collegeNames"] = college
        docs = list(db.workshop_sessions.find(query).sort("date", 1))
        upcoming, completed = [], []
        for d in docs:
            start_dt = _session_start_dt_local(d)
            p = _student_workshop_public(d, student)
            finished = start_dt is not None and datetime.now() >= start_dt + timedelta(hours=5)
            (completed if finished else upcoming).append(p)
        return ok({"upcoming": upcoming, "completed": completed})

    # ==========================================================
    # EVENTS — Meetings & Events feed on the student dashboard.
    # Mirrors the super_admin db.meetings collection (admin module):
    # non-cancelled events visible to everyone, sorted by date/time.
    # Status/join-window auto-advances server-side from the stored
    # schedule (same wall-clock convention as superadmin._meeting_status,
    # kept server-side so the browser is never the time source).
    # Online events without a link yet are flagged "needs link".
    # ==========================================================
    def _event_status(doc):
        if doc.get("status") == "cancelled":
            return "cancelled"
        date_str, time_str = doc.get("date"), doc.get("time")
        if not (date_str and time_str):
            return "upcoming"
        try:
            start_dt = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %H:%M")
        except ValueError:
            return "upcoming"
        end_dt = start_dt + timedelta(minutes=int(doc.get("durationMinutes") or 60))
        now_dt = datetime.now()
        if now_dt < start_dt:
            return "upcoming"
        if now_dt > end_dt:
            return "completed"
        return "ongoing"

    @bp.route("/events", methods=["GET"])
    @role_required("student")
    def my_events():
        """Meetings & Events feed for the student dashboard — every
        non-cancelled meeting/event, plus whether this student already
        joined (meeting attendance is tracked separately from the class
        attendance module in db.meeting_attendance)."""
        student = _current_student()
        if not student:
            return error("Student not found.", 404)
        docs = list(db.meetings.find({"status": {"$ne": "cancelled"}}).sort("date", 1))
        joined_ids = {
            r["meetingId"] for r in
            db.meeting_attendance.find({"studentId": student["_id"]}, {"meetingId": 1})
        }
        events = []
        for d in docs:
            status = _event_status(d)
            needs_link = d.get("type") == "online" and not d.get("meetingLink")
            label, badge_class = {
                "ongoing": ("Ongoing", "badge-teal"),
                "upcoming": ("Upcoming", "badge-amber"),
                "completed": ("Completed", "badge-ink"),
            }.get(status, ("Upcoming", "badge-amber"))
            events.append({
                "id": str(d["_id"]),
                "title": d.get("title"),
                "description": d.get("description") or "",
                "date": d.get("date"),
                "time": d.get("time"),
                "dateLabel": f"{d.get('date') or '—'} · {d.get('time') or '—'}"
                             if d.get("date") else "—",
                "durationMinutes": d.get("durationMinutes") or 60,
                "type": d.get("type") or "offline",
                "venue": d.get("venue"),
                "meetingLink": d.get("meetingLink"),
                "needsLink": needs_link,
                "mandatory": bool(d.get("mandatory")),
                "status": status,
                "label": label,
                "badgeClass": badge_class,
                "joined": d["_id"] in joined_ids,
            })
        return ok({"events": events, "total": len(events)})

    @bp.route("/events/<event_id>/join", methods=["POST"])
    @role_required("student")
    def join_event(event_id):
        """Record a student joining a meeting/event (db.meeting_attendance,
        tracked separately from the class attendance module). For an online
        event the meeting link is returned; for offline, the venue. Joining
        is idempotent — a second join doesn't create a duplicate."""
        student = _current_student()
        if not student:
            return error("Student not found.", 404)
        oid = to_object_id(event_id)
        if not oid:
            return error("Invalid event id.", 404)
        meeting = db.meetings.find_one({"_id": oid, "status": {"$ne": "cancelled"}})
        if not meeting:
            return error("Event not found.", 404)
        status = _event_status(meeting)
        if status != "ongoing":
            return error("This event is not currently open to join.")
        link = meeting.get("meetingLink") if meeting.get("type") == "online" else None
        if meeting.get("type") == "online" and not link:
            return error("The meeting link has not been shared yet.")
        db.meeting_attendance.update_one(
            {"meetingId": oid, "studentId": student["_id"]},
            {"$setOnInsert": {"meetingId": oid, "studentId": student["_id"], "joinedAt": now()}},
            upsert=True,
        )
        return ok({
            "eventId": str(oid),
            "type": meeting.get("type"),
            "meetingLink": link,
            "venue": meeting.get("venue"),
        }, message="Joining confirmed. See you there!")

    return bp
