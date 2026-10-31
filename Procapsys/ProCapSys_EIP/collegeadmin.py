"""
============================================================
 collegeadmin.py — College Admin Dashboard backend
============================================================
Registered at /api/collegeadmin. Read-only Assessment Management for
the College Admin role, scoped to that admin's own assigned college.

Reuses the exact same data the Trainer/Super Admin "Create Quiz"
feature already writes (db.quizzes / db.quiz_attempts, see
quiz_module.py + student.py) — nothing here is a separate/parallel
data source, so College Admin can never see numbers that disagree
with what Trainers and students actually created/submitted.

Every route is @role_required("college_admin") and every route in
this file is a GET — there is intentionally no create/update/delete
here. College Admin can view assessments and student responses for
their college; they cannot create, edit, delete, publish, schedule,
or modify marks. That authority stays with Trainer/Super Admin
(quiz_module.py) exactly as before.
"""

from flask import Blueprint, request
from flask_jwt_extended import get_jwt_identity

from quiz_common import (
    ok, error, role_required, to_object_id, iso_utc, serialize,
    cohort_counts, student_cohort_label, ENTRY_LEVEL, VALID_COHORTS,
    attendance_summary, now,
)
from quiz_module import serialize_quiz, normalize_quiz_college_names


def init_collegeadmin(db):
    bp = Blueprint("collegeadmin", __name__)

    users = db.users
    quizzes = db.quizzes
    quiz_attempts = db.quiz_attempts

    def _admin_doc():
        return users.find_one({"_id": to_object_id(get_jwt_identity())})

    def _admin_college():
        return (_admin_doc() or {}).get("college")

    def _quiz_in_college_scope(quiz, college):
        """Mirrors student.py's _quiz_eligible college check: an empty
        collegeNames/colleges list means the quiz targets every college,
        otherwise the admin's college name must appear in that list."""
        allowed = quiz.get("collegeNames") or quiz.get("colleges") or []
        if not allowed:
            return True
        allowed_l = {str(c).strip().lower() for c in allowed}
        return str(college).strip().lower() in allowed_l

    # ==========================================================
    # ASSESSMENT MANAGEMENT — list, scoped to the admin's college
    # ==========================================================
    @bp.route("/assessments", methods=["GET"])
    @role_required("college_admin")
    def list_assessments():
        college = _admin_college()
        if not college:
            return ok({"assessments": [], "college": None})

        rows = []
        for doc in quizzes.find({}).sort("createdOn", -1):
            doc = normalize_quiz_college_names(db, doc)
            if not _quiz_in_college_scope(doc, college):
                continue
            s = serialize_quiz(doc)
            created_by = doc.get("createdBy") or {}
            response_count = quiz_attempts.count_documents({
                "quizId": doc["_id"],
                "status": "submitted",
                "college": college,
            })
            rows.append({
                "id": s["id"],
                "name": s.get("title"),
                "createdBy": created_by.get("name") or created_by.get("role") or "—",
                "status": s.get("status"),
                "totalQuestions": s.get("totalQuestions"),
                "totalMarks": s.get("totalMarks"),
                "scheduledDate": s.get("startDateTime"),
                "endDate": s.get("endDateTime"),
                "responseCount": response_count,
            })
        return ok({"assessments": rows, "college": college})

    # ==========================================================
    # STUDENT RESPONSES — one assessment's submitted attempts,
    # scoped to the admin's college. Read-only: marks/results are
    # entered/validated by Trainer only (see trainer.py).
    # ==========================================================
    @bp.route("/assessments/<assessment_id>/responses", methods=["GET"])
    @role_required("college_admin")
    def assessment_responses(assessment_id):
        college = _admin_college()
        oid = to_object_id(assessment_id)
        if not oid:
            return error("Invalid assessment id.", 404)

        quiz = quizzes.find_one({"_id": oid})
        if not quiz:
            return error("Assessment not found.", 404)
        quiz = normalize_quiz_college_names(db, quiz)

        if not college or not _quiz_in_college_scope(quiz, college):
            return error("Assessment not found.", 404)

        cursor = quiz_attempts.find({
            "quizId": oid,
            "status": "submitted",
            "college": college,
        }).sort("submittedAt", -1)

        responses = []
        for a in cursor:
            overall = a.get("overall") or {}
            submitted_at = a.get("submittedAt")
            responses.append({
                "attemptId": str(a["_id"]),
                "studentId": str(a["studentId"]) if a.get("studentId") else None,
                "studentName": a.get("studentName"),
                "registerNumber": a.get("studentRollNumber"),
                "department": a.get("department"),
                "marksObtained": overall.get("marksObtained"),
                "totalMarks": overall.get("totalMarks"),
                "percentage": overall.get("percentage"),
                "submissionStatus": "Submitted",
                "submittedAt": iso_utc(submitted_at),
            })

        return ok({
            "assessment": {
                "id": str(quiz["_id"]),
                "name": quiz.get("title"),
            },
            "responses": responses,
        })

    # ==========================================================
    # STUDENTS / COHORTS — live, scoped to the admin's own college.
    #
    # Every value below is read fresh from db.users / db.placement_rules
    # on every request — nothing is cached in this file or precomputed
    # at login. A student's `cohort` field is written exclusively by the
    # shared cohort engine in quiz_common.py (check_and_generate_cohort /
    # recompute_cohort_from_quiz_results / the Placement Rules bulk
    # recalculation), so College Admin always sees the exact same A/B/C/
    # Entry Level assignment Student, Trainer, and Super Admin see for
    # that student — same source, no second copy, no stale snapshot.
    # ==========================================================
    @bp.route("/students", methods=["GET"])
    @role_required("college_admin")
    def list_students():
        """Full student roster for this admin's college, with each
        student's CURRENT cohort. No cohort math happens here — this
        only reads db.users.cohort (via student_cohort_label), which the
        shared cohort engine keeps up to date for every student."""
        college = _admin_college()
        if not college:
            return ok({"students": [], "college": None})

        cohort_filter = (request.args.get("cohort") or "").strip()
        query = {"role": "student", "college": college}
        if cohort_filter == ENTRY_LEVEL:
            query["$or"] = [{"cohort": None}, {"cohort": {"$exists": False}}]
        elif cohort_filter in VALID_COHORTS:
            query["cohort"] = cohort_filter

        docs = list(users.find(query, {"passwordHash": 0}).sort("createdAt", -1))

        # Attendance per student, computed once for the whole roster via a
        # single aggregation on db.attendance (never stored/trusted — same
        # shared source the Student Dashboard and Super Admin use).
        oids = [d["_id"] for d in docs]
        att_stats = {}
        if oids:
            for row in db.attendance.aggregate([
                {"$match": {"studentId": {"$in": oids}}},
                {"$group": {
                    "_id": "$studentId",
                    "present": {"$sum": {"$cond": [{"$eq": ["$status", "present"]}, 1, 0]}},
                    "absent": {"$sum": {"$cond": [{"$eq": ["$status", "absent"]}, 1, 0]}},
                    "lastUpdated": {"$max": {"$ifNull": ["$updatedAt", "$markedAt"]}},
                }},
            ]):
                att_stats[row["_id"]] = row

        rows = []
        for s in docs:
            a = att_stats.get(s["_id"], {})
            present = a.get("present", 0)
            absent = a.get("absent", 0)
            total = present + absent
            rows.append({
                "id": str(s["_id"]),
                "name": s.get("fullName") or s.get("name"),
                "rollNumber": s.get("rollNumber"),
                "department": s.get("department"),
                "email": s.get("email"),
                "college": s.get("college"),
                "cohort": student_cohort_label(s),
                "finalEmployabilityScore": s.get("finalEmployabilityScore"),
                "baselineAssessmentScore": s.get("baselineAssessmentScore"),
                "interviewScore": s.get("interviewScore"),
                "cohortAssignedAt": iso_utc(s.get("cohortAssignedAt")),
                "attendancePresent": present,
                "attendanceAbsent": absent,
                "attendanceTotal": total,
                "attendancePct": round(present / total * 100, 2) if total else 0,
                "attendanceLastUpdated": iso_utc(a.get("lastUpdated")),
            })
        return ok({"students": rows, "college": college})

    @bp.route("/students/<student_id>/attendance", methods=["GET"])
    @role_required("college_admin")
    def student_attendance(student_id):
        """Full live attendance for one student in this admin's college —
        summary counts + newest-first history, both from the shared
        attendance_summary() helper (same data the student sees on their
        own dashboard and Super Admin sees in Mark/View Attendance)."""
        college = _admin_college()
        if not college:
            return error("No college assigned to this admin.", 403)
        oid = to_object_id(student_id)
        if not oid:
            return error("Invalid student id.")
        student = users.find_one({"_id": oid, "role": "student", "college": college})
        if not student:
            return error("Student not found in your college.", 404)
        return ok({
            "student": {
                "id": str(student["_id"]),
                "name": student.get("fullName") or student.get("name"),
                "rollNumber": student.get("rollNumber"),
                "cohort": student_cohort_label(student),
            },
            "attendance": attendance_summary(db, oid),
        })

    # ==========================================================
    # STUDENT ASSESSMENT SUMMARY — the six assessment parameters.
    #
    # Exactly the categories used while creating quizzes: the five
    # DB-backed quiz sections (db.quiz_sections — the same source the
    # Create Quiz wizard reads) plus the Manual Interview score the
    # Super Admin/Trainer enters after quiz submission. Every value is
    # computed live from the student's submitted quiz attempts, so the
    # College Admin can never see a parameter the quiz configuration
    # doesn't actually have (and vice versa).
    # ==========================================================
    @bp.route("/students/<student_id>/assessment-summary", methods=["GET"])
    @role_required("college_admin")
    def student_assessment_summary(student_id):
        college = _admin_college()
        if not college:
            return error("No college assigned to this admin.", 403)
        oid = to_object_id(student_id)
        if not oid:
            return error("Invalid student id.")
        student = users.find_one({"_id": oid, "role": "student", "college": college})
        if not student:
            return error("Student not found in your college.", 404)

        sections_docs = list(db.quiz_sections.find({}).sort("name", 1))
        attempts = list(db.quiz_attempts.find({
            "studentId": oid,
            "status": "submitted",
            "college": college,
        }).sort("submittedAt", -1))

        # Per-section average % across this student's attempts. Pooled
        # questions resolve against the quiz's own question list (same
        # logic as compute_quiz_analytics) so the numbers come straight
        # from the DB every time.
        quizzes_cache = {}
        sec_sums = {}
        sec_counts = {}
        interview_scores = []
        overall_scores = []
        for attempt in attempts:
            interview = attempt.get("interviewMarks")
            if isinstance(interview, (int, float)) and not isinstance(interview, bool):
                interview_scores.append(float(interview))
            overall_pct = (attempt.get("overall") or {}).get("percentage")
            if isinstance(overall_pct, (int, float)) and not isinstance(overall_pct, bool):
                overall_scores.append(float(overall_pct))
            qid = attempt.get("quizId")
            if qid not in quizzes_cache:
                quizzes_cache[qid] = db.quizzes.find_one({"_id": qid}) or {}
            quiz = quizzes_cache[qid]
            pool = quiz.get("questions") or []
            answers = attempt.get("answers") or {}
            totals = {}
            correct = {}
            for pq in (attempt.get("pooledQuestions") or []):
                idx = pq.get("poolIndex")
                if idx is None or idx >= len(pool):
                    continue
                q = pool[idx]
                section = q.get("section") or "General"
                totals[section] = totals.get(section, 0) + 1
                raw = answers.get(str(idx))
                given = raw if isinstance(raw, list) else ([raw] if raw not in (None, "") else [])
                try:
                    given_idx = sorted({int(x) for x in given})
                except (TypeError, ValueError):
                    given_idx = []
                correct_idx = sorted(set(q.get("correct") or []))
                if given_idx and given_idx == correct_idx:
                    correct[section] = correct.get(section, 0) + 1
            for section, n in totals.items():
                pct = round((correct.get(section, 0) / n) * 100, 2) if n else 0.0
                sec_sums[section] = sec_sums.get(section, 0) + pct
                sec_counts[section] = sec_counts.get(section, 0) + 1

        section_rows = []
        for sec in sections_docs:
            name = sec.get("name") or "—"
            n = sec_counts.get(name, 0)
            score = round(sec_sums.get(name, 0) / n, 2) if n else None
            section_rows.append({"name": name, "score": score, "attempts": n})

        interview_score = round(sum(interview_scores) / len(interview_scores), 2) if interview_scores else None
        overall_quiz = round(sum(overall_scores) / len(overall_scores), 2) if overall_scores else None

        return ok({
            "student": {
                "id": str(student["_id"]),
                "name": student.get("fullName") or student.get("name"),
                "rollNumber": student.get("rollNumber"),
                "cohort": student_cohort_label(student),
            },
            "sections": section_rows,
            "manualInterview": {
                "score": interview_score,
                "attempts": len(interview_scores),
            },
            "overallQuizScore": overall_quiz,
            "attemptsCount": len(attempts),
        })

    # ==========================================================
    # OVERALL REPORT — the single report for College Admin.
    # Replaces every separate Baseline/Mid/Final report. For every
    # assessment attempt in this admin's college it shows Assessment
    # Name, Quiz Marks, Interview Marks, Average Marks and Final
    # Overall Marks, all computed live from db.quiz_attempts +
    # db.users.finalEmployabilityScore — nothing hardcoded.
    # ==========================================================
    @bp.route("/reports/overall", methods=["GET"])
    @role_required("college_admin")
    def overall_report():
        college = _admin_college()
        if not college:
            return ok({"report": [], "college": None})

        student_id = (request.args.get("studentId") or "").strip()
        oid = to_object_id(student_id) if student_id else None
        query = {"status": "submitted", "college": college}
        if oid:
            query["studentId"] = oid

        attempts = list(db.quiz_attempts.find(query).sort("submittedAt", -1))
        user_scores = {}
        if oid:
            stu = users.find_one({"_id": oid}, {"finalEmployabilityScore": 1})
            if stu:
                user_scores[oid] = stu.get("finalEmployabilityScore")
        else:
            ids = {a.get("studentId") for a in attempts if a.get("studentId")}
            for s in users.find({"_id": {"$in": list(ids)}}, {"finalEmployabilityScore": 1}):
                user_scores[s["_id"]] = s.get("finalEmployabilityScore")

        rows = []
        for a in attempts:
            overall = a.get("overall") or {}
            rows.append({
                "attemptId": str(a["_id"]),
                "studentId": str(a["studentId"]) if a.get("studentId") else None,
                "studentName": a.get("studentName"),
                "rollNumber": a.get("studentRollNumber"),
                "department": a.get("department"),
                "assessmentId": str(a["quizId"]) if a.get("quizId") else None,
                "assessmentName": a.get("quizTitle") or "Quiz",
                "quizMarks": overall.get("percentage"),
                "quizMarksObtained": overall.get("marksObtained"),
                "quizTotalMarks": overall.get("totalMarks"),
                "interviewMarks": a.get("interviewMarks"),
                "averageMarks": a.get("finalAverage"),
                "finalOverallMarks": user_scores.get(a.get("studentId")) if a.get("studentId") else None,
                "submittedAt": iso_utc(a.get("submittedAt")),
            })
        return ok({"report": rows, "college": college, "studentId": student_id})

    @bp.route("/cohorts/counts", methods=["GET"])
    @role_required("college_admin")
    def get_cohort_counts():
        """Cohort A/B/C/Entry Level counts for this admin's college —
        same shared cohort_counts() helper Trainer and Super Admin use,
        so the numbers can never disagree between roles."""
        college = _admin_college()
        if not college:
            return ok({"cohortCounts": {"A": 0, "B": 0, "C": 0, ENTRY_LEVEL: 0}})
        return ok({"cohortCounts": cohort_counts(db, college=college)})

    @bp.route("/cohorts/students", methods=["GET"])
    @role_required("college_admin")
    def list_cohort_students():
        """Students in one cohort (A/B/C/entry_level), scoped to this
        admin's college — used by the Cohort Management drill-down."""
        college = _admin_college()
        cohort = request.args.get("cohort", "").strip()
        if not college:
            return ok({"students": []})
        query = {"role": "student", "college": college}
        if cohort == ENTRY_LEVEL:
            query["$or"] = [{"cohort": None}, {"cohort": {"$exists": False}}]
        elif cohort in VALID_COHORTS:
            query["cohort"] = cohort
        docs = users.find(query, {"passwordHash": 0}).sort("createdAt", -1)
        return ok({"students": [serialize(d) for d in docs]})

    # ==========================================================
    # COLLEGE ATTENDANCE MODULE — read-only, college-scoped.
    #
    # Attendance is marked exclusively by the Super Admin (db.attendance,
    # one document per student per class date). Every record already
    # carries studentId, date, and present/absent status, and — because
    # it's written from the same roster the College Admin can see — is
    # inherently linked to that student's college/department/cohort at
    # the time of marking. Nothing here recomputes or caches: every call
    # re-reads db.attendance fresh, so the moment the Super Admin marks
    # or edits attendance for a student in this admin's college, the
    # next request to any endpoint below reflects it immediately — no
    # separate sync/update step, no manual refresh required beyond
    # reloading the page's data (which the frontend does on every visit).
    # ==========================================================
    def _college_student_ids(college, department=None):
        query = {"role": "student", "college": college}
        if department:
            query["department"] = department
        return [d["_id"] for d in users.find(query, {"_id": 1})]

    @bp.route("/attendance/departments", methods=["GET"])
    @role_required("college_admin")
    def attendance_departments():
        college = _admin_college()
        if not college:
            return ok({"departments": []})
        depts = sorted({
            d for d in users.find({"role": "student", "college": college}).distinct("department")
            if d
        })
        return ok({"departments": depts})

    @bp.route("/attendance/overview", methods=["GET"])
    @role_required("college_admin")
    def attendance_overview():
        """Whole-college attendance, computed live from db.attendance for
        every student in this admin's college — total classes marked,
        present/absent counts, overall %, and this-month stats. No value
        here is ever stored; it's derived fresh from the same records the
        Student Dashboard and Super Admin see."""
        college = _admin_college()
        if not college:
            return ok({"overview": {"total": 0, "present": 0, "absent": 0, "percentage": 0,
                                     "studentsMarked": 0, "thisMonth": {}}})
        department = (request.args.get("department") or "").strip()
        student_ids = _college_student_ids(college, department or None)
        if not student_ids:
            return ok({"overview": {"total": 0, "present": 0, "absent": 0, "percentage": 0,
                                     "studentsMarked": 0, "thisMonth": {}}})
        records = list(db.attendance.find({"studentId": {"$in": student_ids}}))
        present = sum(1 for r in records if r.get("status") == "present")
        absent = sum(1 for r in records if r.get("status") == "absent")
        total = present + absent
        this_month_key = now().strftime("%Y-%m")
        tm_present = sum(1 for r in records if r.get("status") == "present" and str(r.get("date", ""))[:7] == this_month_key)
        tm_absent = sum(1 for r in records if r.get("status") == "absent" and str(r.get("date", ""))[:7] == this_month_key)
        tm_total = tm_present + tm_absent
        return ok({"overview": {
            "total": total,
            "present": present,
            "absent": absent,
            "percentage": round(present / total * 100, 2) if total else 0,
            "studentsMarked": len({r["studentId"] for r in records}),
            "totalStudents": len(student_ids),
            "thisMonth": {
                "month": this_month_key,
                "total": tm_total,
                "present": tm_present,
                "absent": tm_absent,
                "percentage": round(tm_present / tm_total * 100, 2) if tm_total else 0,
            },
        }})

    @bp.route("/attendance/dates", methods=["GET"])
    @role_required("college_admin")
    def attendance_dates():
        """Distinct class dates (sessions) marked for this college, newest
        first, each with its present/absent split — lets the admin pick
        a specific class/session to inspect without guessing a date."""
        college = _admin_college()
        if not college:
            return ok({"dates": []})
        department = (request.args.get("department") or "").strip()
        student_ids = _college_student_ids(college, department or None)
        if not student_ids:
            return ok({"dates": []})
        pipeline = [
            {"$match": {"studentId": {"$in": student_ids}}},
            {"$group": {
                "_id": "$date",
                "present": {"$sum": {"$cond": [{"$eq": ["$status", "present"]}, 1, 0]}},
                "absent": {"$sum": {"$cond": [{"$eq": ["$status", "absent"]}, 1, 0]}},
            }},
            {"$sort": {"_id": -1}},
            {"$limit": 60},
        ]
        rows = [{"date": r["_id"], "present": r["present"], "absent": r["absent"],
                 "total": r["present"] + r["absent"]} for r in db.attendance.aggregate(pipeline)]
        return ok({"dates": rows})

    @bp.route("/attendance/records", methods=["GET"])
    @role_required("college_admin")
    def college_attendance_records():
        """Every student's status for one class date (session), scoped to
        this admin's college — mirrors the Super Admin's View Records
        screen but restricted to the admin's own college, so the admin
        immediately sees exactly what the Super Admin marked, with no
        separate update/sync process."""
        college = _admin_college()
        if not college:
            return error("No college assigned to this admin.", 403)
        date_str = (request.args.get("date") or "").strip()
        if not date_str:
            return error("A date (YYYY-MM-DD) is required.")
        department = (request.args.get("department") or "").strip()

        query = {"role": "student", "college": college}
        if department:
            query["department"] = department
        students = list(users.find(query).sort("fullName", 1))
        student_ids = [d["_id"] for d in students]

        records = list(db.attendance.find({"date": date_str, "studentId": {"$in": student_ids}}))
        record_by_student = {r["studentId"]: r for r in records}

        rows = []
        for student in students:
            marked = record_by_student.get(student["_id"])
            rows.append({
                "studentId": str(student["_id"]),
                "studentName": student.get("fullName") or "—",
                "rollNumber": student.get("rollNumber") or "—",
                "department": student.get("department") or "—",
                "cohort": student_cohort_label(student),
                "status": marked["status"] if marked else "not_marked",
                "markedBy": (marked.get("markedBy") or {}).get("name") if marked else None,
                "markedAt": iso_utc(marked.get("markedAt")) if marked else None,
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

    @bp.route("/analytics/cohort-distribution", methods=["GET"])
    @role_required("college_admin")
    def cohort_distribution():
        """Department x Cohort matrix for this admin's college, computed
        live from db.users on every call — used by the Cohort Analytics
        and Analytics pages' department/cohort breakdown."""
        college = _admin_college()
        if not college:
            return ok({"departments": [], "matrix": {}})
        pipeline = [
            {"$match": {"role": "student", "college": college}},
            {"$group": {
                "_id": {
                    "department": "$department",
                    "cohort": {"$ifNull": ["$cohort", ENTRY_LEVEL]},
                },
                "count": {"$sum": 1},
            }},
        ]
        matrix = {}
        departments = set()
        for row in users.aggregate(pipeline):
            dept = row["_id"].get("department") or "Unassigned"
            cohort = row["_id"].get("cohort")
            cohort = cohort if cohort in VALID_COHORTS else ENTRY_LEVEL
            departments.add(dept)
            matrix.setdefault(dept, {"A": 0, "B": 0, "C": 0, ENTRY_LEVEL: 0})
            matrix[dept][cohort] = row["count"]
        return ok({"departments": sorted(departments), "matrix": matrix})

    return bp
