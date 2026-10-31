"""
assessments.py — Assessment Question Banks & Templates (read-only)
==================================================================
Registered at /api/assessments. Fixes the frontend's question-banks /
templates URL gap: the Super Admin dashboard already calls
GET /api/assessments/question-banks and GET /api/assessments/templates,
but until now nothing served those URLs, so the Super Admin Assessments
page could never render Question Bank or Template data.

Both endpoints are computed live from real MongoDB collections:

  - /question-banks  -> every section/category group in db.question_bank
                        with live question counts + difficulty split.
  - /templates       -> the platform's configured assessment sections
                        (db.quiz_sections, the same DB-backed list the
                        create-quiz wizard uses), annotated with how many
                        bank questions are actually available per section.

Only role_required users (super_admin / college_admin / trainer) can read
them. There are no write routes here — question-bank writes only ever
happen internally from quiz_module.py's create/update-quiz flow.
"""

from flask import Blueprint

from quiz_common import ok, role_required, iso_utc


_VALID_DIFFICULTIES = ("easy", "medium", "hard")


def init_assessments(db):
    bp = Blueprint("assessments", __name__)

    @bp.route("/question-banks", methods=["GET"])
    @role_required("super_admin", "college_admin", "trainer")
    def list_question_banks():
        """Every question-bank group (section + category) with live counts.
        Each bank: { id, name, category, questions, difficulty } where
        questions is the total active questions and difficulty is the
        { easy, medium, hard } split."""
        pipeline = [
            {"$match": {"active": True}},
            {"$group": {
                "_id": {"section": "$section", "category": "$category"},
                "questions": {"$sum": 1},
                "difficulty": {"$push": "$difficulty"},
            }},
            {"$sort": {"_id.section": 1, "_id.category": 1}},
        ]
        banks = []
        for row in db.question_bank.aggregate(pipeline):
            diff = {"easy": 0, "medium": 0, "hard": 0}
            for d in row.get("difficulty", []):
                key = (d or "").lower()
                if key in diff:
                    diff[key] += 1
            section = row["_id"]["section"] or "Uncategorised"
            category = row["_id"]["category"] or section
            banks.append({
                "id": f"{section}::{category}",
                "name": section,
                "category": category,
                "questions": row.get("questions", 0),
                "difficulty": diff,
            })
        return ok({
            "questionBanks": banks,
            "totalQuestions": sum(b["questions"] for b in banks),
        })

    @bp.route("/templates", methods=["GET"])
    @role_required("super_admin", "college_admin", "trainer")
    def list_templates():
        """The platform's configured assessment sections (db.quiz_sections)
        as assessment templates, each annotated with how many bank
        questions are actually available per section + difficulty. Real
        DB-backed data — never hardcoded sample values."""
        sections = list(db.quiz_sections.find({}).sort("name", 1))
        avail = {}
        for row in db.question_bank.aggregate([
            {"$match": {"active": True}},
            {"$group": {"_id": {"section": "$section", "difficulty": "$difficulty"}, "count": {"$sum": 1}}},
        ]):
            section = row["_id"]["section"]
            difficulty = (row["_id"]["difficulty"] or "").lower()
            if difficulty not in _VALID_DIFFICULTIES:
                continue
            avail.setdefault(section, {"easy": 0, "medium": 0, "hard": 0})[difficulty] = row["count"]

        templates = []
        for s in sections:
            name = s.get("name") or "Untitled Section"
            diff = avail.get(name, {"easy": 0, "medium": 0, "hard": 0})
            templates.append({
                "id": str(s["_id"]),
                "name": name,
                "type": "Question Bank",
                "questions": sum(diff.values()),
                "difficulty": diff,
                "configuredAt": iso_utc(s.get("createdAt")),
            })
        return ok({"templates": templates})

    return bp
