"""
============================================================
 question_bank.py — Permanent Question Bank (Create Quiz, Part 5-8)
============================================================
Every question a Trainer authors or bulk-uploads (once it passes
validation) is permanently stored here, in its own collection,
independent of whatever quiz it was first entered for. This is
deliberately a SEPARATE collection from:

  - db.questions      — quiz_common.py's shared bank for the
                         cohort/placement baseline-mid-final exam
                         engine (a different feature entirely).
  - db.quizzes.questions — the inline, per-quiz snapshot stored on
                         each quiz document itself (quiz_module.py).
                         A quiz's own inline list is a frozen COPY
                         taken at publish time; QuestionBank is the
                         evergreen, reusable source pool it can be
                         drawn from.

Collection: db.question_bank — one document per unique question.
Suggested fields (spec, Part 5):
  question_id, section, category, topic, difficulty, question,
  optionA..D, correctAnswer, explanation, uploadedByTrainer,
  uploadedAt, createdBy, status, tags, version, sourceQuiz, active

Populated automatically by quiz_module.py's create/update-quiz flow
(Trainer only, both Manual Entry and Bulk Upload) — never called
directly from the frontend. Consumed by Super Admin's "Question
Bank" quiz-creation mode (Part 6/7/8), which draws a fresh random
set from here at publish time instead of asking for new questions.
"""

import hashlib
import logging

from quiz_common import ok, error, role_required, now, serialize

logger = logging.getLogger("question_bank")

_VALID_DIFFICULTIES = {"easy", "medium", "hard"}


# ------------------------------------------------------------
# Dedup key (Part 5: "avoid duplicates by using a suitable
# uniqueness check — question text + section + options or a
# generated hash").
# ------------------------------------------------------------
def compute_question_hash(section, text, options):
    """Stable fingerprint for a question: section + normalized question
    text + normalized, ORDER-PRESERVING option list. Two rows that only
    differ in marks/explanation/difficulty still count as the exact same
    question for dedup purposes, per spec — difficulty is intentionally
    NOT part of the hash so re-uploading the same question under a
    different difficulty is still caught as a duplicate rather than
    silently creating a second bank entry.
    """
    norm_section = (section or "").strip().lower()
    norm_text = " ".join((text or "").strip().lower().split())
    norm_options = "|".join(" ".join((o or "").strip().lower().split()) for o in (options or []))
    raw = f"{norm_section}::{norm_text}::{norm_options}"
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def sync_questions_to_bank(db, actor, questions, source_quiz_title=None, source_quiz_id=None):
    """Insert every one of `questions` (already-validated, normalized
    question dicts in quiz_module's internal shape) into db.question_bank,
    skipping any that already exist (by hash). Called once per successful
    Trainer create/update of a Manual Entry or Bulk Upload quiz — never
    for Super Admin's Question Bank mode (those questions are already IN
    the bank; re-inserting them would just be a no-op duplicate anyway,
    but it's skipped explicitly to avoid the extra DB round trip).

    Never raises — a bank-sync failure must never block the quiz itself
    from being saved. Any error is logged and swallowed.

    Returns (inserted_count, skipped_duplicate_count).
    """
    if not questions:
        return 0, 0
    try:
        bank = db.question_bank
        inserted, skipped = 0, 0
        for q in questions:
            options = q.get("options") or []
            section = q.get("section")
            text = q.get("text")
            difficulty = (q.get("difficulty") or "").strip().capitalize()
            if not text or not section or difficulty.lower() not in _VALID_DIFFICULTIES:
                # Shouldn't happen — validate_and_normalize() already
                # enforced this — but never let a malformed record either
                # crash the sync or silently corrupt the bank.
                logger.warning(
                    "sync_questions_to_bank: skipping malformed question (section=%r, difficulty=%r, has_text=%s)",
                    section, difficulty, bool(text),
                )
                continue
            qhash = compute_question_hash(section, text, options)
            if bank.find_one({"questionHash": qhash}):
                skipped += 1
                continue
            doc = {
                "questionHash": qhash,
                "section": section,
                "category": q.get("category") or section,
                "topic": q.get("topic") or "",
                "difficulty": difficulty,
                "question": text,
                "optionA": options[0] if len(options) > 0 else "",
                "optionB": options[1] if len(options) > 1 else "",
                "optionC": options[2] if len(options) > 2 else "",
                "optionD": options[3] if len(options) > 3 else "",
                "correctAnswer": q.get("correct") or [],
                "type": q.get("type") or "single_choice",
                "explanation": q.get("explanation") or "",
                "marks": q.get("marks") or 1,
                "uploadedByTrainer": actor.get("name"),
                "uploadedAt": now(),
                "createdBy": actor,
                "status": "active",
                "tags": [],
                "version": 1,
                "sourceQuiz": source_quiz_title,
                "sourceQuizId": str(source_quiz_id) if source_quiz_id else None,
                "active": True,
            }
            bank.insert_one(doc)
            inserted += 1
        logger.info(
            "sync_questions_to_bank: quiz=%r inserted=%s skipped_duplicate=%s",
            source_quiz_title, inserted, skipped,
        )
        return inserted, skipped
    except Exception:
        logger.exception("sync_questions_to_bank: failed to sync questions for quiz %r — quiz save is unaffected.", source_quiz_title)
        return 0, 0


# ------------------------------------------------------------
# Live availability — always computed fresh from the DB, never cached,
# so the Super Admin's Question Bank wizard step always reflects exactly
# what's really in the bank right now (same "never trust a stale count"
# philosophy the rest of quiz_module.py already follows for Available).
# ------------------------------------------------------------
def bank_availability_summary(db, active_only=True):
    """Returns {section: {"hard": n, "medium": n, "easy": n}} across the
    whole bank, one entry per section that has at least one question."""
    query = {"active": True} if active_only else {}
    pipeline = [
        {"$match": query},
        {"$group": {"_id": {"section": "$section", "difficulty": "$difficulty"}, "count": {"$sum": 1}}},
    ]
    summary = {}
    for row in db.question_bank.aggregate(pipeline):
        section = row["_id"]["section"]
        difficulty = (row["_id"]["difficulty"] or "").lower()
        if difficulty not in _VALID_DIFFICULTIES:
            continue
        bucket = summary.setdefault(section, {"hard": 0, "medium": 0, "easy": 0})
        bucket[difficulty] = row["count"]
    return summary


def validate_question_bank_config(db, raw_config):
    """Part 8 — validation before publishing a Question-Bank-sourced quiz.

    Unlike Manual Entry / Bulk Upload's _validate_section_config (which
    requires Available to EXACTLY match an authored pool), Question Bank
    mode only requires Available (the live bank count) to be >= what the
    admin wants to Display — the bank is a big shared reusable pool, not
    a per-quiz authored set. Available is always taken from the live bank
    count here, never trusted from the client.

    Returns (clean_config, section_distribution, difficulty_distribution,
             questions_available, questions_displayed, errors), same
    shape as quiz_module._validate_section_config so the rest of
    validate_and_normalize() can treat both modes identically.
    """
    errors = []
    if not isinstance(raw_config, dict) or not raw_config:
        return {}, {}, {}, 0, 0, ["Section Distribution is required — configure at least one section."]

    def _int(v):
        try:
            n = int(v)
        except (TypeError, ValueError):
            n = 0
        return max(0, n)

    live = bank_availability_summary(db)
    clean_config = {}
    section_distribution = {}
    difficulty_totals_display = {"easy": 0, "medium": 0, "hard": 0}
    total_available = 0
    total_display = 0

    for raw_name, cfg in raw_config.items():
        name = str(raw_name).strip()
        if not name or not isinstance(cfg, dict):
            continue
        bank_bucket = live.get(name, {"hard": 0, "medium": 0, "easy": 0})
        h_d, m_d, e_d = _int(cfg.get("hardDisplay")), _int(cfg.get("mediumDisplay")), _int(cfg.get("easyDisplay"))
        h_a, m_a, e_a = bank_bucket["hard"], bank_bucket["medium"], bank_bucket["easy"]

        for label, avail, disp in (("Hard", h_a, h_d), ("Medium", m_a, m_d), ("Easy", e_a, e_d)):
            if disp > avail:
                errors.append(
                    f'"{name}" section — only {avail} {label} question(s) available in the Question Bank. '
                    f"Requested: {disp}. Please upload more {label} questions, or reduce the Display count."
                )

        avail_total = h_a + m_a + e_a
        disp_total = h_d + m_d + e_d
        clean_config[name] = {
            "hardAvailable": h_a, "mediumAvailable": m_a, "easyAvailable": e_a,
            "hardDisplay": h_d, "mediumDisplay": m_d, "easyDisplay": e_d,
            "questionsAvailable": avail_total, "questionsToDisplay": disp_total,
        }
        total_available += avail_total
        total_display += disp_total
        difficulty_totals_display["hard"] += h_d
        difficulty_totals_display["medium"] += m_d
        difficulty_totals_display["easy"] += e_d
        if disp_total > 0:
            section_distribution[name] = disp_total

    if total_display == 0:
        errors.append("Configure at least one section with Questions to Display.")

    difficulty_distribution = {k: v for k, v in difficulty_totals_display.items() if v}
    return clean_config, section_distribution, difficulty_distribution, total_available, total_display, errors


def draw_questions_from_bank(db, clean_config):
    """Part 7 — fresh, secure server-side random draw from the bank for
    ONE quiz, at the moment it's published. Every publish call re-runs
    this, so two quizzes (or the same quiz re-published later) can and
    will get different random subsets even with an identical config.
    Duplicate questions WITHIN this one draw are impossible (selection
    is without replacement, per section+difficulty cell); the SAME bank
    question can, by design, appear in other quizzes later (Part 10:
    "allow reuse across different quizzes").

    Returns a list of question dicts already shaped like the
    `questions` array quiz_module.py stores on a quiz document (same
    keys as clean_questions in validate_and_normalize), each carrying
    questionBankId for traceability back to its bank source.
    """
    import random

    chosen = []
    for section, cfg in clean_config.items():
        for level, count_key in (("hard", "hardDisplay"), ("medium", "mediumDisplay"), ("easy", "easyDisplay")):
            count = cfg.get(count_key) or 0
            if count <= 0:
                continue
            candidates = list(db.question_bank.find({
                "section": section, "difficulty": level.capitalize(), "active": True,
            }))
            random.shuffle(candidates)
            for doc in candidates[:count]:
                options = [doc.get("optionA", ""), doc.get("optionB", ""), doc.get("optionC", ""), doc.get("optionD", "")]
                chosen.append({
                    "text": doc.get("question", ""),
                    "options": options,
                    "correct": doc.get("correctAnswer") or [],
                    "type": doc.get("type") or "single_choice",
                    "section": section,
                    "difficulty": level.capitalize(),
                    "marks": doc.get("marks") or 1,
                    "explanation": doc.get("explanation") or "",
                    "questionBankId": str(doc["_id"]),
                })
    random.shuffle(chosen)
    logger.info("draw_questions_from_bank: drew %s question(s) across %s section(s)", len(chosen), len(clean_config))
    return chosen


# ------------------------------------------------------------
# Blueprint — read-only endpoints the Question Bank wizard step needs.
# Mounted at /api/admin only (Question Bank creation mode is Super Admin
# only, per spec Part 4/6); also mounted at /api/trainer read-only so a
# Trainer can see what's already in the shared bank if that's ever
# surfaced in the UI later, without granting any write access.
# ------------------------------------------------------------
def init_question_bank(db, scope):
    from flask import Blueprint
    bp = Blueprint(f"question_bank_{scope}", __name__)
    role = "trainer" if scope == "trainer" else "super_admin"

    @bp.route("/question-bank/summary", methods=["GET"])
    @role_required("trainer", "super_admin")
    def summary():
        return ok({"summary": bank_availability_summary(db)})

    @bp.route("/question-bank", methods=["GET"])
    @role_required("trainer", "super_admin")
    def list_bank():
        from flask import request
        section = request.args.get("section")
        difficulty = request.args.get("difficulty")
        query = {"active": True}
        if section:
            query["section"] = section
        if difficulty:
            query["difficulty"] = difficulty.capitalize()
        limit = min(int(request.args.get("limit", 200) or 200), 1000)
        docs = list(db.question_bank.find(query).sort("uploadedAt", -1).limit(limit))
        return ok({"questions": [serialize(d) for d in docs], "count": db.question_bank.count_documents(query)})

    return bp
