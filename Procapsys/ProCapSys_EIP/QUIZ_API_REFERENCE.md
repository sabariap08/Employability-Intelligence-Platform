# Quiz Management & Dashboard API Reference

New modules added this session (all follow login.py's `init_*(db=db)` factory
pattern and are registered in `app.py`):

- **`quiz_common.py`** — shared engine: cohort rules, placement rules,
  Excel parsing, random question selection, scoring. Nothing role-specific
  lives here; it's the single source of truth every other module calls into.
- **`superadmin.py`** → `/api/admin/*`
- **`trainer.py`** → `/api/trainer/*`
- **`student.py`** → `/api/student/*`

All routes require `Authorization: Bearer <JWT>` (same auth as the existing
`/api/auth` routes) and are role-gated with `quiz_common.role_required(...)`.

---

## Cohort generation (2-stage — see "ADDITIONAL UPDATE" spec)

A student's `cohort` field stays `null` (Entry Level) until **both**:
1. a **baseline** assessment has been submitted and scored, and
2. a manual interview has been scored,

at which point `quiz_common.check_and_generate_cohort()` combines the two
scores into a **Final Employability Score** using weights from
`db.placement_rules`, maps it to a cohort using the ranges in
`db.placement_rules`, and writes `cohort`, `finalEmployabilityScore`,
`cohortAssignedAt` onto the student's user doc.

**No score range or weight is hardcoded anywhere in the code.** Every
calculation re-reads `db.placement_rules` (seeded with a sane default only
the very first time the collection is empty). Super Admin edits via
`PUT /api/admin/placement-rules` take effect immediately for every future
calculation, and — if `recalculate: true` is sent — retroactively for every
student who already has both scores recorded.

## New collections

| Collection             | Purpose |
|-------------------------|---------|
| `questions`             | Question bank (from master Excel upload) |
| `assessments`           | Assessment definitions (cohort target incl. Entry Level, per-section counts, schedule) |
| `assessment_attempts`   | One doc per student attempt: questions served, answers, scores |
| `manual_interviews`     | One doc per scheduled/completed interview, incl. score |
| `placement_rules`       | Single active doc: cohort score ranges + assessment/interview weights |

`users` gained: `cohort`, `baselineAssessmentScore`, `interviewScore`,
`finalEmployabilityScore`, `cohortAssignedAt` (all `null` at signup).

---

## Super Admin — `/api/admin`

| Method & Path | Purpose |
|---|---|
| `POST /questions/upload` | Upload ONE master `.xlsx` (multipart field `file`). Every sheet = a section. Parses, validates, stores every question in `db.questions`. |
| `GET /questions/stats` | Live per-section question counts. |
| `POST /assessments` | Create assessment. `cohortTarget`: `"A" \| "B" \| "C" \| "all" \| "entry_level"`. `sectionCounts`: `{"Communication": 10, ...}`. |
| `GET /assessments` | List all (feeds the Results-page dropdown). |
| `GET/DELETE /assessments/<id>` | Fetch / remove one. |
| `GET /cohorts/counts` | `{A, B, C, entry_level}` counts, DB-computed. |
| `GET /cohorts/students?cohort=` | Student list for a cohort (incl. `entry_level`). |
| `GET /quiz-management/summary` | Live counts: questions, assessments, attempts, cohort breakdown. |
| `GET /quiz-responses?limit=` | Recent submitted attempts with student + assessment names joined in. |
| `GET /manual-interview` | All manual interviews, every college. |
| `GET /placement-rules` | Current placement rules (cohort ranges + weights). |
| `PUT /placement-rules` | Update ranges/weights; optional `recalculate: true`. |
| `GET /dashboard/charts?cohort=` | Skill Radar / Score-by-Category / Overall / Trend, computed live from `assessment_attempts`. |

## Trainer — `/api/trainer` (all scoped to the trainer's own `college`)

Same assessment + cohort + chart shape as Super Admin, plus:

| Method & Path | Purpose |
|---|---|
| `GET /dashboard/insights` | Dashboard widgets computed live from `assessment_attempts` + `workshop_sessions`: `readinessTrend` (weekly avg), `departmentRanking` (top 5, best-per-student avg), `skillRadar` (cohort axes), `improvementRing` (avg consecutive-assessment delta), `upcomingSessions` (this trainer's assigned, not-yet-started sessions). |
| `POST /manual-interview/<studentId>/schedule` | Step 1: schedule/reschedule an interview. |
| `POST /manual-interview/<interviewId>/score` | Step 2: record score 0–100 → feeds cohort generation. |

## Student — `/api/student`

| Method & Path | Purpose |
|---|---|
| `GET /assessments` | `{upcoming, pending, attempted, completed}`, computed from cohort (incl. Entry Level) + schedule + availability window. |
| `POST /assessments/<id>/start` | Generates a fresh random question set (per admin's `sectionCounts`), creates/resumes an attempt. |
| `POST /attempts/<attemptId>/submit` | `{answers: {questionId: option}}` → stores responses, scores, records baseline score toward cohort generation. |
| `GET /cohort-status` | Current cohort (or `entry_level`), both component scores, final score. |
| `GET /results` | Every completed assessment (dropdown source). |
| `GET /results/<assessmentId>` | Full result: section scores, overall, cohort. |
| `GET /dashboard/charts` | This student's Skill Radar / Category % / Overall / Trend. |
| `GET /results/<assessmentId>/report` | PDF report (falls back to structured JSON if `reportlab` isn't installed) — all real DB values. |

---

## Not yet done (next step)

This session focused entirely on the **backend** (fully implemented and
smoke-tested end-to-end, including the two-stage cohort flow and dynamic
placement-rule recalculation). The existing frontend HTML files
(`student.html`, `trainer.html`, `super_admin.html`) were **not modified** —
they still contain their original hardcoded sample data and are not yet
wired to these new endpoints. Wiring each dashboard's JS to call these APIs
(and removing the hardcoded arrays/objects) is the remaining work, and can
proceed file-by-file in the priority order given: Student → Trainer →
Super Admin, without any backend changes.
