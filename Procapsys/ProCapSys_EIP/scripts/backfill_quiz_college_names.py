"""
One-time backfill: fix quiz documents whose collegeNames array actually
holds college ids (the bug fixed in quiz_module.validate_and_normalize —
see the long comment there for the full root-cause explanation).

This is OPTIONAL. student.py and quiz_module.py already self-heal every
affected quiz automatically the first time it's read (see
quiz_module.normalize_quiz_college_names), so nothing breaks if you never
run this. Run it if you'd rather proactively clean up every affected
record in one pass (e.g. for a clean audit log / clean data export)
instead of relying on lazy per-request healing.

Usage:
    python scripts/backfill_quiz_college_names.py            # dry run, prints what would change
    python scripts/backfill_quiz_college_names.py --apply    # actually writes the fix

Safe to run multiple times — already-correct documents are left untouched.
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pymongo import MongoClient  # noqa: E402
from quiz_module import normalize_quiz_college_names, to_object_id  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--apply", action="store_true", help="Write the fix (default is dry-run).")
    parser.add_argument("--mongo-uri", default=os.environ.get("MONGO_URI", "mongodb://localhost:27017"))
    parser.add_argument("--db-name", default=os.environ.get("MONGO_DB_NAME", "eip"))
    args = parser.parse_args()

    client = MongoClient(args.mongo_uri)
    db = client[args.db_name]

    affected = []
    for quiz in db.quizzes.find({"collegeNames": {"$exists": True, "$ne": []}}):
        names = quiz.get("collegeNames") or []
        if names and all(to_object_id(n) for n in names):
            affected.append(quiz)

    if not affected:
        print("No affected quizzes found — nothing to do.")
        return

    print(f"Found {len(affected)} quiz(zes) with ids stored as collegeNames:")
    for quiz in affected:
        print(f"  - {quiz.get('title', 'Untitled')} ({quiz['_id']}): {quiz.get('collegeNames')}")

    if not args.apply:
        print("\nDry run only — re-run with --apply to write the fix.")
        return

    for quiz in affected:
        fixed = normalize_quiz_college_names(db, quiz)
        print(f"  fixed {quiz['_id']} -> {fixed.get('collegeNames')}")
    print(f"\nDone — {len(affected)} quiz(zes) updated.")


if __name__ == "__main__":
    main()
