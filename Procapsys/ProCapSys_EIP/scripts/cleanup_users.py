"""
============================================================
 scripts/cleanup_users.py — Phase 1 clean-database step
============================================================
Deletes every Student, Trainer, and College Admin record from
db.users so testing starts from a clean slate. Super Admin is NOT
a database row (it's the hardcoded account in login.py / .env), so
there is nothing to "keep" for it — this script simply never
touches anything outside db.users.

Also clears db.otp_verifications so no stale OTP tokens survive
the wipe.

Run from the eip-backend/ directory:

    python scripts/cleanup_users.py            # interactive confirmation
    python scripts/cleanup_users.py --yes       # skip confirmation (CI/scripts)

Uses the same MONGO_URI / DB_NAME as the running app (reads .env).
============================================================
"""

import os
import sys

from dotenv import load_dotenv
from pymongo import MongoClient

load_dotenv()

MONGO_URI = os.getenv("MONGO_URI", "mongodb://localhost:27017")
DB_NAME = os.getenv("DB_NAME", "eip")


def main():
    skip_confirm = "--yes" in sys.argv or "-y" in sys.argv

    client = MongoClient(MONGO_URI, serverSelectionTimeoutMS=5000)
    db = client[DB_NAME]

    to_delete = db.users.count_documents({"role": {"$in": ["student", "trainer", "college_admin"]}})
    print(f"[cleanup] Database: '{DB_NAME}' at {MONGO_URI}")
    print(f"[cleanup] Found {to_delete} student/trainer/college_admin record(s) to delete.")

    if to_delete == 0:
        print("[cleanup] Nothing to do. Database is already clean.")
    else:
        if not skip_confirm:
            answer = input("Proceed with deletion? This cannot be undone. [y/N]: ").strip().lower()
            if answer != "y":
                print("[cleanup] Aborted. No changes made.")
                return

        result = db.users.delete_many({"role": {"$in": ["student", "trainer", "college_admin"]}})
        print(f"[cleanup] Deleted {result.deleted_count} user record(s).")

    otp_result = db.otp_verifications.delete_many({})
    print(f"[cleanup] Cleared {otp_result.deleted_count} stale OTP record(s).")

    remaining_admins = db.users.count_documents({"role": "super_admin"})
    print(f"[cleanup] Super Admin is a hardcoded account (not a DB row) — "
          f"{remaining_admins} super_admin row(s) left untouched by design.")
    print("[cleanup] Done. Database is clean.")


if __name__ == "__main__":
    main()
