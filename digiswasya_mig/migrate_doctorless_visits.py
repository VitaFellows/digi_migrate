"""
migrate_doctorless_visits.py
============================
Insert-only backfill for appdetails_id=42 (Narsala) healthcases that the normal
visit migration skipped because they have NO doctor (no prescription doctor and
no hc.doctor_id). These are registration-only / incomplete cases.

This script inserts a visit for each such healthcase with
assigned_doctor_user_id = NULL, reusing visit.py's exact field mapping and its
INSERT ... ON CONFLICT (id) DO NOTHING. It performs NO updates and NO deletes.

Target set (computed live):
  appid=42 healthcases whose patient IS migrated (patient.legacy_id present)
  AND whose healthcase legacy_id is NOT already in visit.legacy_id.
Since the normal migration already migrated every doctor-resolvable case, the
remaining set is exactly the doctor-less ones.

Usage:
  python migrate_doctorless_visits.py            # preview only (no writes)
  python migrate_doctorless_visits.py --commit   # insert the visits
"""
import os
import sys
import argparse
import importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sqlalchemy import text
from sqlalchemy.orm import Session
from config.db import get_legacy_engine, get_new_engine
from utils.id_gen import SafeIDGenerator, get_migrated_legacy_ids
from utils.logger import get_logger

# Reuse the exact, already-vetted visit field mapping + INSERT from visit.py
_spec = importlib.util.spec_from_file_location(
    "visit_mod", os.path.join(os.path.dirname(__file__), "migration", "visit", "visit.py"))
V = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(V)

log = get_logger("doctorless_visits")
TARGET_APPDETAILS_ID = int(os.getenv("TARGET_APPDETAILS_ID", "42"))

_FETCH_SQL = text("""
    SELECT DISTINCT ON (hc.id)
        hc.id, hc.problem, hc.status, hc.created_at, hc.lastupdated,
        hc.person_id, hc.appdetails_id
    FROM "HealthCase_healthcase" hc
    WHERE hc.appdetails_id = :a AND hc.person_id IS NOT NULL
    ORDER BY hc.id
""")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--commit", action="store_true", help="Actually insert (otherwise preview only)")
    args = ap.parse_args()

    legacy_engine = get_legacy_engine()
    new_engine = get_new_engine()

    center_code_to_uuid = V.load_center_code_to_uuid(new_engine)
    person_to_patient = V.load_person_id_to_patient_uuid(new_engine)
    narsala_center = center_code_to_uuid.get("DS-TMC-009")
    coordinators = V.load_coordinator_ids(new_engine, narsala_center)
    already = get_migrated_legacy_ids(new_engine, "visit")

    with legacy_engine.connect() as conn:
        rows = conn.execute(_FETCH_SQL, {"a": TARGET_APPDETAILS_ID}).mappings().all()
        comment_rows = conn.execute(V._FETCH_COMMENTS_SQL).fetchall()
    comment_lookup = {int(r[0]): r[1] for r in comment_rows}

    targets = []
    for row in rows:
        legacy_hc = str(row["id"])
        if legacy_hc in already:
            continue  # already migrated -> skip (idempotent)
        patient_uuid = person_to_patient.get(str(row["person_id"]))
        if patient_uuid is None:
            continue  # patient not migrated -> out of scope
        targets.append((row, patient_uuid))

    log.info("Doctor-less Narsala healthcases to insert as visits (NULL doctor): %d", len(targets))
    for row, pid in targets:
        log.info("  healthcase_id=%s person_id=%s patient_id=%s problem=%r status=%s",
                 row["id"], row["person_id"], pid, (row["problem"] or "")[:40], row["status"])

    if not targets:
        log.info("Nothing to insert.")
        return

    if not args.commit:
        log.info("[PREVIEW] No rows written. Re-run with --commit to insert these %d visit(s).", len(targets))
        return

    id_gen = SafeIDGenerator(new_engine, table="visit")
    batch = []
    for row, patient_uuid in targets:
        center_uuid = V.resolve_center_id(row.get("appdetails_id"), center_code_to_uuid)
        coordinator_uuid = V.pick_random_coordinator(coordinators)
        comment_text = comment_lookup.get(int(row["id"]))
        vitals = V.build_vitals_json(comment_text)
        record = V.build_visit_record(
            row, id_gen.next(), patient_uuid, center_uuid, coordinator_uuid,
            None,            # assigned_doctor_user_id = NULL
            vitals, comment_text,
        )
        batch.append(record)

    payload = [V._prepare_row_for_insert(r) for r in batch]
    with new_engine.connect() as c:
        before = c.execute(text("SELECT COUNT(*) FROM visit")).scalar()
    with Session(new_engine) as s:
        s.execute(V.INSERT_SQL, payload)   # ON CONFLICT (id) DO NOTHING — insert only
        s.commit()
    with new_engine.connect() as c:
        after = c.execute(text("SELECT COUNT(*) FROM visit")).scalar()
    log.info("Inserted %d doctor-less visits. visit count %d -> %d (+%d)",
             len(batch), before, after, after - before)


if __name__ == "__main__":
    main()
