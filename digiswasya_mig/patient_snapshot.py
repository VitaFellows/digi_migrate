"""
Patient migration snapshot / verification tool.

Run BEFORE and AFTER the patient migration. It records:
  - destination patient counts (total, native vs migrated, this-source)
  - a fingerprint of the NATIVE (non-migrated) rows that must remain untouched
  - source eligibility count (appdetails_id = 42)

Usage:
  python patient_snapshot.py before   # snapshot to patient_snapshot_before.json
  python patient_snapshot.py after    # snapshot + diff vs the 'before' file
"""
import sys
import json
import os
import hashlib
from sqlalchemy import text
from config.db import get_legacy_engine, get_new_engine

LEGACY_SOURCE = "dgiswasthya_database"
TARGET_APPDETAILS_ID = 42
HERE = os.path.dirname(__file__)


def dest_snapshot():
    eng = get_new_engine()
    with eng.connect() as c:
        total = c.execute(text("SELECT COUNT(*) FROM patient")).scalar()
        native = c.execute(text(
            "SELECT COUNT(*) FROM patient WHERE legacy_id IS NULL")).scalar()
        migrated = c.execute(text(
            "SELECT COUNT(*) FROM patient WHERE legacy_id IS NOT NULL")).scalar()
        this_source = c.execute(
            text("SELECT COUNT(*) FROM patient WHERE legacy_source = :s"),
            {"s": LEGACY_SOURCE}).scalar()
        # Fingerprint of native rows: a checksum + the set of native ids.
        # If ANY native row changes/disappears, this fingerprint changes.
        native_ids = [str(r[0]) for r in c.execute(text(
            "SELECT id::text FROM patient WHERE legacy_id IS NULL ORDER BY id"
            )).fetchall()]
        max_migrated_at = c.execute(text(
            "SELECT MAX(migrated_at)::text FROM patient WHERE legacy_id IS NOT NULL"
            )).scalar()
    return {
        "total": total,
        "native_unmigrated": native,
        "migrated": migrated,
        "this_source": this_source,
        "max_migrated_at": max_migrated_at,
        "native_id_count": len(native_ids),
        "native_ids_sample": native_ids[:20],
        "native_ids_hash": hashlib.sha256("|".join(native_ids).encode()).hexdigest(),
    }


def source_eligible():
    eng = get_legacy_engine()
    with eng.connect() as c:
        total = c.execute(text(
            'SELECT COUNT(*) FROM "Person_persondetails"')).scalar()
        eligible = c.execute(text("""
            SELECT COUNT(DISTINCT person_id)
            FROM "HealthCase_healthcase"
            WHERE appdetails_id = :a AND person_id IS NOT NULL
        """), {"a": TARGET_APPDETAILS_ID}).scalar()
    return {"source_total_persons": total, "source_eligible_appid42": eligible}


def main():
    phase = sys.argv[1] if len(sys.argv) > 1 else "before"
    snap = {"phase": phase, "destination": dest_snapshot(), "source": source_eligible()}
    out = os.path.join(HERE, f"patient_snapshot_{phase}.json")
    with open(out, "w") as f:
        json.dump(snap, f, indent=2)

    d, s = snap["destination"], snap["source"]
    print(f"\n=== PATIENT SNAPSHOT [{phase}] ===")
    print(f"SOURCE  total persons         : {s['source_total_persons']}")
    print(f"SOURCE  eligible (appid=42)    : {s['source_eligible_appid42']}")
    print(f"DEST    total patients         : {d['total']}")
    print(f"DEST    native (legacy_id NULL): {d['native_unmigrated']}   <-- must NOT change")
    print(f"DEST    migrated (legacy_id set): {d['migrated']}")
    print(f"DEST    this source            : {d['this_source']}")
    print(f"saved -> {out}")

    if phase == "after":
        bpath = os.path.join(HERE, "patient_snapshot_before.json")
        if not os.path.exists(bpath):
            print("\n[!] No 'before' snapshot found — cannot diff.")
            return
        before = json.load(open(bpath))
        bd = before["destination"]
        print("\n=== DIFF (after - before) ===")
        print(f"total     : {bd['total']} -> {d['total']}  (+{d['total']-bd['total']})")
        print(f"migrated  : {bd['migrated']} -> {d['migrated']}  (+{d['migrated']-bd['migrated']})")
        native_ok = (bd['native_unmigrated'] == d['native_unmigrated']
                     and bd['native_ids_hash'] == d['native_ids_hash'])
        print(f"native    : {bd['native_unmigrated']} -> {d['native_unmigrated']}  "
              f"(delta {d['native_unmigrated']-bd['native_unmigrated']})")
        print(f"\nNATIVE ROWS UNTOUCHED: {'YES ✓' if native_ok else 'NO ✗  <-- INVESTIGATE'}")


if __name__ == "__main__":
    main()
