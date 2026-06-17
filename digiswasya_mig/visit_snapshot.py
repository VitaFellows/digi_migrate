"""
Visit migration snapshot / verification tool.

Run BEFORE and AFTER the visit migration. Records destination visit counts,
a deterministic fingerprint of NATIVE (non-migrated) rows that must stay
untouched, and source eligibility. Also checks every migrated visit links to
a migrated patient (never a native patient).

Usage:
  python visit_snapshot.py before
  python visit_snapshot.py after
"""
import sys
import json
import os
import hashlib
from sqlalchemy import text
from config.db import get_legacy_engine, get_new_engine

LEGACY_SOURCE = "digiswasthya_database"   # NB: visit.py uses this spelling
TARGET_APPDETAILS_ID = 42
HERE = os.path.dirname(__file__)


def dest_snapshot():
    eng = get_new_engine()
    with eng.connect() as c:
        total = c.execute(text("SELECT COUNT(*) FROM visit")).scalar()
        native = c.execute(text(
            "SELECT COUNT(*) FROM visit WHERE legacy_id IS NULL")).scalar()
        migrated = c.execute(text(
            "SELECT COUNT(*) FROM visit WHERE legacy_id IS NOT NULL")).scalar()
        this_source = c.execute(
            text("SELECT COUNT(*) FROM visit WHERE legacy_source = :s"),
            {"s": LEGACY_SOURCE}).scalar()
        native_ids = [str(r[0]) for r in c.execute(text(
            "SELECT id::text FROM visit WHERE legacy_id IS NULL ORDER BY id"
            )).fetchall()]
        # Integrity guard: every migrated visit MUST point at a migrated patient.
        migrated_visits_on_native_patient = c.execute(text("""
            SELECT COUNT(*) FROM visit v
            JOIN patient p ON p.id = v.patient_id
            WHERE v.legacy_id IS NOT NULL AND p.legacy_id IS NULL
        """)).scalar()
    return {
        "total": total,
        "native_unmigrated": native,
        "migrated": migrated,
        "this_source": this_source,
        "native_id_count": len(native_ids),
        "native_ids_sample": native_ids[:20],
        "native_ids_hash": hashlib.sha256("|".join(native_ids).encode()).hexdigest(),
        "migrated_visits_on_native_patient": migrated_visits_on_native_patient,
    }


def source_eligible():
    eng = get_legacy_engine()
    with eng.connect() as c:
        total = c.execute(text(
            'SELECT COUNT(*) FROM "HealthCase_healthcase"')).scalar()
        eligible = c.execute(text("""
            SELECT COUNT(DISTINCT id)
            FROM "HealthCase_healthcase"
            WHERE appdetails_id = :a AND person_id IS NOT NULL
        """), {"a": TARGET_APPDETAILS_ID}).scalar()
    return {"source_total_healthcases": total, "source_eligible_appid42": eligible}


def main():
    phase = sys.argv[1] if len(sys.argv) > 1 else "before"
    snap = {"phase": phase, "destination": dest_snapshot(), "source": source_eligible()}
    out = os.path.join(HERE, f"visit_snapshot_{phase}.json")
    with open(out, "w") as f:
        json.dump(snap, f, indent=2)

    d, s = snap["destination"], snap["source"]
    print(f"\n=== VISIT SNAPSHOT [{phase}] ===")
    print(f"SOURCE  total healthcases      : {s['source_total_healthcases']}")
    print(f"SOURCE  eligible (appid=42)    : {s['source_eligible_appid42']}")
    print(f"DEST    total visits           : {d['total']}")
    print(f"DEST    native (legacy_id NULL): {d['native_unmigrated']}   <-- must NOT change")
    print(f"DEST    migrated (legacy_id set): {d['migrated']}")
    print(f"DEST    this source            : {d['this_source']}")
    print(f"DEST    migrated visits on NATIVE patient: {d['migrated_visits_on_native_patient']}  (must be 0)")
    print(f"saved -> {out}")

    if phase == "after":
        bpath = os.path.join(HERE, "visit_snapshot_before.json")
        if not os.path.exists(bpath):
            print("\n[!] No 'before' snapshot found — cannot diff.")
            return
        before = json.load(open(bpath))
        bd = before["destination"]
        print("\n=== DIFF (after - before) ===")
        print(f"total     : {bd['total']} -> {d['total']}  (+{d['total']-bd['total']})")
        print(f"migrated  : {bd['migrated']} -> {d['migrated']}  (+{d['migrated']-bd['migrated']})")
        print(f"native    : {bd['native_unmigrated']} -> {d['native_unmigrated']}  "
              f"(delta {d['native_unmigrated']-bd['native_unmigrated']})")
        native_ok = (bd['native_unmigrated'] == d['native_unmigrated']
                     and bd['native_ids_hash'] == d['native_ids_hash'])
        no_native_links = d['migrated_visits_on_native_patient'] == 0
        print(f"\nNATIVE ROWS UNTOUCHED        : {'YES ✓' if native_ok else 'NO ✗  <-- INVESTIGATE'}")
        print(f"NO VISITS ON NATIVE PATIENTS : {'YES ✓' if no_native_links else 'NO ✗  <-- INVESTIGATE'}")


if __name__ == "__main__":
    main()
