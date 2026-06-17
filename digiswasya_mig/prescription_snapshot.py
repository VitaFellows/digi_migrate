"""
Prescription migration snapshot / verification tool.

Run BEFORE and AFTER. Records destination prescription counts, a deterministic
fingerprint of NATIVE rows that must stay untouched, source eligibility, and an
integrity guard that every migrated prescription links to a MIGRATED visit
(never a native visit).

Usage:
  python prescription_snapshot.py before
  python prescription_snapshot.py after
"""
import sys
import json
import os
import hashlib
from sqlalchemy import text
from config.db import get_legacy_engine, get_new_engine

LEGACY_SOURCE = "digiswasthya_database"
TARGET_APPDETAILS_ID = 42
HERE = os.path.dirname(__file__)


def dest_snapshot():
    eng = get_new_engine()
    with eng.connect() as c:
        total = c.execute(text("SELECT COUNT(*) FROM prescription")).scalar()
        native = c.execute(text(
            "SELECT COUNT(*) FROM prescription WHERE legacy_id IS NULL")).scalar()
        migrated = c.execute(text(
            "SELECT COUNT(*) FROM prescription WHERE legacy_id IS NOT NULL")).scalar()
        this_source = c.execute(
            text("SELECT COUNT(*) FROM prescription WHERE legacy_source = :s"),
            {"s": LEGACY_SOURCE}).scalar()
        native_ids = [str(r[0]) for r in c.execute(text(
            "SELECT id::text FROM prescription WHERE legacy_id IS NULL ORDER BY id"
            )).fetchall()]
        # Integrity guard: every migrated prescription MUST point at a migrated visit.
        migrated_presc_on_native_visit = c.execute(text("""
            SELECT COUNT(*) FROM prescription pr
            JOIN visit v ON v.id = pr.visit_id
            WHERE pr.legacy_id IS NOT NULL AND v.legacy_id IS NULL
        """)).scalar()
    return {
        "total": total,
        "native_unmigrated": native,
        "migrated": migrated,
        "this_source": this_source,
        "native_id_count": len(native_ids),
        "native_ids_sample": native_ids[:20],
        "native_ids_hash": hashlib.sha256("|".join(native_ids).encode()).hexdigest(),
        "migrated_presc_on_native_visit": migrated_presc_on_native_visit,
    }


def source_eligible():
    eng = get_legacy_engine()
    with eng.connect() as c:
        # DISTINCT ON (healthcase_id): one prescription per eligible healthcase
        eligible = c.execute(text("""
            SELECT COUNT(DISTINCT p.healthcase_id)
            FROM "BodyVitals_prescription" p
            WHERE p.healthcase_id IS NOT NULL AND p.person_id IS NOT NULL
              AND p.healthcase_id IN (
                  SELECT id FROM "HealthCase_healthcase" WHERE appdetails_id = :a
              )
        """), {"a": TARGET_APPDETAILS_ID}).scalar()
    return {"source_eligible_distinct_healthcases": eligible}


def main():
    phase = sys.argv[1] if len(sys.argv) > 1 else "before"
    snap = {"phase": phase, "destination": dest_snapshot(), "source": source_eligible()}
    out = os.path.join(HERE, f"prescription_snapshot_{phase}.json")
    with open(out, "w") as f:
        json.dump(snap, f, indent=2)

    d, s = snap["destination"], snap["source"]
    print(f"\n=== PRESCRIPTION SNAPSHOT [{phase}] ===")
    print(f"SOURCE  eligible (1/healthcase): {s['source_eligible_distinct_healthcases']}")
    print(f"DEST    total prescriptions    : {d['total']}")
    print(f"DEST    native (legacy_id NULL): {d['native_unmigrated']}   <-- must NOT change")
    print(f"DEST    migrated (legacy_id set): {d['migrated']}")
    print(f"DEST    this source            : {d['this_source']}")
    print(f"DEST    migrated presc on NATIVE visit: {d['migrated_presc_on_native_visit']}  (must be 0)")
    print(f"saved -> {out}")

    if phase == "after":
        bpath = os.path.join(HERE, "prescription_snapshot_before.json")
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
        no_native_links = d['migrated_presc_on_native_visit'] == 0
        print(f"\nNATIVE ROWS UNTOUCHED        : {'YES ✓' if native_ok else 'NO ✗  <-- INVESTIGATE'}")
        print(f"NO PRESC ON NATIVE VISITS    : {'YES ✓' if no_native_links else 'NO ✗  <-- INVESTIGATE'}")


if __name__ == "__main__":
    main()
