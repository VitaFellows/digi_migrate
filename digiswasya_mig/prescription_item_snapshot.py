"""
PrescriptionItem migration snapshot / verification tool — LIVE DB aware.

The destination is a LIVE database: the app inserts prescriptionitem rows
continuously. So we do NOT assert "row count is unchanged". Instead we prove
the stronger, concurrency-safe property:

  Every row that existed at BASELINE is still present AND byte-for-byte
  unchanged after our migration (no modify, no delete).

New rows that appear are expected; we classify them as:
  - OURS  : item links to a migrated prescription (legacy_source=digiswasthya_database)
  - LIVE  : item links to a native prescription (ordinary app activity)

Usage:
  python prescription_item_snapshot.py before
  python prescription_item_snapshot.py after
"""
import sys
import json
import os
from sqlalchemy import text
from config.db import get_new_engine

LEGACY_SOURCE = "digiswasthya_database"
TARGET_APPDETAILS_ID = 42
HERE = os.path.dirname(__file__)

_MIGRATED_LINK = (
    "pi.prescription_id IN (SELECT id FROM prescription "
    f"WHERE legacy_id IS NOT NULL AND legacy_source='{LEGACY_SOURCE}')"
)


def _content_hash_for_ids(conn, ids):
    """Deterministic md5 over the full row content of the given ids."""
    if not ids:
        return "0", 0
    row = conn.execute(text("""
        SELECT md5(coalesce(string_agg(rh, '' ORDER BY rid), '')) AS h, COUNT(*) AS n
        FROM (
            SELECT id::text AS rid, md5(prescriptionitem::text) AS rh
            FROM prescriptionitem WHERE id::text = ANY(:ids)
        ) s
    """), {"ids": ids}).first()
    return row[0], row[1]


def dest_snapshot():
    eng = get_new_engine()
    with eng.connect() as c:
        total = c.execute(text("SELECT COUNT(*) FROM prescriptionitem")).scalar()
        migrated_linked = c.execute(text(
            f"SELECT COUNT(*) FROM prescriptionitem pi WHERE {_MIGRATED_LINK}")).scalar()
        all_ids = [str(r[0]) for r in c.execute(text(
            "SELECT id::text FROM prescriptionitem")).fetchall()]
        content_hash, n = _content_hash_for_ids(c, all_ids)
    return {
        "total": total,
        "migrated_linked": migrated_linked,
        "baseline_ids": all_ids,
        "baseline_content_hash": content_hash,
    }


def main():
    phase = sys.argv[1] if len(sys.argv) > 1 else "before"
    eng = get_new_engine()

    if phase == "before":
        snap = {"phase": phase, "destination": dest_snapshot()}
        out = os.path.join(HERE, "prescription_item_snapshot_before.json")
        with open(out, "w") as f:
            json.dump(snap, f)
        d = snap["destination"]
        print(f"\n=== PRESCRIPTION_ITEM BASELINE ===")
        print(f"total items           : {d['total']}")
        print(f"migrated-linked (ours): {d['migrated_linked']}")
        print(f"protected baseline ids: {len(d['baseline_ids'])}  (content-hashed)")
        print(f"saved -> {out}")
        return

    # phase == after: verify baseline rows untouched, classify new rows
    bpath = os.path.join(HERE, "prescription_item_snapshot_before.json")
    if not os.path.exists(bpath):
        print("[!] No 'before' snapshot — cannot verify.")
        return
    before = json.load(open(bpath))["destination"]
    base_ids = before["baseline_ids"]

    with eng.connect() as c:
        # 1. how many baseline ids still present?
        present = c.execute(text(
            "SELECT COUNT(*) FROM prescriptionitem WHERE id::text = ANY(:ids)"),
            {"ids": base_ids}).scalar()
        # 2. recompute content hash over the SAME baseline ids
        cur_hash, _ = _content_hash_for_ids(c, base_ids)
        total_now = c.execute(text("SELECT COUNT(*) FROM prescriptionitem")).scalar()
        migrated_now = c.execute(text(
            f"SELECT COUNT(*) FROM prescriptionitem pi WHERE {_MIGRATED_LINK}")).scalar()

    missing = len(base_ids) - present
    untouched = (missing == 0 and cur_hash == before["baseline_content_hash"])
    added_total = total_now - before["total"]
    added_ours = migrated_now - before["migrated_linked"]
    added_live = added_total - added_ours

    print(f"\n=== PRESCRIPTION_ITEM VERIFY (live-DB aware) ===")
    print(f"baseline protected rows  : {len(base_ids)}")
    print(f"  still present          : {present}")
    print(f"  missing (deleted)      : {missing}   (must be 0)")
    print(f"  content hash match     : {'YES' if cur_hash==before['baseline_content_hash'] else 'NO'}")
    print(f"total items   : {before['total']} -> {total_now}  (+{added_total})")
    print(f"  added by OUR migration : {added_ours}")
    print(f"  added by LIVE app      : {added_live}  (concurrent activity, expected)")
    print(f"\nEXISTING ROWS UNTOUCHED (no modify/delete): {'YES ✓' if untouched else 'NO ✗  <-- INVESTIGATE'}")


if __name__ == "__main__":
    main()
