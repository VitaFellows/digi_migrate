"""
migrate_center.py
=================
Per-center orchestrator for the legacy-DB -> AWS RDS migration, reusing the
exact insert-only / source-tagged / native-safe migration functions.

For each target appdetails_id it runs, in order:
  1. patient    (migration/patient/mig_patient.py)
  2. visit      (migration/visit/visit.py)
  3. prescription (migration/prescription/prescription.py)
  4. prescription_item (migration/prescription_item/prescription_item.py)
  5. doctor-less visit backfill (migrate_doctorless_visits.py logic)
and verifies, via WRITE-ATTRIBUTION, that no native row (legacy_source/legacy_id
IS NULL) was modified or deleted on patient/visit/prescription.

Modules are (re)loaded fresh per center AFTER setting TARGET_APPDETAILS_ID in the
environment, so each picks up the correct center.

Usage:
  python migrate_center.py 67            # one center
  python migrate_center.py 45 24 41 ...  # several, in sequence
"""
import os
import sys
import json
import importlib.util
from sqlalchemy import text
from config.db import get_new_engine

HERE = os.path.dirname(os.path.abspath(__file__))
NATIVE_TABLES = ("patient", "visit", "prescription")


def _load(name, relpath):
    spec = importlib.util.spec_from_file_location(name, os.path.join(HERE, relpath))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _native_ids(conn, table):
    return set(str(r[0]) for r in conn.execute(
        text(f"SELECT id::text FROM {table} WHERE legacy_id IS NULL")).fetchall())


def verify_native(before_ids):
    eng = get_new_engine()
    ok = True
    with eng.connect() as c:
        for t in NATIVE_TABLES:
            base = before_ids[t]
            if not base:
                print(f"   [{t}] no baseline native rows"); continue
            present = c.execute(text(
                f"SELECT COUNT(*) FROM {t} WHERE id::text = ANY(:ids)"), {"ids": list(base)}).scalar()
            flipped = c.execute(text(
                f"SELECT COUNT(*) FROM {t} WHERE id::text = ANY(:ids) AND legacy_source IS NOT NULL"),
                {"ids": list(base)}).scalar()
            good = (present == len(base) and flipped == 0)
            ok = ok and good
            print(f"   [{t}] baseline={len(base)} present={present} flipped={flipped} -> "
                  f"{'UNTOUCHED ✓' if good else 'INVESTIGATE ✗'}")
    return ok


def migrate_one(appid: int):
    print("\n" + "=" * 70)
    print(f"  CENTER appdetails_id={appid}")
    print("=" * 70)
    os.environ["TARGET_APPDETAILS_ID"] = str(appid)

    # baseline native id sets BEFORE any write
    eng = get_new_engine()
    with eng.connect() as c:
        before = {t: _native_ids(c, t) for t in NATIVE_TABLES}
    print(f"  baseline native rows: " + ", ".join(f"{t}={len(before[t])}" for t in NATIVE_TABLES))

    # load fresh module instances so each reads TARGET_APPDETAILS_ID from env
    pat = _load(f"pat_{appid}", "migration/patient/mig_patient.py")
    vis = _load(f"vis_{appid}", "migration/visit/visit.py")
    pre = _load(f"pre_{appid}", "migration/prescription/prescription.py")
    itm = _load(f"itm_{appid}", "migration/prescription_item/prescription_item.py")
    dls = _load(f"dls_{appid}", "migrate_doctorless_visits.py")

    print("\n-- patient --");        id_map = {}; pat.migrate_patient(id_map, limit=None)
    print("\n-- visit --");          vis.migrate_visit(limit=None)
    print("\n-- prescription --");   pre.migrate_prescription(limit=None)
    print("\n-- prescription_item --"); itm.migrate_prescription_items(limit=None)

    # doctor-less visit backfill (insert-only, NULL doctor) for this center
    print("\n-- doctor-less visit backfill --")
    sys.argv = ["migrate_doctorless_visits.py", "--commit"]
    dls.main()

    print(f"\n-- native verification (appid={appid}) --")
    ok = verify_native(before)
    print(f"  CENTER {appid} native integrity: {'PASS ✓' if ok else 'FAIL ✗'}")
    return ok


def main():
    appids = [int(a) for a in sys.argv[1:]]
    if not appids:
        sys.exit("usage: python migrate_center.py <appid> [<appid> ...]")
    results = {}
    for a in appids:
        results[a] = migrate_one(a)
    print("\n" + "=" * 70)
    print("  SUMMARY")
    for a, ok in results.items():
        print(f"   appid={a}: native {'PASS ✓' if ok else 'FAIL ✗'}")


if __name__ == "__main__":
    main()
