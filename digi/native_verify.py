"""
Native-data integrity verifier — WRITE-ATTRIBUTION method (live-DB safe).

The Excel migration ALWAYS stamps legacy_source on any row it inserts or updates.
Native rows have legacy_source IS NULL. So the rigorous, concurrency-proof test
that "our migration did not touch native data" is:

  Every row that had legacy_source IS NULL at baseline must STILL have
  legacy_source IS NULL afterwards (none flipped by us) AND still exist (none
  deleted). Concurrent live-app edits don't set legacy_source, so they don't
  trip this check — unlike a raw content hash.

Usage:
  python native_verify.py <table> before
  python native_verify.py <table> after
  (tables with a legacy_source column: patient, visit, prescription)
"""
import sys, json, os, urllib.parse
from pathlib import Path
from sqlalchemy import create_engine, text

HERE = Path(__file__).resolve().parent


def engine():
    url = [l.split("=", 1)[1].strip() for l in (HERE / ".env").read_text().splitlines()
           if l.startswith("DATABASE_URL=")][0]
    pfx, rest = url.split("://", 1); creds, host = rest.rsplit("@", 1); u, p = creds.split(":", 1)
    return create_engine(f"postgresql+psycopg2://{u}:{urllib.parse.quote(p)}@{host}", pool_pre_ping=True)


def main():
    table = sys.argv[1]
    phase = sys.argv[2] if len(sys.argv) > 2 else "before"
    eng = engine()
    f = HERE / f"native_verify_{table}_before.json"

    with eng.connect() as c:
        if phase == "before":
            ids = [str(r[0]) for r in c.execute(text(
                f"SELECT id::text FROM {table} WHERE legacy_source IS NULL")).fetchall()]
            json.dump({"ids": ids}, open(f, "w"))
            print(f"[{table}] baseline native (legacy_source IS NULL) rows protected: {len(ids)}")
            return

        base = json.load(open(f))["ids"]
        present = c.execute(text(
            f"SELECT COUNT(*) FROM {table} WHERE id::text = ANY(:ids)"), {"ids": base}).scalar()
        flipped = c.execute(text(
            f"SELECT COUNT(*) FROM {table} WHERE id::text = ANY(:ids) AND legacy_source IS NOT NULL"),
            {"ids": base}).scalar()
    missing = len(base) - present
    ok = (missing == 0 and flipped == 0)
    print(f"=== [{table}] NATIVE WRITE-ATTRIBUTION VERIFY ===")
    print(f"  baseline native rows : {len(base)}")
    print(f"  still present        : {present}   (missing/deleted: {missing}, must be 0)")
    print(f"  flipped by our run   : {flipped}   (legacy_source now set — must be 0)")
    print(f"  NATIVE UNTOUCHED BY MIGRATION: {'YES ✓' if ok else 'NO ✗  <-- INVESTIGATE'}")


if __name__ == "__main__":
    main()
