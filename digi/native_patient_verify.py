"""
Native-patient integrity verifier for the Excel→RDS load — LIVE DB aware.

Proves the non-negotiable rule: NO native patient row (legacy_source IS NULL)
is modified or deleted by the Excel migration. Robust to concurrent live-app
inserts (we scope the after-check to the exact baseline native IDs).

Also reports what happened to migration rows (excel_digiswasthya / *_database).

Usage:
  python native_patient_verify.py before
  python native_patient_verify.py after
"""
import sys, json, os
from sqlalchemy import text
import urllib.parse

HERE = os.path.dirname(__file__)


def _engine():
    # Read DATABASE_URL the same way the migration does.
    from pathlib import Path
    envp = Path(HERE) / ".env"
    url = None
    if envp.exists():
        for line in envp.read_text().splitlines():
            line = line.strip()
            if line.startswith("DATABASE_URL="):
                url = line.split("=", 1)[1].strip()
    url = url or os.getenv("DATABASE_URL")
    if "://" in url and "@" in url:
        prefix, rest = url.split("://", 1)
        creds, host = rest.rsplit("@", 1)
        if ":" in creds:
            u, p = creds.split(":", 1)
            url = f"{prefix}://{u}:{urllib.parse.quote(p)}@{host}"
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg2://" + url[len("postgresql://"):]
    from sqlalchemy import create_engine
    return create_engine(url, pool_pre_ping=True)


def _content_hash(conn, ids):
    if not ids:
        return "0"
    return conn.execute(text("""
        SELECT md5(coalesce(string_agg(rh, '' ORDER BY rid), '')) FROM (
            SELECT id::text rid, md5(patient::text) rh
            FROM patient WHERE id::text = ANY(:ids)
        ) s
    """), {"ids": ids}).scalar()


def snapshot():
    eng = _engine()
    with eng.connect() as c:
        native_ids = [str(r[0]) for r in c.execute(text(
            "SELECT id::text FROM patient WHERE legacy_id IS NULL")).fetchall()]
        h = _content_hash(c, native_ids)
        by_src = {(r[0] or "(native NULL)"): r[1] for r in c.execute(text(
            "SELECT legacy_source, COUNT(*) FROM patient GROUP BY legacy_source")).fetchall()}
    return {"native_ids": native_ids, "native_hash": h, "by_source": by_src}


def main():
    phase = sys.argv[1] if len(sys.argv) > 1 else "before"
    if phase == "before":
        snap = snapshot()
        json.dump(snap, open(os.path.join(HERE, "native_patient_before.json"), "w"))
        print(f"\n=== NATIVE PATIENT BASELINE ===")
        print(f"native rows protected : {len(snap['native_ids'])}")
        print(f"by legacy_source      : {snap['by_source']}")
        return

    before = json.load(open(os.path.join(HERE, "native_patient_before.json")))
    base_ids = before["native_ids"]
    eng = _engine()
    with eng.connect() as c:
        present = c.execute(text(
            "SELECT COUNT(*) FROM patient WHERE id::text = ANY(:ids)"), {"ids": base_ids}).scalar()
        cur_hash = _content_hash(c, base_ids)
        by_src = {(r[0] or "(native NULL)"): r[1] for r in c.execute(text(
            "SELECT legacy_source, COUNT(*) FROM patient GROUP BY legacy_source")).fetchall()}

    missing = len(base_ids) - present
    untouched = (missing == 0 and cur_hash == before["native_hash"])
    print(f"\n=== NATIVE PATIENT VERIFY (live-DB aware) ===")
    print(f"baseline native rows   : {len(base_ids)}")
    print(f"  still present        : {present}")
    print(f"  missing (deleted)    : {missing}   (must be 0)")
    print(f"  content hash match   : {'YES' if cur_hash==before['native_hash'] else 'NO'}")
    print(f"\nby legacy_source  before -> after:")
    for k in sorted(set(before['by_source']) | set(by_src)):
        print(f"   {k:24s} {before['by_source'].get(k,0)} -> {by_src.get(k,0)}")
    print(f"\nNATIVE PATIENTS UNTOUCHED (no modify/delete): {'YES ✓' if untouched else 'NO ✗  <-- INVESTIGATE'}")


if __name__ == "__main__":
    main()
