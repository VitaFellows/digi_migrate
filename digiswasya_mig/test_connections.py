"""Quick connectivity check for both migration databases."""
import sys
from sqlalchemy import text
from config.db import get_legacy_engine, get_new_engine


def check(label, engine_factory):
    print(f"\n[{label}]")
    try:
        engine = engine_factory()
    except KeyError as e:
        print(f"  SKIP — env var {e} not set")
        return False
    try:
        with engine.connect() as conn:
            ver = conn.execute(text("select version()")).scalar()
            db = conn.execute(text("select current_database()")).scalar()
            print(f"  OK  — connected to '{db}'")
            print(f"        {ver.split(',')[0]}")
        return True
    except Exception as e:
        print(f"  FAIL — {type(e).__name__}: {str(e).splitlines()[0]}")
        return False


if __name__ == "__main__":
    ok_src = check("SOURCE (Supabase / legacy)", get_legacy_engine)
    ok_dst = check("DESTINATION (AWS RDS / new)", get_new_engine)
    sys.exit(0 if (ok_src and ok_dst) else 1)
