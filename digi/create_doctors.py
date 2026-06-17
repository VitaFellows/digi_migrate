"""
create_doctors.py
=================
Bulk-create DOCTOR users in the AWS RDS (DigiAarogyaSaarathi) from an Excel
file that contains, at minimum, a full name and a phone number per doctor.

What it does
------------
- Reads the Excel (local .xlsx path or a public Google Sheets URL).
- Detects the name + phone columns flexibly (optional: specialization, qualifications).
- AUTO-GENERATES a unique Doctor ID (employee_id), e.g. DOC0001, DOC0002 …
  (continues after the highest existing DOC#### so it never collides).
- Fills the NOT-NULL user fields: role=DOCTOR, first_name, last_name, phone,
  email (auto, unique), password_hash (UNUSABLE placeholder — no login until reset),
  is_active=TRUE, created_at=now, is_available_now=FALSE.
- Tags every created row with legacy_source='excel_doctor_import' for traceability.
- INSERT-ONLY. Never updates/deletes existing users. Skips a row if a doctor with
  the same phone already exists (configurable).

Safety
------
- Dry-run by default unless you pass --commit. --preview is an explicit dry-run.
- Wraps all inserts in ONE transaction; any error rolls back the whole batch.
- Writes a report CSV (created_doctors_report.csv) mapping name → doctor_id → uuid.

Usage
-----
  python create_doctors.py --excel doctors.xlsx --preview          # show plan, no writes
  python create_doctors.py --excel doctors.xlsx --commit           # actually create
  python create_doctors.py --excel <google-sheet-url> --commit
  python create_doctors.py --excel doctors.xlsx --limit 5 --commit
  python create_doctors.py --excel doctors.xlsx --id-prefix DSDOC --commit
"""
import argparse
import csv
import io
import os
import re
import sys
import uuid
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd
from sqlalchemy import create_engine, text

HERE = Path(__file__).resolve().parent

NAME_COL_CANDIDATES = [
    "full name", "fullname", "name", "doctor name", "doctor's name",
    "doctors name", "doctor", "dr name",
]
PHONE_COL_CANDIDATES = [
    "phone", "phone number", "phone no", "contact", "contact no", "contact no.",
    "contact number", "mobile", "mobile number", "mobile no", "cell",
]
SPEC_COL_CANDIDATES = ["specialization", "speciality", "specialty"]
QUAL_COL_CANDIDATES = ["qualification", "qualifications", "degree"]

UNUSABLE_PWD_PREFIX = "!migrated-no-login:"   # clearly not a valid hash -> no login
LEGACY_SOURCE = "excel_doctor_import"
EMAIL_DOMAIN = "doctor.migrated.local"

# Names matching any of these (case-insensitive) are treated as test/dummy accounts
# and skipped. Extend with --exclude-name on the CLI.
DEFAULT_EXCLUDE_PATTERNS = [r"\btest", r"\bdemo\b", r"\bdummy\b", r"arpit\s+doc"]


# --------------------------------------------------------------------------- #
# Connection
# --------------------------------------------------------------------------- #
def get_engine():
    url = None
    envp = HERE / ".env"
    if envp.exists():
        for line in envp.read_text().splitlines():
            line = line.strip()
            if line.startswith("DATABASE_URL="):
                url = line.split("=", 1)[1].strip()
    url = url or os.getenv("DATABASE_URL")
    if not url:
        sys.exit("[ERROR] DATABASE_URL not found in digi/.env or environment.")
    # URL-encode password (handles @, # etc.)
    if "://" in url and "@" in url:
        prefix, rest = url.split("://", 1)
        creds, host = rest.rsplit("@", 1)
        if ":" in creds:
            u, p = creds.split(":", 1)
            url = f"{prefix}://{u}:{urllib.parse.quote(p)}@{host}"
    if url.startswith("postgresql://"):
        url = "postgresql+psycopg2://" + url[len("postgresql://"):]
    return create_engine(url, pool_pre_ping=True)


# --------------------------------------------------------------------------- #
# Excel loading (local path or public Google Sheet URL)
# --------------------------------------------------------------------------- #
def load_excel(path_or_url: str, sheet):
    s = str(path_or_url).strip().strip('"').strip("'")
    if s.startswith("http://") or s.startswith("https://"):
        m = re.search(r"/d/([a-zA-Z0-9-_]+)", s) or re.search(r"[?&]id=([a-zA-Z0-9-_]+)", s)
        if m:
            s = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
        req = urllib.request.Request(s, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as r:
            src = io.BytesIO(r.read())
    else:
        p = Path(s)
        if not p.exists():
            sys.exit(f"[ERROR] Excel file not found: {p}")
        src = p
    return pd.read_excel(src, sheet_name=sheet if sheet is not None else 0, header=0)


def find_col(df, candidates):
    norm = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        if cand in norm:
            return norm[cand]
    # loose contains-match
    for cand in candidates:
        for low, original in norm.items():
            if cand in low:
                return original
    return None


# --------------------------------------------------------------------------- #
# Field helpers
# --------------------------------------------------------------------------- #
def split_name(full_name: str):
    parts = [p for p in re.split(r"\s+", str(full_name).strip()) if p]
    if not parts:
        return None, None
    if len(parts) == 1:
        return parts[0], "."          # last_name is NOT NULL; use "." sentinel
    return parts[0], " ".join(parts[1:])


def normalize_phone(phone):
    if phone is None:
        return None
    digits = "".join(c for c in str(phone) if c.isdigit())
    if not digits:
        return None
    return digits[-10:] if len(digits) > 10 else digits


def make_id_generator(existing_ids, prefix, width):
    pat = re.compile(rf"^{re.escape(prefix)}(\d+)$", re.IGNORECASE)
    max_n = 0
    for eid in existing_ids:
        m = pat.match(str(eid or "").strip())
        if m:
            max_n = max(max_n, int(m.group(1)))
    counter = {"n": max_n}
    used = set(str(e).strip().lower() for e in existing_ids if e)

    def _next():
        while True:
            counter["n"] += 1
            candidate = f"{prefix}{counter['n']:0{width}d}"
            if candidate.lower() not in used:
                used.add(candidate.lower())
                return candidate
    return _next


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Bulk-create DOCTOR users in AWS RDS from Excel.")
    ap.add_argument("--excel", required=True, help="Local .xlsx path or public Google Sheets URL")
    ap.add_argument("--sheet", default=None, help="Sheet name or index (default: first sheet)")
    ap.add_argument("--limit", type=int, default=None, help="Only process first N rows")
    ap.add_argument("--id-prefix", default="DOC", help="Doctor ID prefix (default: DOC)")
    ap.add_argument("--id-width", type=int, default=4, help="Doctor ID number width (default: 4 -> DOC0001)")
    ap.add_argument("--commit", action="store_true", help="Actually write to the DB (otherwise dry-run)")
    ap.add_argument("--preview", action="store_true", help="Explicit dry-run; no writes")
    ap.add_argument("--allow-duplicate-phone", action="store_true",
                    help="Create even if a user with the same phone already exists")
    ap.add_argument("--exclude-name", action="append", default=[],
                    help="Extra case-insensitive regex; names matching it are skipped (repeatable)")
    ap.add_argument("--no-default-excludes", action="store_true",
                    help="Disable the built-in test/demo/dummy name exclusions")
    args = ap.parse_args()

    exclude_patterns = list(args.exclude_name)
    if not args.no_default_excludes:
        exclude_patterns += DEFAULT_EXCLUDE_PATTERNS
    exclude_re = [re.compile(p, re.IGNORECASE) for p in exclude_patterns]

    dry_run = args.preview or not args.commit
    sheet = args.sheet
    if sheet is not None and str(sheet).isdigit():
        sheet = int(sheet)

    df = load_excel(args.excel, sheet)
    df = df.dropna(how="all").copy()

    name_col = find_col(df, NAME_COL_CANDIDATES)
    phone_col = find_col(df, PHONE_COL_CANDIDATES)
    spec_col = find_col(df, SPEC_COL_CANDIDATES)
    qual_col = find_col(df, QUAL_COL_CANDIDATES)

    if not name_col or not phone_col:
        print(f"[ERROR] Could not detect required columns. Found columns: {list(df.columns)}")
        print(f"        name_col={name_col!r}  phone_col={phone_col!r}")
        sys.exit(1)
    print(f"Detected columns -> name: {name_col!r} | phone: {phone_col!r}"
          f"{f' | specialization: {spec_col!r}' if spec_col else ''}")

    eng = get_engine()
    with eng.connect() as c:
        existing_emp = [r[0] for r in c.execute(text(
            'SELECT employee_id FROM "user" WHERE employee_id IS NOT NULL')).fetchall()]
        existing_email = set(str(r[0]).strip().lower() for r in c.execute(text(
            'SELECT email FROM "user" WHERE email IS NOT NULL')).fetchall())
        existing_phones = set(normalize_phone(r[0]) for r in c.execute(text(
            'SELECT phone FROM "user" WHERE phone IS NOT NULL')).fetchall())

    next_id = make_id_generator(existing_emp, args.id_prefix, args.id_width)

    def _clean(v):
        return None if (v is None or (isinstance(v, float) and pd.isna(v))) else v

    rows = df.to_dict("records")
    planned, skipped = [], []
    seen_phones_this_run = set()
    now = datetime.now(timezone.utc)

    count = 0
    for rec in rows:
        raw_name = _clean(rec.get(name_col))
        raw_phone = _clean(rec.get(phone_col))
        if raw_name is None or str(raw_name).strip() == "":
            continue
        count += 1
        if args.limit and count > args.limit:
            count -= 1
            break

        first, last = split_name(raw_name)
        phone = normalize_phone(raw_phone)
        reason = None
        if any(rx.search(str(raw_name)) for rx in exclude_re):
            reason = "excluded (test/dummy name)"
        elif not phone:
            reason = "missing/invalid phone"
        elif phone in existing_phones and not args.allow_duplicate_phone:
            reason = f"phone {phone} already exists in user table"
        elif phone in seen_phones_this_run and not args.allow_duplicate_phone:
            reason = f"phone {phone} duplicated within the Excel file"

        if reason:
            skipped.append({"name": str(raw_name).strip(), "phone": phone, "reason": reason})
            continue

        seen_phones_this_run.add(phone)
        emp_id = next_id()
        email = f"{emp_id.lower()}@{EMAIL_DOMAIN}"
        if email in existing_email:
            email = f"{emp_id.lower()}.{uuid.uuid4().hex[:6]}@{EMAIL_DOMAIN}"
        existing_email.add(email)

        planned.append({
            "id": str(uuid.uuid4()),
            "role": "DOCTOR",
            "first_name": first,
            "last_name": last,
            "phone": phone,
            "email": email,
            "password_hash": UNUSABLE_PWD_PREFIX + uuid.uuid4().hex,
            "is_active": True,
            "created_at": now,
            "is_available_now": False,
            "specialization": (str(_clean(rec.get(spec_col))).strip() if spec_col and _clean(rec.get(spec_col)) else None),
            "qualifications": (str(_clean(rec.get(qual_col))).strip() if qual_col and _clean(rec.get(qual_col)) else None),
            "employee_id": emp_id,
            "legacy_source": LEGACY_SOURCE,
            "_display_name": str(raw_name).strip(),
        })

    # ---- Report ----
    print(f"\nPlanned to create: {len(planned)} doctor(s) | Skipped: {len(skipped)}")
    print("-" * 72)
    for p in planned:
        print(f"  {p['employee_id']:>10s}  {p['_display_name']:<28s} {p['phone']}")
    if skipped:
        print("\nSkipped:")
        for s in skipped:
            print(f"  - {s['name']:<28s} {s['phone'] or '(no phone)'}  → {s['reason']}")

    report_path = HERE / "created_doctors_report.csv"
    with open(report_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["status", "doctor_id", "name", "phone", "email", "uuid"])
        for p in planned:
            w.writerow(["planned" if dry_run else "created",
                        p["employee_id"], p["_display_name"], p["phone"], p["email"], p["id"]])
        for s in skipped:
            w.writerow(["skipped", "", s["name"], s["phone"] or "", "", s["reason"]])
    print(f"\nReport written -> {report_path}")

    if dry_run:
        print("\n[DRY RUN] No rows written. Re-run with --commit to create these doctors.")
        return

    if not planned:
        print("\nNothing to insert.")
        return

    insert_sql = text("""
        INSERT INTO "user" (
            id, role, first_name, last_name, phone, email, password_hash,
            is_active, created_at, is_available_now, specialization, qualifications,
            employee_id, legacy_source
        ) VALUES (
            :id, :role, :first_name, :last_name, :phone, :email, :password_hash,
            :is_active, :created_at, :is_available_now, :specialization, :qualifications,
            :employee_id, :legacy_source
        )
    """)
    payload = [{k: v for k, v in p.items() if not k.startswith("_")} for p in planned]
    with eng.begin() as conn:
        before = conn.execute(text('SELECT COUNT(*) FROM "user" WHERE role=\'DOCTOR\'')).scalar()
        conn.execute(insert_sql, payload)
        after = conn.execute(text('SELECT COUNT(*) FROM "user" WHERE role=\'DOCTOR\'')).scalar()
    print(f"\n[SUCCESS] DOCTOR users: {before} -> {after}  (+{after - before})")
    print(f"Created {len(planned)} doctor(s). Details in {report_path}")


if __name__ == "__main__":
    main()
