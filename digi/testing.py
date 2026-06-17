"""
Standalone direct Supabase migration script with sequential commits and Update-or-Insert Mode.

This script:
1. Migrates and commits the 'patient' table first in its own transaction block.
2. Uses a custom government_id generation logic: first three letters of first name +
   first three letters of last name + date of birth year.
3. Pauses and prompts the user before migrating the remaining transactional tables.
4. Adds a date-based collision filter with a +/- 1-day safety margin.
5. In the collision dates, operates in Update-or-Insert mode: if a patient/visit/prescription
   already exists, it compares their fields and updates empty/null columns in the database 
   with Excel values (including merging missing vitals keys in vitals_json).
"""

from __future__ import annotations

import os
import re
import sys
import json
import uuid
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from difflib import SequenceMatcher
from datetime import datetime, date, timedelta

import pandas as pd

# SQLAlchemy dependency check
try:
    from sqlalchemy import create_engine, text
except Exception as exc:
    raise RuntimeError(
        "SQLAlchemy is required for Supabase DB migration. Install sqlalchemy and a PostgreSQL driver (psycopg2-binary)."
    ) from exc

# Import configurations, cleaners, transformers, and helpers from the existing codebase
try:
    from cleaner import detect_duplicates, clean_date, clean_text
    from config import (
        EXCEL_ONLY_CUTOFF_DATE,
        INPUT_EXCEL_PATH,
        INPUT_SHEET_NAME,
        LEGACY_SOURCE,
        MIGRATION_SYSTEM_USER_ID,
        FALLBACK_DOCTOR_USER_ID,
        TENANT_ID,
        IS_DEMO,
    )
    from transformer import (
        transform_patient,
        transform_visit,
        transform_prescription,
        transform_prescription_items,
        transform_followup_schedules,
    )
except ImportError:
    # Handle import fallbacks if run inside the directory
    sys.path.append(str(Path(__file__).resolve().parent))
    from cleaner import detect_duplicates, clean_date, clean_text
    from config import (
        EXCEL_ONLY_CUTOFF_DATE,
        INPUT_EXCEL_PATH,
        INPUT_SHEET_NAME,
        LEGACY_SOURCE,
        MIGRATION_SYSTEM_USER_ID,
        FALLBACK_DOCTOR_USER_ID,
        TENANT_ID,
        IS_DEMO,
    )
    from transformer import (
        transform_patient,
        transform_visit,
        transform_prescription,
        transform_prescription_items,
        transform_followup_schedules,
    )

import cleaner
import transformer
from cleaner import clean_vitals as original_clean_vitals

# Legacy DB 1 year 4 months collision range config (Adjust these dates as needed)
COLLISION_START_DATE_STR = "2023-01-01"
COLLISION_END_DATE_STR = "2024-05-01"


# - MONKEYPATCHES AND CLEANING EXTENSIONS -------------------------------------

def parse_age_years_months(age_val: Any) -> tuple[int | None, int | None]:
    """
    Parses raw age value from Excel and returns (age_years, age_months).
    If it indicates months (e.g. '11 months', '4 m'), sets age_years to 0.
    """
    if age_val is None or (isinstance(age_val, float) and pd.isna(age_val)):
        return None, None
    
    age_str = str(age_val).strip().lower()
    if age_str in ("", "nan", "nat", "null", "none", "-"):
        return None, None
        
    # Extract first number in string
    match = re.search(r"([0-9.]+)", age_str)
    if not match:
        return None, None
        
    num_val = float(match.group(1))
    
    # Check if string contains month indicators
    is_months = False
    if "month" in age_str or "mnth" in age_str:
        is_months = True
    elif re.search(r"\b\d+\s*m\b", age_str) or re.search(r"\b\d+\s*m$", age_str):
        is_months = True
        
    if is_months:
        return 0, int(round(num_val))
    else:
        return int(round(num_val)), None


def my_clean_vitals(row: dict, uid: str, log: list) -> dict:
    """
    Overridden clean_vitals to parse BMI from Excel and insert it into vitals_json.
    """
    vitals = original_clean_vitals(row, uid, log)
    
    # Extract BMI using transformer._row_lookup and cleaner._safe_float
    bmi_val = cleaner._safe_float(transformer._row_lookup(row, "BMI", "BMI (do not fill manually)", "bmidonotfillmanually"))
    if bmi_val is not None and not pd.isna(bmi_val):
        if float(bmi_val).is_integer():
            vitals["bmi"] = int(bmi_val)
        else:
            vitals["bmi"] = round(float(bmi_val), 1)
    else:
        vitals["bmi"] = None
        
    return vitals

# Save original lookup and define our enhanced lookup wrapper
original_row_lookup = transformer._row_lookup

def my_row_lookup(row: dict, *candidates: str, default: Any = None) -> Any:
    # Build enhanced list of candidates
    enhanced_candidates = list(candidates)
    
    is_uid = any(c in candidates for c in ("UID", "Id", "ID"))
    if is_uid:
        for extra in ("ma ", "CONSULTATION ID\n", "CONSULTATION ID\nDS1", "CONSULTATION ID\nDS5", "CONSULTATION ID", "CONSULTATION ID DS1", "CONSULTATION ID DS5", "PATIENT ID"):
            if extra not in enhanced_candidates:
                enhanced_candidates.append(extra)
                
    is_phone = any("PHONE" in str(c).upper() or "CONTACT" in str(c).upper() or "MOBILE" in str(c).upper() for c in candidates)
    if is_phone:
        for extra in ("CONTACT NO.", "CONTACT NO", "PHONE  NO.", "PHONE NO.", "Contact No.", "Contact no.", "Phone", "Phone No.", "Phone No", "PHONE", "Mobile"):
            if extra not in enhanced_candidates:
                enhanced_candidates.append(extra)
                
    is_age = any("AGE" in str(c).upper() for c in candidates)
    if is_age:
        for extra in ("AGE", "Age", "age"):
            if extra not in enhanced_candidates:
                enhanced_candidates.append(extra)
                
    is_date = any("DATE" in str(c).upper() for c in candidates)
    if is_date:
        for extra in ("DATE", "Date", "date", "Visit Date", "VISIT DATE", "Consultation Date", "DATE OF VISIT"):
            if extra not in enhanced_candidates:
                enhanced_candidates.append(extra)

    res = original_row_lookup(row, *enhanced_candidates, default=default)
    
    # If UID lookup failed, check Unnamed: 0 or other unnamed columns
    if is_uid and (res is None or str(res).strip() == "" or pd.isna(res)):
        val = row.get("Unnamed: 0")
        if val and not pd.isna(val):
            val_str = str(val).strip()
            if "DS" in val_str or val_str.startswith("DS"):
                return val_str
            for k, v in row.items():
                if "unnamed" in str(k).lower() and v and not pd.isna(v) and ("DS" in str(v) or str(v).strip().startswith("DS")):
                    return str(v).strip()
                    
    return res

def resolve_uid(row: dict) -> str:
    uid = my_row_lookup(row, "UID", "Id", "ID")
    return str(uid).strip() if uid and not pd.isna(uid) else ""

# Enhanced Dynamic Medicine Columns support
def my_infer_block_start_indices(keys: list[str]) -> list[int]:
    starts = []
    valid_norms = {
        "medicine1", "medicine", "medicinename", "med",
        "medicine2", "name",
        "medicine3", "name1",
        "medicine4", "name2",
        "medicine5", "medicinename4", "name3", "name4", "medicinename1", "medicinename2", "medicinename3", "medicinement"
    }
    for idx, key in enumerate(keys):
        key_norm = _norm(key)
        if key_norm in valid_norms:
            starts.append(idx)
    return starts

def parse_multiline_medicines(cell_value: str) -> list[dict]:
    parsed_items = []
    lines = [line.strip() for line in cell_value.split("\n") if line.strip()]
    for line in lines:
        if re.match(r"^[✓✗\s—\-+_|:;.,]+$", line):
            continue
            
        parts = [p.strip() for p in line.split("\t") if p.strip()]
        if not parts:
            continue
            
        med_name = parts[0]
        if re.match(r"^[✓✗\s—\-+_|:;.,]+$", med_name):
            continue
            
        dosage = "1"
        frequency = "As directed"
        duration = None
        notes = None
        
        if len(parts) > 1:
            if re.match(r"^[✓✗\s—\-+_|:;.,]+$", parts[1]):
                pass
            else:
                notes = parts[1]
                freq_match = re.search(r"\b(od|bd|tds|tid|qid|sos|hs|prn|once|twice|daily|weekly)\b", notes.lower())
                if freq_match:
                    frequency = freq_match.group(1).upper()
                    
        freq_match_in_name = re.search(r"\b(od|bd|tds|tid|qid|sos|hs|prn|once|twice|daily|weekly)\b", med_name.lower())
        if freq_match_in_name:
            frequency = freq_match_in_name.group(1).upper()
            
        parsed_items.append({
            "medication_name": med_name,
            "dosage": dosage,
            "frequency": frequency,
            "duration": duration,
            "notes": notes,
        })
    return parsed_items

from transformer import build_frequency_string, clean_medicine_name, _blank_row, _extract_block_value

def my_transform_prescription_items(row: dict, prescription_uuid: str, log: list) -> list[dict]:
    items: list[dict] = []
    try:
        keys = list(row.keys())
        values = list(row.values())
        
        multiline_processed = False
        for k, v in row.items():
            if pd.isna(v) or not isinstance(v, str):
                continue
            v_str = str(v)
            if "\n" in v_str and ("\t" in v_str or "✓" in v_str or "✗" in v_str or "—" in v_str):
                parsed = parse_multiline_medicines(v_str)
                if parsed:
                    for p_item in parsed:
                        item = _blank_row("prescriptionitem")
                        item.update({
                            "id": str(uuid.uuid4()),
                            "prescription_id": prescription_uuid,
                            "medication_name": p_item["medication_name"],
                            "dosage": p_item["dosage"],
                            "frequency": p_item["frequency"],
                            "duration": p_item["duration"],
                            "quantity": None,
                            "unit": None,
                            "notes": p_item["notes"],
                            "formulary_id": None,
                            "tenant_id": TENANT_ID,
                        })
                        items.append(item)
                    multiline_processed = True
                    
        if multiline_processed:
            return items

        start_indices = my_infer_block_start_indices(keys)
        seen_block_starts = []
        for start_index in start_indices:
            if start_index in seen_block_starts:
                continue
            seen_block_starts.append(start_index)
            
            medication_name = clean_medicine_name(_extract_block_value(keys, values, start_index, 0))
            if medication_name is None:
                continue
                
            if isinstance(medication_name, str) and "\n" in medication_name and ("\t" in medication_name or "✓" in medication_name or "✗" in medication_name):
                parsed = parse_multiline_medicines(medication_name)
                for p_item in parsed:
                    item = _blank_row("prescriptionitem")
                    item.update({
                        "id": str(uuid.uuid4()),
                        "prescription_id": prescription_uuid,
                        "medication_name": p_item["medication_name"],
                        "dosage": p_item["dosage"],
                        "frequency": p_item["frequency"],
                        "duration": p_item["duration"],
                        "quantity": None,
                        "unit": None,
                        "notes": p_item["notes"],
                        "formulary_id": None,
                        "tenant_id": TENANT_ID,
                    })
                    items.append(item)
                continue

            dosage = clean_text(_extract_block_value(keys, values, start_index, 1)) or "1"
            unit = clean_text(_extract_block_value(keys, values, start_index, 2))
            morning = _extract_block_value(keys, values, start_index, 3)
            afternoon = _extract_block_value(keys, values, start_index, 4)
            evening = _extract_block_value(keys, values, start_index, 5)
            remark = _extract_block_value(keys, values, start_index, 6)
            duration = clean_text(_extract_block_value(keys, values, start_index, 7))
            frequency = build_frequency_string(morning, afternoon, evening, remark) or "As directed"
            notes = clean_text(remark, max_len=1000)
            
            item = _blank_row("prescriptionitem")
            item.update({
                "id": str(uuid.uuid4()),
                "prescription_id": prescription_uuid,
                "medication_name": medication_name,
                "dosage": dosage,
                "frequency": frequency,
                "duration": duration,
                "quantity": None,
                "unit": unit,
                "notes": notes,
                "formulary_id": None,
                "tenant_id": TENANT_ID,
            })
            items.append(item)
            
        return items
    except Exception as exc:
        uid = str(my_row_lookup(row, "UID", default="") or "")
        log.append({"uid": uid, "field": "prescriptionitem", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        return items

# Apply monkeypatches
cleaner.clean_vitals = my_clean_vitals
transformer.clean_vitals = my_clean_vitals
transformer._row_lookup = my_row_lookup
cleaner._row_lookup = my_row_lookup
transformer.transform_prescription_items = my_transform_prescription_items
_row_lookup = my_row_lookup


# - ENVIRONMENT AND UTILITIES ------------------------------------------------

def _load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, value = stripped.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        os.environ.setdefault(key, value)


import io
import urllib.request

def get_google_drive_download_url(url: str) -> str:
    """
    Converts a standard Google Drive or Google Sheets sharing link into a direct
    download link for Excel format (.xlsx).
    """
    url_str = str(url).strip().strip('"').strip("'")
    if "drive.google.com" in url_str or "docs.google.com" in url_str:
        file_id_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url_str)
        if file_id_match:
            file_id = file_id_match.group(1)
            return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        id_match = re.search(r"[?&]id=([a-zA-Z0-9-_]+)", url_str)
        if id_match:
            file_id = id_match.group(1)
            return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    return url_str


def load_excel_from_path(excel_path_or_url: str | Path) -> io.BytesIO | Path:
    """
    Given a local file path or an HTTP/HTTPS URL, returns an in-memory BytesIO stream
    for URLs or a Path object for local files.
    """
    path_str = str(excel_path_or_url).strip().strip('"').strip("'")
    if path_str.startswith("http://") or path_str.startswith("https://"):
        url = get_google_drive_download_url(path_str)
        print(f"      [HTTP] Downloading Excel spreadsheet from: {url}")
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'}
        )
        with urllib.request.urlopen(req) as response:
            return io.BytesIO(response.read())
    return Path(path_str)


@dataclass
class DbConfig:
    database_url: str
    excel_path: Path | str
    sheet_name: int | str
    schema: str = "public"


def _read_config() -> DbConfig:
    import urllib.parse
    script_dir = Path(__file__).resolve().parent
    project_root = script_dir.parent
    
    # Load environment files
    _load_env_file(script_dir / ".env")
    _load_env_file(project_root / ".env")

    database_url = os.getenv("DATABASE_URL")
    if not database_url or "[YOUR_SUPABASE_" in database_url:
        print("\n[WARNING] DATABASE_URL is missing or contains placeholder values in your .env file!")
        print("Please edit the .env file in your workspace root and add your actual Supabase PostgreSQL connection string.\n")
        sys.exit(1)

    # Auto-encode password to handle special characters (e.g. @ in password)
    if "://" in database_url:
        prefix, rest = database_url.split("://", 1)
        if "@" in rest:
            creds, host_part = rest.rsplit("@", 1)
            if ":" in creds:
                user, password = creds.split(":", 1)
                encoded_password = urllib.parse.quote(password)
                database_url = f"{prefix}://{user}:{encoded_password}@{host_part}"

    env_excel_path = os.getenv("INPUT_EXCEL_PATH")
    env_workbook_path = os.getenv("MIGRATION_WORKBOOK_PATH")
    
    excel_path_val = None
    
    # 1. If INPUT_EXCEL_PATH is set to something other than "sample.xlsx", prioritize it
    if env_excel_path and env_excel_path != "sample.xlsx":
        excel_path_val = env_excel_path
        
    # 2. If MIGRATION_WORKBOOK_PATH is set and points to an existing file or is a URL, use it
    if not excel_path_val and env_workbook_path:
        env_wb_str = str(env_workbook_path).strip().strip('"').strip("'")
        if env_wb_str.startswith("http://") or env_wb_str.startswith("https://"):
            excel_path_val = env_wb_str
        else:
            p_wb = Path(env_wb_str)
            if p_wb.exists() and p_wb.is_file():
                excel_path_val = env_workbook_path
            
    # 3. Fallback to default
    if not excel_path_val:
        excel_path_val = env_excel_path if env_excel_path else INPUT_EXCEL_PATH
        
    excel_path_str = str(excel_path_val).strip().strip('"').strip("'")
    if excel_path_str.startswith("http://") or excel_path_str.startswith("https://"):
        print(f"      [CONFIG] Resolved Excel spreadsheet source URL to: {excel_path_str}")
        schema = os.getenv("DB_SCHEMA", "public")
        return DbConfig(database_url=database_url, excel_path=excel_path_str, sheet_name=INPUT_SHEET_NAME, schema=schema)

    excel_path = Path(excel_path_val)
    if not excel_path.is_absolute():
        candidate1 = project_root / excel_path
        candidate2 = script_dir / excel_path
        if candidate1.exists():
            excel_path = candidate1
        elif candidate2.exists():
            excel_path = candidate2
        else:
            excel_path = candidate1

    print(f"      [CONFIG] Resolved Excel spreadsheet source path to: {excel_path.resolve()}")
    schema = os.getenv("DB_SCHEMA", "public")
    return DbConfig(database_url=database_url, excel_path=excel_path, sheet_name=INPUT_SHEET_NAME, schema=schema)


def _table(schema: str, name: str) -> str:
    return f'"{schema}"."{name}"' if schema else f'"{name}"'


def _fetch_one(conn, sql: str, params: dict[str, Any]) -> Any:
    return conn.execute(text(sql), params).mappings().first()


def _norm(value) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalnum())


def _find_column(df: pd.DataFrame, *candidates: str):
    normalized = {_norm(column): column for column in df.columns}
    for candidate in candidates:
        candidate_norm = _norm(candidate)
        if candidate_norm in normalized:
            return normalized[candidate_norm]
    return None


def _first_existing(row: dict[str, Any], *fields: str) -> str | None:
    for field in fields:
        value = clean_text(row.get(field))
        if value:
            return value
    return None


def clean_row_for_db(row: dict[str, Any]) -> dict[str, Any]:
    cleaned = {}
    for k, v in row.items():
        if pd.isna(v):
            cleaned[k] = None
        else:
            if isinstance(v, str):
                v_stripped = v.strip()
                cleaned[k] = None if v_stripped.lower() in ("nan", "nat", "null", "none", "") else v_stripped
            else:
                cleaned[k] = v
    return cleaned


# - FUZZY CENTER RESOLUTION AND COORDINATOR ROTATION -------------------------

def _norm_text(value: Any) -> str:
    return clean_text(value).lower() if clean_text(value) else ""


def _center_aliases(value: Any) -> list[str]:
    text = clean_text(value)
    if not text:
        return []
    lowered = text.lower()
    lowered = lowered.replace("ngp", " ")
    lowered = lowered.replace("telemedicine", " ")
    lowered = lowered.replace("centre", " ")
    lowered = lowered.replace("center", " ")
    lowered = lowered.replace("digiswasthya", " ")
    lowered = lowered.replace("vijay nagar", "vijaynagar")
    parts = [part.strip() for part in re.split(r"[,/\-]", lowered) if part.strip()]
    aliases = set()
    for part in parts:
        cleaned = _norm_text(part)
        if cleaned:
            aliases.add(cleaned)
    collapsed = _norm_text(lowered)
    if collapsed:
        aliases.add(collapsed)
    return list(aliases)


def _resolve_center_row_for_sheet(center_rows: list[dict[str, Any]], sheet_ref: int | str | None) -> dict[str, Any] | None:
    if not center_rows:
        return None
    if isinstance(sheet_ref, str):
        stripped = sheet_ref.strip()
        # Special fallback mapping for Bihar sheet
        if stripped.lower() == "bihar":
            for row in center_rows:
                name = str(row.get("name", "")).lower()
                if "sahebganj" in name:
                    return row
        if stripped.isdigit():
            index = int(stripped)
            if 0 <= index < len(center_rows):
                return center_rows[index]
        normalized = _norm_text(stripped)
        sheet_aliases = _center_aliases(stripped)
        for row in center_rows:
            row_aliases = _center_aliases(row.get("short_name")) + _center_aliases(row.get("name")) + _center_aliases(row.get("center_code"))
            row_aliases = [alias for alias in dict.fromkeys(row_aliases) if alias]
            for alias in row_aliases:
                if alias == normalized or alias in sheet_aliases or normalized in alias:
                    return row
            for alias in row_aliases:
                for sheet_alias in sheet_aliases:
                    ratio = SequenceMatcher(None, alias, sheet_alias).ratio()
                    if ratio >= 0.88:
                        return row
    elif isinstance(sheet_ref, int):
        if 0 <= sheet_ref < len(center_rows):
            return center_rows[sheet_ref]
        return center_rows[0]
    return center_rows[0]


def _build_center_coordinator_rotation(user_rows: list[dict[str, Any]]) -> tuple[dict[str, list[str]], dict[str, int]]:
    coordinators_by_center: dict[str, list[str]] = defaultdict(list)
    counters: dict[str, int] = defaultdict(int)
    for row in user_rows:
        role = _norm_text(row.get("role"))
        center_id = clean_text(row.get("center_id"))
        user_id = clean_text(row.get("id"))
        if role in ("coordinator", "staff", "admin") and center_id and user_id:
            coordinators_by_center[center_id].append(user_id)
    return coordinators_by_center, counters


def _pick_center_coordinator(center_id: str | None, coordinators_by_center: dict[str, list[str]], counters: dict[str, int], fallback_user_id: str, all_coordinators: list[str]) -> str:
    if not center_id:
        if all_coordinators:
            index = counters["global_coordinators"] % len(all_coordinators)
            counters["global_coordinators"] += 1
            return all_coordinators[index]
        return fallback_user_id
        
    coordinator_ids = coordinators_by_center.get(center_id, [])
    if not coordinator_ids:
        if all_coordinators:
            index = counters["global_coordinators"] % len(all_coordinators)
            counters["global_coordinators"] += 1
            return all_coordinators[index]
        return fallback_user_id
        
    index = counters[center_id] % len(coordinator_ids)
    counters[center_id] += 1
    return coordinator_ids[index]


# - DEDUPLICATION AND DATABASE CONSTRAINTS CHECKS -----------------------------

def generate_custom_government_id(name: str | None, age: Any, visit_date_val: Any, uid: str) -> str:
    """
    Generate custom government ID:
    First three letters of first name + first three letters of last name + DOB year.
    DOB year is calculated as: visit_year - age_years.
    """
    # Fallback default values
    visit_year = 2025
    
    # 1. Parse visit date to extract year
    if visit_date_val:
        parsed_date = clean_date(visit_date_val, "DATE", "temp", [])
        if parsed_date:
            visit_year = parsed_date.year
            
    # 2. Parse age to subtract from visit year
    dob_year = visit_year
    age_years, age_months = parse_age_years_months(age)
    if age_years is not None:
        dob_year = visit_year - age_years

    # 3. Clean and parse name
    if not name or str(name).strip().lower() in ("nan", "nat", "null", "none", "unknown"):
        # Unique fallback for unknown names to avoid collision
        return f"unkunk{dob_year}_{uid[-8:]}".lower()

    # Split name by space
    parts = [p.strip() for p in str(name).split() if p.strip()]
    if not parts:
        return f"unkunk{dob_year}_{uid[-8:]}".lower()

    first_name = parts[0]
    last_name = parts[-1] if len(parts) > 1 else ""

    # Keep only alphabetic characters
    first_clean = "".join(c for c in first_name if c.isalpha()).lower()
    last_clean = "".join(c for c in last_name if c.isalpha()).lower()

    f_part = first_clean[:3]
    if len(f_part) < 3:
        f_part = (f_part + "xxx")[:3]

    l_part = last_clean[:3]
    if not l_part:
        l_part = "unk"
    elif len(l_part) < 3:
        l_part = (l_part + "xxx")[:3]

    return f"{f_part}{l_part}{dob_year}"


def clean_to_digits(phone: Any) -> str:
    if not phone:
        return ""
    return "".join(c for c in str(phone) if c.isdigit())


def phone_numbers_match(phone1: Any, phone2: Any) -> bool:
    digits1 = clean_to_digits(phone1)[-10:]
    digits2 = clean_to_digits(phone2)[-10:]
    if not digits1 and not digits2:
        return True
    return digits1 == digits2


def _find_all_matching_patient_ids(conn, schema: str, row: dict[str, Any]) -> list[str]:
    """
    Look-up to match patient profiles based on case-insensitive full_name
    and phone number.
    Returns all matching patient UUIDs.
    """
    full_name = row.get("full_name")
    excel_phone = row.get("phone") or row.get("phone_last10")
    
    if full_name:
        # Standardize whitespace in parameter
        clean_full_name = " ".join(str(full_name).strip().split())
        # Use TRIM and REGEXP_REPLACE to collapse any extra spaces in database name column
        sql = f"SELECT id, phone, phone_last10 FROM {_table(schema, 'patient')} WHERE LOWER(TRIM(REGEXP_REPLACE(full_name, '\\s+', ' ', 'g'))) = LOWER(:full_name)"
        candidates = [dict(r) for r in conn.execute(text(sql), {"full_name": clean_full_name}).mappings().all()]
        
        matched_ids = []
        for cand in candidates:
            # Match Phone number
            cand_phone = cand.get("phone_last10") or cand.get("phone")
            if phone_numbers_match(excel_phone, cand_phone):
                matched_ids.append(str(cand["id"]))
                
        return matched_ids
    return []


def _find_matching_patient_ids_with_age_fallback(conn, schema: str, row: dict[str, Any]) -> list[str]:
    """
    Find matched patient IDs by Name + Phone + Age.
    If none found, fallback to Name + Phone only.
    """
    full_name = row.get("full_name")
    excel_phone = row.get("phone") or row.get("phone_last10")
    excel_age_years = row.get("age_years")
    excel_age_months = row.get("age_months")
    
    if not full_name:
        return []
        
    # Collapsing multiple spaces
    clean_full_name = " ".join(str(full_name).strip().split())
    
    sql = f"""
        SELECT id, phone, phone_last10, age_years, age_months 
        FROM {_table(schema, 'patient')} 
        WHERE LOWER(TRIM(REGEXP_REPLACE(full_name, '\\s+', ' ', 'g'))) = LOWER(:full_name)
    """
    candidates = [dict(r) for r in conn.execute(text(sql), {"full_name": clean_full_name}).mappings().all()]
    
    # Filter candidates by phone
    phone_matched = []
    for cand in candidates:
        cand_phone = cand.get("phone_last10") or cand.get("phone")
        if phone_numbers_match(excel_phone, cand_phone):
            phone_matched.append(cand)
            
    if not phone_matched:
        return []
        
    # First check: Name + Phone + Age
    age_matched_ids = []
    for cand in phone_matched:
        cand_years = cand.get("age_years")
        cand_months = cand.get("age_months")
        
        # Match age_years (handle None safely)
        years_match = (cand_years == excel_age_years)
        
        excel_months_clean = excel_age_months if excel_age_months is not None else 0
        cand_months_clean = cand_months if cand_months is not None else 0
        months_match = (cand_months_clean == excel_months_clean)
        
        if years_match and months_match:
            age_matched_ids.append(str(cand["id"]))
            
    if age_matched_ids:
        return age_matched_ids
        
    # Fallback to Name + Phone
    return [str(cand["id"]) for cand in phone_matched]


def _find_existing_patient_id(conn, schema: str, row: dict[str, Any]) -> str | None:
    matched = _find_all_matching_patient_ids(conn, schema, row)
    return matched[0] if matched else None


def _validate_not_null_fields(table: str, row: dict[str, Any], required_fields: list[str], uid: str, log: list) -> None:
    for field in required_fields:
        val = row.get(field)
        if val is None or str(val).strip() == "" or str(val).lower() in ("nan", "nat", "null", "none"):
            log.append({
                "uid": uid,
                "field": f"{table}.{field}",
                "issue": "REQUIRED_FIELD_EMPTY",
                "raw": str(val) if val is not None else None,
                "detail": f"Field '{field}' in table '{table}' is empty/null which violates NOT NULL constraint!"
            })


def _ensure_patient(conn, cfg: DbConfig, row: dict[str, Any], cache: dict[str, str]) -> str:
    """Deduplicates and inserts a single patient profile into Supabase."""
    row = clean_row_for_db(row)
    row = filter_columns_for_db("patient", row)
    
    full_name = str(row.get('full_name') or '').strip().lower()
    phone = row.get('phone_last10') or row.get('phone') or ''
    phone_clean = clean_to_digits(phone)[-10:]
    key = f"{full_name}|{phone_clean}"
    
    if key in cache:
        return cache[key]
        
    # 2. Live Supabase database check
    existing_id = _find_existing_patient_id(conn, cfg.schema, row)
    if existing_id:
        cache[key] = existing_id
        return existing_id

        
    # 3. New Patient Registration
    if not row.get("id"):
        row["id"] = str(uuid.uuid4())
        
    columns = list(row.keys())
    placeholders = ", ".join(f":{col}" for col in columns)
    column_sql = ", ".join(f'"{col}"' for col in columns)
    sql = f"INSERT INTO {_table(cfg.schema, 'patient')} ({column_sql}) VALUES ({placeholders}) RETURNING id"
    
    result = conn.execute(text(sql), row).mappings().first()
    patient_id = str(result["id"])
    cache[key] = patient_id
    return patient_id


def _ensure_legacy_row(conn, cfg: DbConfig, table: str, row: dict[str, Any]) -> str:
    """Inserts a transactional record checking unique keys legacy_id & legacy_source."""
    row = clean_row_for_db(row)
    row = filter_columns_for_db(table, row)
    legacy_id = row.get("legacy_id")
    legacy_source = row.get("legacy_source")
    
    if legacy_id and legacy_source:
        sql = f"SELECT id FROM {_table(cfg.schema, table)} WHERE legacy_id=:legacy_id AND legacy_source=:legacy_source LIMIT 1"
        existing = _fetch_one(conn, sql, {"legacy_id": legacy_id, "legacy_source": legacy_source})
        if existing:
            return str(existing["id"])
            
    if not row.get("id"):
        row["id"] = str(uuid.uuid4())
        
    columns = list(row.keys())
    placeholders = ", ".join(f":{col}" for col in columns)
    column_sql = ", ".join(f'"{col}"' for col in columns)
    sql = f"INSERT INTO {_table(cfg.schema, table)} ({column_sql}) VALUES ({placeholders}) RETURNING id"
    result = conn.execute(text(sql), row).mappings().first()
    return str(result["id"])


def clean_doctor_name(val: Any) -> str:
    if not val:
        return ""
    text = str(val).lower().strip()
    text = re.sub(r"^(dr|doctor|ms|miss|mrs|mr)\b\.?\s*", "", text)
    cleaned = "".join(c for c in text if c.isalnum())
    # Handle known spelling corrections/aliases
    corrections = {
        "abhijitgupta": "abhijeetgupta",
        "shivangi": "shivangiyadav",
        "bhuvneshchaturvedicentre": "bhuvneshchaturvedi",
    }
    return corrections.get(cleaned, cleaned)


# - UPDATE MODE HELPER --------------------------------------------------------

DB_TABLE_COLUMNS: dict[str, set[str]] = {}

def filter_columns_for_db(table: str, row: dict[str, Any]) -> dict[str, Any]:
    cols = DB_TABLE_COLUMNS.get(table)
    if cols:
        return {k: v for k, v in row.items() if k in cols}
    return row


def _has_actual_updates(conn, schema: str, table: str, record_id: str, transformed_row: dict[str, Any]) -> bool:
    """
    Checks if there are any actual field updates (NULL -> value) between the database and transformed Excel row.
    Returns True if at least one field has an update, False otherwise.
    """
    try:
        transformed_row = filter_columns_for_db(table, transformed_row)
        sql = f"SELECT * FROM {_table(schema, table)} WHERE id = :id LIMIT 1"
        db_row = conn.execute(text(sql), {"id": record_id}).mappings().first()
        if not db_row:
            return False
            
        # For patient records, only allow updates if existing legacy_source is "Digiswasthya_Database"
        if table == "patient" and db_row.get("legacy_source") != "Digiswasthya_Database":
            return False
            
        for col, excel_val in transformed_row.items():
            if col in ("id", "created_at", "migrated_at", "legacy_id", "legacy_source", "tenant_id", "is_demo"):
                continue
                
            db_val = db_row.get(col)
            
            # Special merging for vitals_json (JSONB)
            if col == "vitals_json":
                db_vitals = {}
                if db_val:
                    try:
                        db_vitals = json.loads(db_val) if isinstance(db_val, str) else db_val
                    except Exception:
                        pass
                excel_vitals = {}
                if excel_val:
                    try:
                        excel_vitals = json.loads(excel_val) if isinstance(excel_val, str) else excel_val
                    except Exception:
                        pass
                        
                if isinstance(db_vitals, dict) and isinstance(excel_vitals, dict):
                    for k, v in excel_vitals.items():
                        if (db_vitals.get(k) is None or str(db_vitals.get(k)).strip() == "") and (v is not None and str(v).strip() != ""):
                            return True
                continue
                
            # Standard column blank check (treating placeholder values as empty)
            db_val_str = str(db_val).strip() if db_val is not None else ""
            db_is_empty = (
                db_val is None or 
                db_val_str == "" or 
                (col == "mother_tongue" and db_val_str in ("0", "0.0")) or
                (col == "chief_complaint" and db_val_str.lower() == "not recorded") or
                (col == "assessment" and db_val_str.lower() == "migrated from digiswasthya excel record")
            )
            excel_val_str = str(excel_val).strip() if excel_val is not None else ""
            excel_is_valid = (
                excel_val is not None and 
                excel_val_str != "" and 
                not (col == "mother_tongue" and excel_val_str in ("0", "0.0")) and
                not (col == "chief_complaint" and excel_val_str.lower() == "not recorded") and
                not (col == "assessment" and excel_val_str.lower() == "migrated from digiswasthya excel record")
            )
            if db_is_empty and excel_is_valid:
                return True
                
        return False
    except Exception:
        return False


def _update_record_fields(conn, schema: str, table: str, record_id: str, transformed_row: dict[str, Any], uid: str, log: list, force_legacy_source: bool = False) -> bool:
    """
    Queries the existing database record and updates fields that are NULL/empty in the DB 
    but have valid values in the Excel transformed row.
    Returns True if updates were executed, False otherwise.
    """
    try:
        transformed_row = filter_columns_for_db(table, transformed_row)
        sql = f"SELECT * FROM {_table(schema, table)} WHERE id = :id LIMIT 1"
        db_row = conn.execute(text(sql), {"id": record_id}).mappings().first()
        if not db_row:
            return False
            
        # For patient records, only allow updates if existing legacy_source is "Digiswasthya_Database"
        if table == "patient" and db_row.get("legacy_source") != "Digiswasthya_Database":
            return False
            
        update_fields = {}
        has_actual_field_updates = False

        for col, excel_val in transformed_row.items():
            if col in ("id", "created_at", "migrated_at", "legacy_id", "legacy_source", "tenant_id", "is_demo"):
                continue
                
            db_val = db_row.get(col)
            
            # Special merging for vitals_json (JSONB)
            if col == "vitals_json":
                db_vitals = {}
                if db_val:
                    try:
                        db_vitals = json.loads(db_val) if isinstance(db_val, str) else db_val
                    except Exception:
                        pass
                excel_vitals = {}
                if excel_val:
                    try:
                        excel_vitals = json.loads(excel_val) if isinstance(excel_val, str) else excel_val
                    except Exception:
                        pass
                        
                if isinstance(db_vitals, dict) and isinstance(excel_vitals, dict):
                    merged_vitals = {**db_vitals}
                    vitals_updated = False
                    for k, v in excel_vitals.items():
                        if (db_vitals.get(k) is None or str(db_vitals.get(k)).strip() == "") and (v is not None and str(v).strip() != ""):
                            merged_vitals[k] = v
                            vitals_updated = True
                    if vitals_updated:
                        update_fields[col] = json.dumps(merged_vitals, default=str)
                        has_actual_field_updates = True
                        missing_keys = [k for k, v in excel_vitals.items() if db_vitals.get(k) is None or str(db_vitals.get(k)).strip() == ""]
                        print(f"      [UPDATE] Table '{table}', ID: {record_id} -> field 'vitals_json' merged missing keys: {missing_keys}")
                        log.append({
                            "uid": uid,
                            "field": f"{table}.vitals_json",
                            "issue": f"UPDATED_{table.upper()}_FIELD",
                            "raw": json.dumps(excel_vitals),
                            "detail": f"Merged missing vitals keys: {missing_keys}"
                        })
                continue
                
            # Standard column blank check (treating placeholder values as empty)
            db_val_str = str(db_val).strip() if db_val is not None else ""
            db_is_empty = (
                db_val is None or 
                db_val_str == "" or 
                (col == "mother_tongue" and db_val_str in ("0", "0.0")) or
                (col == "chief_complaint" and db_val_str.lower() == "not recorded") or
                (col == "assessment" and db_val_str.lower() == "migrated from digiswasthya excel record")
            )
            excel_val_str = str(excel_val).strip() if excel_val is not None else ""
            excel_is_valid = (
                excel_val is not None and 
                excel_val_str != "" and 
                not (col == "mother_tongue" and excel_val_str in ("0", "0.0")) and
                not (col == "chief_complaint" and excel_val_str.lower() == "not recorded") and
                not (col == "assessment" and excel_val_str.lower() == "migrated from digiswasthya excel record")
            )
            if db_is_empty and excel_is_valid:
                update_fields[col] = excel_val
                has_actual_field_updates = True
                print(f"      [UPDATE] Table '{table}', ID: {record_id} -> field '{col}' is empty/0 in DB, filling with Excel value: '{excel_val}'")
                log.append({
                    "uid": uid,
                    "field": f"{table}.{col}",
                    "issue": f"UPDATED_{table.upper()}_FIELD",
                    "raw": str(excel_val),
                    "detail": f"Updated empty/0 database field '{col}' in table '{table}' with Excel value"
                })
                
        should_update_legacy_source = False
        if table in ("patient", "visit", "prescription"):
            if has_actual_field_updates or force_legacy_source:
                if db_row.get("legacy_source") != "excel+database":
                    should_update_legacy_source = True
                    
        if should_update_legacy_source:
            update_fields["legacy_source"] = "excel+database"
            print(f"      [UPDATE] Table '{table}', ID: {record_id} -> field 'legacy_source' set to 'excel+database'")
            log.append({
                "uid": uid,
                "field": f"{table}.legacy_source",
                "issue": f"UPDATED_{table.upper()}_FIELD",
                "raw": "excel+database",
                "detail": f"Updated legacy_source to 'excel+database' for matched {table}"
            })
            
        if update_fields:
            set_clauses = ", ".join(f'"{col}" = :{col}' for col in update_fields.keys())
            update_sql = f"UPDATE {_table(schema, table)} SET {set_clauses} WHERE id = :id"
            conn.execute(text(update_sql), {**update_fields, "id": record_id})
            return True
            
        return False
    except Exception as exc:
        log.append({
            "uid": uid,
            "field": f"{table}_update",
            "issue": "UPDATE_ERROR",
            "raw": None,
            "detail": f"Failed to perform update check on table '{table}': {exc}"
        })
        return False


def _get_collision_range(start_str: str, end_str: str) -> tuple[date | None, date | None]:
    if not start_str or not end_str:
        return None, None
    try:
        start = clean_date(start_str, "start_date", "config", [])
        end = clean_date(end_str, "end_date", "config", [])
        if start and end:
            return start - timedelta(days=1), end + timedelta(days=1)
    except Exception:
        pass
    return None, None


# - LOGGING SYSTEM ------------------------------------------------------------

def save_execution_log(cfg, stats, warnings, phase_b_completed=True):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    if isinstance(cfg.excel_path, str) and (cfg.excel_path.startswith("http://") or cfg.excel_path.startswith("https://")):
        log_dir = Path("output")
    else:
        log_dir = cfg.excel_path.parent / "output"
        if not log_dir.exists():
            log_dir = Path("output")
    os.makedirs(log_dir, exist_ok=True)
    
    log_file = log_dir / f"supabase_update_mode_log_{timestamp}.txt"
    
    # Categorize logs
    duplicates = []
    not_null_warnings = []
    cleaning_warnings = []
    updates = []
    
    for w in warnings:
        issue = str(w.get("issue") or "")
        if issue.startswith("SKIPPED_"):
            duplicates.append(w)
        elif issue.startswith("UPDATED_"):
            updates.append(w)
        elif issue == "REQUIRED_FIELD_EMPTY":
            not_null_warnings.append(w)
        else:
            cleaning_warnings.append(w)
            
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n")
        f.write(f"   SUPABASE UPDATE-MODE SEQUENTIAL RUN EXECUTION LOG - {timestamp}\n")
        f.write("=" * 60 + "\n\n")
        
        f.write("--- SUMMARY STATISTICS ---\n")
        for k, v in stats.items():
            f.write(f"{k:<35}: {v}\n")
        f.write("\n")

        f.write("--- FIELD UPDATES PERFORMED (COLLISION DATES) ---\n")
        if not updates:
            f.write("No fields were updated in existing records during this run.\n")
        else:
            f.write(f"Total Fields Updated: {len(updates)}\n\n")
            f.write(f"{'UID':<15} | {'Table.Column':<25} | {'Updated Value':<25} | {'Detail':<45}\n")
            f.write("-" * 125 + "\n")
            for w in updates:
                uid = str(w.get("uid") or "N/A")
                field = str(w.get("field") or "N/A")
                raw = str(w.get("raw") or "N/A")
                detail = str(w.get("detail") or "N/A")
                f.write(f"{uid:<15} | {field:<25} | {raw:<25} | {detail:<45}\n")
        f.write("\n\n")
        
        f.write("--- DUPLICATES DETECTED & SKIPPED/REUSED ---\n")
        if not duplicates:
            f.write("No duplicate records detected or skipped during this run.\n")
        else:
            f.write(f"Total Duplicate Actions: {len(duplicates)}\n\n")
            f.write(f"{'UID':<15} | {'Table/Field':<15} | {'Action/Issue':<28} | {'Identifier/Value':<20} | {'Detail':<45}\n")
            f.write("-" * 125 + "\n")
            for w in duplicates:
                uid = str(w.get("uid") or "N/A")
                field = str(w.get("field") or "N/A")
                issue = str(w.get("issue") or "N/A")
                raw = str(w.get("raw") or "N/A")
                detail = str(w.get("detail") or "N/A")
                f.write(f"{uid:<15} | {field:<15} | {issue:<28} | {raw:<20} | {detail:<45}\n")
        f.write("\n\n")
        
        f.write("--- REQUIRED NOT-NULL FIELD CONSTRAINT WARNINGS ---\n")
        if not not_null_warnings:
            f.write("No missing required (NOT NULL) fields observed.\n")
        else:
            f.write(f"Total Required Empty: {len(not_null_warnings)}\n\n")
            f.write(f"{'UID':<15} | {'Table.Column':<20} | {'Issue':<25} | {'Raw Value':<15} | {'Detail':<45}\n")
            f.write("-" * 125 + "\n")
            for w in not_null_warnings:
                uid = str(w.get("uid") or "N/A")
                field = str(w.get("field") or "N/A")
                issue = str(w.get("issue") or "N/A")
                raw = str(w.get("raw") or "N/A")
                detail = str(w.get("detail") or "N/A")
                f.write(f"{uid:<15} | {field:<20} | {issue:<25} | {raw:<15} | {detail:<45}\n")
        f.write("\n\n")
        
        f.write("--- DATA CLEANING & VALIDATION WARNINGS ---\n")
        if not cleaning_warnings:
            f.write("No other cleaning warnings or validation issues observed!\n")
        else:
            f.write(f"Total Cleaning Warnings: {len(cleaning_warnings)}\n\n")
            f.write(f"{'UID':<15} | {'Field':<15} | {'Issue':<25} | {'Raw Value':<20} | {'Detail':<45}\n")
            f.write("-" * 125 + "\n")
            for w in cleaning_warnings:
                uid = str(w.get("uid") or "N/A")
                field = str(w.get("field") or "N/A")
                issue = str(w.get("issue") or "N/A")
                raw = str(w.get("raw") or "N/A")
                detail = str(w.get("detail") or "N/A")
                f.write(f"{uid:<15} | {field:<15} | {issue:<25} | {raw:<20} | {detail:<45}\n")
                
    print(f"\n[Log] Execution log successfully saved to: {log_file}")


# - MAIN MIGRATION PIPELINE ---------------------------------------------------

def main():
    print("=" * 60)
    print("   DIRECT SUPABASE UPDATE-MODE DATA MIGRATION")
    print("=" * 60)

    cfg = _read_config()
    is_url = isinstance(cfg.excel_path, str) and (cfg.excel_path.startswith("http://") or cfg.excel_path.startswith("https://"))
    if not is_url:
        if not cfg.excel_path.exists():
            print(f"[ERROR] ERROR: Raw source Excel spreadsheet not found at: {cfg.excel_path}")
            print("Please verify the configuration or place the Excel sheet in the workspace.")
            sys.exit(1)

    # Initialize Date-based Split Ranges (Commented out/disabled for full testing update/skip mode)
    # collision_start, collision_end = _get_collision_range(COLLISION_START_DATE_STR, COLLISION_END_DATE_STR)
    print("Running in full update/skip testing mode: all rows are processed for update checks regardless of date.")

    print(f"Connecting to database at {cfg.database_url.split('@')[-1]}...")
    engine = create_engine(cfg.database_url, pool_pre_ping=True)

    stats = {
        "migration_status": "STARTED",
        "total_excel_rows_loaded": 0,
        "blank_uids_dropped": 0,
        "excel_duplicates_dropped": 0,
        "date_filtered_rows_dropped": 0,
        "total_migratable_records": 0,
        "new_patients_inserted": 0,
        "existing_patients_reused": 0,
        "patients_skipped_not_found": 0,
        "visits_inserted": 0,
        "visits_skipped_duplicate": 0,
        "prescriptions_inserted": 0,
        "prescriptions_skipped_duplicate": 0,
        "medication_items_inserted": 0,
        "medication_items_skipped_duplicate": 0,
        "followups_inserted": 0,
        "followups_skipped_duplicate": 0,
    }
    migration_log = []

    db_centers = []
    db_users = []
    user_index_by_employee = {}
    user_index_by_legacy = {}
    user_index_by_clean_name = {}
    db_doctors = []
    all_coordinators = []
    coordinators_by_center = {}
    coordinator_counters = {}
    resolved_center_id = None
    df = pd.DataFrame()

    # Pre-flight metadata fetch
    try:
        with engine.begin() as conn:
            print("\n[1/5] Fetching live metadata from Supabase database...")
            
            # Centers
            center_sql = f"SELECT id, name, center_code, district, state FROM {_table(cfg.schema, 'center')}"
            db_centers = [dict(r) for r in conn.execute(text(center_sql)).mappings().all()]
            print(f"      Loaded {len(db_centers)} clinic centers from Supabase center table.")
            
            # Users
            user_sql = f"SELECT id, role, first_name, last_name, employee_id, legacy_id, center_id FROM {_table(cfg.schema, 'user')}"
            db_users = [dict(r) for r in conn.execute(text(user_sql)).mappings().all()]
            print(f"      Loaded {len(db_users)} user accounts (doctors/coordinators) from Supabase user table.")
            
            if not db_centers:
                print("[ERROR] ERROR: Pre-populated 'center' table is empty in Supabase! Please set up at least one center before running.")
                sys.exit(1)
            if not db_users:
                print("[ERROR] ERROR: Pre-populated 'user' table is empty in Supabase! Please set up doctors/coordinators before running.")
                sys.exit(1)

            for u in db_users:
                emp_id = _norm_text(u.get("employee_id"))
                leg_id = _norm_text(u.get("legacy_id"))
                
                full_name = f"{u.get('first_name') or ''} {u.get('last_name') or ''}"
                clean_db_name = clean_doctor_name(full_name)
                
                if emp_id:
                    user_index_by_employee[emp_id] = u
                if leg_id:
                    user_index_by_legacy[leg_id] = u
                if clean_db_name:
                    user_index_by_clean_name[clean_db_name] = u
                if str(u.get("role")).upper() == "DOCTOR":
                    db_doctors.append(u)
                if str(u.get("role")).upper() in ("COORDINATOR", "STAFF", "ADMIN"):
                    all_coordinators.append(str(u["id"]))

            coordinators_by_center, coordinator_counters = _build_center_coordinator_rotation(db_users)
            
            # Query database column metadata dynamically
            for table_name in ("patient", "visit", "prescription", "prescriptionitem", "followup_schedule"):
                col_sql = """
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_schema = :schema AND table_name = :table
                """
                col_res = conn.execute(text(col_sql), {"schema": cfg.schema, "table": table_name})
                DB_TABLE_COLUMNS[table_name] = {r[0] for r in col_res.all()}
            print("      Loaded database table columns dynamically for schema safety.")
    except Exception as exc:
        print(f"\n[ERROR] CRITICAL ERROR OCCURRED during metadata fetch: {exc}")
        sys.exit(1)

    # Load and clean excel data from all selected sheets
    if is_url:
        print(f"\n[2/5] Loading source Excel spreadsheet from URL: {cfg.excel_path}...")
    else:
        print(f"\n[2/5] Loading source Excel spreadsheet: {cfg.excel_path.resolve()}...")
    
    excel_source = load_excel_from_path(cfg.excel_path)
    xl = pd.ExcelFile(excel_source)
    sheet_names = xl.sheet_names
    
    if isinstance(cfg.sheet_name, str) and cfg.sheet_name.lower() != "all" and cfg.sheet_name in sheet_names:
        sheets_to_process = [cfg.sheet_name]
    elif isinstance(cfg.sheet_name, int) and cfg.sheet_name < len(sheet_names) and cfg.sheet_name != 0:
        sheets_to_process = [sheet_names[cfg.sheet_name]]
    else:
        sheets_to_process = sheet_names
        
    print(f"      Selected sheets to process: {sheets_to_process}")
    
    excel_row_data = []
    total_excel_rows_loaded = 0
    blank_uids_dropped = 0
    excel_duplicates_dropped = 0
    date_filtered_rows_dropped = 0
    
    for sheet_name in sheets_to_process:
        print(f"\n   --- Processing sheet: '{sheet_name}' ---")
        if hasattr(excel_source, "seek"):
            excel_source.seek(0)
        sheet_df = pd.read_excel(excel_source, sheet_name=sheet_name, header=0)
        sheet_df = sheet_df.dropna(how="all").copy()
        
        rows_loaded = len(sheet_df)
        total_excel_rows_loaded += rows_loaded
        print(f"      Loaded {rows_loaded} rows from sheet '{sheet_name}'.")
        
        if rows_loaded == 0:
            continue
            
        # Resolve center ID for this sheet
        resolved_center = _resolve_center_row_for_sheet(db_centers, sheet_name)
        resolved_center_id = str(resolved_center.get("id")) if resolved_center else db_centers[0]["id"]
        print(f"      Mapped sheet '{sheet_name}' -> Center: '{resolved_center.get('name')}' ({resolved_center_id})")
        
        # Resolve UID column name dynamically
        uid_column = _find_column(sheet_df, "UID", "ma ", "CONSULTATION ID", "CONSULTATION ID\n", "CONSULTATION ID\nDS1", "CONSULTATION ID\nDS5", "PATIENT ID", "Id", "ID", "Unnamed: 0")
        if uid_column is not None:
            uid_mask = sheet_df[uid_column].apply(lambda value: not pd.isna(value) and str(value).strip() != "")
            dropped = int((~uid_mask).sum())
            if dropped:
                print(f"      Dropped {dropped} rows missing UID values in column '{uid_column}'.")
                blank_uids_dropped += dropped
            sheet_df = sheet_df.loc[uid_mask].copy()
        else:
            print(f"      [WARNING] Could not identify a UID/Consultation ID column for sheet '{sheet_name}'.")
            
        # Detect duplicates
        rows_before_dedup = len(sheet_df)
        sheet_df = detect_duplicates(sheet_df, migration_log)
        excel_duplicates_dropped += (rows_before_dedup - len(sheet_df))
        
        # Date Cutoff
        if EXCEL_ONLY_CUTOFF_DATE:
            date_column = _find_column(sheet_df, "DATE", "Date", "Visit Date", "Consultation Date", "DATE OF VISIT")
            if date_column is not None:
                cutoff = pd.to_datetime(EXCEL_ONLY_CUTOFF_DATE).date()
                parsed_dates = pd.to_datetime(sheet_df[date_column], errors="coerce", dayfirst=True)
                keep_mask = parsed_dates.dt.date < cutoff
                filtered_out = int((~keep_mask).sum())
                sheet_df = sheet_df.loc[keep_mask].copy()
                print(f"      Date Cutoff: removed {filtered_out} rows.")
                date_filtered_rows_dropped += filtered_out
                
        # Transform and package row data for this sheet
        for idx, (_, row) in enumerate(sheet_df.iterrows(), start=1):
            raw = row.to_dict()
            
            # Resolve doctor
            doctor_legacy_id = _first_existing(raw, "DOCTOR'S ID", "DOCTOR ID", "Doctor ID")
            doctor_row = None
            if doctor_legacy_id:
                norm_doc_id = _norm_text(doctor_legacy_id)
                doctor_row = user_index_by_employee.get(norm_doc_id) or user_index_by_legacy.get(norm_doc_id)
            if not doctor_row:
                doctor_name_col = _find_column(sheet_df, "DOCTOR'S NAME", "DOCTOR NAME", "Doctor Name")
                if doctor_name_col:
                    excel_doc_name = raw.get(doctor_name_col)
                    clean_excel_name = clean_doctor_name(excel_doc_name)
                    if clean_excel_name:
                        doctor_row = user_index_by_clean_name.get(clean_excel_name)
            
            doctor_uuid = str(doctor_row.get("id")) if doctor_row else None
            
            prescription_doctor_uuid = doctor_uuid
            if not prescription_doctor_uuid:
                if db_doctors:
                    prescription_doctor_uuid = db_doctors[0]["id"]
                else:
                    prescription_doctor_uuid = db_users[0]["id"]
                    
            created_by_user_id = _pick_center_coordinator(resolved_center_id, coordinators_by_center, coordinator_counters, db_users[0]["id"], all_coordinators)
            
            # Transform patient
            patient_row = transform_patient(raw, migration_log, center_id=resolved_center_id)
            
            raw_age = _row_lookup(raw, "AGE", "Age")
            age_years, age_months = parse_age_years_months(raw_age)
            patient_row["age_years"] = age_years
            patient_row["age_months"] = age_months
            
            raw_date = _row_lookup(raw, "DATE", "Date")
            uid = resolve_uid(raw)
            if not uid:
                uid = f"unassigned_{uuid.uuid4().hex[:8]}"
            
            custom_gov_id = generate_custom_government_id(patient_row.get("full_name"), raw_age, raw_date, uid)
            patient_row["government_id"] = custom_gov_id
            
            excel_row_data.append({
                "raw": raw,
                "uid": uid,
                "patient_row": patient_row,
                "orig_patient_id": patient_row["id"],
                "doctor_uuid": doctor_uuid,
                "prescription_doctor_uuid": prescription_doctor_uuid,
                "created_by_user_id": created_by_user_id,
                "resolved_center_id": resolved_center_id,
                "is_collision": True,  # Always True to perform update-or-insert check
            })
            
    stats["total_excel_rows_loaded"] = total_excel_rows_loaded
    stats["blank_uids_dropped"] = blank_uids_dropped
    stats["excel_duplicates_dropped"] = excel_duplicates_dropped
    stats["date_filtered_rows_dropped"] = date_filtered_rows_dropped
    stats["total_migratable_records"] = len(excel_row_data)
    
    print(f"\nTotal aggregated records to migrate: {len(excel_row_data)}")

    # 3. Phase A: Clean, transform, and insert patients table live (Transaction 1)
    print("\n[3/5] PHASE A: PROCESSING AND INSERTING PATIENTS TABLE...")
    patient_id_map: dict[str, str] = {}
    inserted_patients = 0
    updated_patients_count = 0
    first_5_batch = excel_row_data[:5]
    remaining_batch = excel_row_data[5:]

    def process_patient_batch(batch, start_idx):
        nonlocal inserted_patients, updated_patients_count
        with engine.begin() as conn:
            for idx, item in enumerate(batch, start=start_idx):
                uid = item["uid"]
                patient_row = item["patient_row"]
                
                # Check for duplicates in Supabase database
                matched_ids = _find_all_matching_patient_ids(conn, cfg.schema, patient_row)
                
                if matched_ids:
                    # 1. Check if any matched record has any actual field updates
                    any_row_has_updates = any(
                        _has_actual_updates(conn, cfg.schema, "patient", pid, patient_row)
                        for pid in matched_ids
                    )
                    
                    # 2. Update all matched duplicate database records
                    for pid in matched_ids:
                        _update_record_fields(
                            conn, 
                            cfg.schema, 
                            "patient", 
                            pid, 
                            patient_row, 
                            uid, 
                            migration_log,
                            force_legacy_source=any_row_has_updates
                        )
                        
                    # Map to the first matched patient ID for downstream tables
                    resolved_id = matched_ids[0]
                    patient_id_map[patient_row["id"]] = resolved_id
                    item["resolved_patient_ids"] = matched_ids
                    updated_patients_count += len(matched_ids)
                    stats["existing_patients_reused"] += 1
                    
                    migration_log.append({
                        "uid": uid,
                        "field": "patient",
                        "issue": "SKIPPED_PATIENT_DUPLICATE",
                        "raw": patient_row.get("full_name"),
                        "detail": f"Patient matched {len(matched_ids)} database records by Name + Phone. Updated all. Reused ID: {resolved_id}"
                    })
                else:
                    # 2. Insert new patient record if not found
                    if not patient_row.get("id"):
                        patient_row["id"] = str(uuid.uuid4())
                        
                    columns = list(patient_row.keys())
                    placeholders = ", ".join(f":{col}" for col in columns)
                    column_sql = ", ".join(f'"{col}"' for col in columns)
                    sql = f"INSERT INTO {_table(cfg.schema, 'patient')} ({column_sql}) VALUES ({placeholders}) RETURNING id"
                    
                    result = conn.execute(text(sql), patient_row).mappings().first()
                    inserted_id = str(result["id"])
                    
                    patient_id_map[patient_row["id"]] = inserted_id
                    item["resolved_patient_ids"] = [inserted_id]
                    inserted_patients += 1
                    stats["new_patients_inserted"] += 1
                    
                    migration_log.append({
                        "uid": uid,
                        "field": "patient",
                        "issue": "INSERTED_PATIENT_NEW",
                        "raw": patient_row.get("full_name"),
                        "detail": f"Patient '{patient_row.get('full_name')}' not found in database. Registered new patient with ID: {inserted_id}"
                    })
                    
                if idx % 100 == 0 or idx == len(excel_row_data):
                    print(f"      Processed {idx}/{len(excel_row_data)} patients...")

    try:
        # Process first 5 records
        print("\n--> Migrating first 5 patient records...")
        process_patient_batch(first_5_batch, 1)
        print("First 5 patient records committed successfully.")

        # Pause and ask if we should continue
        is_non_interactive = not sys.stdin.isatty() or os.getenv("NON_INTERACTIVE") == "1"
        if remaining_batch:
            if is_non_interactive:
                print("Running in non-interactive mode. Auto-proceeding with remaining patient records...")
                choice_pat = "y"
            else:
                choice_pat = "n"
                try:
                    choice_pat = input(
                        f"\nFirst 5 patient records have been processed and committed.\n"
                        f"Do you want to proceed with migrating the remaining {len(remaining_batch)} patients? (y/n): "
                    ).strip().lower()
                except (EOFError, KeyboardInterrupt):
                    pass
            
            if choice_pat not in ("y", "yes"):
                print("\nGracefully stopping as requested. First 5 patients committed.")
                stats["migration_status"] = "COMPLETED_FIRST_5_PATIENTS_ONLY"
                save_execution_log(cfg, stats, migration_log, phase_b_completed=False)
                return

            print("\n--> Migrating remaining patient records...")
            process_patient_batch(remaining_batch, 6)

        print("\n[SUCCESS] PATIENT TABLE MIGRATION COMPLETE!")
        print(f"      - New Patients Inserted: {inserted_patients}")
        print(f"      - Existing Patient Records Updated: {updated_patients_count}")
        print(f"      - Total Mapped Patients: {len(patient_id_map)}")

    except Exception as exc:
        print(f"\n[ERROR] CRITICAL ERROR OCCURRED during Patient Migration: {exc}")
        print("Patient transaction has been rolled back. No changes were committed.")
        stats["migration_status"] = f"FAILED_PHASE_A: {type(exc).__name__}"
        migration_log.append({
            "uid": "SYSTEM",
            "field": "patients_transaction",
            "issue": "MIGRATION_FAILED",
            "raw": str(exc),
            "detail": f"Patient phase migration failure: {str(exc)}"
        })
        save_execution_log(cfg, stats, migration_log, phase_b_completed=False)
        raise exc

    # Patients have been successfully committed to Supabase!
    # 4. Interactive Terminal Validation Prompt
    print("\n" + "=" * 60)
    
    # Auto-proceed if standard input is not a TTY or NON_INTERACTIVE is set to '1'
    is_non_interactive = not sys.stdin.isatty() or os.getenv("NON_INTERACTIVE") == "1"
    
    if is_non_interactive:
        print("Running in non-interactive mode. Auto-proceeding with migrating remaining tables...")
        choice = "y"
    else:
        choice = "n"
        try:
            choice = input(
                "Patient filling completed and committed successfully to Supabase!\n"
                "Do you want to proceed with migrating the remaining tables (visits, prescriptions, prescription items, follow-up schedules)? (y/n): "
            ).strip().lower()
        except (EOFError, KeyboardInterrupt):
            pass
    print("=" * 60)

    if choice not in ("y", "yes"):
        print("\nGracefully stopping as requested. Patient database changes committed.")
        print("Transactional tables (visit, prescription, etc.) skipped. Exiting script.")
        stats["migration_status"] = "COMPLETED_PHASE_A_ONLY"
        save_execution_log(cfg, stats, migration_log, phase_b_completed=False)
        return

    # 5. Phase B: Run transactional tables live mapping (visits, prescriptions, prescription items, follow-up schedules)
    print("\n[5/5] PHASE B: PROCESSING AND MIGRATING VISITS, PRESCRIPTIONS, AND TRANSACTIONAL TABLES...")
    
    visit_id_map = {}
    inserted_visits = 0
    skipped_visits = 0

    prescription_id_map = {}
    inserted_rx = 0
    skipped_rx = 0

    inserted_items = 0
    skipped_items = 0

    inserted_followups = 0
    skipped_followups = 0

    def process_phase_b_batch(batch, start_idx):
        nonlocal inserted_visits, skipped_visits, inserted_rx, skipped_rx, inserted_items, skipped_items, inserted_followups, skipped_followups
        with engine.begin() as conn:
            # --- LOOP 1: VISITS ---
            print(f"\nMigrating 'visit' table for batch starting at index {start_idx}...")
            for offset, item in enumerate(batch):
                idx = start_idx + offset
                raw = item["raw"]
                orig_pat_id = item["orig_patient_id"]
                if orig_pat_id not in patient_id_map:
                    # Skip visit because parent patient was not found/matched in DB
                    skipped_visits += 1
                    stats["visits_skipped_duplicate"] += 1
                    continue
                supabase_patient_id = patient_id_map[orig_pat_id]
                is_collision = item["is_collision"]
                
                visit_row = transform_visit(
                    raw, 
                    supabase_patient_id, 
                    item["doctor_uuid"], 
                    raw.get("DATE", raw.get("Date")), 
                    migration_log, 
                    center_id=item["resolved_center_id"], 
                    created_by_user_id=item["created_by_user_id"]
                )
                
                # Validate NOT NULL fields
                _validate_not_null_fields(
                    "visit",
                    visit_row,
                    ["patient_id", "center_id", "created_by_user_id", "created_at", "status"],
                    item["uid"],
                    migration_log
                )
                
                patient_row = item["patient_row"]
                pat_ids = _find_matching_patient_ids_with_age_fallback(conn, cfg.schema, patient_row)
                if pat_ids:
                    supabase_patient_id = pat_ids[0]
                else:
                    supabase_patient_id = patient_id_map[orig_pat_id]
                    pat_ids = [supabase_patient_id]
                
                # Update mappings
                item["resolved_patient_ids"] = pat_ids
                patient_id_map[orig_pat_id] = supabase_patient_id
                visit_row["patient_id"] = supabase_patient_id
                
                # Check for duplicate visits based on patient_ids, doctor_uuid, and visit_date
                sql_dup = f"""
                    SELECT id FROM {_table(cfg.schema, 'visit')}
                    WHERE patient_id IN :patient_ids
                      AND assigned_doctor_user_id IS NOT DISTINCT FROM :doctor_uuid
                      AND CAST(created_at AS date) = CAST(:visit_date AS date)
                """
                params_dup = {
                    "patient_ids": tuple(pat_ids),
                    "doctor_uuid": item["doctor_uuid"],
                    "visit_date": visit_row.get("created_at")
                }
                candidates = [dict(r) for r in conn.execute(text(sql_dup), params_dup).mappings().all()]
                matched_visit_ids = [str(cand["id"]) for cand in candidates]
                item["matched_visit_ids"] = matched_visit_ids
                
                if matched_visit_ids:
                    item["visit_existed"] = True
                    # Do NOT update duplicate visit or map/store it for downstream insertions.
                    # This ensures no visits or downstream records (prescriptions, followup schedule) are updated/inserted.
                    skipped_visits += 1
                    stats["visits_skipped_duplicate"] += 1
                    migration_log.append({
                        "uid": item["uid"],
                        "field": "visit",
                        "issue": "SKIPPED_VISIT_DUPLICATE",
                        "raw": str(visit_row.get("visit_date")),
                        "detail": f"Visit matched {len(matched_visit_ids)} database records by Patient + Doctor + Date. Skipped update and all downstream tables processing."
                    })
                else:
                    item["visit_existed"] = False
                    # 2. Insert new visit record
                    resolved_supabase_id = _ensure_legacy_row(conn, cfg, "visit", visit_row)
                    visit_id_map[visit_row["id"]] = resolved_supabase_id
                    item["supabase_visit_id"] = resolved_supabase_id
                    item["matched_visit_ids"] = [resolved_supabase_id]
                    
                    inserted_visits += 1
                    stats["visits_inserted"] += 1
                    
                if idx % 100 == 0 or idx == start_idx + len(batch) - 1:
                    print(f"      Processed {idx}/{len(excel_row_data)} visits...")
            print(f"      Visits complete: {inserted_visits} new, {skipped_visits} skipped.")
            
            # --- LOOP 2: PRESCRIPTIONS ---
            print(f"\nMigrating 'prescription' table for batch starting at index {start_idx}...")
            for offset, item in enumerate(batch):
                idx = start_idx + offset
                raw = item["raw"]
                if "supabase_visit_id" not in item:
                    # Skip prescription because parent visit was skipped
                    skipped_rx += 1
                    stats["prescriptions_skipped_duplicate"] += 1
                    continue
                supabase_visit_id = item["supabase_visit_id"]
                is_collision = item["is_collision"]
                
                rx_row = transform_prescription(
                    raw, 
                    supabase_visit_id, 
                    item["prescription_doctor_uuid"], 
                    raw.get("DATE", raw.get("Date")), 
                    migration_log
                )
                
                _validate_not_null_fields(
                    "prescription",
                    rx_row,
                    ["visit_id", "doctor_user_id", "created_at"],
                    item["uid"],
                    migration_log
                )
                
                visit_ids = item.get("matched_visit_ids", [supabase_visit_id])
                
                # Check for duplicate prescriptions based on visit_ids
                sql_rx_dup = f"SELECT id FROM {_table(cfg.schema, 'prescription')} WHERE visit_id IN :visit_ids"
                rx_candidates = [dict(r) for r in conn.execute(text(sql_rx_dup), {"visit_ids": tuple(visit_ids)}).mappings().all()]
                matched_rx_ids = [str(cand["id"]) for cand in rx_candidates]
                item["matched_prescription_ids"] = matched_rx_ids
                
                if matched_rx_ids:
                    # 1. Update all matched duplicate prescriptions
                    for rx_id in matched_rx_ids:
                        _update_record_fields(conn, cfg.schema, "prescription", rx_id, rx_row, item["uid"], migration_log)
                        
                    resolved_supabase_id = matched_rx_ids[0]
                    prescription_id_map[rx_row["id"]] = resolved_supabase_id
                    item["supabase_prescription_id"] = resolved_supabase_id
                    
                    skipped_rx += 1
                    stats["prescriptions_skipped_duplicate"] += 1
                    migration_log.append({
                        "uid": item["uid"],
                        "field": "prescription",
                        "issue": "SKIPPED_PRESCRIPTION_DUPLICATE",
                        "raw": str(rx_row.get("created_at")),
                        "detail": f"Prescription matched {len(matched_rx_ids)} database records by Visit ID. Updated all. Reused ID: {resolved_supabase_id}"
                    })
                else:
                    # 2. Insert new prescription record
                    resolved_supabase_id = _ensure_legacy_row(conn, cfg, "prescription", rx_row)
                    prescription_id_map[rx_row["id"]] = resolved_supabase_id
                    item["supabase_prescription_id"] = resolved_supabase_id
                    item["matched_prescription_ids"] = [resolved_supabase_id]
                    
                    inserted_rx += 1
                    stats["prescriptions_inserted"] += 1
                    
                if idx % 100 == 0 or idx == start_idx + len(batch) - 1:
                    print(f"      Processed {idx}/{len(excel_row_data)} prescriptions...")
            print(f"      Prescriptions complete: {inserted_rx} new, {skipped_rx} skipped.")

            # --- LOOP 3: PRESCRIPTION ITEMS ---
            print(f"\nMigrating 'prescriptionitem' table for batch starting at index {start_idx}...")
            for offset, item in enumerate(batch):
                idx = start_idx + offset
                raw = item["raw"]
                if "supabase_prescription_id" not in item:
                    # Skip prescription items because parent prescription was skipped
                    continue
                if item.get("visit_existed", False):
                    # Skip prescription items because visit already existed in DB
                    continue
                supabase_rx_id = item["supabase_prescription_id"]
                
                items_list = transform_prescription_items(raw, supabase_rx_id, migration_log)
                for item_row in items_list:
                    item_row = clean_row_for_db(item_row)
                    
                    # Validate NOT NULL fields
                    _validate_not_null_fields(
                        "prescriptionitem",
                        item_row,
                        ["prescription_id", "medication_name", "dosage", "frequency"],
                        item["uid"],
                        migration_log
                    )
                    
                    # Since visit/prescription is brand new, there are no duplicates in DB
                    matched_item_ids = []
                    
                    if not matched_item_ids:
                        if not item_row.get("id"):
                            item_row["id"] = str(uuid.uuid4())
                        item_row = filter_columns_for_db("prescriptionitem", item_row)
                        columns = list(item_row.keys())
                        placeholders = ", ".join(f":{col}" for col in columns)
                        column_sql = ", ".join(f'"{col}"' for col in columns)
                        conn.execute(text(f"INSERT INTO {_table(cfg.schema, 'prescriptionitem')} ({column_sql}) VALUES ({placeholders})"), item_row)
                        inserted_items += 1
                        stats["medication_items_inserted"] += 1
                    else:
                        # Skip updating duplicate record
                        skipped_items += 1
                        stats["medication_items_skipped_duplicate"] += 1
                        migration_log.append({
                            "uid": item["uid"],
                            "field": "prescriptionitem",
                            "issue": "SKIPPED_ITEM_DUPLICATE",
                            "raw": item_row.get("medication_name"),
                            "detail": f"Item '{item_row.get('medication_name')}' matched {len(matched_item_ids)} database records by Prescription + Name + Details. Skipping update."
                        })
                        
                if idx % 100 == 0 or idx == start_idx + len(batch) - 1:
                    print(f"      Processed {idx}/{len(excel_row_data)} medication lists...")
            print(f"      Medication Items complete: {inserted_items} new, {skipped_items} skipped.")

            # --- LOOP 4: FOLLOW-UP SCHEDULES ---
            print(f"\nMigrating 'followup_schedule' table for batch starting at index {start_idx}...")
            for offset, item in enumerate(batch):
                idx = start_idx + offset
                raw = item["raw"]
                orig_pat_id = item["orig_patient_id"]
                if orig_pat_id not in patient_id_map or "supabase_prescription_id" not in item:
                    # Skip followup schedule because parent patient/prescription was skipped
                    continue
                if item.get("visit_existed", False):
                    # Skip followup schedules because visit already existed in DB
                    continue
                supabase_patient_id = patient_id_map[orig_pat_id]
                supabase_rx_id = item["supabase_prescription_id"]
                
                followups_list = transform_followup_schedules(raw, supabase_patient_id, supabase_rx_id, migration_log, item.get("supabase_visit_id"))
                for f_row in followups_list:
                    f_row = clean_row_for_db(f_row)
                    if f_row.get("resolution"):
                        f_row["resolution"] = f_row["resolution"][:40]
                    if f_row.get("status"):
                        f_row["status"] = f_row["status"][:20]
                    
                    # Validate NOT NULL fields
                    _validate_not_null_fields(
                        "followup_schedule",
                        f_row,
                        ["patient_id", "source_prescription_id", "sequence_no", "scheduled_date", "status"],
                        item["uid"],
                        migration_log
                    )
                    
                    # Since visit/prescription is brand new, there are no duplicates in DB
                    matched_f_ids = []
                    
                    if not matched_f_ids:
                        if not f_row.get("id"):
                            f_row["id"] = str(uuid.uuid4())
                        f_row = filter_columns_for_db("followup_schedule", f_row)
                        columns = list(f_row.keys())
                        placeholders = ", ".join(f":{col}" for col in columns)
                        column_sql = ", ".join(f'"{col}"' for col in columns)
                        conn.execute(text(f"INSERT INTO {_table(cfg.schema, 'followup_schedule')} ({column_sql}) VALUES ({placeholders})"), f_row)
                        inserted_followups += 1
                        stats["followups_inserted"] += 1
                    else:
                        # Skip updating duplicate record
                        skipped_followups += 1
                        stats["followups_skipped_duplicate"] += 1
                        migration_log.append({
                            "uid": item["uid"],
                            "field": "followup_schedule",
                            "issue": "SKIPPED_FOLLOWUP_DUPLICATE",
                            "raw": str(f_row.get("sequence_no")),
                            "detail": f"Followup schedule matched {len(matched_f_ids)} database records by Patient + Prescription + Sequence. Skipping update."
                        })
                if idx % 100 == 0 or idx == start_idx + len(batch) - 1:
                    print(f"      Processed {idx}/{len(excel_row_data)} follow-up plans...")
            print(f"      Followups complete: {inserted_followups} new, {skipped_followups} skipped.")

    try:
        # Process first 5 records of Phase B
        print("\n--> Migrating first 5 visits and downstream records...")
        process_phase_b_batch(first_5_batch, 1)
        print("First 5 visits and downstream records committed successfully.")

        # Pause and ask if we should continue
        is_non_interactive = not sys.stdin.isatty() or os.getenv("NON_INTERACTIVE") == "1"
        if remaining_batch:
            if is_non_interactive:
                print("Running in non-interactive mode. Auto-proceeding with migrating remaining records...")
                choice_rem = "y"
            else:
                choice_rem = "n"
                try:
                    choice_rem = input(
                        f"\nFirst 5 visits and downstream records have been processed and committed.\n"
                        f"Do you want to proceed with migrating the remaining {len(remaining_batch)} records? (y/n): "
                    ).strip().lower()
                except (EOFError, KeyboardInterrupt):
                    pass
            
            if choice_rem not in ("y", "yes"):
                print("\nGracefully stopping as requested. First 5 visits committed.")
                stats["migration_status"] = "COMPLETED_FIRST_5_VISITS_ONLY"
                save_execution_log(cfg, stats, migration_log, phase_b_completed=True)
                return

            print("\n--> Migrating remaining visits and downstream records...")
            process_phase_b_batch(remaining_batch, 6)

        print("\n" + "=" * 60)
        print("  SUCCESS: DIRECT SUPABASE UPDATE-MODE DATA MIGRATION COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        
        stats["migration_status"] = "SUCCESS_FULL"
        save_execution_log(cfg, stats, migration_log, phase_b_completed=True)

    except Exception as exc:
        print(f"\n[ERROR] CRITICAL ERROR OCCURRED during Phase B Transactional Migration: {exc}")
        print("Transactional changes (visits, prescriptions, items, follow-ups) have been rolled back.")
        print("Patient changes remain committed.")
        stats["migration_status"] = f"FAILED_PHASE_B: {type(exc).__name__}"
        migration_log.append({
            "uid": "SYSTEM",
            "field": "phase_b_transactional_tables",
            "issue": "MIGRATION_FAILED",
            "raw": str(exc),
            "detail": f"Phase B migration failure: {str(exc)}"
        })
        save_execution_log(cfg, stats, migration_log, phase_b_completed=False)
        raise exc


if __name__ == "__main__":
    main()
