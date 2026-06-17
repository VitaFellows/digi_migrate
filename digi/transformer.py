# transformer.py
# Transforms cleaned data into rows ready for each DB table.
# Each function takes a cleaned row dict and returns a dict matching the DB table schema.
# All column names match the new application DB exactly.
# UUID generation happens here using uuid.uuid4().
#
# INSERT ORDER (FK dependency chain - NEVER change this order):
# 1. center     - no FK deps, must be pre-inserted manually (see config.py)
# 2. user       - FK -> center
# 3. patient    - FK -> center
# 4. visit      - FK -> patient, center, user
# 5. prescription - FK -> visit, user
# 6. prescriptionitem - FK -> prescription (up to 5 per Excel row)
# 7. followup_schedule - FK -> patient, prescription (up to 5 per Excel row)
#
# AWS DB INSERT CODE is written as comments - not executed.
# All output goes to the output Excel workbook for review first.

from datetime import date, datetime, time, timezone
import json
import re
import uuid
from typing import Any

try:
    from cleaner import *
    from cleaner import _is_blank, _safe_float, _safe_int
    from config import (
        CENTER_ID,
        FALLBACK_DOCTOR_USER_ID,
        LEGACY_SOURCE,
        MIGRATION_SYSTEM_USER_ID,
        TENANT_ID,
        IS_DEMO,
    )
    from government_id import generate_government_id
    from output_builder import TABLE_COLUMNS
except ImportError:
    from .cleaner import *
    from .cleaner import _is_blank, _safe_float, _safe_int
    from .config import (
        CENTER_ID,
        FALLBACK_DOCTOR_USER_ID,
        LEGACY_SOURCE,
        MIGRATION_SYSTEM_USER_ID,
        TENANT_ID,
        IS_DEMO,
    )
    from .government_id import generate_government_id
    from .output_builder import TABLE_COLUMNS

# - DOCTOR LOOKUP TABLE -------------------------------------------------------
_doctor_uuid_cache: dict[str, str] = {}


def get_or_create_doctor_uuid(doctor_legacy_id: str | None) -> str:
    """
    Look up doctor UUID from cache.
    """
    if not doctor_legacy_id:
        return FALLBACK_DOCTOR_USER_ID
    if doctor_legacy_id not in _doctor_uuid_cache:
        _doctor_uuid_cache[doctor_legacy_id] = str(uuid.uuid4())
    return _doctor_uuid_cache[doctor_legacy_id]


# - GENERIC HELPERS -----------------------------------------------------------

def _blank_row(table_name: str) -> dict:
    return {column: None for column in TABLE_COLUMNS[table_name]}


def _norm(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def _row_lookup(row: dict, *candidates: str, default: Any = None) -> Any:
    normalized = {}
    for key, value in row.items():
        key_norm = _norm(key)
        if key_norm not in normalized:
            normalized[key_norm] = value
    for candidate in candidates:
        candidate_norm = _norm(candidate)
        if candidate_norm in normalized:
            value = normalized[candidate_norm]
            if not _is_blank(value):
                return value
    # Fallback for merged/noisy headers where candidate text is a subset.
    for candidate in candidates:
        candidate_norm = _norm(candidate)
        if not candidate_norm:
            continue
        for key_norm, value in normalized.items():
            if candidate_norm in key_norm:
                if not _is_blank(value):
                    return value
    return default


def _has_referral_value(value: Any) -> bool:
    cleaned = clean_text(value)
    if cleaned is None:
        return False
    return cleaned.lower() not in {"no", "nil", "n/a", "na", "none", "null", "-"}


def _row_items(row: dict) -> list[tuple[str, Any]]:
    return list(row.items())


def _row_values(row: dict) -> list[Any]:
    return list(row.values())


def _iso_datetime(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        text = clean_text(value)
        if text is None:
            return None
        for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                parsed = datetime.strptime(text, fmt)
                if parsed.tzinfo is None:
                    parsed = parsed.replace(tzinfo=timezone.utc)
                return parsed.astimezone(timezone.utc).isoformat()
            except Exception:
                pass
        cleaned_date = clean_date(text, "datetime", "transform", [])
        if cleaned_date is not None:
            return datetime.combine(cleaned_date, time.min, tzinfo=timezone.utc).isoformat()
        return None
    if isinstance(value, datetime):
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.astimezone(timezone.utc).isoformat()
    if isinstance(value, date):
        return datetime.combine(value, time.min, tzinfo=timezone.utc).isoformat()
    cleaned_date = clean_date(value, "datetime", "transform", [])
    if cleaned_date is not None:
        return datetime.combine(cleaned_date, time.min, tzinfo=timezone.utc).isoformat()
    return None


def _safe_run(log: list, uid: str | None, field: str, func, fallback=None):
    try:
        return func()
    except Exception as exc:
        log.append({"uid": uid, "field": field, "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        return fallback


def discover_column_positions(df):
    """
    Print the source columns with indices and infer likely medicine/follow-up blocks.
    """
    print("[discover_column_positions] Source columns:")
    for index, column in enumerate(df.columns.tolist()):
        print(f"  {index:>3}: {column}")

    medicine_starts = []
    for index, column in enumerate(df.columns.tolist()):
        normalised = _norm(column)
        if re.fullmatch(r"medicine\d+", normalised):
            medicine_starts.append(index)
    if medicine_starts:
        print(f"[discover_column_positions] Likely medicine block starts: {medicine_starts}")
    return {"medicine_starts": medicine_starts}


# - PATIENT TRANSFORMER -------------------------------------------------------

def transform_patient(row: dict, log: list, center_id: str | None = None) -> dict:
    """
    Transform a cleaned Excel row into a patient table row dict.
    """
    try:
        patient = _blank_row("patient")
        uid = str(_row_lookup(row, "UID", "Id", default="") or "")
        raw_name = _row_lookup(row, "PATIENT NAME", "Patient Name", "Name", "PATIENT")
        name_cleaned = clean_name(raw_name, uid, log)
        name = name_cleaned or "Unknown"
        if raw_name is None or name_cleaned is None:
            log.append({"uid": uid, "field": "full_name", "issue": "MISSING_PATIENT_NAME", "raw": raw_name, "detail": "Fallback to Unknown"})

        raw_phone = _row_lookup(row, "CONTACT NO.", "CONTACT NO", "PHONE", "Phone", "Mobile")
        cleaned_phone = clean_phone(raw_phone, uid, log)
        sex = clean_sex(_row_lookup(row, "GENDER", "Sex", "SEX"), uid, log)
        age_years = _safe_run(log, uid, "age_years", lambda: _safe_int(_row_lookup(row, "AGE", "Age")))
        address = clean_text(_row_lookup(row, "ADDRESS", "Address"), max_len=1000)
        district = clean_text(_row_lookup(row, "DISTRICT", "District"), max_len=100)
        state = clean_text(_row_lookup(row, "STATE", "State"), max_len=100)
        occupation = clean_occupation(_row_lookup(row, "OCCUPATION", "Occupation"))
        category = clean_category(_row_lookup(row, "CATEGORY", "Category"))
        mother_tongue = clean_text(_row_lookup(row, "MOTHER TONGUE", "Mother Tongue"), max_len=100)
        created_date = clean_date(_row_lookup(row, "DATE", "Date", "VISIT DATE", "Visit Date"), "DATE", uid, log)
        created_at = datetime.combine(created_date, time.min, tzinfo=timezone.utc).isoformat() if created_date else None

        patient.update({
            "id": str(uuid.uuid4()),
            "full_name": name,
            "phone": cleaned_phone,
            "sex": sex,
            "age_years": age_years,
            "address_text": address,
            "address_line_1": address,
            "district": district,
            "state": state,
            "occupation": occupation,
            "category": category,
            "mother_tongue": mother_tongue,
            "government_id": generate_government_id(name, cleaned_phone, uid, _row_lookup(row, "AGE", "Age")),
            "center_id": center_id or CENTER_ID,
            "created_at": created_at,
            "patient_status": "CURRENT",
            "is_demo": IS_DEMO,
            "tenant_id": TENANT_ID,
            "phone_last10": cleaned_phone[-10:] if cleaned_phone else None,
            "legacy_id": uid,
            "legacy_source": LEGACY_SOURCE,
            "migrated_at": datetime.now(timezone.utc).isoformat(),
        })
        return patient
    except Exception as exc:
        uid = str(_row_lookup(row, "UID", default="") or "")
        log.append({"uid": uid, "field": "patient", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        patient = _blank_row("patient")
        patient.update({"id": str(uuid.uuid4()), "full_name": "Unknown", "center_id": center_id or CENTER_ID, "created_at": datetime.now(timezone.utc).isoformat(), "patient_status": "CURRENT", "is_demo": IS_DEMO, "legacy_id": uid, "legacy_source": LEGACY_SOURCE, "migrated_at": datetime.now(timezone.utc).isoformat(), "tenant_id": TENANT_ID})
        return patient


# - USER (DOCTOR) TRANSFORMER -------------------------------------------------

def transform_doctor(doctor_legacy_id: str, doctor_name: str | None, speciality: str | None, log: list) -> dict:
    """
    Transform doctor info from Excel into a user table row.
    """
    try:
        user = _blank_row("user")
        first_name, last_name = split_doctor_name(clean_name(doctor_name, doctor_legacy_id, log) if doctor_name is not None else None)
        user.update({
            "id": get_or_create_doctor_uuid(doctor_legacy_id),
            "role": "DOCTOR",
            "first_name": first_name,
            "last_name": last_name,
            "employee_id": doctor_legacy_id,
            "phone": f"PLACEHOLDER-{doctor_legacy_id}" if doctor_legacy_id else None,
            "email": f"{doctor_legacy_id}@placeholder.internal" if doctor_legacy_id else None,
            "password_hash": "PLACEHOLDER_HASH",
            "center_id": CENTER_ID,
            "is_active": True,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "specialization": clean_text(speciality, max_len=255),
            "qualifications": None,
            "is_available_now": False,
            "notes": None,
            "avatar_url": None,
            "is_away": False,
            "away_since": None,
            "fcm_token": None,
            "google_meet_link": None,
            "zoom_link": None,
            "preferred_video_platform": None,
            "signature_url": None,
            "licence_number": None,
            "pin_code": None,
            "per_consultation_fee": None,
            "other_charges": None,
            "government_id": None,
            "doctor_review_preference": False,
            "allowed_video_platforms": None,
            "tenant_id": TENANT_ID,
            "ltc_suggest_unplanned_chronic": False,
            "ltc_suggest_frequent_flyer": False,
            "legacy_id": doctor_legacy_id,
            "legacy_source": LEGACY_SOURCE,
        })
        return user
    except Exception as exc:
        log.append({"uid": doctor_legacy_id, "field": "doctor", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        user = _blank_row("user")
        user.update({"id": get_or_create_doctor_uuid(doctor_legacy_id), "role": "DOCTOR", "first_name": "Unknown", "last_name": "Doctor", "employee_id": doctor_legacy_id, "phone": f"PLACEHOLDER-{doctor_legacy_id}", "email": f"{doctor_legacy_id}@placeholder.internal", "password_hash": "PLACEHOLDER_HASH", "center_id": CENTER_ID, "is_active": True, "created_at": datetime.now(timezone.utc).isoformat(), "is_available_now": False, "is_away": False, "doctor_review_preference": False, "ltc_suggest_unplanned_chronic": False, "ltc_suggest_frequent_flyer": False, "legacy_id": doctor_legacy_id, "legacy_source": LEGACY_SOURCE, "tenant_id": TENANT_ID})
        return user


# - VISIT TRANSFORMER --------------------------------------------------------

def transform_visit(row: dict, patient_uuid: str, doctor_uuid: str, visit_date: Any, log: list, center_id: str | None = None, created_by_user_id: str | None = None) -> dict:
    """
    Transform Excel row into a visit table row.
    """
    try:
        visit = _blank_row("visit")
        uid = str(_row_lookup(row, "UID", default="") or "")
        visit_dt = _iso_datetime(visit_date)
        vitals_dict = clean_vitals(row, uid, log)
        # Check if we have Kathaicha-style complaints (multi-complaint sheets)
        keys = list(row.keys())
        complaint1_idx = -1
        for idx, key in enumerate(keys):
            if _norm(key) == "complaint1":
                complaint1_idx = idx
                break

        if complaint1_idx != -1:
            complaints_list = []
            primary_duration = None
            for comp_num in range(1, 5):
                base_idx = complaint1_idx + (comp_num - 1) * 3
                if base_idx < len(keys):
                    text_val = clean_text(row.get(keys[base_idx]))
                    dur_num_val = row.get(keys[base_idx + 1]) if (base_idx + 1) < len(keys) else None
                    dur_unit_val = row.get(keys[base_idx + 2]) if (base_idx + 2) < len(keys) else None
                    
                    dur_num_clean = clean_text(dur_num_val)
                    dur_unit_clean = clean_text(dur_unit_val)
                    
                    is_dur_num_text = False
                    if dur_num_clean:
                        try:
                            float(dur_num_clean)
                        except ValueError:
                            if len(dur_num_clean) > 3 or re.search(r'[a-zA-Z\s]{4,}', dur_num_clean):
                                is_dur_num_text = True
                    
                    if is_dur_num_text:
                        complaint_text = dur_num_clean
                        duration_str = None
                    else:
                        complaint_text = text_val
                        duration_parts = []
                        if dur_num_clean:
                            duration_parts.append(dur_num_clean)
                        if dur_unit_clean:
                            duration_parts.append(dur_unit_clean)
                        duration_str = " ".join(duration_parts) if duration_parts else None
                    
                    if complaint_text:
                        if duration_str:
                            complaints_list.append(f"{complaint_text} ({duration_str})")
                            if comp_num == 1:
                                primary_duration = duration_str
                        else:
                            complaints_list.append(complaint_text)
            
            chief_complaint = ", ".join(complaints_list) if complaints_list else "Not recorded"
            duration_val = primary_duration
        else:
            chief_complaint = clean_text(
                _row_lookup(
                    row,
                    "CHIEF COMPLAIN HISTORY",
                    "CHIEF COMPLAINT",
                    "Chief Complaint",
                    "Advice//Recommendation CHIEF COMPLAIN HISTORY",
                    "Unnamed: 28",
                ),
                max_len=2000,
            ) or "Not recorded"
            # Prioritize Since/Duration columns if present in the sheet
            has_direct_duration_col = False
            duration_col_key = None
            for k in keys:
                k_norm = _norm(k)
                if k_norm in ("since", "duration"):
                    has_direct_duration_col = True
                    duration_col_key = k
                    break

            if has_direct_duration_col:
                duration_val = clean_text(row.get(duration_col_key))
            else:
                duration_val = None
                chief_candidates = {
                    "chiefcomplainhistory",
                    "chiefcomplainthistory",
                    "chiefcomplaint",
                    "advicerecommendationchiefcomplainhistory",
                    "unnamed28",
                    "unnamed29",
                    "adviceclinicalnotesrecommendationchiefcomplainhistory",
                    "recommendationchiefcomplainhistory"
                }
                chief_idx = -1
                for idx, k in enumerate(keys):
                    if _norm(k) in chief_candidates:
                        if not k.lower().startswith("unnamed"):
                            chief_idx = idx
                            break
                if chief_idx != -1 and chief_idx + 1 < len(keys):
                    next_key = keys[chief_idx + 1]
                    if "unnamed" in str(next_key).lower():
                        duration_val = clean_text(row.get(next_key))

        visit.update({
            "id": str(uuid.uuid4()),
            "patient_id": patient_uuid,
            "center_id": center_id or CENTER_ID,
            "created_by_user_id": created_by_user_id or MIGRATION_SYSTEM_USER_ID,
            "assigned_doctor_user_id": doctor_uuid,
            "status": "COMPLETED",
            "chief_complaint": chief_complaint,
            "complaint_duration": clean_text(duration_val, max_len=255) if duration_val else None,
            "history_notes": None,
            "is_reconsultation": clean_consultation_type(_row_lookup(row, "New/Re-Consultation", "CONSULTATION TYPE", "Consultation Type")),
            "vitals_json": json.dumps(vitals_dict or {}, default=str),
            "created_at": visit_dt,
            "updated_at": visit_dt,
            "completed_at": visit_dt,
            "tenant_id": TENANT_ID,
            "legacy_id": uid if uid else str(uuid.uuid4()),
            "legacy_source": LEGACY_SOURCE,
            "consultation_type": None,
        })
        return visit
    except Exception as exc:
        uid = str(_row_lookup(row, "UID", default="") or "")
        log.append({"uid": uid, "field": "visit", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        visit = _blank_row("visit")
        visit.update({"id": str(uuid.uuid4()), "patient_id": patient_uuid, "center_id": center_id or CENTER_ID, "created_by_user_id": created_by_user_id or MIGRATION_SYSTEM_USER_ID, "status": "COMPLETED", "chief_complaint": "Not recorded", "is_reconsultation": False, "vitals_json": json.dumps({}, default=str), "created_at": datetime.now(timezone.utc).isoformat(), "updated_at": datetime.now(timezone.utc).isoformat(), "completed_at": datetime.now(timezone.utc).isoformat(), "tenant_id": TENANT_ID, "legacy_id": uid, "legacy_source": LEGACY_SOURCE})
        return visit


# - PRESCRIPTION TRANSFORMER -------------------------------------------------

def transform_prescription(row: dict, visit_uuid: str, doctor_uuid: str, visit_date: Any, log: list) -> dict:
    """
    Transform Excel row into a prescription table row.
    """
    try:
        prescription = _blank_row("prescription")
        uid = str(_row_lookup(row, "UID", default="") or "")
        advice = clean_text(_row_lookup(row, "Advice//Recommendation", "Advice / Recommendation", "Advice", "Recommendation"), max_len=4000)
        
        confirmed_raw = _row_lookup(row, "CONFIRM DIAGNOSIS", "Confirmed Diagnosis", "DIAGNOSIS", "Diagnosis")
        confirmed_clean = clean_diagnosis(confirmed_raw)
        
        provisional_raw = _row_lookup(row, "Provisional diagnosis", "Provisional Diagnosis", "provisional daignosis", "Provisional Daignosis", "TYPES OF DISEASE", "Types of Disease")
        provisional_clean = clean_diagnosis(provisional_raw)
        
        if confirmed_clean:
            assessment = f"Confirmed: {confirmed_clean}"
        elif provisional_clean:
            assessment = f"Provisional: {provisional_clean}"
        else:
            assessment = advice or "Migrated from DigiSwasthya Excel record"
        referral_name = clean_text(
            _row_lookup(
                row,
                "Name of the Tertiary Treatment Recommened (If Yes)",
                "Name of the Tertiary Treatment Recommened",
                "Name of Tertiary Treatment Recommended (If Yes)",
                "Name of Tertiary Treatment Recommended",
                "TERTIARY TREATMENT",
                "Referral Treatment Name",
            ),
            max_len=1000,
        )
        specialist_name = clean_text(
            _row_lookup(
                row,
                "Name of the 2nd Specialist Consultation (If Recommended by Doctor)",
                "Name of 2nd Specialist Consultation",
                "SECOND SPECIALIST",
                "Specialist Name",
            ),
            max_len=1000,
        )
        instruction_text = advice
        patient_instructions = clean_text(
            _row_lookup(
                row,
                "Advice",
                "Advice near follow-ups section",
                "Advice Near Follow Ups",
                "Patient Instructions",
                "Follow Up Instructions",
            ),
            max_len=4000,
        )
        investigation_notes = []
        for n in range(1, 7):
            note = clean_text(_row_lookup(row, f"Test Referrals {n}", f"Test Referral {n}", f"Referral {n}"), max_len=500)
            if note:
                investigation_notes.append(note)
        follow_up_at = _row_lookup(row, "Follow Up -1 Date", "Follow Up 1 Date", "FollowUp 1 Date")
        prescription.update({
            "id": str(uuid.uuid4()),
            "visit_id": visit_uuid,
            "doctor_user_id": doctor_uuid or FALLBACK_DOCTOR_USER_ID,
            "assessment": assessment,
            "provisional_diagnosis": clean_diagnosis(_row_lookup(row, "Provisional diagnosis", "Provisional Diagnosis", "provisional daignosis", "Provisional Daignosis", "TYPES OF DISEASE", "Types of Disease")),
            "confirmed_diagnosis": clean_diagnosis(_row_lookup(row, "CONFIRM DIAGNOSIS", "Confirmed Diagnosis", "DIAGNOSIS", "Diagnosis")),
            "is_referral_recommended": _has_referral_value(referral_name) or _has_referral_value(specialist_name),
            "referral_treatment_name": referral_name,
            "referral_specialist_name": specialist_name,
            "instructions": instruction_text,
            "follow_up_at": _iso_datetime(follow_up_at),
            "created_at": _iso_datetime(visit_date),
            "investigation_notes": " | ".join(investigation_notes) if investigation_notes else None,
            "deleted_at": None,
            "patient_instructions": patient_instructions,
            "legacy_id": uid if uid else str(uuid.uuid4()),
            "legacy_source": LEGACY_SOURCE,
        })
        return prescription
    except Exception as exc:
        uid = str(_row_lookup(row, "UID", default="") or "")
        log.append({"uid": uid, "field": "prescription", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        prescription = _blank_row("prescription")
        prescription.update({"id": str(uuid.uuid4()), "visit_id": visit_uuid, "doctor_user_id": doctor_uuid or FALLBACK_DOCTOR_USER_ID, "assessment": "Migrated from DigiSwasthya Excel record", "is_referral_recommended": False, "created_at": _iso_datetime(visit_date), "legacy_id": uid if uid else str(uuid.uuid4()), "legacy_source": LEGACY_SOURCE})
        return prescription


# - PRESCRIPTION ITEM TRANSFORMER --------------------------------------------

def _infer_block_start_indices(keys: list[str], block_label: str) -> list[int]:
    starts = []
    for idx, key in enumerate(keys):
        key_norm = _norm(key)
        if block_label == "medicine":
            if re.fullmatch(r"medicine\d+", key_norm):
                starts.append(idx)
        elif block_label in key_norm:
            starts.append(idx)
    return starts


def _extract_block_value(keys: list[str], values: list[Any], start_index: int, offset: int) -> Any:
    index = start_index + offset
    if 0 <= index < len(values):
        return values[index]
    return None


def transform_prescription_items(row: dict, prescription_uuid: str, log: list) -> list[dict]:
    """
    Transform medicine blocks (1-5) from Excel row into prescriptionitem rows.
    """
    items: list[dict] = []
    try:
        keys = list(row.keys())
        values = list(row.values())
        start_indices = _infer_block_start_indices(keys, "medicine")
        if not start_indices:
            for index, key in enumerate(keys):
                if _norm(key).startswith("med"):
                    start_indices.append(index)
        seen_block_starts = []
        for start_index in start_indices:
            if start_index in seen_block_starts:
                continue
            seen_block_starts.append(start_index)
            if len(seen_block_starts) > 5:
                break
            medication_name = clean_medicine_name(_extract_block_value(keys, values, start_index, 0))
            if medication_name is None:
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
            })
            items.append(item)
        return items
    except Exception as exc:
        uid = str(_row_lookup(row, "UID", default="") or "")
        log.append({"uid": uid, "field": "prescriptionitem", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        return items


# - FOLLOW-UP SCHEDULE TRANSFORMER -------------------------------------------

def _find_followup_block_value(row: dict, block_no: int, kind: str) -> Any:
    keys = list(row.keys())
    values = list(row.values())

    date_index = None
    patterns = [
        f"followup{block_no}date",
        f"followup-{block_no}date",
        f"followup{block_no}",
    ]
    for idx, key in enumerate(keys):
        key_norm = _norm(key)
        if f"followup-{block_no}date" in key_norm or f"followup{block_no}date" in key_norm:
            date_index = idx
            break
    if date_index is None:
        # Defensive fallback: for block 1 headers often have no numeric suffix in exports.
        if block_no == 1:
            for idx, key in enumerate(keys):
                key_norm = _norm(key)
                if "followup" in key_norm and "date" in key_norm and "-2" not in str(key) and ".1" not in str(key):
                    date_index = idx
                    break
    if date_index is None:
        return None

    if kind == "date":
        return values[date_index]
    if kind == "status":
        return values[date_index + 1] if date_index + 1 < len(values) else None
    if kind == "remark":
        return values[date_index + 2] if date_index + 2 < len(values) else None
    return None


def _get_final_remark(row: dict) -> str | None:
    keys = list(row.keys())
    # Find the index of the first follow-up column to avoid matching medicine remarks
    first_followup_idx = -1
    for idx, k in enumerate(keys):
        k_norm = _norm(k)
        if "followup" in k_norm:
            first_followup_idx = idx
            break
            
    limit = first_followup_idx if first_followup_idx != -1 else 0
    for idx in range(len(keys) - 1, limit - 1, -1):
        k = keys[idx]
        k_norm = _norm(k)
        if k_norm in ("finalremark", "finalremarks", "remark"):
            val = row.get(k)
            if not _is_blank(val):
                return clean_text(val)
    return None


def _get_patient_feedback(row: dict) -> str | None:
    keys = list(row.keys())
    # Find the index of the first follow-up column to avoid matching any feedback columns before follow-ups
    first_followup_idx = -1
    for idx, k in enumerate(keys):
        k_norm = _norm(k)
        if "followup" in k_norm:
            first_followup_idx = idx
            break
            
    limit = first_followup_idx if first_followup_idx != -1 else 0
    for idx in range(limit, len(keys)):
        k = keys[idx]
        k_norm = _norm(k)
        if k_norm in ("patientfeedbackduringfollowup", "patientfeedback", "feedback"):
            val = row.get(k)
            if not _is_blank(val):
                return clean_text(val)
    return None


def transform_followup_schedules(row: dict, patient_uuid: str, prescription_uuid: str, log: list, visit_uuid: str | None = None) -> list[dict]:
    """
    Transform follow-up blocks (1-5) from Excel row into followup_schedule rows.
    """
    schedules: list[dict] = []
    try:
        block_data = []
        last_non_empty = None
        for block_no in range(1, 6):
            scheduled_date = clean_date(_find_followup_block_value(row, block_no, "date"), f"Follow Up -{block_no} Date", str(_row_lookup(row, "UID", default="") or ""), log)
            status = clean_followup_status(_find_followup_block_value(row, block_no, "status"))
            notes = clean_text(_find_followup_block_value(row, block_no, "remark"), max_len=500)
            if scheduled_date is not None:
                last_non_empty = block_no
            block_data.append((block_no, scheduled_date, status, notes))

        recovery_status = clean_text(_row_lookup(row, "PATIENT RECOVERY STATUS", "Patient Recovery Status"), max_len=500)
        final_status = clean_text(_row_lookup(row, "Final Follow Up Status", "FINAL FOLLOW UP STATUS", "Follow Up Final Status"), max_len=500)
        patient_feedback = clean_text(_get_patient_feedback(row), max_len=1000)
        final_remark = clean_text(_get_final_remark(row), max_len=1000)

        for block_no, scheduled_date, status, notes in block_data:
            if scheduled_date is None:
                continue
            schedule = _blank_row("followup_schedule")
            resolution = None
            resolution_notes = None
            if last_non_empty == block_no:
                resolution_parts = []
                if recovery_status:
                    resolution_parts.append(f"Recovery: {recovery_status}")
                if final_status:
                    resolution_parts.append(f"Final: {final_status}")
                resolution = " | ".join(resolution_parts) if resolution_parts else None
                notes_parts = []
                if patient_feedback:
                    notes_parts.append(patient_feedback)
                if final_remark:
                    notes_parts.append(final_remark)
                resolution_notes = " | ".join(notes_parts) if notes_parts else None
            schedule.update({
                "id": str(uuid.uuid4()),
                "treatment_plan_id": None,
                "sequence_no": block_no,
                "scheduled_date": _iso_datetime(scheduled_date),
                "status": status,
                "completed_visit_id": visit_uuid,
                "assigned_coordinator_id": None,
                "reminder_sent_at": None,
                "rescheduled_from_id": None,
                "cancelled_reason": None,
                "notes": notes,
                "created_at": datetime.now(timezone.utc).isoformat(),
                "updated_at": datetime.now(timezone.utc).isoformat(),
                "patient_id": patient_uuid,
                "source_prescription_id": prescription_uuid,
                "resolution": resolution,
                "resolution_notes": resolution_notes,
                "resolved_by_user_id": None,
                "resolved_at": None,
            })
            schedules.append(schedule)
        return schedules
    except Exception as exc:
        uid = str(_row_lookup(row, "UID", default="") or "")
        log.append({"uid": uid, "field": "followup_schedule", "issue": "TRANSFORM_ERROR", "raw": None, "detail": str(exc)})
        return schedules
