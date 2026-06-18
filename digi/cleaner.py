# cleaner.py
# All data cleaning functions - applied BEFORE transformation.
# Each function takes a raw value and returns a cleaned value (or None).
# All functions are pure (no side effects) and safe to call with None input.
# A cleaning_log list is passed around to record every issue found.

import hashlib
import re
from datetime import date, datetime
from typing import Any

import pandas as pd

try:
    from dateutil import parser as dateutil_parser
except Exception:
    dateutil_parser = None


BLANK_VALUES = {"", "na", "n/a", "nil", "none", "null", "-"}


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if pd.isna(value):
        return True
    if isinstance(value, str):
        return value.strip().lower() in BLANK_VALUES
    return False


def _clean_whitespace(value: Any) -> str:
    text = str(value).replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _normalise_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def _log(log: list, uid: str | None, field: str, issue: str, raw: Any, detail: str | None = None):
    log.append({
        "uid": uid,
        "field": field,
        "issue": issue,
        "raw": raw,
        "detail": detail,
    })


def _safe_float(value: Any) -> float | None:
    if _is_blank(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = _clean_whitespace(value)
    text = text.replace(",", "")
    try:
        return float(text)
    except Exception:
        return None


def _safe_int(value: Any) -> int | None:
    if _is_blank(value):
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if pd.isna(value):
            return None
        return int(round(value))
    text = _clean_whitespace(value)
    text = re.sub(r"[^0-9\-]", "", text)
    if not text or text in {"-", "--"}:
        return None
    try:
        return int(text)
    except Exception:
        return None


# - DATE CLEANING -------------------------------------------------------------

def clean_date(value: Any, field_name: str, uid: str, log: list) -> date | None:
    """
    Convert any date representation to a Python date object.
    """
    if _is_blank(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.date()

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            result = pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))
            return result.date()
        except Exception as exc:
            _log(log, uid, field_name, "INVALID_DATE_SERIAL", value, str(exc))
            return None

    if isinstance(value, str):
        text = _clean_whitespace(value)
        if text.lower() in BLANK_VALUES:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except Exception:
                pass
        if text.isdigit():
            try:
                result = pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(text))
                return result.date()
            except Exception:
                pass
        if dateutil_parser is not None:
            try:
                return dateutil_parser.parse(text, dayfirst=True).date()
            except Exception:
                pass
        _log(log, uid, field_name, "INVALID_DATE", value, "Unable to parse date")
        return None

    _log(log, uid, field_name, "INVALID_DATE", value, f"Unsupported type: {type(value).__name__}")
    return None


# - PHONE CLEANING ------------------------------------------------------------

def clean_phone(value: Any, uid: str, log: list) -> str | None:
    """
    Normalize phone number to 10-digit string.
    """
    if _is_blank(value):
        return None

    if isinstance(value, bool):
        _log(log, uid, "phone", "INVALID_PHONE", value)
        return None

    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        if float(value).is_integer():
            digits = str(int(value))
        else:
            digits = re.sub(r"\D", "", str(value))
    else:
        text = _clean_whitespace(value)
        if text.endswith(".0") and text.replace(".0", "").isdigit():
            digits = text[:-2]
        else:
            digits = re.sub(r"\D", "", text)

    if digits in {"1", "0"} or len(digits) < 7:
        _log(log, uid, "phone", "INVALID_PHONE", value)
        return None
    if len(digits) == 10:
        return digits
    if len(digits) == 11 and digits.startswith("0"):
        return digits[-10:]
    if len(digits) == 12 and digits.startswith("91"):
        return digits[-10:]
    _log(log, uid, "phone", "INVALID_PHONE", value)
    return None


# - NAME CLEANING -------------------------------------------------------------

def clean_name(value: Any, uid: str, log: list) -> str | None:
    """
    Clean patient or doctor name.
    """
    if _is_blank(value):
        _log(log, uid, "name", "EMPTY_NAME", value)
        return None
    text = _clean_whitespace(value)
    if len(text) > 500:
        _log(log, uid, "name", "NAME_TRUNCATED", value)
        text = text[:500]
    if text.lower() in BLANK_VALUES:
        _log(log, uid, "name", "EMPTY_NAME", value)
        return None
    return text


def split_doctor_name(full_name: str | None) -> tuple[str, str]:
    """
    Split full doctor name into (first_name, last_name).
    """
    if full_name is None:
        return ("Unknown", "Doctor")
    text = _clean_whitespace(full_name)
    text = re.sub(r"^dr\.?\s*", "", text, flags=re.IGNORECASE)
    text = _clean_whitespace(text)
    if not text:
        return ("Unknown", "Doctor")
    parts = text.split(" ", 1)
    if len(parts) == 1:
        return (parts[0], ".")
    return (parts[0], parts[1])


# - GENERAL TEXT CLEANING -----------------------------------------------------

def clean_text(value: Any, max_len: int | None = None) -> str | None:
    """
    General text field cleaner.
    """
    if _is_blank(value):
        return None
    text = _clean_whitespace(value)
    if text.lower() in BLANK_VALUES or text == "-":
        return None
    if max_len is not None and len(text) > max_len:
        text = text[: max(0, max_len - 3)] + "..." if max_len >= 3 else text[:max_len]
    return text


def clean_occupation(value: Any) -> str | None:
    """
    Clean and normalize occupation values.
    """
    TYPO_MAP = {
        "retaird": "Retired",
        "self empoyeed": "Self Employed",
        "self employed": "Self Employed",
        "house wife": "Home Maker",
        "pvt job": "Private Job",
        "private job": "Private Job",
    }
    cleaned = clean_text(value)
    if cleaned is None:
        return None
    mapped = TYPO_MAP.get(cleaned.lower())
    return mapped if mapped else cleaned


def clean_category(value: Any) -> str | None:
    """
    Normalize caste/category values.
    """
    cleaned = clean_text(value)
    if cleaned is None:
        return None
    lowered = cleaned.lower()
    if lowered in {"chrisition", "christion"}:
        return "Christian"
    if cleaned.upper() in {"OBC", "SC", "VJNT", "ST", "EWS", "NT", "NT1", "NT2", "NT3"}:
        return cleaned.upper()
    return cleaned[:1].upper() + cleaned[1:].lower() if cleaned.islower() or cleaned.isupper() else cleaned


# - DIAGNOSIS CLEANING --------------------------------------------------------

def clean_diagnosis(value: Any) -> str | None:
    """
    Clean diagnosis text.
    """
    cleaned = clean_text(value)
    if cleaned is None:
        return None
    if cleaned.lower() in {"na", "n/a", "nil", "nil.", "nil ", "nilnil"}:
        return None
    return cleaned


# - CONSULTATION TYPE ---------------------------------------------------------

def clean_consultation_type(value: Any) -> bool:
    """
    Parse New/Re-Consultation value into boolean for is_reconsultation.
    """
    if _is_blank(value):
        return False
    text = clean_text(value)
    if text is None:
        return False
    return "re" in text.lower()


# - MEDICINE BLOCK CLEANING ---------------------------------------------------

def clean_medicine_name(value: Any) -> str | None:
    """
    Clean medication name.
    """
    cleaned = clean_text(value)
    if cleaned is None:
        return None
    if cleaned.lower() in {"nil", "na", "n/a", "none", "-", "no", "not applicable", "null"}:
        return None
    return cleaned


def build_frequency_string(morning: Any, afternoon: Any, evening: Any, remark: Any) -> str:
    """
    Build a human-readable frequency string from MAE columns.
    """
    remark_text = clean_text(remark)
    if remark_text:
        if re.fullmatch(r"(?i)(OD|BD|TID|QID|SOS|PRN|HS|ODD|STAT|ONCE|TWICE|THRICE|DAILY|WEEKLY|MONTHLY|PRN\.?|HS\.?|QHS|QOD|QID\.?|BID|TDS)", remark_text.strip()):
            return remark_text.strip()

    def flag(value: Any) -> int:
        num = _safe_int(value)
        if num is None:
            return 0
        return 1 if num > 0 else 0

    m = flag(morning)
    a = flag(afternoon)
    e = flag(evening)

    combo = (m, a, e)
    mapping = {
        (1, 0, 0): "Once daily (Morning)",
        (0, 1, 0): "Once daily (Afternoon)",
        (0, 0, 1): "Once daily (Evening)",
        (1, 0, 1): "Twice daily (Morning + Evening)",
        (1, 1, 0): "Twice daily (Morning + Afternoon)",
        (0, 1, 1): "Twice daily (Afternoon + Evening)",
        (1, 1, 1): "Three times daily",
        (0, 0, 0): "As directed",
    }
    return mapping.get(combo, f"M:{m} A:{a} E:{e}")


# - FOLLOW-UP STATUS ----------------------------------------------------------

def clean_followup_status(value: Any) -> str:
    """
    Normalize follow-up status to DB values.
    """
    text = clean_text(value)
    if text is None:
        return "COMPLETED"
    lowered = text.lower()
    if "done" in lowered or "complet" in lowered:
        return "COMPLETED"
    if "pend" in lowered:
        return "PENDING"
    return "COMPLETED"


# - SEX / GENDER --------------------------------------------------------------

def clean_sex(value: Any, uid: str, log: list) -> str | None:
    """
    Normalize sex/gender to DB enum values.
    """
    if _is_blank(value):
        return None
    text = clean_text(value)
    if text is None:
        return None
    lowered = text.lower()
    mapping = {
        "m": "MALE",
        "male": "MALE",
        "f": "FEMALE",
        "female": "FEMALE",
        "o": "OTHER",
        "other": "OTHER",
    }
    if lowered in mapping:
        return mapping[lowered]
    _log(log, uid, "sex", "UNKNOWN_SEX", value)
    return None


def _row_lookup(row: dict, *candidates: str, default: Any = None) -> Any:
    normalized = {}
    for key, value in row.items():
        key_norm = _normalise_key(key)
        if key_norm not in normalized:
            normalized[key_norm] = value
    for candidate in candidates:
        candidate_norm = _normalise_key(candidate)
        if candidate_norm in normalized:
            value = normalized[candidate_norm]
            if not _is_blank(value):
                return value
    # Fallback for merged/noisy headers where candidate text is a subset.
    for candidate in candidates:
        candidate_norm = _normalise_key(candidate)
        if not candidate_norm or len(candidate_norm) < 3:
            continue
        for key_norm, value in normalized.items():
            if candidate_norm in key_norm:
                if not _is_blank(value):
                    return value
    return default


# - VITALS CLEANING -----------------------------------------------------------

def clean_vitals(row: dict, uid: str, log: list) -> dict:
    """
    Build the vitals_json dict from raw Excel row values.
    """
    normalized = {_normalise_key(key): value for key, value in row.items()}
    def lookup(*candidates):
        return _row_lookup(row, *candidates)

    bp_value = clean_text(lookup("BP", "Blood Pressure", "B.P.", "BloodPressure"))
    
    # SPO2 check: read as float first to handle percentage values correctly
    raw_spo2 = _safe_float(lookup("SPO2", "SpO2", "Oxygen Saturation", "SpO2 %"))
    spo2_val = None
    if raw_spo2 is not None:
        if 0.0 < raw_spo2 <= 1.0:
            spo2_val = int(round(raw_spo2 * 100))
        else:
            spo2_val = int(round(raw_spo2))

    pulse = _safe_int(lookup("Pulse", "PULSE", "Heart Rate"))
    temp_val = _safe_float(lookup("TEMP", "Temperature", "Temp", "Temperature (F)", "Temp F", "Body Temperature"))
    height = _safe_float(lookup("HEIGHT", "Height", "Height (cm)", "Height CM"))
    weight = _safe_float(lookup("WEIGHT", "Weight", "Weight (kg)", "Weight KG"))
    sugar_val = _safe_float(lookup("SUGAR", "Sugar", "Blood Sugar", "RBS", "FBS", "Sugar (mg/dl)", "Sugar mg/dl"))

    def to_number(val):
        if val is None:
            return None
        if isinstance(val, (int, float)):
            if float(val).is_integer():
                return int(val)
            return round(float(val), 1)
        return val

    vitals = {
        "bp": bp_value or None,
        "spo2": to_number(spo2_val),
        "pulse": to_number(pulse),
        "temp_f": to_number(temp_val),
        "height_cm": to_number(height),
        "weight_kg": to_number(weight),
        "sugar_mg_dl": to_number(sugar_val)
    }

    return vitals


# - DUPLICATE DETECTION -------------------------------------------------------

def detect_duplicates(df: pd.DataFrame, log: list) -> pd.DataFrame:
    """
    Detect and remove duplicate rows from the source Excel DataFrame.
    """
    if df is None or df.empty:
        return df

    columns = {_normalise_key(col): col for col in df.columns}
    uid_col = columns.get("uid") or columns.get("id")
    patient_col = columns.get("patientname") or columns.get("name") or columns.get("patient") or columns.get("patientfullname")
    date_col = columns.get("date") or columns.get("visitdate") or columns.get("consultationdate")
    doctor_col = columns.get("doctorsid") or columns.get("doctorid") or columns.get("doctorlegacyid") or columns.get("doctorsidno")

    filtered = df.copy()

    if uid_col is not None:
        seen: dict[Any, int] = {}
        keep_mask = []
        for idx, value in enumerate(filtered[uid_col].tolist()):
            if _is_blank(value):
                keep_mask.append(True)
                continue
            key = str(value).strip()
            if key in seen:
                _log(log, str(value), "UID", "DUPLICATE_UID", value, f"Duplicate of row {seen[key] + 1}")
                keep_mask.append(False)
            else:
                seen[key] = idx
                keep_mask.append(True)
        filtered = filtered.loc[keep_mask].copy()

    if patient_col is not None and date_col is not None and doctor_col is not None:
        pair_seen: set[tuple[str, str, str]] = set()
        for _, row in filtered.iterrows():
            patient = clean_text(row.get(patient_col))
            visit_date = clean_date(row.get(date_col), "date", str(row.get(uid_col) if uid_col else None), log)
            doctor = clean_text(row.get(doctor_col))
            if patient and visit_date and doctor:
                key = (patient.lower(), visit_date.isoformat(), doctor.lower())
                if key in pair_seen:
                    _log(log, str(row.get(uid_col) if uid_col else None), "row", "POSSIBLE_DUPLICATE", None, "Same patient name + date + doctor")
                else:
                    pair_seen.add(key)

    return filtered
