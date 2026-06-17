# government_id.py
# Generates a synthetic government_id for patient records
# Real government IDs (Aadhaar, voter ID etc.) were not collected in the Excel data.
# We generate a deterministic identifier from available fields so the value is:
#   - Reproducible (same input always produces same output)
#   - Unique enough to distinguish patients
#   - Clearly synthetic (prefixed with GEN-)

import re
from typing import Any


_identity_to_uid: dict[tuple[str, str, str], str] = {}


def _normalize_text(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() in {"NULL", "NONE", "NA", "N/A"}:
        return None
    return re.sub(r"\s+", " ", text).upper()


def _normalize_phone(value: Any) -> str | None:
    if value is None:
        return None
    digits = re.sub(r"\D", "", str(value))
    if digits in {"", "0", "1"}:
        return None
    return digits


def _normalize_age(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.upper() in {"NULL", "NONE", "NA", "N/A"}:
        return None
    try:
        age = int(float(text))
    except Exception:
        return None
    if age < 0 or age > 120:
        return None
    return str(age)


def generate_government_id(
    full_name: str | None,
    phone: str | None,
    uid: str | None,
    age: Any | None = None,
) -> str | None:
    """
    Generate a patient government_id from the Excel UID.

    Priority:
    1. If full_name, phone, and age are all available:
       reuse the earliest UID seen for that identity group
    2. If UID exists:
       use the Excel UID as the government_id
    3. If nothing usable is available:
       return None

    The duplicate rule is based on the earliest row encountered in the workbook.
    """
    normalized_name = _normalize_text(full_name)
    normalized_phone = _normalize_phone(phone)
    normalized_age = _normalize_age(age)
    normalized_uid = _normalize_text(uid)

    if normalized_name and normalized_name != "UNKNOWN" and normalized_phone and normalized_age:
        identity_key = (normalized_name, normalized_phone, normalized_age)
        existing_uid = _identity_to_uid.get(identity_key)
        if existing_uid:
            return existing_uid
        if normalized_uid:
            _identity_to_uid[identity_key] = str(uid).strip()
            return str(uid).strip()

    if normalized_uid:
        return str(uid).strip()

    return None
