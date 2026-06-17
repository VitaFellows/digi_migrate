# config.py
# Central configuration — edit these values before running

# ── INPUT ──────────────────────────────────────────────────────────────────
INPUT_EXCEL_PATH = "sample.xlsx"             # Path to the DigiSwasthya source Excel file
INPUT_SHEET_NAME = "Narsala"                  # Sheet index (0 = first sheet) or sheet name string

# ── REFERENCE CSV TABLES ───────────────────────────────────────────────────
CENTERS_CSV_PATH = "01_centers.csv"
DOCTORS_CSV_PATH = "02_doctors.csv"
COORDINATORS_CSV_PATH = "03_coordinators.csv"
USER_CENTER_LINKS_CSV_PATH = "04_user_center_links.csv"   # Not used by the migration logic
COORDINATOR_DOCTOR_LINKS_CSV_PATH = "05_coordinator_doctor_links.csv"

# ── OUTPUT ─────────────────────────────────────────────────────────────────
OUTPUT_EXCEL_PATH = "../output/migration_output.xlsx"   # Output workbook path

# ── LEGACY TRACKING ─────────────────────────────────────────────────────────
LEGACY_SOURCE = "excel_digiswasthya"          # Stored in legacy_source column of every table
                                               # The old-DB migration team uses a DIFFERENT value
                                               # e.g. "legacy_db_digiswasthya"
                                               # This lets us always know which rows came from Excel

# ── CENTER PRE-SETUP (ask Ramana for these values before running) ───────────
CENTER_ID = "REPLACE_WITH_ACTUAL_UUID"        # Fallback UUID when centers CSV is unavailable — BLOCKER #1
CENTER_NAME = "DigiSwasthya"                  # Exact name — confirm with NGO
CENTER_CODE = "DIGISWASTHYA_001"              # Must be UNIQUE across all centers in the system
CENTER_DISTRICT = None                        # Fill if known
CENTER_STATE = None                           # Fill if known
CENTER_TENANT_ID = None                        # BLOCKER #9 — get from Ramana

# ── SYSTEM USER PRE-SETUP (migration script user) ──────────────────────────
MIGRATION_SYSTEM_USER_ID = "REPLACE_WITH_ACTUAL_UUID"   # UUID of system user — BLOCKER #2
# This user is the "created_by_user_id" for all migrated visits
# Must be pre-inserted into the user table before running this script

# ── FALLBACK DOCTOR (if doctor lookup fails) ────────────────────────────────
FALLBACK_DOCTOR_USER_ID = "REPLACE_WITH_ACTUAL_UUID"    # BLOCKER #3 — ask Ramana

# ── DATE RANGE FILTER (CRITICAL for overlap handling) ───────────────────────
# The old-DB migration covers a date range (e.g. Aug 2024 to Nov 2025).
# Set this to the START date of the old DB's data range.
# Excel rows ON OR AFTER this date will be SKIPPED (they come from old DB migration instead).
# Set to None to disable filtering (process all rows).
EXCEL_ONLY_CUTOFF_DATE = None   # Example: "2024-08-01" — confirm exact date with Ramana and old-DB team

# ── TENANT ──────────────────────────────────────────────────────────────────
TENANT_ID = None                              # Set once confirmed

# ── FLAGS ────────────────────────────────────────────────────────────────────
IS_DEMO = False                               # Set True only for test runs on sample data
