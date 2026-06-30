import os
import sys
import re
import difflib
import openpyxl
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Any, Set
from sqlalchemy import create_engine, text

# Import project config & cleaning utilities
migration_dir = Path(__file__).resolve().parent
sys.path.append(str(migration_dir))

try:
    from direct_migrate_supabase_update_mode import _read_config, clean_doctor_name, _find_column, _norm_text
except ImportError:
    from .direct_migrate_supabase_update_mode import _read_config, clean_doctor_name, _find_column, _norm_text


def fetch_database_doctors(engine, schema="public") -> List[Dict[str, Any]]:
    with engine.connect() as conn:
        sql = text(f"SELECT id, role, first_name, last_name, employee_id, legacy_id FROM {schema}.user")
        all_users = [dict(r) for r in conn.execute(sql).mappings().all()]
    # Filter for doctors or users with medical roles
    return all_users


def extract_unique_doctors_from_excel(excel_path: Path) -> Tuple[Set[str], List[Tuple[str, str, int]]]:
    unique_names = set()
    occurrences = []
    
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        
        # 1. Extract from Data Validation Dropdowns if present
        if ws.data_validations and ws.data_validations.dataValidation:
            for dv in ws.data_validations.dataValidation:
                if dv.formula1 and isinstance(dv.formula1, str):
                    f = dv.formula1.replace('"', '').strip()
                    if ',' in f and not f.startswith('='):
                        for name in f.split(','):
                            n = name.strip()
                            if n and len(n) > 2:
                                unique_names.add(n)

    # 2. Extract from actual data rows using pandas
    xl = pd.ExcelFile(excel_path)
    for sheet in xl.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet, header=0)
        df_clean = df.dropna(how="all").copy()
        doc_col = _find_column(df_clean, "DOCTOR'S NAME", "DOCTOR NAME", "Doctor Name", "Doctor's Name")
        if doc_col:
            for idx, row in df_clean.iterrows():
                val = row.get(doc_col)
                if val and not pd.isna(val):
                    s = str(val).strip()
                    if s:
                        unique_names.add(s)
                        occurrences.append((s, sheet, idx + 1))
                        
    return unique_names, occurrences


def calculate_similarity_confidence(excel_clean: str, db_doctor: Dict[str, Any]) -> float:
    db_full_name = f"{db_doctor.get('first_name') or ''} {db_doctor.get('last_name') or ''}".strip()
    db_clean = clean_doctor_name(db_full_name)
    
    if not excel_clean or not db_clean:
        return 0.0
        
    if excel_clean == db_clean:
        return 100.0
        
    # SequenceMatcher similarity ratio
    ratio = difflib.SequenceMatcher(None, excel_clean, db_clean).ratio() * 100.0
    
    # Partial token matching boost
    excel_tokens = set(re.findall(r'\w+', excel_clean))
    db_tokens = set(re.findall(r'\w+', db_clean))
    if excel_tokens and db_tokens:
        overlap = len(excel_tokens.intersection(db_tokens)) / max(len(excel_tokens), len(db_tokens))
        token_score = overlap * 100.0
        ratio = max(ratio, token_score)
        
    return round(ratio, 1)


def audit_excel_doctors(excel_file_path: str):
    p = Path(excel_file_path)
    if not p.exists():
        print(f"[ERROR] Excel file not found at: {excel_file_path}")
        return

    print("================================================================================")
    print(f"      PRE-MIGRATION DOCTOR MATCHING AUDIT FOR: {p.name}")
    print("================================================================================\n")

    os.environ["INPUT_EXCEL_PATH"] = str(p.resolve())
    os.environ["MIGRATION_WORKBOOK_PATH"] = str(p.resolve())
    cfg = _read_config()
    engine = create_engine(cfg.database_url)
    db_users = fetch_database_doctors(engine, cfg.schema)

    # Build DB Doctor Lookups
    user_index_by_employee = {}
    user_index_by_legacy = {}
    user_index_by_clean_name = {}

    for u in db_users:
        emp_id = _norm_text(u.get("employee_id"))
        leg_id = _norm_text(u.get("legacy_id"))
        full_name = f"{u.get('first_name') or ''} {u.get('last_name') or ''}"
        clean_db_name = clean_doctor_name(full_name)
        if emp_id: user_index_by_employee[emp_id] = u
        if leg_id: user_index_by_legacy[leg_id] = u
        if clean_db_name: user_index_by_clean_name[clean_db_name] = u

    unique_doc_names, occurrences = extract_unique_doctors_from_excel(p)
    print(f"Discovered {len(unique_doc_names)} unique doctor names in workbook across data cells & dropdowns.\n")

    audit_results = []

    for doc_name in sorted(list(unique_doc_names)):
        clean_name = clean_doctor_name(doc_name)
        matched_user = user_index_by_clean_name.get(clean_name)
        count = sum(1 for d, s, r in occurrences if d == doc_name)
        sample_locs = [f"{s} (r{r})" for d, s, r in occurrences if d == doc_name][:2]
        sample_str = ", ".join(sample_locs) if sample_locs else "Dropdown Only"

        if matched_user:
            db_name = f"{matched_user.get('first_name')} {matched_user.get('last_name')}"
            audit_results.append({
                "excel_name": doc_name,
                "status": "EXACT_MATCH",
                "matched_db_doctor": db_name,
                "db_id": matched_user['id'],
                "confidence": 100.0,
                "occurrences": count,
                "samples": sample_str
            })
        else:
            # Rank candidates by fuzzy confidence
            candidates = []
            for u in db_users:
                score = calculate_similarity_confidence(clean_name, u)
                if score > 40.0:
                    db_name = f"{u.get('first_name')} {u.get('last_name')}"
                    candidates.append((score, db_name, u['id']))
                    
            candidates.sort(key=lambda x: x[0], reverse=True)
            
            if candidates and candidates[0][0] >= 65.0:
                best_score, best_name, best_id = candidates[0]
                audit_results.append({
                    "excel_name": doc_name,
                    "status": "HIGH_CONFIDENCE_SUGGESTION",
                    "matched_db_doctor": f"{best_name} (Suggested)",
                    "db_id": best_id,
                    "confidence": best_score,
                    "occurrences": count,
                    "samples": sample_str
                })
            else:
                top_suggestion = f"{candidates[0][1]} ({candidates[0][0]}%)" if candidates else "None"
                audit_results.append({
                    "excel_name": doc_name,
                    "status": "UNMAPPED_NEEDS_CREATION",
                    "matched_db_doctor": top_suggestion,
                    "db_id": "N/A",
                    "confidence": candidates[0][0] if candidates else 0.0,
                    "occurrences": count,
                    "samples": sample_str
                })

    # Print Formatted CLI Report Table and Save to File
    report_lines = []
    report_lines.append("================================================================================")
    report_lines.append(f"      PRE-MIGRATION DOCTOR MATCHING AUDIT FOR: {p.name}")
    report_lines.append("================================================================================\n")
    report_lines.append(f"{'EXCEL DOCTOR NAME':<32} | {'STATUS':<26} | {'MATCHED / SUGGESTED DB DOCTOR':<32} | {'CONF.':<6} | {'ROWS'}")
    report_lines.append("-" * 110)
    for r in audit_results:
        report_lines.append(f"{r['excel_name']:<32} | {r['status']:<26} | {r['matched_db_doctor']:<32} | {r['confidence']:<5.1f}% | {r['occurrences']}")

    exact_c = sum(1 for r in audit_results if r['status'] == "EXACT_MATCH")
    sugg_c = sum(1 for r in audit_results if r['status'] == "HIGH_CONFIDENCE_SUGGESTION")
    unmap_c = sum(1 for r in audit_results if r['status'] == "UNMAPPED_NEEDS_CREATION")

    report_lines.append("\n================================================================================")
    report_lines.append("                              AUDIT SUMMARY                                    ")
    report_lines.append("================================================================================")
    report_lines.append(f"Total Unique Doctor Names Analyzed: {len(audit_results)}")
    report_lines.append(f"  - Exact Matches to DB Users     : {exact_c}")
    report_lines.append(f"  - High-Confidence Suggestions    : {sugg_c}")
    report_lines.append(f"  - Unmapped (Needs DB Creation)  : {unmap_c}")
    report_lines.append("================================================================================")

    # Print to console
    for line in report_lines:
        print(line)

    # Save to text file
    report_txt_path = p.parent / "doctor_audit_report.txt"
    try:
        report_txt_path.write_text("\n".join(report_lines), encoding="utf-8")
        print(f"\n[SUCCESS] Audit report saved successfully to: {report_txt_path.name}")
    except Exception as e:
        print(f"\n[WARNING] Could not save audit report to text file: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        file_p = sys.argv[1]
    else:
        try:
            cfg = _read_config()
            file_p = str(cfg.excel_path)
        except Exception:
            file_p = "DS6_and_Gujarat.xlsx"
    audit_excel_doctors(file_p)
