"""
Excel Phone Number Search Utility

This script searches all columns across all sheets in an Excel file (default: example.xlsx)
for rows matching any of the specified phone numbers. 

Features:
- Normalizes both the target search list and the cells to their last 10 digits to
  ensure matching is robust against different formats (e.g. +91 74596-54851, 07459654851, 7459654851.0).
- Searches all columns (does not restrict to specific columns).
- Outputs the matched value, sheet name, sheet number (1-indexed), column header, and the full row details.
"""

import os
import sys
from typing import Any, List, Dict
import pandas as pd

# =====================================================================
# CONFIGURATION
# =====================================================================
# Place the path to your source Excel file here
EXCEL_FILE_PATH = 


SEARCH_PHONES = [
   "9769465241",
   "918318441688",
   "917903059001",
   "919604778091",
   "919999999999",
   "919999999999",
   "919999999999",
   "919999999999",
   "919999999999",
   "919999999999",
   "916209891922",
   "919209775807",
   "918767593817",
   "919209775807",
   "917739412689",
   "919699338465",
   "917499619783",
   "918208026908",
   "918208026908",
   "918196841679",
   "918767593817",
   "917767075507",
   "918196841679",
   "919665682967",
   "919665682967",
   "919199121223",
   "919525337174",
   "919511854887",
   "919766447946",
   "919518750562",
   "919370183788",
   "917499618495",
   "918788073815",
   "918788073815",
   "919272100260",
   "919272100260",
   "919272100260",
   "917620038542",
   "917972953731",
   "919922550826",
   "918806550242"
]
# =====================================================================


def clean_to_last10_digits(value: Any) -> str | None:
    """
    Extracts only the digits from a cell value and returns the last 10 digits
    if it contains enough digit characters. Handles floats, decimals, and string formats.
    """
    if pd.isna(value) or value is None:
        return None
        
    val_str = str(value).strip()
    
    # Handle pandas reading floats like '7459654851.0'
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
        
    # Remove any non-numeric characters
    digits = "".join(char for char in val_str if char.isdigit())
    
    if len(digits) >= 10:
        return digits[-10:]
        
    return None


def search_excel(file_path: str, targets: List[str]):
    print("=" * 70)
    print("           EXCEL CELL-LEVEL PHONE SEARCH UTILITY")
    print("=" * 70)
    
    # Resolve absolute path for feedback
    abs_path = os.path.abspath(file_path)
    if not os.path.exists(file_path):
        print(f"[ERROR] Excel file not found at: {abs_path}")
        print("Please check the EXCEL_FILE_PATH variable inside the script.")
        return

    # Normalize targets to their last 10 digits for matching
    normalized_targets: Dict[str, str] = {}
    for t in targets:
        cleaned = clean_to_last10_digits(t)
        if cleaned:
            normalized_targets[cleaned] = t
        else:
            # If input is already 10 digits but has letters/formatting
            digits_only = "".join(c for c in str(t) if c.isdigit())
            if len(digits_only) >= 10:
                normalized_targets[digits_only[-10:]] = t
            else:
                print(f"[Warning] Target phone number '{t}' is too short (< 10 digits). Skipping.")

    if not normalized_targets:
        print("[ERROR] No valid 10-digit target phone numbers provided for search.")
        return

    print(f"Loading Excel file: {file_path}")
    print(f"Target phone search list (normalized): {list(normalized_targets.values())}\n")

    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
    except Exception as e:
        print(f"[ERROR] Failed to read Excel sheets: {e}")
        return

    total_matches = 0
    match_records = []

    for sheet_idx, sheet_name in enumerate(sheet_names, start=1):
        print(f"Scanning Sheet #{sheet_idx}: '{sheet_name}'...")
        try:
            # dtype=str prevents pandas from converting phone numbers to scientific notation (e.g. 7.46e+09)
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
        except Exception as e:
            print(f"  [Warning] Failed to read sheet '{sheet_name}': {e}")
            continue

        # Drop entirely empty rows
        df_cleaned = df.dropna(how="all")

        for row_idx, (_, row) in enumerate(df_cleaned.iterrows(), start=1):
            row_dict = row.to_dict()
            
            # Look through all columns in the row
            for col_name, cell_value in row_dict.items():
                cleaned_cell = clean_to_last10_digits(cell_value)
                
                # Check if cell value matches any normalized target
                if cleaned_cell and cleaned_cell in normalized_targets:
                    original_target = normalized_targets[cleaned_cell]
                    total_matches += 1
                    
                    record = {
                        "match_no": total_matches,
                        "target_searched": original_target,
                        "matched_in_cell": str(cell_value),
                        "sheet_number": sheet_idx,
                        "sheet_name": sheet_name,
                        "row_number_in_sheet": row_idx,
                        "column_name": col_name,
                        "full_row": row_dict
                    }
                    match_records.append(record)
                    
                    print(f"  --> MATCH FOUND!")
                    print(f"      Matched Search Phone : {original_target}")
                    print(f"      Cell Raw Value       : {cell_value}")
                    print(f"      Sheet                : #{sheet_idx} ('{sheet_name}')")
                    print(f"      Row (in sheet)       : {row_idx}")
                    print(f"      Column Name          : '{col_name}'")
                    print(f"      Row Data Preview     : {list(row_dict.items())[:4]}...")
                    print("-" * 50)

    print("\n" + "=" * 70)
    print("                     SEARCH SUMMARY")
    print("=" * 70)
    print(f"Total Sheets Scanned : {len(sheet_names)}")
    print(f"Total Matches Found  : {total_matches}")
    print("=" * 70)

    if total_matches > 0:
        # Prompt user if they want to export results to a JSON or text format
        print("\nAll matches are printed above.")
        
        # Save a summary report file
        report_path = "phone_search_results.txt"
        try:
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(f"Excel Phone Search Results\n")
                f.write(f"Source Excel File: {abs_path}\n")
                f.write(f"Search list: {list(normalized_targets.values())}\n")
                f.write(f"Total Matches: {total_matches}\n")
                f.write("=" * 80 + "\n\n")
                for r in match_records:
                    f.write(f"Match #{r['match_no']}\n")
                    f.write(f"  Target Searched     : {r['target_searched']}\n")
                    f.write(f"  Matched Value in DB : {r['matched_in_cell']}\n")
                    f.write(f"  Sheet Number        : #{r['sheet_number']} ('{r['sheet_name']}')\n")
                    f.write(f"  Row Number          : {r['row_number_in_sheet']}\n")
                    f.write(f"  Column Header       : '{r['column_name']}'\n")
                    f.write(f"  Full Row Data       :\n")
                    for k, v in r['full_row'].items():
                        f.write(f"    - {k}: {v}\n")
                    f.write("-" * 80 + "\n")
            print(f"[Report] Full report with complete row data saved to: {os.path.abspath(report_path)}")
        except Exception as e:
            print(f"[Warning] Failed to write report file: {e}")


if __name__ == "__main__":
    search_excel(EXCEL_FILE_PATH, SEARCH_PHONES)
