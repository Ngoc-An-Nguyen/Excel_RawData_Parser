# File: excel_reader.py
# Purpose: Contains functions to load raw Excel data, handling name variations and sanitization.

import pandas as pd
import os
import re  # Import regex for sanitization


# --- Sanitization Helper ---
# Copied from Smart_deleter to make loader aware
def sanitize_sheet_name(name):
    """Removes invalid Excel sheet characters and truncates to 31 chars."""
    # Remove invalid characters: \ / ? * [ ]
    sanitized = re.sub(r'[\\/?*\[\]]', ' ', name)
    # Truncate to 31 characters (Excel limit)
    return sanitized[:31]


# --- Single Sheet Loader (Unchanged) ---
def load_raw_excel(file_path, sheet_index=0):
    # ... (function code unchanged) ...
    print(f"Attempting to load single sheet (index {sheet_index}) from: {file_path}")
    if not os.path.exists(file_path): print(f"Error: Input file not found at: {file_path}"); return None
    try:
        excel_file = pd.ExcelFile(file_path)
        if not excel_file.sheet_names: print(f"Error: No sheets found."); return None
        if sheet_index >= len(excel_file.sheet_names): print(f"Error: Sheet index out of range."); return None
        sheet_name_to_load = excel_file.sheet_names[sheet_index]
        print(f"Loading sheet '{sheet_name_to_load}' using header=None")
        df = excel_file.parse(sheet_name_to_load, header=None)
        print(f"Successfully loaded single sheet '{sheet_name_to_load}'.")
        return df
    except Exception as e:
        print(f"Error loading sheet index {sheet_index}: {e}"); return None


# --- Multi-Sheet Loader with Variations & Sanitization Check ---
def load_specified_sheets_with_variations(file_path, canonical_sheet_names, name_variations_map):
    """
    Loads multiple sheets based on canonical names, using a map for variations/typos
    AND checking for sanitized versions of canonical names.
    """
    print(f"--- Loading sheets (Variations + Sanitization Check) from: {file_path} ---")
    loaded_sheets_data = {}
    matched_actual_sheets = set()

    if not os.path.exists(file_path): print(f"Error: Input file not found: {file_path}"); return loaded_sheets_data

    try:
        excel_file = pd.ExcelFile(file_path)
        available_sheets_original_case = excel_file.sheet_names
        available_lower_to_orig = {s.lower(): s for s in available_sheets_original_case}
        print(f"Available sheets in file: {available_sheets_original_case}")

        loaded_count = 0;
        missing_canonical_names = []

        for canonical_name in canonical_sheet_names:
            found_match = False
            variations_to_check = name_variations_map.get(canonical_name, [])
            if canonical_name.lower() not in variations_to_check:
                variations_to_check.append(canonical_name.lower())

            # 1. Check variations map first
            for variation in variations_to_check:
                if variation in available_lower_to_orig:
                    original_actual_name = available_lower_to_orig[variation]
                    if original_actual_name not in matched_actual_sheets:
                        try:
                            print(
                                f"  Found '{canonical_name}' via variation '{variation}' as sheet '{original_actual_name}'. Loading...")
                            df = excel_file.parse(original_actual_name, header=None)
                            loaded_sheets_data[canonical_name] = df
                            matched_actual_sheets.add(original_actual_name)
                            loaded_count += 1;
                            found_match = True;
                            break
                        except Exception as e_sheet:
                            print(f"    *Error* loading sheet '{original_actual_name}': {e_sheet}"); break

            # 2. If no match via variations, check SANITIZED canonical name
            if not found_match:
                sanitized_lower = sanitize_sheet_name(canonical_name).lower()
                # Also add sanitized lower to variations_to_check if different, to cover all bases
                if sanitized_lower not in variations_to_check:
                    variations_to_check.append(sanitized_lower)  # Check this too now

                if sanitized_lower in available_lower_to_orig:
                    original_actual_name = available_lower_to_orig[sanitized_lower]
                    if original_actual_name not in matched_actual_sheets:
                        try:
                            print(
                                f"  Found '{canonical_name}' via SANITIZED name as sheet '{original_actual_name}'. Loading...")
                            df = excel_file.parse(original_actual_name, header=None)
                            loaded_sheets_data[canonical_name] = df
                            matched_actual_sheets.add(original_actual_name)
                            loaded_count += 1;
                            found_match = True
                        except Exception as e_sheet:
                            print(f"    *Error* loading sheet '{original_actual_name}': {e_sheet}");

            # 3. If still no match
            if not found_match:
                print(
                    f"  *Warning*: Could not find a match for requested sheet '{canonical_name}' using variations {variations_to_check}.")
                missing_canonical_names.append(canonical_name)

        print(f"--- Finished loading. Successfully loaded {loaded_count} sheets. ---")
        if missing_canonical_names: print(f"--- Could not find matches for: {missing_canonical_names} ---")

    except Exception as e_file:
        print(f"Error processing Excel file '{file_path}': {e_file}"); return loaded_sheets_data

    return loaded_sheets_data


# --- Example usage (Unchanged) ---
if __name__ == "__main__":
    # ... (rest of the testing block remains the same) ...
    test_multi_file_path = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\Data_SKU_Hao_Hao_2M2025.xlsx"
    canonical_names = ['Brand', 'Manufacturer', 'Vietnam Off in.SR', 'TT Off VN in.SR', 'Off Urban in.SR', 'Off Rural',
                       'MT VN', 'TT Off NE/NW in.SR', 'TT Off RRD in.SR', 'TT Off NCC in.SR', 'TT Off SCC in.SR',
                       'TT Off CH in.SR', 'TT Off SE in.SR', 'TT Off MKD in.SR']
    variations = {'Brand': ['brand', 'brands'], 'Manufacturer': ['manufacturer', 'manufactuer'],
                  'Vietnam Off in.SR': ['vietnam off in.sr'],
                  'TT Off VN in.SR': ['tt off vn in.sr', 'tt off vietnam in.sr'],
                  'Off Urban in.SR': ['off urban in.sr'], 'Off Rural': ['off rural'], 'MT VN': ['mt vn', 'mt vietnam'],
                  'TT Off NE/NW in.SR': ['tt off ne/nw in.sr', 'ne nw'],
                  'TT Off RRD in.SR': ['tt off rrd in.sr', 'rrd'], 'TT Off NCC in.SR': ['tt off ncc in.sr', 'ncc'],
                  'TT Off SCC in.SR': ['tt off scc in.sr', 'scc'], 'TT Off CH in.SR': ['tt off ch in.sr', 'ch'],
                  'TT Off SE in.SR': ['tt off se in.sr', 'se'], 'TT Off MKD in.SR': ['tt off mkd in.sr', 'mkd']}
    print("\n--- Testing load_specified_sheets_with_variations ---")
    loaded_data = load_specified_sheets_with_variations(test_multi_file_path, canonical_names, variations)
    if loaded_data:
        print("\n--- Load Test Summary (Using Canonical Names) ---")
        print(f"Successfully loaded {len(loaded_data)} sheets referenced by canonical name:")
        for canonical_name_key, df_sample in loaded_data.items(): print(
            f"  - Canonical Name: '{canonical_name_key}', Shape: {df_sample.shape}")
    else:
        print("\nLoad test failed or loaded no sheets.")
    print("--- End of excel_reader.py test ---")