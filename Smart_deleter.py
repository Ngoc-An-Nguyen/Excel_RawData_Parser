# File: Smart_deleter.py
# Purpose: Selects sheet, cleans cols, allows flexible deletion, sanitizes output sheet names.

import pandas as pd
import os
import numpy as np
from datetime import datetime
import copy
import re # Import regex module for sanitization

# --- Import utility functions ---
try:
    from excel_reader import load_specified_sheets_with_variations
    from data_utils import (find_data_categories, find_topic_block,
                           find_unique_topic_titles, find_entity_rows,
                           find_time_columns, parse_month_code)
except ImportError as e: print(f"Error: Could not import required functions: {e}"); exit()

# --- Configuration ---
INPUT_FILE_PATH = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\Data_SKU_Hao_Hao_2M2025.xlsx"
OUTPUT_DIR = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\log"
CANONICAL_NAMES = [
    'Brand', 'Manufacturer', 'Vietnam Off in.SR', 'TT Off VN in.SR', 'Off Urban in.SR',
    'Off Rural', 'MT VN', 'TT Off NE/NW in.SR', 'TT Off RRD in.SR', 'TT Off NCC in.SR',
    'TT Off SCC in.SR', 'TT Off CH in.SR', 'TT Off SE in.SR', 'TT Off MKD in.SR' ]
VARIATIONS = { # Sheet variations
    'Brand': ['brand', 'brands'], 'Manufacturer': ['manufacturer', 'manufactuer'],
    'Vietnam Off in.SR': ['vietnam off in.sr'], 'TT Off VN in.SR': ['tt off vn in.sr', 'tt off vietnam in.sr'],
    'Off Urban in.SR': ['off urban in.sr'], 'Off Rural': ['off rural'], 'MT VN': ['mt vn', 'mt vietnam'],
    'TT Off NE/NW in.SR': ['tt off ne/nw in.sr', 'ne nw'], 'TT Off RRD in.SR': ['tt off rrd in.sr', 'rrd'],
    'TT Off NCC in.SR': ['tt off ncc in.sr', 'ncc'], 'TT Off SCC in.SR': ['tt off scc in.sr', 'scc'],
    'TT Off CH in.SR': ['tt off ch in.sr', 'ch'], 'TT Off SE in.SR': ['tt off se in.sr', 'se'],
    'TT Off MKD in.SR': ['tt off mkd in.sr', 'mkd'] }
ENTITY_TYPE_MAP = {
    'Brand': 'Brand', 'Manufacturer': 'Manufacturer', 'Vietnam Off in.SR':'SKU', 'TT Off VN in.SR':'SKU',
    'Off Urban in.SR':'SKU', 'Off Rural':'SKU', 'MT VN':'SKU', 'TT Off NE/NW in.SR':'SKU',
    'TT Off RRD in.SR':'SKU', 'TT Off NCC in.SR':'SKU', 'TT Off SCC in.SR':'SKU',
    'TT Off CH in.SR':'SKU', 'TT Off SE in.SR':'SKU', 'TT Off MKD in.SR':'SKU' }
SHEETS_WITH_INTERNAL_CATS = {'Brand', 'Manufacturer'}
TOPIC_VARIATIONS_MAP = {
    'volume': "Volume ('unit/pack)", 'value': "Value ('000 000 VND)", 'revenue': "Value ('000 000 VND)",
    'share': "% Volume share", 'market share': "% Volume share", '% share': "% Volume share",
    'total cov': "% Do phu ve so luong tiem - Num Total Stock Dist", 'total coverage': "% Do phu ve so luong tiem - Num Total Stock Dist",
    'wtd': "Weighted Distribution Handling", 'wtd.': "Weighted Distribution Handling", 'coverage': "Weighted Distribution Handling",
    'sppd': "Volume SPPD (Volume/Wtd)", "volume ('unit/pack)": "Volume ('unit/pack)",
    "value ('000 000 vnd)": "Value ('000 000 VND)", "% volume share": "% Volume share",
    "% do phu ve so luong tiem - num total stock dist": "% Do phu ve so luong tiem - Num Total Stock Dist",
    "weighted distribution handling": "Weighted Distribution Handling", "volume sppd (volume/wtd)": "Volume SPPD (Volume/Wtd)"
}

# --- Helper Functions ---
def get_canonical_topic(topic_input, available_topics_list):
    # (Function unchanged)
    topic_input_lower = topic_input.lower()
    canonical_topic = None
    if topic_input_lower in TOPIC_VARIATIONS_MAP:
        mapped_topic = TOPIC_VARIATIONS_MAP[topic_input_lower]
        if mapped_topic in available_topics_list: canonical_topic = mapped_topic
        else: print(f"Warn: Variation '{topic_input}' maps to '{mapped_topic}' (not available).")
    if canonical_topic is None and topic_input in available_topics_list:
         canonical_topic = topic_input; print(f"(Recognized direct match: '{topic_input}')")
    return canonical_topic

def sanitize_sheet_name(name):
    """Removes invalid Excel sheet characters and truncates to 31 chars."""
    # Remove invalid characters: \ / ? * [ ]
    sanitized = re.sub(r'[\\/?*\[\]]', ' ', name)
    # Truncate to 31 characters (Excel limit)
    return sanitized[:31]

# --- Main Script Logic ---
if __name__ == "__main__":
    print("--- Starting Smart Deleter (Multi-Sheet Session) ---")

    # --- Step 1: Load ALL Relevant Sheets ---
    loaded_data_dict = load_specified_sheets_with_variations(INPUT_FILE_PATH, CANONICAL_NAMES, VARIATIONS)
    if not loaded_data_dict: print("Error: No relevant sheets loaded."); exit()

    # --- Step 2: Create Working Copy & Change Log ---
    modified_data_dict = copy.deepcopy(loaded_data_dict)
    change_log = []

    # --- Step 3: Main Loop (Outer - Sheet Selection) ---
    while True:
        print("\n" + "#"*60); print("## MAIN MENU - Select Sheet or Action ##"); print("#"*60)
        # (Display Change Log)
        if change_log:
            print("--- Recent Changes Made ---"); limit = 5
            for log_entry in change_log[-limit:]: print(f"- {log_entry}")
            if len(change_log) > limit: print(f"  (...and {len(change_log)-limit} more changes)")
            print("-" * 25)
        # (List available sheets)
        available_sheet_keys = sorted(list(modified_data_dict.keys()))
        print("Available sheets to modify:"); [print(f"  {i+1}. {key} ({modified_data_dict[key].shape[0]}r, {modified_data_dict[key].shape[1]}c)") for i, key in enumerate(available_sheet_keys)]
        print("\nActions:"); print("  save - Save all changes"); print("  exit - Exit without saving"); print("#"*60)
        outer_choice = input("Enter sheet number, 'save', or 'exit': ").strip().lower()

        if outer_choice == 'exit': print("Exiting without saving."); break
        if outer_choice == 'save': print("Proceeding to save..."); break

        # --- Process Sheet Selection ---
        selected_sheet_key = None
        try:
            choice_idx = int(outer_choice) - 1
            if 0 <= choice_idx < len(available_sheet_keys): selected_sheet_key = available_sheet_keys[choice_idx]
            else: print("Invalid sheet number."); continue
        except ValueError: print("Invalid input."); continue

        # --- Selected a sheet ---
        print(f"\n--- Modifying Sheet: '{selected_sheet_key}' ---")
        df_current = modified_data_dict[selected_sheet_key]
        ENTITY_TERM = ENTITY_TYPE_MAP.get(selected_sheet_key, 'Entity')

        # --- Conditional Column Cleaning ---
        print("Analyzing current sheet structure...")
        is_sku_sheet = selected_sheet_key not in SHEETS_WITH_INTERNAL_CATS
        df_to_analyze = df_current
        if is_sku_sheet:
             time_cols_map_initial = find_time_columns(df_to_analyze); orig_cols = list(df_to_analyze.columns)
             if time_cols_map_initial:
                  time_col_indices_orig = sorted(list(time_cols_map_initial.values()))
                  time_cols_to_keep_objs = [orig_cols[idx] for idx in time_col_indices_orig if idx in orig_cols]
                  cols_to_keep_objs = [orig_cols[0]] + time_cols_to_keep_objs if 0 in orig_cols else time_cols_to_keep_objs
                  if set(cols_to_keep_objs) != set(orig_cols) and len(cols_to_keep_objs) > 1 :
                       print("Performing column cleaning..."); df_to_analyze = df_to_analyze[cols_to_keep_objs].copy(); print(f"  New shape: {df_to_analyze.shape}")
        # else: print("Skipping cleaning for Brand/Manufacturer.")

        # --- Analyze Structure ---
        internal_categories = find_data_categories(df_to_analyze)
        available_topics = find_unique_topic_titles(df_to_analyze)
        entity_rows_map = find_entity_rows(df_to_analyze)
        available_entities = sorted(list(entity_rows_map.keys()))
        time_cols_map = find_time_columns(df_to_analyze)
        available_months = sorted(list(time_cols_map.keys()), key=lambda m: parse_month_code(m) or datetime.min)
        has_internal_categories = bool(internal_categories)
        print(f"  Analysis: {len(internal_categories)} Channels, {len(available_topics)} Topics, {len(available_entities)} {ENTITY_TERM}s, {len(available_months)} Months.")
        # ... Warnings ...

        df_modified_sheet = df_to_analyze.copy()

        # --- Inner Loop (Sheet Modification) ---
        while True:
            print("\n" + "="*50); print(f"  MODIFYING '{selected_sheet_key}'"); print("="*50)
            print(f"(Current state: {len(df_modified_sheet)} rows, {len(df_modified_sheet.columns)} columns)")
            print("-"*50); print("Actions for this sheet:")
            print("  1. Delete topic (e.g., Volume, Value) everywhere")
            if has_internal_categories: print("  2. Delete topic (e.g., Volume, Value) within specific channel"); print("  3. Delete entire channel")
            else: print("  -- (Options 2 & 3 N/A) --")
            print(f"  4. Keep Specific {ENTITY_TERM}s (Delete Others)"); print("  5. Keep Specific Time Range (Delete Others)")
            print("  back - Return to main sheet selection"); print("="*50)
            allowed_choices_num = ['1', '4', '5'];
            if has_internal_categories: allowed_choices_num.extend(['2', '3'])
            allowed_choices_inner = sorted(allowed_choices_num) + ['back']; prompt_choices = f"{', '.join(allowed_choices_inner[:-1])}, {allowed_choices_inner[-1]}"
            inner_choice = input(f"Enter action ({prompt_choices}): ").strip().lower()

            if inner_choice == 'back': print(f"--- Finished modifying '{selected_sheet_key}' for now. ---"); break
            if inner_choice not in allowed_choices_num: print("Invalid choice."); continue

            indices_to_drop = set(); columns_to_drop = set(); log_entry = None

            # --- Deletion Logic (Options 1-5) ---
            # (No changes needed within the logic itself, only added logging)
            if inner_choice == '1' or (inner_choice == '2' and has_internal_categories):
                 if not available_topics: print("Error: No topics."); continue
                 print("\nAvailable Topics:"); [print(f"  {i+1}. {t}") for i, t in enumerate(available_topics)]
                 topic_input = input("Enter Topic name/variation: ").strip()
                 canonical_topic_to_delete = get_canonical_topic(topic_input, available_topics)
                 if canonical_topic_to_delete is None: print(f"Error: Topic '{topic_input}' not recognized."); continue
                 print(f"Targeting topic: '{canonical_topic_to_delete}'")
                 if inner_choice == '1':
                     log_details = f"Topic '{canonical_topic_to_delete}' everywhere"
                     found_count = 0; cat_source = internal_categories if has_internal_categories else {None: (None, None)}
                     for cat_name, (cat_start, cat_end) in cat_source.items():
                         start_label, end_label = find_topic_block(df_modified_sheet, canonical_topic_to_delete, cat_start, cat_end)
                         if start_label is not None: found_count += 1;
                         try: start_loc = df_modified_sheet.index.get_loc(start_label); end_loc = len(df_modified_sheet) if isinstance(end_label, (int, np.integer)) and end_label > df_modified_sheet.index[-1] else df_modified_sheet.index.get_loc(end_label); indices_to_drop.update(list(df_modified_sheet.iloc[start_loc:end_loc].index))
                         except: pass
                     if found_count > 0: log_entry = f"Deleted {log_details}"
                     else: print("Note: Topic not found.")
                 elif inner_choice == '2':
                      print("\nAvailable Channels:"); cat_list = list(internal_categories.keys()); [print(f"  {i+1}. {n}") for i, n in enumerate(cat_list)]
                      try:
                          cat_choice_idx = int(input("Choose channel number: ").strip()) - 1; chosen_cat_name = cat_list[cat_choice_idx]; cat_start, cat_end = internal_categories[chosen_cat_name]
                          start_label, end_label = find_topic_block(df_modified_sheet, canonical_topic_to_delete, cat_start, cat_end)
                          if start_label is not None:
                              try: start_loc = df_modified_sheet.index.get_loc(start_label); end_loc = len(df_modified_sheet) if isinstance(end_label, (int, np.integer)) and end_label > df_modified_sheet.index[-1] else df_modified_sheet.index.get_loc(end_label); indices_to_drop.update(list(df_modified_sheet.iloc[start_loc:end_loc].index)); log_entry = f"Deleted Topic '{canonical_topic_to_delete}' from Channel '{chosen_cat_name}'"
                              except: pass
                          else: print("Note: Topic not found in channel.")
                      except (ValueError, IndexError) as e: print(f"Invalid choice: {e}")
            elif inner_choice == '3' and has_internal_categories:
                 print("\nAvailable Channels:"); cat_list = list(internal_categories.keys()); [print(f"  {i+1}. {n}") for i, n in enumerate(cat_list)]
                 try:
                     cat_choice_idx = int(input("Choose channel number: ").strip()) - 1; chosen_cat_name = cat_list[cat_choice_idx]; cat_start, cat_end = internal_categories[chosen_cat_name]
                     if input(f"Delete channel '{chosen_cat_name}'? (yes/no): ").lower() == 'yes':
                          try: start_loc = df_modified_sheet.index.get_loc(cat_start); end_loc = len(df_modified_sheet) if isinstance(cat_end, (int, np.integer)) and cat_end > df_modified_sheet.index[-1] else df_modified_sheet.index.get_loc(cat_end); indices_to_drop.update(list(df_modified_sheet.iloc[start_loc:end_loc].index)); log_entry = f"Deleted Channel '{chosen_cat_name}'"
                          except: pass
                     else: print("Cancelled.")
                 except (ValueError, IndexError) as e: print(f"Invalid choice: {e}")
            elif inner_choice == '4':
                 if not available_entities: print(f"Error: No {ENTITY_TERM}s."); continue
                 entity_lookup_lower = {b.lower(): b for b in available_entities}; print(f"\nAvailable {ENTITY_TERM}s:"); [print(f"  - {e}") for e in available_entities]
                 keep_input = input(f"Enter comma-separated {ENTITY_TERM}s to KEEP: ").strip(); user_keep_lower = {e.strip().lower() for e in keep_input.split(',') if e.strip()}
                 if not user_keep_lower: print("None specified."); continue
                 valid_keep_lower = user_keep_lower.intersection(entity_lookup_lower.keys()); entities_to_keep_orig = {entity_lookup_lower[e_low] for e_low in valid_keep_lower}
                 if not entities_to_keep_orig: print("No matches found."); continue
                 entities_to_delete_orig = set(available_entities) - entities_to_keep_orig; print(f"Keeping: {', '.join(sorted(list(entities_to_keep_orig)))}")
                 if not entities_to_delete_orig: print("All kept."); continue
                 print(f"Deleting: {', '.join(sorted(list(entities_to_delete_orig)))}"); rows_for_deleted = set()
                 for name_orig in entities_to_delete_orig:
                     if name_orig in entity_rows_map: rows_for_deleted.update(entity_rows_map[name_orig])
                 indices_to_drop.update([idx for idx in rows_for_deleted if idx in df_modified_sheet.index])
                 if indices_to_drop: log_entry = f"Kept only specific {ENTITY_TERM}s"
            elif inner_choice == '5':
                 if not time_cols_map: print("Error: No time columns."); continue
                 print("\nAvailable Months:"); print(f"  From {available_months[0]} to {available_months[-1]}")
                 start_m = input(f"Start Month ({available_months[0]}): ").strip().upper(); end_m = input(f"End Month ({available_months[-1]}): ").strip().upper()
                 start_dt = parse_month_code(start_m); end_dt = parse_month_code(end_m)
                 if not start_dt or not end_dt or start_dt > end_dt: print("Invalid range."); continue
                 cols_to_keep_idx = set(); months_kept = set()
                 for month, col_idx in time_cols_map.items():
                     m_dt = parse_month_code(month)
                     if m_dt and start_dt <= m_dt <= end_dt: cols_to_keep_idx.add(col_idx); months_kept.add(month)
                 if not cols_to_keep_idx: print("No columns in range."); continue
                 print(f"Keeping range: {start_m} to {end_m} ({len(months_kept)} months)")
                 cols_to_drop_idx = set(time_cols_map.values()) - cols_to_keep_idx
                 print(f"Deleting {len(cols_to_drop_idx)} time columns."); columns_to_drop.update(cols_to_drop_idx)
                 if columns_to_drop: log_entry = f"Kept only time range {start_m}-{end_m}"

            # --- Perform Deletions & Update State ---
            try:
                made_changes = False
                if indices_to_drop:
                    valid_indices = sorted([idx for idx in indices_to_drop if idx in df_modified_sheet.index], reverse=True)
                    if valid_indices: print(f"\nDropping {len(valid_indices)} rows..."); df_modified_sheet = df_modified_sheet.drop(index=valid_indices); made_changes = True; print("Row deletion successful.")
                if columns_to_drop:
                    current_cols = df_modified_sheet.columns
                    # Find columns by their *positional index* in the current dataframe
                    valid_cols_to_drop_objs = [current_cols[idx] for idx in columns_to_drop if idx < len(current_cols)]
                    if valid_cols_to_drop_objs: print(f"\nDropping {len(valid_cols_to_drop_objs)} columns..."); df_modified_sheet = df_modified_sheet.drop(columns=valid_cols_to_drop_objs); made_changes = True; print("Column deletion successful.")

                if made_changes:
                    df_modified_sheet = df_modified_sheet.reset_index(drop=True)
                    modified_data_dict[selected_sheet_key] = df_modified_sheet # Update main dict
                    if log_entry: change_log.append(f"{log_entry} in Sheet '{selected_sheet_key}'")
                    print("\nRe-analyzing structure...")
                    internal_categories = find_data_categories(df_modified_sheet)
                    available_topics = find_unique_topic_titles(df_modified_sheet)
                    entity_rows_map = find_entity_rows(df_modified_sheet)
                    available_entities = sorted(list(entity_rows_map.keys()))
                    time_cols_map = find_time_columns(df_modified_sheet)
                    available_months = sorted(list(time_cols_map.keys()), key=lambda m: parse_month_code(m) or datetime.min)
                    has_internal_categories = bool(internal_categories)
                    print(f"  Updated: {len(internal_categories)} Channels, {len(available_topics)} Topics, {len(available_entities)} {ENTITY_TERM}s, {len(available_months)} Months.")
                elif log_entry is None and inner_choice in allowed_choices_num: print("No changes made for this operation.")
            except Exception as e: print(f"Error during update: {e}")
        # End of inner loop

    # --- Step 6: Save ALL Modified Sheets ---
    if outer_choice == 'save':
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"Smart_Deleter_Output_Session_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)
        try:
            print(f"\nSaving ALL sheets ({len(modified_data_dict)} total) to: {output_path}")
            os.makedirs(OUTPUT_DIR, exist_ok=True)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name_key in modified_data_dict.keys():
                    # --- Sanitize sheet name before saving ---
                    sanitized_name = sanitize_sheet_name(sheet_name_key)
                    print(f"  Saving sheet: '{sheet_name_key}' as '{sanitized_name}'...")
                    modified_data_dict[sheet_name_key].to_excel(writer, sheet_name=sanitized_name, index=False, header=False)
            print("Save complete.")
            # Optional: Save change log
            if change_log:
                 log_path = os.path.join(OUTPUT_DIR, f"Smart_Deleter_Log_{timestamp}.txt")
                 try:
                      with open(log_path, 'w') as f: f.write("\n".join(change_log))
                      print(f"Change log saved to: {log_path}")
                 except Exception as log_e: print(f"Error saving change log: {log_e}")

        except Exception as e: print(f"\nError saving file: {e}")

    print("\n--- Smart_deleter script finished. ---")