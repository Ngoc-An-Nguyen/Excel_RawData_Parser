# File: data_organizer.py
# Purpose: Loads multiple sheets, parses to a unified tidy format,
#          allows filtering, and formats/saves the result (typically wide).

import pandas as pd
import os
import numpy as np
from datetime import datetime
import copy
import re

# --- Import utility functions ---
try:
    from excel_reader import load_specified_sheets_with_variations
    from data_utils import (parse_sheet_to_tidy, # The main parser
                           find_unique_topic_titles, find_entity_rows, # For filter options
                           find_time_columns, parse_month_code)
except ImportError as e: print(f"Error: Could not import required functions: {e}"); exit()

# --- Configuration ---
# Define paths for both potential inputs
RAW_INPUT_FILE_PATH = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\Data_SKU_Hao_Hao_2M2025.xlsx"
# Assume deleter output has a predictable pattern or user knows the path.
# For now, let's hardcode a placeholder - user would change this.
# Better: Find latest 'Smart_Deleter_Output_Session_*.xlsx' in OUTPUT_DIR.
DELETER_OUTPUT_SEARCH_DIR = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\log"
# Placeholder - will try to find latest deleter output later if selected
DELETER_INPUT_FILE_PATH_PLACEHOLDER = os.path.join(DELETER_OUTPUT_SEARCH_DIR, "LATEST_DELETER_OUTPUT.xlsx")

OUTPUT_DIR = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\log"
OUTPUT_FILENAME_BASE = "Organized_Data" # Will add details later

# Canonical names, variations map, entity types, category map (needed for loading and parsing)
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
CATEGORY_ABBR_MAP = { # Abbr sheet name -> Full Category Name
    'NE NW': 'TT Off NE/NW in.SR', 'RRD': 'TT Off RRD in.SR', 'NCC': 'TT Off NCC in.SR',
    'SCC': 'TT Off SCC in.SR', 'CH': 'TT Off CH in.SR', 'SE': 'TT Off SE in.SR',
    'MKD': 'TT Off MKD in.SR' }

# --- Helper to find latest deleter output ---
def find_latest_deleter_output(search_dir):
    latest_file = None
    latest_time = 0
    try:
        for filename in os.listdir(search_dir):
            if filename.startswith("Smart_Deleter_Output_Session_") and filename.endswith(".xlsx"):
                file_path = os.path.join(search_dir, filename)
                try:
                    file_time = os.path.getmtime(file_path)
                    if file_time > latest_time:
                        latest_time = file_time
                        latest_file = file_path
                except OSError:
                    continue # Skip files we can't get time for
    except FileNotFoundError:
        print(f"Warning: Deleter output search directory not found: {search_dir}")
    except Exception as e:
        print(f"Error searching for deleter output: {e}")
    return latest_file

# --- Helper to sanitize sheet names for saving ---
def sanitize_sheet_name(name):
    sanitized = re.sub(r'[\\/?*\[\]]', ' ', name)
    return sanitized[:31]

# --- Main Script Logic ---
if __name__ == "__main__":
    print("--- Starting Data Organizer (Multi-Sheet) ---")

    # --- Step 1: Choose Input Source ---
    print("\nChoose input data source:")
    print("  1. Organize from Raw Excel File")
    print("  2. Organize from Latest Smart Deleter Output")
    input_choice = input("Enter choice (1 or 2): ").strip()

    input_path = None
    input_desc = ""
    if input_choice == '1':
        input_path = RAW_INPUT_FILE_PATH
        input_desc = "Raw File"
        print(f"Using Raw Input File: {input_path}")
    elif input_choice == '2':
        print(f"Searching for latest deleter output in: {DELETER_OUTPUT_SEARCH_DIR}")
        input_path = find_latest_deleter_output(DELETER_OUTPUT_SEARCH_DIR)
        if input_path:
            input_desc = "Deleter Output"
            print(f"Found Latest Deleter Output: {input_path}")
        else:
            print(f"Error: No suitable Smart Deleter output file found in {DELETER_OUTPUT_SEARCH_DIR}.")
            print(f"Please ensure a file like 'Smart_Deleter_Output_Session_*.xlsx' exists or choose option 1.")
            exit()
    else:
        print("Invalid choice. Exiting.")
        exit()

    if not input_path or not os.path.exists(input_path):
         print(f"Error: Selected input file path does not exist: {input_path}")
         exit()


    # --- Step 2: Load Data ---
    # Load using variations map, providing canonical names
    loaded_data_dict = load_specified_sheets_with_variations(input_path, CANONICAL_NAMES, VARIATIONS)
    if not loaded_data_dict: print("Error: No sheets were loaded from the input file."); exit()

    # --- Step 3: Parse All Sheets to Tidy Format ---
    all_parsed_data = []
    print("\n--- Parsing Loaded Sheets ---")
    for canonical_sheet_name, raw_df in loaded_data_dict.items():
        entity_t = ENTITY_TYPE_MAP.get(canonical_sheet_name, 'Unknown')
        parsed_df = parse_sheet_to_tidy(canonical_sheet_name, raw_df, entity_t, CATEGORY_ABBR_MAP)
        if parsed_df is not None:
            all_parsed_data.append(parsed_df)
            print(f"  Successfully parsed sheet: '{canonical_sheet_name}'")
        else:
            print(f"  *Warning*: Failed to parse sheet: '{canonical_sheet_name}'")

    # --- Step 4: Combine into Master Tidy DF ---
    if not all_parsed_data: print("\nError: No data was successfully parsed from any sheet. Exiting."); exit()

    print("\n--- Combining Parsed Data ---")
    master_tidy_df = pd.concat(all_parsed_data, ignore_index=True)
    master_tidy_df['Month_dt'] = master_tidy_df['Month'].apply(parse_month_code)
    print(f"Master Tidy DataFrame created with {master_tidy_df.shape[0]} rows and {master_tidy_df.shape[1]} columns.")
    print("Columns:", list(master_tidy_df.columns))

    # --- Step 5: User Filtering Loop ---
    filtered_df = master_tidy_df.copy() # Start with all data
    active_filters = {} # Dictionary to store active filters

    while True:
        print("\n" + "="*50); print("  FILTERING MENU"); print("="*50)
        print(f"(Current DataFrame: {len(filtered_df)} rows)")
        print("--- Active Filters ---")
        if active_filters:
            for key, value in active_filters.items(): print(f"  - {key}: {value}")
        else: print("  (None)")
        print("----------------------")
        print("Options:")
        print("  1. Filter by Level (Brand/Mfr/SKU)")
        print("  2. Filter by Data Category")
        print("  3. Filter by Data Topic")
        print("  4. Filter by Entity (Name)")
        print("  5. Filter by Time Range")
        print("  reset - Reset all filters")
        print("  save - Proceed to Format & Save")
        print("  exit - Exit without saving")
        print("="*50)

        choice = input("Enter choice: ").strip().lower()

        if choice == 'exit': print("Exiting without saving."); exit()
        if choice == 'save': print("Proceeding to format and save..."); break
        if choice == 'reset':
            filtered_df = master_tidy_df.copy()
            active_filters = {}
            print("Filters reset.")
            continue

        try:
            # --- Filter Logic ---
            if choice == '1': # Level
                levels = sorted(filtered_df['Level'].unique())
                print("Available Levels:"); [print(f"  {i+1}. {lvl}") for i, lvl in enumerate(levels)]
                idx = int(input("Enter number of Level to keep: ").strip()) - 1
                if 0 <= idx < len(levels):
                    keep_level = levels[idx]
                    filtered_df = filtered_df[filtered_df['Level'] == keep_level].copy()
                    active_filters['Level'] = keep_level
                    print(f"Filtered by Level: {keep_level}")
                else: print("Invalid number.")

            elif choice == '2': # Data Category
                cats = sorted(filtered_df['Data Category'].unique())
                selected_cats = select_categories_interactive(cats) # Use helper
                if selected_cats:
                    filtered_df = filtered_df[filtered_df['Data Category'].isin(selected_cats)].copy()
                    active_filters['Data Category'] = selected_cats
                    print(f"Filtered by {len(selected_cats)} Categories.")
                else: print("No categories selected for filtering.")

            elif choice == '3': # Data Topic
                topics = sorted(filtered_df['Data Topic'].unique())
                print("Available Topics:"); [print(f"  {i+1}. {t}") for i, t in enumerate(topics)]
                t_indices = input("Enter comma-separated numbers of Topics to KEEP: ").strip()
                chosen_indices = {int(i.strip()) - 1 for i in t_indices.split(',') if i.strip()}
                keep_topics = [topics[i] for i in chosen_indices if 0 <= i < len(topics)]
                if keep_topics:
                    filtered_df = filtered_df[filtered_df['Data Topic'].isin(keep_topics)].copy()
                    active_filters['Data Topic'] = keep_topics
                    print(f"Filtered by {len(keep_topics)} Topics.")
                else: print("No valid topics selected.")

            elif choice == '4': # Entity
                entities = sorted(filtered_df['Entity'].unique())
                if len(entities) > 50: # Show sample if too many
                     print("Available Entities (Sample):"); [print(f"  - {e}") for e in entities[:50]]
                     print("  ...")
                else: print("Available Entities:"); [print(f"  - {e}") for e in entities]
                entity_input = input("Enter comma-separated Entity names to KEEP (case-insensitive): ").strip()
                keep_lower = {e.strip().lower() for e in entity_input.split(',') if e.strip()}
                if keep_lower:
                     # Perform case-insensitive filtering
                     filtered_df = filtered_df[filtered_df['Entity'].str.lower().isin(keep_lower)].copy()
                     active_filters['Entity'] = list(keep_lower) # Store lowercase for display
                     print(f"Filtered by {len(keep_lower)} Entities.")
                else: print("No entities specified.")

            elif choice == '5': # Time Range
                months = sorted(filtered_df['Month'].unique(), key=lambda m: parse_month_code(m) or datetime.min)
                print("Available Months:"); print(f"  From {months[0]} to {months[-1]}")
                start_m = input("Start Month: ").strip().upper(); end_m = input("End Month: ").strip().upper()
                start_dt = parse_month_code(start_m); end_dt = parse_month_code(end_m)
                if start_dt and end_dt and start_dt <= end_dt:
                    # Filter using the datetime column
                    filtered_df = filtered_df[(filtered_df['Month_dt'] >= start_dt) & (filtered_df['Month_dt'] <= end_dt)].copy()
                    active_filters['Time Range'] = f"{start_m} to {end_m}"
                    print(f"Filtered by Time Range: {start_m} to {end_m}")
                else: print("Invalid date range or format.")

            else: print("Invalid choice.")

        except (ValueError, IndexError) as e:
            print(f"Invalid input: {e}. Please try again.")
        except Exception as e:
             print(f"An error occurred during filtering: {e}")


    # --- Step 6: Format and Save ---
    if filtered_df.empty:
        print("No data remaining after filtering. Nothing to save.")
        exit()

    # Determine output format
    output_df = None
    output_format_desc = ""
    unique_levels = filtered_df['Level'].unique()

    if len(unique_levels) == 1:
        print(f"\nData filtered to a single level ('{unique_levels[0]}'). Attempting to pivot to wide format...")
        try:
            pivot_index = ['Data Category', 'Entity', 'Month', 'Month_dt']
            pivot_cols = 'Data Topic'
            pivot_vals = 'Value'

            # Ensure index columns exist
            pivot_index = [col for col in pivot_index if col in filtered_df.columns]

            wide_df = pd.pivot_table(filtered_df,
                                     index=pivot_index,
                                     columns=pivot_cols,
                                     values=pivot_vals)
            wide_df = wide_df.reset_index().rename_axis(columns=None)

            # Reorder columns: Index cols first, then topic cols alphabetically
            id_cols_final = pivot_index
            topic_cols_final = sorted([col for col in wide_df.columns if col not in id_cols_final])
            final_col_order = id_cols_final[:1] + topic_cols_final + id_cols_final[1:] # Cat, Topics..., Entity, Month, Month_dt
            final_col_order = [col for col in final_col_order if col in wide_df.columns] # Safety check

            output_df = wide_df[final_col_order].sort_values(by=['Data Category', 'Entity', 'Month_dt'])
            output_format_desc = "WideFormat"
            print("Pivoting successful.")
        except Exception as e:
            print(f"Pivoting failed: {e}. Saving in long format.")
            output_df = filtered_df.sort_values(by=['Data Category', 'Level', 'Entity', 'Data Topic', 'Month_dt'])
            output_format_desc = "LongFormat_PivotFailed"
    else:
        print("\nData contains multiple levels. Saving in standard long format.")
        output_df = filtered_df.sort_values(by=['Data Category', 'Level', 'Entity', 'Data Topic', 'Month_dt'])
        output_format_desc = "LongFormat_MultiLevel"

    # --- Saving ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filter_desc = "_".join(active_filters.keys()).replace(" ", "") if active_filters else "Unfiltered"
    output_filename = f"{OUTPUT_FILENAME_BASE}_{output_format_desc}_{filter_desc}_{timestamp}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        print(f"\nSaving final data ({len(output_df)} rows) to: {output_path}")
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        output_df.to_excel(output_path, sheet_name=f'Organized_{output_format_desc}', index=False)
        print("Save complete.")
    except Exception as e: print(f"\nError saving file: {e}")

    print("\n--- Data Organizer script finished. ---")