# File: data_utils.py (Replace ONLY the find_time_columns function)

import pandas as pd
import numpy as np
from datetime import datetime

# --- find_data_categories function remains the same ---
def find_data_categories(df):
    """
    Identifies Data Category headers within sheets like 'Brand'/'Manufacturer'.
    (Stricter Logic - Excludes rows that look like Topic Headers)
    """
    categories = {}
    category_starts = []

    if 0 not in df.columns: return categories

    for index, row in df.iterrows():
        val_col0 = row.iloc[0]
        val_col1 = row.iloc[1] if len(row) > 1 else None # Safely get col 1

        is_potential_category = False
        is_also_topic_header = False # Flag to check if it might be a topic header

        # --- Initial Category Checks (Col 0 Text, Col 1 Blank/Year) ---
        if pd.notna(val_col0):
            str_col0 = str(val_col0).strip()
            if str_col0 != '' and pd.isna(pd.to_numeric(str_col0, errors='coerce')):
                # Check Col 1 condition (Blank or Year)
                is_col1_blank = pd.isna(val_col1) or str(val_col1).strip() == ''
                is_col1_year = False
                if isinstance(val_col1, (str, int, float)):
                    str_val1 = str(val_col1).strip().split('.')[0]
                    if len(str_val1) == 4 and str_val1.isdigit():
                        try:
                            if 2000 < int(str_val1) < 2050: is_col1_year = True
                        except ValueError: pass
                if is_col1_blank or is_col1_year:
                     # Tentatively meets category criteria based on first two cols
                     is_potential_category = True

                # --- Check if it's ALSO a Topic Header (MMMYY check) ---
                # Scan columns starting from index 1 for MMMYY pattern
                for col_idx in range(1, len(row)):
                    col_val = row.iloc[col_idx]
                    if isinstance(col_val, str):
                        month_code = col_val.strip().lstrip("'")
                        if len(month_code) == 5 and month_code[:3].isalpha() and month_code[3:].isdigit():
                            is_also_topic_header = True
                            break # Found MMMYY, it's a topic header, no need to scan further

        # --- Final Decision ---
        # Add ONLY if it met initial criteria AND was NOT identified as a topic header
        if is_potential_category and not is_also_topic_header:
            # Exclude specific known non-category rows like "Update to:"
            # You might need to add more specific exclusions if needed
            if "update to:" not in str_col0.lower():
                if not any(d['name'] == str_col0 for d in category_starts):
                     category_starts.append({'name': str_col0, 'index': index})


    if not category_starts:
        print("Info (find_data_categories): No rows matching category header criteria found.")
        return categories

    # Sort by index to ensure correct order before determining end points
    category_starts.sort(key=lambda x: x['index'])

    # Determine end indices
    num_categories = len(category_starts)
    for i, start_info in enumerate(category_starts):
        start_index_label = start_info['index']
        category_name = start_info['name']
        end_index_label = None
        if i + 1 < num_categories:
            end_index_label = category_starts[i+1]['index']
        else:
            end_index_label = df.index[-1] + 1 # Goes to end of DataFrame
        categories[category_name] = (start_index_label, end_index_label)
        # print(f"  Detected Category: '{category_name}' starting at index {start_index_label}") # Keep commented unless debugging

    return categories

# --- find_topic_block function remains the same ---
def find_topic_block(df, topic_title, search_start_label=None, search_end_label=None):
    start_index_label = None; end_index_label = None
    try:
        search_df = df; start_loc_orig = 0; scan_end_loc = len(df)
        if search_start_label is not None:
             start_loc_orig = df.index.get_loc(search_start_label)
             if search_end_label is not None:
                  if isinstance(search_end_label, (int, np.integer)) and search_end_label > df.index[-1]: scan_end_loc = len(df)
                  else: scan_end_loc = df.index.get_loc(search_end_label)
                  search_df = df.iloc[start_loc_orig:scan_end_loc]
             else: search_df = df.iloc[start_loc_orig:]; scan_end_loc = len(df)
        elif search_end_label is not None:
              if isinstance(search_end_label, (int, np.integer)) and search_end_label > df.index[-1]: scan_end_loc = len(df)
              else: scan_end_loc = df.index.get_loc(search_end_label)
              search_df = df.iloc[:scan_end_loc]
        if search_df.empty: return None, None
        if 0 not in search_df.columns: return None, None
        title_matches = search_df[search_df[0].astype(str) == str(topic_title)]
        if not title_matches.empty:
            start_index_label = title_matches.index[0]
            start_loc_in_orig_df = df.index.get_loc(start_index_label)
            current_check_loc = start_loc_in_orig_df + 1
            while current_check_loc < scan_end_loc:
                current_actual_index = df.index[current_check_loc]
                row = df.loc[current_actual_index]; first_cell_value = row.iloc[0]
                value_in_col1 = row.iloc[1] if len(row) > 1 else None
                is_blank_row = row.isnull().all() or all(str(s).strip() == "" for s in row)
                is_next_topic_header = False
                if pd.notna(first_cell_value):
                    first_cell_str = str(first_cell_value).strip()
                    if first_cell_str != '' and pd.isna(pd.to_numeric(first_cell_str, errors='coerce')):
                         is_month_code = False
                         if isinstance(value_in_col1, str):
                             val_str = value_in_col1.strip().lstrip("'") # Handle apostrophe here too
                             if len(val_str) == 5 and val_str[:3].isalpha() and val_str[3:].isdigit(): is_month_code = True
                         if is_month_code: is_next_topic_header = True
                if is_blank_row or is_next_topic_header: end_index_label = current_actual_index; break
                current_check_loc += 1
            else: end_index_label = search_end_label if search_end_label is not None else df.index[-1] + 1
        return start_index_label, end_index_label
    except KeyError as e: print(f"Error (find_topic_block): Label '{e}' not found."); return None, None
    except Exception as e: print(f"Error in find_topic_block: {e}"); return None, None

# --- find_unique_topic_titles function remains the same ---
def find_unique_topic_titles(df):
    unique_titles = set();
    if 0 not in df.columns or 1 not in df.columns: return []
    for index, row in df.iterrows():
        val_col0 = row.iloc[0]; val_col1 = row.iloc[1]
        is_topic_header = False
        if pd.notna(val_col0):
            str_col0 = str(val_col0).strip()
            if str_col0 != '' and pd.isna(pd.to_numeric(str_col0, errors='coerce')):
                # Scan across columns starting from 1 to find the first MMMYY
                first_month_idx = -1
                for c_idx in range(1, len(row)):
                     cell_val = row.iloc[c_idx]
                     if isinstance(cell_val, str):
                          month_code = cell_val.strip().lstrip("'")
                          if len(month_code) == 5 and month_code[:3].isalpha() and month_code[3:].isdigit():
                               first_month_idx = c_idx
                               break # Found first month
                if first_month_idx != -1: # Only count if a month code was found in the row
                    is_topic_header = True

        if is_topic_header: unique_titles.add(str_col0)
    return sorted(list(unique_titles))


# --- find_entity_rows function (Corrected for Location within Topic Blocks) ---
def find_entity_rows(df):
    """
    Identifies rows likely containing entity data (e.g., Brand, SKU names),
    excluding known header types and ensuring entities fall within identified topic blocks.
    """
    entity_rows = {}
    if 0 not in df.columns: return entity_rows # Need at least column 0

    # --- Step 1: Identify all Topic Blocks ---
    topic_block_ranges = [] # Store tuples of (start_loc, end_loc) for topic data rows
    topic_header_indices = set()
    unique_topics = find_unique_topic_titles(df)
    # Find ranges for ALL topics first to define valid entity zones
    for topic in unique_topics:
        start_label, end_label = find_topic_block(df, topic) # Search entire df for each topic
        if start_label is not None:
            try:
                topic_header_indices.add(start_label)
                start_loc = df.index.get_loc(start_label)
                # End location logic needs care with labels vs iloc ranges
                if end_label is None: # Goes to end of df
                     end_loc = len(df)
                elif isinstance(end_label, (int, np.integer)) and end_label > df.index[-1]:
                     end_loc = len(df) # Handle case where end label is beyond actual index
                else:
                     try:
                         end_loc = df.index.get_loc(end_label)
                     except KeyError: # End label might have been deleted, go to end
                          end_loc = len(df)

                # The actual entity rows are AFTER the header (start_loc + 1) and BEFORE the end_loc
                if start_loc + 1 < end_loc:
                     topic_block_ranges.append((start_loc + 1, end_loc)) # Use iloc ranges for easier checking
                # print(f"Debug: Topic '{topic}' block data range (iloc): {start_loc + 1} to {end_loc}") # Debug
            except KeyError:
                 print(f"Warning (find_entity_rows): Could not get location for topic header label '{start_label}'.")
            except Exception as e:
                 print(f"Warning (find_entity_rows): Error processing block range for topic '{topic}': {e}")

    # --- Step 2: Identify Category Headers (if any, less critical for this bug but good practice) ---
    category_header_indices = set()
    # Reuse the header identification logic from the previous version's first pass
    for index, row in df.iterrows():
         if index in topic_header_indices: continue # Already identified
         val_col0 = row.iloc[0]; val_col1 = row.iloc[1] if 1 in row.index else None
         if pd.notna(val_col0):
             str_col0 = str(val_col0).strip()
             if str_col0 != '' and pd.isna(pd.to_numeric(str_col0, errors='coerce')):
                 # Check Category Header
                 is_col1_blank = pd.isna(val_col1) or str(val_col1).strip() == ''
                 is_col1_year = False
                 if isinstance(val_col1, (str, int, float)):
                     str_val1 = str(val_col1).strip().split('.')[0]
                     if len(str_val1) == 4 and str_val1.isdigit():
                         try:
                            if 2000 < int(str_val1) < 2050: is_col1_year = True
                         except ValueError: pass
                 if is_col1_blank or is_col1_year:
                      # Add stricter check from find_data_categories to avoid double-counting topics
                      is_also_topic_header = False
                      for c_idx in range(1, len(row)):
                           cell_val = row.iloc[c_idx]
                           if isinstance(cell_val, str):
                                month_code = cell_val.strip().lstrip("'")
                                if len(month_code) == 5 and month_code[:3].isalpha() and month_code[3:].isdigit():
                                     is_also_topic_header = True; break
                      if not is_also_topic_header:
                           category_header_indices.add(index)

    # --- Step 3: Combine All Header Indices ---
    all_header_indices = topic_header_indices.union(category_header_indices)

    # --- Step 4: Iterate Through Rows for Entity Check (with Location Constraint) ---
    time_cols_map = find_time_columns(df) # Get time columns again to check for data presence
    time_col_indices = list(time_cols_map.values()) if time_cols_map else []
    data_start_col_index = time_col_indices[0] if time_col_indices else (1 if 1 in df.columns else 0)

    for index, row in df.iterrows():
        # Check 1: Is it a known header row?
        if index in all_header_indices: continue

        # Check 2: Does it look structurally like an entity row? (Text in col 0, data exists after)
        val_col0 = row.iloc[0]
        looks_like_entity = False
        entity_name = ""
        if pd.notna(val_col0):
            entity_name = str(val_col0).strip()
            if entity_name != '' and pd.isna(pd.to_numeric(entity_name, errors='coerce')):
                # Check for likely data presence
                has_likely_data = False
                for col_idx in range(data_start_col_index, len(row)):
                    cell_value = row.iloc[col_idx]
                    if pd.isna(cell_value):
                        if not time_col_indices or col_idx in time_col_indices: has_likely_data = True; break
                    elif isinstance(cell_value, (int, float, np.number)): has_likely_data = True; break
                    elif isinstance(cell_value, str):
                        test_val = cell_value.strip()
                        if test_val == '-': has_likely_data = True; break
                        if pd.to_numeric(test_val, errors='coerce') is not None: has_likely_data = True; break
                if has_likely_data:
                    looks_like_entity = True

        if not looks_like_entity: continue # Doesn't meet structural criteria

        # Check 3: Is the row's location *within* any identified topic block data range?
        is_within_a_topic_block = False
        try:
            current_row_loc = df.index.get_loc(index)
            for block_start_loc, block_end_loc in topic_block_ranges:
                if block_start_loc <= current_row_loc < block_end_loc:
                    is_within_a_topic_block = True
                    break # Found a valid block
        except KeyError:
             # Should not happen if iterating df.iterrows, but safety first
             print(f"Warning (find_entity_rows): Index '{index}' not found during location check.")
             continue

        # --- Add entity ONLY if ALL checks pass ---
        if is_within_a_topic_block:
             if entity_name not in entity_rows: entity_rows[entity_name] = []
             entity_rows[entity_name].append(index)
             # print(f"Debug: Identified '{entity_name}' at index {index} (iloc {current_row_loc}) as Entity within topic block.") # Debug
        # else:
             # print(f"Debug: Skipping row {index} ('{entity_name}') as not within a topic block range.") # Debug


    if not entity_rows:
         print(f"Warning (find_entity_rows): No rows identified as entities using location constraint.")
    # else:
        # print(f"Debug (find_entity_rows): Found {len(entity_rows)} unique entities.")

    return entity_rows


# --- CORRECTED find_time_columns function ---
def find_time_columns(df):
    """
    Identifies columns containing time period data (e.g., 'JAN22'),
    scanning across rows to find the start. Handles leading apostrophe.
    """
    time_cols = {}
    if 0 not in df.columns: return time_cols # Need at least col 0

    first_month_col_index = -1
    header_row_index = -1

    # Iterate through rows to find a likely topic header row
    # print("Debug: Scanning rows for topic header...") # Debug
    for index, row in df.iterrows():
        val_col0 = row.iloc[0]
        # Check if Col 0 looks like a topic title (non-blank, non-numeric text)
        if pd.notna(val_col0):
            str_col0 = str(val_col0).strip()
            if str_col0 != '' and pd.isna(pd.to_numeric(str_col0, errors='coerce')):
                # print(f"Debug: Potential topic '{str_col0}' at index {index}. Scanning columns...") # Debug
                # Found a potential topic row. Now scan its columns for the first MMMYY.
                for col_idx in range(1, len(row)): # Start scan from col 1
                    col_val = row.iloc[col_idx]
                    if isinstance(col_val, str):
                        # Strip whitespace and potential leading apostrophe
                        month_code = col_val.strip().lstrip("'")
                        if len(month_code) == 5 and month_code[:3].isalpha() and month_code[3:].isdigit():
                             # Found the first MMMYY column in this row!
                             # Optional: Verify the next column for robustness
                             is_robust = False
                             if col_idx + 1 < len(row):
                                 next_col_val = row.iloc[col_idx + 1]
                                 if isinstance(next_col_val, str):
                                      next_month_code = next_col_val.strip().lstrip("'")
                                      if len(next_month_code) == 5 and next_month_code[:3].isalpha() and next_month_code[3:].isdigit():
                                           is_robust = True # Confirmed sequence starts
                                 #else: consider is_robust = True if it's NaN? Depends on file ending.
                             else: # It was the last column, accept it
                                  is_robust = True

                             if is_robust:
                                  first_month_col_index = col_idx
                                  header_row_index = index
                                  # print(f"Debug: Found time start at row {index}, col {col_idx} ('{month_code}')") # Debug
                                  break # Stop scanning columns in THIS ROW
                                #else: continue scanning columns in this row if robustness check fails

                # If we found the start column in the loop above, break the outer row loop
                if first_month_col_index != -1:
                    break # Stop scanning ROWS

    # Check if we ever found a suitable header row and starting column
    if header_row_index == -1:
        print(f"Warning (find_time_columns): No row found with a valid topic title and subsequent 'MMMYY' columns.")
        return time_cols

    # Extract all MMMYY columns starting from the identified index IN THAT HEADER ROW
    print(f"Info (find_time_columns): Extracting time columns starting from index {first_month_col_index} using header row {header_row_index}.")
    header_row = df.loc[header_row_index]
    for col_idx in range(first_month_col_index, len(header_row)):
        col_value = header_row.iloc[col_idx]
        if isinstance(col_value, str):
            month_code = col_value.strip().lstrip("'") # Handle apostrophe again
            if len(month_code) == 5 and month_code[:3].isalpha() and month_code[3:].isdigit():
                 time_cols[month_code] = col_idx
            else:
                 # Stop if the sequence breaks (e.g., hits a blank or different format)
                 # print(f"Debug: Stopping time column extraction at index {col_idx}, value '{col_value}'") # Debug
                 break
        # Stop if not a string, unless it's NaN which signifies end of data in that row
        elif not pd.isna(col_value):
             # print(f"Debug: Stopping time column extraction at index {col_idx}, non-string value '{col_value}'") # Debug
             break

    if not time_cols:
         print("Warning (find_time_columns): Found header row but failed to extract MMMYY columns from it.")
    # else:
         # print(f"Debug: Found {len(time_cols)} time columns: {list(time_cols.items())[:5]}") # Debug

    return time_cols

# --- parse_month_code function remains the same ---
def parse_month_code(code):
    try: return datetime.strptime(code, '%b%y')
    except (ValueError, TypeError): return None

# --- parse_sheet_to_tidy function remains the same ---
def parse_sheet_to_tidy(sheet_name, df_raw, entity_type, category_map={}, canonical_sheet_to_category=None):
    print(f"\n--- Parsing Sheet: '{sheet_name}' (Entities: {entity_type}) ---")
    tidy_data_list = []
    level = entity_type
    time_cols = find_time_columns(df_raw) # Call corrected function
    topics_in_sheet = find_unique_topic_titles(df_raw)
    if not time_cols: print(f"  Warning: No time columns identified for sheet '{sheet_name}'. Skipping."); return None
    if not topics_in_sheet: print(f"  Warning: No topics found for sheet '{sheet_name}'. Skipping."); return None

    category_from_sheet = None; uses_internal_categories = False
    if sheet_name in category_map: category_from_sheet = category_map[sheet_name]
    elif sheet_name in ['Off Urban in.SR', 'Off Rural', 'MT VN', 'Vietnam Off in.SR', 'TT Off VN in.SR'] or sheet_name.startswith("TT Off "): category_from_sheet = sheet_name
    elif sheet_name in ['Brand', 'Manufacturer']: uses_internal_categories = True
    else: uses_internal_categories = True # Default assumption for unknown sheets

    if category_from_sheet: print(f"  Using category: '{category_from_sheet}'")
    elif uses_internal_categories: print(f"  Using internal categories.")

    if category_from_sheet:
        for topic in topics_in_sheet:
            topic_start_label, topic_end_label = find_topic_block(df_raw, topic)
            if topic_start_label is not None:
                try:
                    start_loc = df_raw.index.get_loc(topic_start_label)
                    if isinstance(topic_end_label, (int, np.integer)) and topic_end_label > df_raw.index[-1]: end_loc = len(df_raw)
                    else: end_loc = df_raw.index.get_loc(topic_end_label)
                    df_block_slice = df_raw.iloc[start_loc:end_loc]
                    entity_rows = df_block_slice.iloc[1:]
                    month_col_indices = list(time_cols.values())
                    valid_month_col_indices = [idx for idx in month_col_indices if idx in df_block_slice.columns]
                    valid_month_names = [m for m, idx in time_cols.items() if idx in valid_month_col_indices]
                    if not valid_month_col_indices: continue
                    for index, row in entity_rows.iterrows():
                        entity_name_raw = row.iloc[0]
                        if pd.isna(entity_name_raw) or str(entity_name_raw).strip() == '': continue
                        entity_name = str(entity_name_raw).strip()
                        entity_data = row.iloc[valid_month_col_indices]
                        for month, value in zip(valid_month_names, entity_data):
                            cleaned_value = pd.to_numeric(str(value).replace('-', ''), errors='coerce')
                            tidy_data_list.append({'Data Category': category_from_sheet, 'Data Topic': topic,'Level': level, 'Entity': entity_name,'Month': month, 'Value': cleaned_value})
                except KeyError as e: print(f"  Error slicing/parsing block for '{topic}' (label {e}).")
                except Exception as e: print(f"  Error processing block '{topic}': {e}")
    elif uses_internal_categories:
        internal_categories = find_data_categories(df_raw)
        if not internal_categories: print(f"  Error: Expected internal categories for '{sheet_name}' but found none."); return None
        for cat_name, (cat_start, cat_end) in internal_categories.items():
            # print(f"    Processing internal category: '{cat_name}'") # Can be verbose
            for topic in topics_in_sheet:
                topic_start_label, topic_end_label = find_topic_block(df_raw, topic, cat_start, cat_end)
                if topic_start_label is not None:
                    try:
                        start_loc = df_raw.index.get_loc(topic_start_label)
                        if isinstance(topic_end_label, (int, np.integer)) and topic_end_label > df_raw.index[-1]: end_loc_topic = len(df_raw)
                        else: end_loc_topic = df_raw.index.get_loc(topic_end_label)
                        if isinstance(cat_end, (int, np.integer)) and cat_end > df_raw.index[-1]: end_loc_cat = len(df_raw)
                        else: end_loc_cat = df_raw.index.get_loc(cat_end)
                        end_loc = min(end_loc_topic, end_loc_cat)
                        df_block_slice = df_raw.iloc[start_loc:end_loc]
                        entity_rows = df_block_slice.iloc[1:]
                        month_col_indices = list(time_cols.values())
                        valid_month_col_indices = [idx for idx in month_col_indices if idx in df_block_slice.columns]
                        valid_month_names = [m for m, idx in time_cols.items() if idx in valid_month_col_indices]
                        if not valid_month_col_indices: continue
                        for index, row in entity_rows.iterrows():
                            entity_name_raw = row.iloc[0]
                            if pd.isna(entity_name_raw) or str(entity_name_raw).strip() == '': continue
                            entity_name = str(entity_name_raw).strip()
                            entity_data = row.iloc[valid_month_col_indices]
                            for month, value in zip(valid_month_names, entity_data):
                                cleaned_value = pd.to_numeric(str(value).replace('-', ''), errors='coerce')
                                tidy_data_list.append({'Data Category': cat_name, 'Data Topic': topic, 'Level': level, 'Entity': entity_name, 'Month': month, 'Value': cleaned_value})
                    except KeyError as e: print(f"    Error slicing/parsing block '{topic}' in '{cat_name}' (label {e}).")
                    except Exception as e: print(f"    Error processing block '{topic}' in '{cat_name}': {e}")

    if not tidy_data_list: print(f"  --- No data parsed for sheet '{sheet_name}'. ---"); return None
    print(f"  --- Parsed {len(tidy_data_list)} points for sheet '{sheet_name}'. ---")
    return pd.DataFrame(tidy_data_list)


# --- Optional Testing Block ---
if __name__ == "__main__":
    # (Testing block remains the same as previous version)
    # ... It will now call the corrected find_time_columns ...
    try: from excel_reader import load_specified_sheets_with_variations
    except ImportError: load_specified_sheets_with_variations = None
    if load_specified_sheets_with_variations:
        test_file_path = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\Data_SKU_Hao_Hao_2M2025.xlsx"
        canonical_names = [
            'Brand', 'Manufacturer', 'Vietnam Off in.SR', 'TT Off VN in.SR', 'Off Urban in.SR',
            'Off Rural', 'MT VN', 'TT Off NE/NW in.SR', 'TT Off RRD in.SR', 'TT Off NCC in.SR',
            'TT Off SCC in.SR', 'TT Off CH in.SR', 'TT Off SE in.SR', 'TT Off MKD in.SR' ]
        variations = { # Copied from excel_reader test for consistency
            'Brand': ['brand', 'brands'], 'Manufacturer': ['manufacturer', 'manufactuer'],
            'Vietnam Off in.SR': ['vietnam off in.sr'], 'TT Off VN in.SR': ['tt off vn in.sr', 'tt off vietnam in.sr'],
            'Off Urban in.SR': ['off urban in.sr'], 'Off Rural': ['off rural'], 'MT VN': ['mt vn', 'mt vietnam'],
            'TT Off NE/NW in.SR': ['tt off ne/nw in.sr', 'ne nw'], 'TT Off RRD in.SR': ['tt off rrd in.sr', 'rrd'],
            'TT Off NCC in.SR': ['tt off ncc in.sr', 'ncc'], 'TT Off SCC in.SR': ['tt off scc in.sr', 'scc'],
            'TT Off CH in.SR': ['tt off ch in.sr', 'ch'], 'TT Off SE in.SR': ['tt off se in.sr', 'se'],
            'TT Off MKD in.SR': ['tt off mkd in.sr', 'mkd'] }
        category_abbr_map = {
            'NE NW': 'TT Off NE/NW in.SR', 'RRD': 'TT Off RRD in.SR', 'NCC': 'TT Off NCC in.SR',
            'SCC': 'TT Off SCC in.SR', 'CH': 'TT Off CH in.SR', 'SE': 'TT Off SE in.SR',
            'MKD': 'TT Off MKD in.SR' }
        entity_type_map = {
            'Brand': 'Brand', 'Manufacturer': 'Manufacturer', 'Vietnam Off in.SR':'SKU', 'TT Off VN in.SR':'SKU',
            'Off Urban in.SR':'SKU', 'Off Rural':'SKU', 'MT VN':'SKU', 'TT Off NE/NW in.SR':'SKU',
            'TT Off RRD in.SR':'SKU', 'TT Off NCC in.SR':'SKU', 'TT Off SCC in.SR':'SKU',
            'TT Off CH in.SR':'SKU', 'TT Off SE in.SR':'SKU', 'TT Off MKD in.SR':'SKU' }

        print("--- Testing data_utils.py with Multi-Sheet Parsing ---")
        loaded_data_dict = load_specified_sheets_with_variations(test_file_path, canonical_names, variations)
        all_parsed_data = []
        if loaded_data_dict:
            for canonical_sheet_name, raw_df in loaded_data_dict.items():
                entity_t = entity_type_map.get(canonical_sheet_name, 'Unknown')
                # Make sure to pass the corrected function results if needed by parse_sheet_to_tidy
                parsed_df = parse_sheet_to_tidy(canonical_sheet_name, raw_df, entity_t, category_abbr_map)
                if parsed_df is not None: all_parsed_data.append(parsed_df)
            if all_parsed_data:
                print("\n--- Combining all parsed sheets ---")
                master_df = pd.concat(all_parsed_data, ignore_index=True)
                print(f"Master DataFrame shape: {master_df.shape}")
                print("Sample of Master DataFrame:")
                print(master_df.sample(min(10, len(master_df))).to_string()) # Show sample, handle small df case
                master_df['Month_dt'] = master_df['Month'].apply(parse_month_code)
                print("\nChecking unique values (sample):")
                if not master_df.empty:
                    print("Levels:", master_df['Level'].unique())
                    print("Categories (Top 5):", master_df['Data Category'].unique()[:5])
                    print("Topics (Top 5):", master_df['Data Topic'].unique()[:5])
                    print("Entities (Top 5):", master_df['Entity'].unique()[:5])
                    print("Months (Top 5):", master_df['Month'].unique()[:5])
                else: print("Master DataFrame is empty.")
            else: print("\nNo sheets were successfully parsed.")
        else: print("\nCould not load sheets for parsing test.")
        print("--- End of data_utils.py test ---")