# File: main_gui.py (Refactored)

import customtkinter as ctk
from tkinter import filedialog, messagebox, Listbox, MULTIPLE, END, ANCHOR, Toplevel, ttk
import os
import copy
import pandas as pd
import numpy as np
from datetime import datetime
import re
import json
import threading # Added for potential future background tasks

# --- Import our backend functions ---
try:
    from excel_reader import load_specified_sheets_with_variations
    from data_utils import (find_data_categories, find_topic_block,
                           find_unique_topic_titles, find_entity_rows,
                           find_time_columns, parse_month_code,
                           parse_sheet_to_tidy)
    from action_logger import ActionLogger
except ImportError as e:
    print(f"Error importing backend modules: {e}")
    messagebox.showerror("Import Error", f"Failed to import backend functions: {e}")
    exit()

# --- Configuration & Constants ---
CANONICAL_NAMES = [ 'Brand', 'Manufacturer', 'Vietnam Off in.SR', 'TT Off VN in.SR', 'Off Urban in.SR', 'Off Rural', 'MT VN', 'TT Off NE/NW in.SR', 'TT Off RRD in.SR', 'TT Off NCC in.SR', 'TT Off SCC in.SR', 'TT Off CH in.SR', 'TT Off SE in.SR', 'TT Off MKD in.SR' ]
VARIATIONS = { 'Brand': ['brand', 'brands'], 'Manufacturer': ['manufacturer', 'manufactuer'], 'Vietnam Off in.SR': ['vietnam off in.sr'], 'TT Off VN in.SR': ['tt off vn in.sr', 'tt off vietnam in.sr'], 'Off Urban in.SR': ['off urban in.sr'], 'Off Rural': ['off rural'], 'MT VN': ['mt vn', 'mt vietnam'], 'TT Off NE/NW in.SR': ['tt off ne/nw in.sr', 'ne nw'], 'TT Off RRD in.SR': ['tt off rrd in.sr', 'rrd'], 'TT Off NCC in.SR': ['tt off ncc in.sr', 'ncc'], 'TT Off SCC in.SR': ['tt off scc in.sr', 'scc'], 'TT Off CH in.SR': ['tt off ch in.sr', 'ch'], 'TT Off SE in.SR': ['tt off se in.sr', 'se'], 'TT Off MKD in.SR': ['tt off mkd in.sr', 'mkd'] }
ENTITY_TYPE_MAP = { 'Brand': 'Brand', 'Manufacturer': 'Manufacturer', 'Vietnam Off in.SR':'SKU', 'TT Off VN in.SR':'SKU', 'Off Urban in.SR':'SKU', 'Off Rural':'SKU', 'MT VN':'SKU', 'TT Off NE/NW in.SR':'SKU', 'TT Off RRD in.SR':'SKU', 'TT Off NCC in.SR':'SKU', 'TT Off SCC in.SR':'SKU', 'TT Off CH in.SR':'SKU', 'TT Off SE in.SR':'SKU', 'TT Off MKD in.SR':'SKU' }
SHEETS_WITH_INTERNAL_CATS = {'Brand', 'Manufacturer'}
TOPIC_VARIATIONS_MAP = { 'volume': "Volume ('unit/pack)", 'value': "Value ('000 000 VND)", 'revenue': "Value ('000 000 VND)", 'share': "% Volume share", 'market share': "% Volume share", '% share': "% Volume share", 'total cov': "% Do phu ve so luong tiem - Num Total Stock Dist", 'total coverage': "% Do phu ve so luong tiem - Num Total Stock Dist", 'wtd': "Weighted Distribution Handling", 'wtd.': "Weighted Distribution Handling", 'coverage': "Weighted Distribution Handling", 'sppd': "Volume SPPD (Volume/Wtd)", "volume ('unit/pack)": "Volume ('unit/pack)", "value ('000 000 vnd)": "Value ('000 000 VND)", "% volume share": "% Volume share", "% do phu ve so luong tiem - num total stock dist": "% Do phu ve so luong tiem - Num Total Stock Dist", "weighted distribution handling": "Weighted Distribution Handling", "volume sppd (volume/wtd)": "Volume SPPD (Volume/Wtd)" }
OUTPUT_DIR = r"C:\Users\anguye18\OneDrive - dentsu\Tài liệu\Excel files\log" # Consider making this configurable
CATEGORY_ABBR_MAP = { 'NE NW': 'TT Off NE/NW in.SR', 'RRD': 'TT Off RRD in.SR', 'NCC': 'TT Off NCC in.SR', 'SCC': 'TT Off SCC in.SR', 'CH': 'TT Off CH in.SR', 'SE': 'TT Off SE in.SR', 'MKD': 'TT Off MKD in.SR'}

# --- Colors ---
BUTTON_FG_COLOR = ("black", "white")
BUTTON_TEXT_COLOR = ("white", "black")
BUTTON_HOVER_COLOR = ("#17385C", "#F2A6A6")
BUTTON_DEFAULT_WIDTH = 120 # Example default width

class App(ctk.CTk):
    # --- Constants for Actions and Keys ---
    ACTION_DELETE_TOPIC_ALL = "1"
    ACTION_DELETE_TOPIC_CHANNEL = "2"
    ACTION_DELETE_CHANNEL = "3"
    ACTION_KEEP_ENTITIES = "4"
    ACTION_KEEP_TIME_RANGE = "5"

    FILTER_KEY_LEVEL = "Level"
    FILTER_KEY_CATEGORY = "Data Category"
    FILTER_KEY_TOPIC = "Data Topic"
    FILTER_KEY_ENTITY = "Entity"
    FILTER_KEY_TIME = "Time Range"

    # Logger action types (already defined in ActionLogger, but useful for clarity here)
    LOG_ACTION_DELETE_TOPIC_ALL = "delete_topic_all"
    LOG_ACTION_DELETE_TOPIC_CHANNEL = "delete_topic_channel"
    LOG_ACTION_DELETE_CHANNEL = "delete_channel"
    LOG_ACTION_KEEP_ENTITIES = "keep_entities"
    LOG_ACTION_KEEP_TIME_RANGE = "keep_time_range"


    def __init__(self):
        super().__init__()
        self.title("Excel Data Processor"); self.geometry("900x650")
        ctk.set_appearance_mode("System"); ctk.set_default_color_theme("blue")

        # --- State variables ---
        self.input_file_path = None
        self.loaded_data_dict = None
        self.working_data_dict = None
        self.action_logger = ActionLogger()
        self.master_tidy_df = None
        self.filtered_df = None
        self.active_filters = {}
        self.filter_widgets = {}
        self.current_analysis_results = None # For Stage 3 sheet analysis
        self.final_output_format = "Long" # Default for Stage 5

        # --- Main layout ---
        self.grid_rowconfigure(0, weight=1); self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1);
        self.container = ctk.CTkFrame(self)
        self.container.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        self.container.grid_rowconfigure(0, weight=1); self.container.grid_columnconfigure(0, weight=1)

        # --- Frame management ---
        self.stage_frames = {}
        self.current_frame_name = None

        # --- Create Initial Frame (Stage 1) ---
        self._create_and_register_frame("Stage1", self.create_stage1_frame)
        self.stage_frames["Stage1"].grid(row=0, column=0, sticky="nsew") # Grid the first frame

        # --- Appearance Toggle ---
        self._create_appearance_toggle()

        # --- Show Initial Frame ---
        self.show_frame("Stage1")

    # --- Helper Methods ---
    def _create_styled_button(self, parent, text, command, width=BUTTON_DEFAULT_WIDTH, **kwargs):
        """Creates a CTkButton with the application's standard styling."""
        return ctk.CTkButton(parent, text=text, command=command, width=width,
                             fg_color=BUTTON_FG_COLOR, text_color=BUTTON_TEXT_COLOR,
                             hover_color=BUTTON_HOVER_COLOR, **kwargs)

    def _create_and_register_frame(self, name, creation_method):
        """Creates a frame using the given method and registers it."""
        frame = creation_method(self.container)
        self.stage_frames[name] = frame
        # Grid the frame initially but don't raise it yet
        frame.grid(row=0, column=0, sticky="nsew")

    def _create_appearance_toggle(self):
        """Creates the Dark/Light mode switch."""
        self.appearance_label = ctk.CTkLabel(self, text="Mode:")
        self.appearance_label.grid(row=1, column=0, padx=(10, 5), pady=5, sticky="w")
        self.appearance_switch_var = ctk.StringVar(value="off")
        self.appearance_switch = ctk.CTkSwitch(self, text="Dark", command=self.change_appearance_mode_event,
                                               variable=self.appearance_switch_var, onvalue="on", offvalue="off")
        self.appearance_switch.grid(row=1, column=0, padx=(60, 10), pady=5, sticky="w")
        # Set initial switch state
        if ctk.get_appearance_mode().lower() == "dark":
            self.appearance_switch.select()
        else:
            self.appearance_switch.deselect()

    def get_canonical_topic(self, topic_input, available_topics_list):
        """Finds the canonical topic name using TOPIC_VARIATIONS_MAP."""
        # (Moved here from global scope for better encapsulation)
        topic_input_lower = topic_input.lower()
        canonical_topic = None
        if topic_input_lower in TOPIC_VARIATIONS_MAP:
            mapped_topic = TOPIC_VARIATIONS_MAP[topic_input_lower]
            if mapped_topic in available_topics_list:
                canonical_topic = mapped_topic
            else:
                 print(f"Warning: Variation '{topic_input}' maps to '{mapped_topic}' (not available).")
        # If no map match or mapped topic not available, check for direct match
        if canonical_topic is None and topic_input in available_topics_list:
             canonical_topic = topic_input
             print(f"(Recognized direct match: '{topic_input}')")
        return canonical_topic

    def _sanitize_sheet_name(self, name):
        """Removes invalid Excel sheet characters and truncates to 31 chars."""
        # Remove invalid characters: \ / ? * [ ]
        sanitized = re.sub(r'[\\/?*\[\]]', ' ', name)
        # Truncate to 31 characters (Excel limit)
        return sanitized[:31]

    # --- Frame Navigation & Updates ---
    def show_frame(self, frame_name):
        """Brings the specified frame to the front."""
        if frame_name in self.stage_frames:
            frame = self.stage_frames[frame_name]
            frame.tkraise()
            self.current_frame_name = frame_name
            print(f"Showing Frame: {frame_name}")
            # Call update methods when specific frames are shown
            if frame_name == "Stage3": self.update_stage3_display()
            if frame_name == "Stage4": self.update_stage4_display() # Update filters/preview
            if frame_name == "Stage5": self._update_stage5_display() # Update save info
        else:
             print(f"Warning: Frame '{frame_name}' not found.")

    def change_appearance_mode_event(self):
        """Toggles between Light and Dark mode."""
        new_mode = "Dark" if self.appearance_switch_var.get() == "on" else "Light"
        ctk.set_appearance_mode(new_mode)

    # --- Stage 1: Load Raw File ---
    def create_stage1_frame(self, parent_container):
        frame = ctk.CTkFrame(parent_container)
        frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(frame, text="Step 1: Select Raw Input File", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 20), sticky="ew")
        ctk.CTkLabel(frame, text="File Path:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.file_path_entry = ctk.CTkEntry(frame, placeholder_text="No file selected", width=300, state="readonly")
        self.file_path_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # Use helper for styled buttons
        browse_button = self._create_styled_button(frame, text="Browse...", command=self.browse_file, width=100)
        browse_button.grid(row=1, column=2, padx=10, pady=10)

        self.load_button = self._create_styled_button(frame, text="Load & Proceed", command=self.load_and_proceed, state="disabled", width=200)
        self.load_button.grid(row=2, column=0, columnspan=3, padx=10, pady=20)

        self.status_label_s1 = ctk.CTkLabel(frame, text="")
        self.status_label_s1.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        return frame

    def browse_file(self):
        filepath = filedialog.askopenfilename(title="Select Raw Excel File", filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")))
        if filepath:
            self.input_file_path = filepath
            self.file_path_entry.configure(state="normal")
            self.file_path_entry.delete(0, "end")
            self.file_path_entry.insert(0, filepath)
            self.file_path_entry.configure(state="readonly")
            self.load_button.configure(state="normal")
            self.status_label_s1.configure(text="")
        else:
            self.input_file_path = None
            self.load_button.configure(state="disabled")

    def load_and_proceed(self):
        if not self.input_file_path: return
        self.status_label_s1.configure(text="Loading sheets...", text_color="gray")
        self.update_idletasks() # Refresh UI to show status
        try:
            # Use the robust excel_reader function
            self.loaded_data_dict = load_specified_sheets_with_variations(
                self.input_file_path, CANONICAL_NAMES, VARIATIONS
            )
            if not self.loaded_data_dict:
                self.status_label_s1.configure(text="Error: No relevant sheets found or loaded.", text_color="red")
                return

            self.working_data_dict = copy.deepcopy(self.loaded_data_dict)
            self.action_logger.reset() # Start fresh log for this session
            self.status_label_s1.configure(text=f"Loading successful! ({len(self.loaded_data_dict)} sheets)", text_color="green")
            print(f"Loaded {len(self.loaded_data_dict)} sheets into working_data_dict.")

            # --- Create Stage 2 if it doesn't exist ---
            if "Stage2" not in self.stage_frames:
                self._create_and_register_frame("Stage2", self.create_stage2_frame)

            # Enable preset browsing now that data is loaded
            if hasattr(self, 'preset_browse_button'):
                self.preset_browse_button.configure(state="normal")
            if hasattr(self, 'preset_path_entry'):
                 self.preset_path_entry.configure(state="disabled") # Keep disabled until browsed

            self.show_frame("Stage2")

        except Exception as e:
            self.status_label_s1.configure(text=f"Error during loading: {e}", text_color="red")
            print(f"Loading error: {e}")
            # Ensure working dict is cleared on error
            self.working_data_dict = None
            self.loaded_data_dict = None

    # --- Stage 2: Choose Path ---
    def create_stage2_frame(self, parent_container):
        frame = ctk.CTkFrame(parent_container)
        frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(frame, text="Step 2: Choose Next Action", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, padx=10, pady=(10, 10), sticky="ew")

        # Display loaded file name safely
        loaded_file_name = os.path.basename(self.input_file_path) if self.input_file_path else "N/A"
        ctk.CTkLabel(frame, text=f"Loaded File: {loaded_file_name}", wraplength=600).grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(frame, text="Clean/delete data before organizing?", font=ctk.CTkFont(size=12)).grid(row=2, column=0, padx=10, pady=(10, 5), sticky="ew")

        # Use helper for styled buttons
        delete_button = self._create_styled_button(frame, text="Clean / Delete Data First", command=self.go_to_stage3, width=250)
        delete_button.grid(row=3, column=0, padx=50, pady=10, sticky="ew")

        organize_button = self._create_styled_button(frame, text="Organize Data Directly", command=self.go_to_stage4, width=250)
        organize_button.grid(row=4, column=0, padx=50, pady=10, sticky="ew")

        # --- Preset Option ---
        ctk.CTkLabel(frame, text="--- OR ---").grid(row=5, column=0, padx=50, pady=5)
        preset_frame = ctk.CTkFrame(frame, fg_color="transparent")
        preset_frame.grid(row=6, column=0, padx=50, pady=5, sticky="ew")
        preset_frame.grid_columnconfigure(1, weight=1)

        self.preset_path_entry = ctk.CTkEntry(preset_frame, placeholder_text="Apply preset: Select .json file ->", state="disabled")
        self.preset_path_entry.grid(row=0, column=0, columnspan=2, padx=(0, 5), sticky="ew")

        # Note: Preset Browse button styling is standard CTk, not the custom one here
        self.preset_browse_button = ctk.CTkButton(preset_frame, text="Browse...", command=self._browse_preset_file, state="disabled", width=100)
        self.preset_browse_button.grid(row=0, column=2, padx=(5, 0))

        self.preset_apply_button = self._create_styled_button(frame, text="Load Preset & Apply", command=self._load_and_apply_preset, state="disabled", width=250)
        self.preset_apply_button.grid(row=7, column=0, padx=50, pady=(5, 10), sticky="ew")

        return frame

    def _browse_preset_file(self):
        filepath = filedialog.askopenfilename(
            title="Select Preset File",
            filetypes=(("JSON files", "*.json"), ("All files", "*.*")),
            initialdir=OUTPUT_DIR # Suggest log directory
        )
        if filepath:
            self.preset_path_entry.configure(state="normal") # Make editable temporarily
            self.preset_path_entry.delete(0, "end")
            self.preset_path_entry.insert(0, filepath)
            self.preset_path_entry.configure(state="readonly") # Set back to readonly
            self.preset_apply_button.configure(state="normal") # Enable apply button
        else:
             # Keep path entry as is if cancelled, but disable apply button
             self.preset_apply_button.configure(state="disabled")

    # --- Preset Loading & Application (Refactored) ---
    def _load_and_apply_preset(self):
        preset_path = self.preset_path_entry.get()
        if not preset_path or not os.path.exists(preset_path):
            messagebox.showerror("Preset Error", "Invalid or non-existent preset file selected.")
            return
        if not self.working_data_dict:
             messagebox.showerror("Data Error", "No raw data loaded to apply preset to.")
             return

        print(f"Loading and applying preset: {preset_path}")
        self._set_stage2_controls_state("disabled") # Disable buttons
        # Add status label if desired: self.status_label_s2.configure(text="Applying preset...")
        self.update_idletasks()

        try:
            with open(preset_path, 'r') as f:
                preset_data = json.load(f)

            # --- Apply Steps ---
            modified_dict = self._apply_preset_deletions(preset_data.get("deletion_steps", []))
            if modified_dict is None: return # Error handled within helper

            self.working_data_dict = modified_dict # Update main dict

            # --- Parse Data ---
            if not self._parse_data_post_preset(): return # Error handled within helper

            # --- Apply Filters ---
            if not self._apply_preset_filters(preset_data.get("filter_settings", {})): return # Error handled within helper

            # --- Transition ---
            print("Preset applied successfully. Transitioning to Stage 5...")
            self.go_to_stage5() # Go directly to save stage

        except Exception as e:
            messagebox.showerror("Preset Application Error", f"Failed to apply preset:\n{e}")
            print(f"Error applying preset: {e}")
        finally:
            # --- Re-enable Controls ---
            self._set_stage2_controls_state("normal")
            # Keep apply disabled until new file browse
            self.preset_apply_button.configure(state="disabled")
            self.preset_path_entry.configure(state="readonly")


    def _set_stage2_controls_state(self, state):
        """Helper to enable/disable all buttons in Stage 2."""
        if "Stage2" in self.stage_frames:
            for child in self.stage_frames["Stage2"].winfo_children():
                 if isinstance(child, ctk.CTkButton): # Standard buttons
                      child.configure(state=state)
                 elif isinstance(child, ctk.CTkFrame): # Check frames containing buttons
                      for grandchild in child.winfo_children():
                           if isinstance(grandchild, ctk.CTkButton):
                                grandchild.configure(state=state)
            # Special handling for preset apply button if needed
            # self.preset_apply_button.configure(state=state if state == "disabled" else "normal" if self.preset_path_entry.get() else "disabled")


    def _apply_preset_deletions(self, deletion_steps):
        """Applies deletion steps from preset data. Returns modified dict or None on error."""
        print("Applying deletion steps from preset...")
        if not self.working_data_dict: return None # Should not happen here but check
        temp_working_dict = copy.deepcopy(self.working_data_dict) # Work on a copy
        self.action_logger.reset() # Reset log before applying preset actions

        for i, step in enumerate(deletion_steps):
            sheet_key = step.get("sheet")
            action_type = step.get("action_type")
            params = step.get("parameters")
            print(f"  Step {i+1}: Action '{action_type}' on sheet '{sheet_key}'")

            if not all([sheet_key, action_type, params is not None]):
                 messagebox.showerror("Preset Error", f"Invalid deletion step format at index {i}: {step}")
                 return None # Stop processing on invalid step
            if sheet_key not in temp_working_dict:
                 print(f"  *Warning*: Sheet '{sheet_key}' specified in preset step {i+1} not found. Skipping.")
                 continue # Skip this step

            df_to_modify = temp_working_dict[sheet_key]
            # Call the helper to apply the step (doesn't interact with GUI/logger)
            modified_df = self._apply_preset_deletion_step_logic(df_to_modify, action_type, params)

            if modified_df is None:
                 messagebox.showerror("Preset Error", f"Failed to apply deletion step {i+1} for sheet '{sheet_key}'. Check console.")
                 return None # Stop processing on step failure
            temp_working_dict[sheet_key] = modified_df.reset_index(drop=True)

        print("Deletion steps applied.")
        return temp_working_dict


    def _parse_data_post_preset(self):
        """Parses sheets after preset deletions. Returns True on success, False on error."""
        print("Parsing sheets after preset deletions...")
        all_parsed_data = []
        try:
            for canonical_sheet_name, raw_df in self.working_data_dict.items():
                 if raw_df.empty: continue
                 entity_t = ENTITY_TYPE_MAP.get(canonical_sheet_name, 'Unknown')
                 # Pass CATEGORY_ABBR_MAP explicitly
                 parsed_df = parse_sheet_to_tidy(canonical_sheet_name, raw_df, entity_t, CATEGORY_ABBR_MAP)
                 if parsed_df is not None: all_parsed_data.append(parsed_df)
                 else: print(f"  *Warning*: Failed parsing sheet post-preset: '{canonical_sheet_name}'")

            if not all_parsed_data:
                 messagebox.showerror("Preset Error", "No data parsed successfully after applying preset deletions.")
                 return False

            self.master_tidy_df = pd.concat(all_parsed_data, ignore_index=True)
            self.master_tidy_df['Month_dt'] = self.master_tidy_df['Month'].apply(parse_month_code)
            self.filtered_df = self.master_tidy_df.copy() # Initialize filtered df
            print(f"Parsing complete. Master Tidy DF: {self.master_tidy_df.shape}")
            return True
        except Exception as parse_e:
             messagebox.showerror("Preset Error", f"Error during parsing after preset deletions: {parse_e}")
             print(f"Error parsing post-preset: {parse_e}")
             return False


    def _apply_preset_filters(self, filter_settings):
        """Applies filter settings from preset data. Returns True on success, False on error."""
        print("Applying filter settings from preset...")
        self.filtered_df = self.master_tidy_df.copy() # Start with all parsed data
        applied_filters_dict = {} # Track filters actually applied

        try:
            # Level Filter
            level_val = filter_settings.get(self.FILTER_KEY_LEVEL)
            if level_val and level_val != "ALL":
                 self.filtered_df = self.filtered_df[self.filtered_df[self.FILTER_KEY_LEVEL] == level_val]
                 applied_filters_dict[self.FILTER_KEY_LEVEL] = level_val

            # Category Filter
            cat_val = filter_settings.get(self.FILTER_KEY_CATEGORY)
            if cat_val: # Assumes list/set
                 self.filtered_df = self.filtered_df[self.filtered_df[self.FILTER_KEY_CATEGORY].isin(cat_val)]
                 applied_filters_dict[self.FILTER_KEY_CATEGORY] = set(cat_val) # Store as set

            # Topic Filter
            topic_val = filter_settings.get(self.FILTER_KEY_TOPIC)
            if topic_val:
                 self.filtered_df = self.filtered_df[self.filtered_df[self.FILTER_KEY_TOPIC].isin(topic_val)]
                 applied_filters_dict[self.FILTER_KEY_TOPIC] = set(topic_val)

            # Entity Filter (Case-insensitive)
            entity_val = filter_settings.get(self.FILTER_KEY_ENTITY)
            if entity_val:
                 entities_lower = {str(e).lower() for e in entity_val} # Ensure strings and lower
                 self.filtered_df = self.filtered_df[self.filtered_df[self.FILTER_KEY_ENTITY].astype(str).str.lower().isin(entities_lower)]
                 applied_filters_dict[self.FILTER_KEY_ENTITY] = set(entity_val) # Store original case

            # Time Filter
            time_str = filter_settings.get(self.FILTER_KEY_TIME)
            if time_str:
                 start_m, end_m = None, None
                 # Simplified parsing logic assuming "START to END", "From START", "Until END"
                 parts = time_str.split(' to ')
                 if len(parts) == 2: start_m, end_m = parts[0].strip(), parts[1].strip()
                 elif 'From ' in time_str: start_m = time_str.replace('From ','').strip()
                 elif 'Until ' in time_str: end_m = time_str.replace('Until ','').strip()

                 start_dt = parse_month_code(start_m) if start_m else None
                 end_dt = parse_month_code(end_m) if end_m else None

                 time_applied_str = None
                 if start_dt and end_dt and start_dt <= end_dt:
                      self.filtered_df = self.filtered_df[(self.filtered_df['Month_dt'] >= start_dt) & (self.filtered_df['Month_dt'] <= end_dt)]
                      time_applied_str = f"{start_m} to {end_m}"
                 elif start_dt:
                      self.filtered_df = self.filtered_df[self.filtered_df['Month_dt'] >= start_dt]
                      time_applied_str = f"From {start_m}"
                 elif end_dt:
                      self.filtered_df = self.filtered_df[self.filtered_df['Month_dt'] <= end_dt]
                      time_applied_str = f"Until {end_m}"

                 if time_applied_str:
                      applied_filters_dict[self.FILTER_KEY_TIME] = time_applied_str
                 elif start_m or end_m: # Only warn if something was specified but couldn't be applied
                      print(f"  *Warning*: Could not apply time range filter from preset: '{time_str}'")


            # Store the applied filters
            self.active_filters = applied_filters_dict
            self.action_logger.set_filter_settings(applied_filters_dict) # Log applied filters
            print(f"Filters applied. Final DF shape: {self.filtered_df.shape}")
            return True

        except Exception as filter_e:
             messagebox.showerror("Preset Error", f"Error applying filter settings from preset: {filter_e}")
             print(f"Error applying preset filters: {filter_e}")
             return False


    def _apply_preset_deletion_step_logic(self, df_in, action_type, params):
        """
        Internal logic to apply a single preset deletion step to a DataFrame.
        Returns modified DataFrame or None on error. No GUI interaction.
        """
        df_modified = df_in.copy()
        indices_to_drop = set()
        columns_to_drop = set() # Use column names/objects

        try:
            if action_type in [self.LOG_ACTION_DELETE_TOPIC_ALL, self.LOG_ACTION_DELETE_TOPIC_CHANNEL]:
                topic = params.get("topic")
                if not topic: raise ValueError("Missing 'topic' parameter")
                current_topics = find_unique_topic_titles(df_modified)
                if topic not in current_topics:
                     print(f"  *Preset Step Info*: Topic '{topic}' not found in current sheet state. Skipping.")
                     return df_modified

                cat_source = {}; channel = None
                if action_type == self.LOG_ACTION_DELETE_TOPIC_ALL:
                    internal_cats = find_data_categories(df_modified)
                    cat_source = internal_cats if internal_cats else {None: (None, None)}
                else: # delete_topic_channel
                    channel = params.get("channel")
                    if not channel: raise ValueError("Missing 'channel' parameter")
                    internal_cats = find_data_categories(df_modified)
                    if channel not in internal_cats:
                        print(f"  *Preset Step Info*: Channel '{channel}' for topic deletion not found. Skipping.")
                        return df_modified
                    cat_source = {channel: internal_cats[channel]}

                for cat_name, (cat_start, cat_end) in cat_source.items():
                    start_label, end_label = find_topic_block(df_modified, topic, cat_start, cat_end)
                    if start_label is not None:
                        try:
                            start_loc = df_modified.index.get_loc(start_label)
                            end_loc_val = len(df_modified) if isinstance(end_label, (int, np.integer)) and end_label > df_modified.index[-1] else df_modified.index.get_loc(end_label)
                            indices_to_drop.update(list(df_modified.iloc[start_loc:end_loc_val].index))
                        except KeyError: pass # Ignore if label already gone

            elif action_type == self.LOG_ACTION_DELETE_CHANNEL:
                channels_to_delete = params.get("channels", [])
                if not channels_to_delete: return df_modified
                internal_cats = find_data_categories(df_modified)
                for channel in channels_to_delete:
                    if channel in internal_cats:
                        cat_start, cat_end = internal_cats[channel]
                        try:
                            start_loc = df_modified.index.get_loc(cat_start)
                            end_loc_val = len(df_modified) if isinstance(cat_end, (int, np.integer)) and cat_end > df_modified.index[-1] else df_modified.index.get_loc(cat_end)
                            indices_to_drop.update(list(df_modified.iloc[start_loc:end_loc_val].index))
                        except KeyError: pass
                    else:
                        print(f"  *Preset Step Info*: Channel '{channel}' to delete not found. Skipping.")

            elif action_type == self.LOG_ACTION_KEEP_ENTITIES:
                 entities_to_keep = set(params.get("entities_to_keep", []))
                 current_entity_map = find_entity_rows(df_modified)
                 all_entities_in_df = set(current_entity_map.keys())

                 entities_to_delete = all_entities_in_df if not entities_to_keep else all_entities_in_df - entities_to_keep
                 if not entities_to_delete: return df_modified

                 rows_for_deleted = set().union(*(current_entity_map[name_orig] for name_orig in entities_to_delete if name_orig in current_entity_map))
                 indices_to_drop.update([idx for idx in rows_for_deleted if idx in df_modified.index])

            elif action_type == self.LOG_ACTION_KEEP_TIME_RANGE:
                start_m, end_m = params.get("start_month"), params.get("end_month")
                if not start_m or not end_m: raise ValueError("Missing time range parameters")
                start_dt, end_dt = parse_month_code(start_m), parse_month_code(end_m)
                if not start_dt or not end_dt or start_dt > end_dt:
                    print(f"  *Preset Step Info*: Invalid time range '{start_m}-{end_m}'. Skipping.")
                    return df_modified

                current_time_map = find_time_columns(df_modified)
                if not current_time_map:
                     print("  *Preset Step Info*: No time columns found. Cannot apply time filter.")
                     return df_modified

                cols_to_keep_indices_time = {idx for month, idx in current_time_map.items() if (m_dt := parse_month_code(month)) and start_dt <= m_dt <= end_dt}
                all_current_time_cols_indices = set(current_time_map.values())
                cols_to_drop_indices = all_current_time_cols_indices - cols_to_keep_indices_time

                if not cols_to_drop_indices: return df_modified # All columns kept

                current_cols = df_modified.columns
                columns_to_drop.update([current_cols[idx] for idx in cols_to_drop_indices if idx < len(current_cols)])

            else:
                 print(f"  *Warning*: Unknown action_type '{action_type}' in preset step. Skipping.")
                 return df_modified

            # --- Perform Drops ---
            if indices_to_drop:
                 valid_indices = sorted([idx for idx in indices_to_drop if idx in df_modified.index], reverse=True)
                 if valid_indices: df_modified = df_modified.drop(index=valid_indices)
            if columns_to_drop:
                 valid_cols_to_drop = [col for col in columns_to_drop if col in df_modified.columns]
                 if valid_cols_to_drop: df_modified = df_modified.drop(columns=valid_cols_to_drop)

            return df_modified

        except Exception as e:
             print(f"  *ERROR* applying preset step ({action_type}): {e}")
             return None # Signal failure

    # --- Stage 3: Deletion Session ---
    def create_stage3_frame(self, parent_container):
        frame = ctk.CTkFrame(parent_container)
        frame.grid_columnconfigure(0, weight=1); frame.grid_columnconfigure(1, weight=2) # Left vs Right pane ratio
        frame.grid_rowconfigure(1, weight=1) # Allow panes to expand vertically

        ctk.CTkLabel(frame, text="Step 3: Deletion Session (Optional)", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 10), sticky="ew")

        # --- Left Pane (Log & Sheet List) ---
        left_pane = ctk.CTkFrame(frame); left_pane.grid(row=1, column=0, padx=(10, 5), pady=10, sticky="nsew")
        left_pane.grid_rowconfigure(1, weight=1); left_pane.grid_rowconfigure(3, weight=2); left_pane.grid_columnconfigure(0, weight=1) # Adjust weights

        ctk.CTkLabel(left_pane, text="Change Log:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=(5,0), sticky="w")
        self.change_log_textbox = ctk.CTkTextbox(left_pane, height=100, state="disabled", wrap="word") # Wrap long lines
        self.change_log_textbox.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        ctk.CTkLabel(left_pane, text="Available Sheets:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=5, pady=(10,0), sticky="w")
        sheet_list_frame = ctk.CTkScrollableFrame(left_pane, label_text="")
        sheet_list_frame.grid(row=3, column=0, padx=5, pady=5, sticky="nsew")
        sheet_list_frame.grid_columnconfigure(0, weight=1);
        self.sheet_list_widget_frame = sheet_list_frame # Reference to populate later
        self.selected_sheet_var_s3 = ctk.StringVar(value=None) # Variable for selected sheet radio button

        # Use helper for styled button
        modify_button = self._create_styled_button(left_pane, text="Modify Selected Sheet", command=self.start_modifying_sheet)
        modify_button.grid(row=4, column=0, padx=5, pady=10)

        # --- Right Pane (Deletion Controls - Populated by start_modifying_sheet) ---
        self.right_pane = ctk.CTkFrame(frame, fg_color="transparent")
        self.right_pane.grid(row=1, column=1, padx=(5, 10), pady=10, sticky="nsew")
        # Initially empty, populated when "Modify Selected Sheet" is clicked

        # --- Bottom Buttons ---
        bottom_frame = ctk.CTkFrame(frame, fg_color="transparent"); bottom_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="ew")
        bottom_frame.grid_columnconfigure(0, weight=1) # Center button
        # Use helper for styled button
        self.proceed_to_org_button = self._create_styled_button(bottom_frame, text="Finish Deleting & Proceed to Organize", command=self.go_to_stage4, width=300)
        self.proceed_to_org_button.pack(pady=5)

        return frame

    def update_change_log_display(self):
        """Updates the change log textbox using ActionLogger."""
        if hasattr(self, 'change_log_textbox') and hasattr(self, 'action_logger'):
            self.change_log_textbox.configure(state="normal")
            self.change_log_textbox.delete("1.0", "end")
            log_text = self.action_logger.get_deletion_log_readable()
            self.change_log_textbox.insert("1.0", log_text)
            self.change_log_textbox.configure(state="disabled")
        elif hasattr(self, 'change_log_textbox'): # Fallback if logger missing
            self.change_log_textbox.configure(state="normal")
            self.change_log_textbox.delete("1.0", "end")
            self.change_log_textbox.insert("1.0", "(Action Logger not initialized)")
            self.change_log_textbox.configure(state="disabled")

    def update_stage3_display(self):
        """Refreshes the sheet list and change log when Stage 3 is shown."""
        self.update_change_log_display()

        # Clear previous radio buttons
        for widget in self.sheet_list_widget_frame.winfo_children(): widget.destroy()
        self.selected_sheet_var_s3.set(None) # Reset selection variable

        if self.working_data_dict:
            available_sheet_keys = sorted(list(self.working_data_dict.keys()))
            if available_sheet_keys:
                 # Select first sheet by default if available
                 self.selected_sheet_var_s3.set(available_sheet_keys[0])
                 for sheet_key in available_sheet_keys:
                     rows, cols = self.working_data_dict[sheet_key].shape
                     label = f"{sheet_key} ({rows}r, {cols}c)"
                     rb = ctk.CTkRadioButton(self.sheet_list_widget_frame, text=label,
                                             variable=self.selected_sheet_var_s3, value=sheet_key)
                     rb.pack(anchor="w", padx=5, pady=2)
            else: ctk.CTkLabel(self.sheet_list_widget_frame, text="(No sheets available)").pack(anchor="w", padx=5)
        else: ctk.CTkLabel(self.sheet_list_widget_frame, text="(No data loaded)").pack(anchor="w", padx=5)

        self.clear_right_pane() # Clear modification options

    # Modify clear_right_pane
    def clear_right_pane(self):
        """Clears the right pane, preserving the title label and main back button."""
        if hasattr(self, 'right_pane'):
            widgets_to_destroy = []
            # Find widgets to destroy
            for widget in self.right_pane.winfo_children():
                # Keep the title label
                if widget == getattr(self, 'modifying_sheet_label_s3', None):
                    continue
                # Keep the main back button
                if widget == getattr(self, 'back_to_main_button_s3', None):
                    continue
                widgets_to_destroy.append(widget)

            for widget in widgets_to_destroy:
                try:
                    # Use pack_forget or grid_forget depending on how widgets are added
                    # Pack is used for action_frame, input_widgets_frame, apply_button, back_button
                    widget.pack_forget()
                except Exception:
                    try:
                        widget.grid_forget()  # In case you use grid for something else later
                    except Exception:
                        pass  # Widget might not be packed or gridded
                try:
                    widget.destroy()
                except Exception as e:
                    print(f"Minor error destroying widget: {e}")  # Log minor errors if needed

        self.current_analysis_results = None  # Clear analysis

        # Reset radio button variable if it exists
        if hasattr(self, 'deletion_action_var'):
            self.deletion_action_var.set(None)

        # Ensure the main back button is packed if it exists and action inputs are cleared
        if hasattr(self, 'back_to_main_button_s3') and self.back_to_main_button_s3:
            # Check if it's managed by pack
            if self.back_to_main_button_s3.winfo_manager() == 'pack':
                # Unpack first in case it was already packed differently
                try:
                    self.back_to_main_button_s3.pack_forget()
                except Exception:
                    pass
                # Repack in its default bottom-center position ONLY if no action selected
                if not hasattr(self, 'deletion_action_var') or not self.deletion_action_var.get():
                    self.back_to_main_button_s3.pack(side="bottom", pady=10, anchor="center")
            # else: handle if managed by grid, if necessary

    def start_modifying_sheet(self):
        """Analyzes the selected sheet and sets up the right pane for deletion actions."""
        selected_sheet_key = self.selected_sheet_var_s3.get()
        if not selected_sheet_key or not self.working_data_dict:
            messagebox.showwarning("Selection Error", "Please select a sheet to modify.")
            return

        # Clear previous dynamic inputs, preserving title label and back button
        self.clear_right_pane()

        print(f"Analyzing structure for sheet: '{selected_sheet_key}'")
        df_current = self.working_data_dict[selected_sheet_key]
        entity_term = ENTITY_TYPE_MAP.get(selected_sheet_key, 'Entity')

        # --- Analyze structure ---
        internal_categories = {}
        if selected_sheet_key in SHEETS_WITH_INTERNAL_CATS:
            print(f"  Sheet '{selected_sheet_key}' expects internal categories. Running analysis...")
            internal_categories = find_data_categories(df_current)
        else:
            print(f"  Sheet '{selected_sheet_key}' does not use internal categories.")

        available_topics = find_unique_topic_titles(df_current)
        entity_rows_map = find_entity_rows(df_current)
        available_entities = sorted(list(entity_rows_map.keys()))
        time_cols_map = find_time_columns(df_current)
        available_months = sorted(list(time_cols_map.keys()), key=lambda m: parse_month_code(m) or datetime.min)
        has_internal_categories = bool(internal_categories)

        # Store analysis results for later use by deletion logic
        self.current_analysis_results = {
            "topics": available_topics, "entities": available_entities, "months": available_months,
            "categories": internal_categories, "entity_term": entity_term, "time_map": time_cols_map
        }

        # --- Warnings based on analysis ---
        if not available_topics: messagebox.showwarning("Analysis Warning",
                                                        f"No topics found in sheet '{selected_sheet_key}'. Deletion by topic may not work.",
                                                        parent=self.right_pane)
        if not available_entities: messagebox.showwarning("Analysis Warning",
                                                          f"No {entity_term}s found in sheet '{selected_sheet_key}'. Entity filtering may not work.",
                                                          parent=self.right_pane)
        if not time_cols_map: messagebox.showwarning("Analysis Warning",
                                                     f"No time columns found in sheet '{selected_sheet_key}'. Time filtering may not work.",
                                                     parent=self.right_pane)
        if selected_sheet_key in SHEETS_WITH_INTERNAL_CATS and not has_internal_categories:
            messagebox.showwarning("Analysis Warning",
                                   f"Expected internal categories in '{selected_sheet_key}' but found none! Category deletion may not work.",
                                   parent=self.right_pane)

        # --- Populate Right Pane Title (Update or Create) ---
        title_text = f"Modifying: {selected_sheet_key}"
        # Check if the label attribute exists AND if the widget itself still exists
        if hasattr(self,
                   'modifying_sheet_label_s3') and self.modifying_sheet_label_s3 and self.modifying_sheet_label_s3.winfo_exists():
            # Configure existing label
            self.modifying_sheet_label_s3.configure(text=title_text)
            # Ensure it's packed at the top if not already
            if not self.modifying_sheet_label_s3.winfo_ismapped():
                self.modifying_sheet_label_s3.pack(pady=(0, 10), anchor="w", fill="x")  # Fill ensures it takes width
        else:
            # Create and store the label if it doesn't exist or was destroyed
            self.modifying_sheet_label_s3 = ctk.CTkLabel(self.right_pane, text=title_text,
                                                         font=ctk.CTkFont(weight="bold"))
            self.modifying_sheet_label_s3.pack(pady=(0, 10), anchor="w", fill="x")  # Pack it at the top

        # --- Populate Right Pane with Actions ---
        action_frame = ctk.CTkFrame(self.right_pane, fg_color="transparent")
        # Pack it below the title label
        action_frame.pack(fill="x", padx=5, pady=5, anchor="n")
        action_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(action_frame, text="Select Action:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5,
                                                                                                pady=(0, 5), sticky="w")

        # Initialize deletion_action_var if it doesn't exist
        if not hasattr(self, 'deletion_action_var'):
            self.deletion_action_var = ctk.StringVar(value=None)
        else:
            self.deletion_action_var.set(None)  # Reset selection for new sheet

        actions = []
        actions.append((self.ACTION_DELETE_TOPIC_ALL, "Delete topic everywhere"))
        if has_internal_categories:
            actions.append((self.ACTION_DELETE_TOPIC_CHANNEL, "Delete topic within specific channel"))
            actions.append((self.ACTION_DELETE_CHANNEL, "Delete entire channel"))
        actions.append((self.ACTION_KEEP_ENTITIES, f"Keep Specific {entity_term}s (Delete others)"))
        actions.append((self.ACTION_KEEP_TIME_RANGE, "Keep Specific Time Range (Delete others)"))

        for i, (val, text) in enumerate(actions):
            rb = ctk.CTkRadioButton(action_frame, text=text, variable=self.deletion_action_var, value=val,
                                    # Use lambda v=val to capture the correct value
                                    command=lambda v=val: self.create_deletion_inputs(v, self.right_pane)
                                    )
            rb.grid(row=i + 1, column=0, padx=10, pady=2, sticky="w")

        # --- Create the main back button if it doesn't exist ---
        # Check if the attribute exists AND if the widget itself still exists
        if not hasattr(self,
                       'back_to_main_button_s3') or not self.back_to_main_button_s3 or not self.back_to_main_button_s3.winfo_exists():
            self.back_to_main_button_s3 = ctk.CTkButton(self.right_pane, text="Cancel / Back to Sheet List",
                                                        command=self.clear_right_pane, width=200,
                                                        fg_color=("gray70", "gray35"), hover_color=("gray60", "gray25"))

        # --- Manage initial packing of the back button ---
        # Ensure it's packed at the bottom only when no action inputs are shown yet
        # Check if the back button exists and if it's currently managed by pack
        if hasattr(self, 'back_to_main_button_s3') and self.back_to_main_button_s3:
            is_currently_packed = False
            try:
                if self.back_to_main_button_s3.winfo_manager() == 'pack':
                    is_currently_packed = True
            except Exception:
                is_currently_packed = False  # If widget was destroyed, winfo_manager fails

            if is_currently_packed and not self.back_to_main_button_s3.winfo_ismapped():  # Packed but not visible
                if not self.deletion_action_var.get():  # Pack only if no radio selected
                    self.back_to_main_button_s3.pack(side="bottom", pady=10, anchor="center")
            elif not is_currently_packed:  # Exists but not packed
                if not self.deletion_action_var.get():  # Pack only if no radio selected
                    self.back_to_main_button_s3.pack(side="bottom", pady=10, anchor="center")
            # If it's already packed and visible, leave it alone initially

    # Modify create_deletion_inputs
    def create_deletion_inputs(self, action_code, parent_frame):
        """Creates specific input widgets AND the Apply button below them."""
        # --- Clear previous dynamic inputs and Apply button ---
        widgets_to_destroy = []
        action_frame_found = False
        # Find widgets below the action radio button frame
        for widget in parent_frame.winfo_children():
            if isinstance(widget, ctk.CTkFrame) and widget.winfo_children():
                if isinstance(widget.winfo_children()[0], ctk.CTkLabel) and widget.winfo_children()[0].cget(
                        "text").startswith("Select Action:"):
                    action_frame_found = True
                    continue  # Keep the action frame itself
            # If we've passed the action frame, mark for destruction
            if action_frame_found:
                # Keep the main back button out of destruction path
                if widget != getattr(self, 'back_to_main_button_s3', None):
                    widgets_to_destroy.append(widget)

        for widget in widgets_to_destroy:
            # Unpack before destroying if using pack
            try:
                widget.pack_forget()
            except Exception:
                pass  # Ignore if not packed or already gone
            widget.destroy()

        # --- Retrieve analysis results ---
        if not self.current_analysis_results:
            messagebox.showerror("Error", "Analysis results not available.", parent=parent_frame)
            return
        analysis = self.current_analysis_results

        # --- Create container for dynamic inputs (packed at top) ---
        input_widgets_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        input_widgets_frame.pack(side="top", fill="x", expand=False, padx=5, pady=5)  # Pack at top

        self.deletion_details = {}  # Store input widgets/variables

        # --- Create Specific Widgets inside input_widgets_frame ---
        # (Pack widgets into input_widgets_frame now)
        if action_code in [self.ACTION_DELETE_TOPIC_ALL, self.ACTION_DELETE_TOPIC_CHANNEL]:
            ctk.CTkLabel(input_widgets_frame, text="Topic to Delete:").pack(anchor="w", padx=5, pady=(10, 0))
            topic_options = ["-- Select Topic --"] + analysis["topics"]
            topic_combobox = ctk.CTkComboBox(input_widgets_frame, values=topic_options, width=300)
            topic_combobox.pack(anchor="w", padx=5, pady=2)
            topic_entry = ctk.CTkEntry(input_widgets_frame, placeholder_text="Or type topic/variation name", width=300)
            topic_entry.pack(anchor="w", padx=5, pady=2)
            self.deletion_details["topic_combobox"] = topic_combobox
            self.deletion_details["topic_entry"] = topic_entry

            if action_code == self.ACTION_DELETE_TOPIC_CHANNEL:
                ctk.CTkLabel(input_widgets_frame, text="From Channel:").pack(anchor="w", padx=5, pady=(10, 0))
                cat_options = ["-- Select Channel --"] + list(analysis["categories"].keys())
                cat_combobox = ctk.CTkComboBox(input_widgets_frame, values=cat_options, width=300)
                cat_combobox.pack(anchor="w", padx=5, pady=2)
                self.deletion_details["category_combobox"] = cat_combobox

        elif action_code == self.ACTION_DELETE_CHANNEL:
            ctk.CTkLabel(input_widgets_frame, text="Channels to DELETE (Check items):").pack(anchor="w", padx=5,
                                                                                             pady=(10, 0))
            scroll_frame = ctk.CTkScrollableFrame(input_widgets_frame, height=150, width=300)
            scroll_frame.pack(anchor="w", padx=5, pady=2, fill="x", expand=True)  # Allow vertical expansion
            scroll_frame.grid_columnconfigure(0, weight=1)
            self.deletion_details["channel_checkbox_vars"] = {}
            if not analysis["categories"]:
                ctk.CTkLabel(scroll_frame, text="(No internal channels found)").pack(anchor="w")
            else:
                for name in analysis["categories"].keys():
                    var = ctk.StringVar(value="off")
                    cb = ctk.CTkCheckBox(scroll_frame, text=name, variable=var, onvalue="on", offvalue="off")
                    cb.pack(anchor="w", padx=5)
                    self.deletion_details["channel_checkbox_vars"][name] = var

        elif action_code == self.ACTION_KEEP_ENTITIES:
            term = analysis["entity_term"] + ("s" if analysis["entity_term"] != "Brand" else "")  # Pluralize
            ctk.CTkLabel(input_widgets_frame, text=f"{term} to KEEP (Check items):").pack(anchor="w", padx=5,
                                                                                          pady=(10, 0))
            scroll_frame = ctk.CTkScrollableFrame(input_widgets_frame, height=150, width=300)
            scroll_frame.pack(anchor="w", padx=5, pady=2, fill="x", expand=True)  # Allow vertical expansion
            scroll_frame.grid_columnconfigure(0, weight=1)
            self.deletion_details["entity_checkbox_vars"] = {}
            if not analysis["entities"]:
                ctk.CTkLabel(scroll_frame, text=f"(No {term} found)").pack(anchor="w")
            else:
                for name in analysis["entities"]:
                    var = ctk.StringVar(value="on")  # Default to KEEPING
                    cb = ctk.CTkCheckBox(scroll_frame, text=name, variable=var, onvalue="on", offvalue="off")
                    cb.pack(anchor="w", padx=5)
                    self.deletion_details["entity_checkbox_vars"][name] = var

        elif action_code == self.ACTION_KEEP_TIME_RANGE:
            ctk.CTkLabel(input_widgets_frame, text="Time Range to KEEP:").pack(anchor="w", padx=5, pady=(10, 0))
            time_entry_frame = ctk.CTkFrame(input_widgets_frame, fg_color="transparent")
            time_entry_frame.pack(fill="x", padx=5)
            start_label = ctk.CTkLabel(time_entry_frame, text="Start (e.g., JAN22):")
            start_label.pack(side="left", padx=(0, 5), pady=2)
            start_entry = ctk.CTkEntry(time_entry_frame, width=80)
            start_entry.pack(side="left", padx=5, pady=2)
            end_label = ctk.CTkLabel(time_entry_frame, text="End (e.g., DEC24):")
            end_label.pack(side="left", padx=(10, 5), pady=2)
            end_entry = ctk.CTkEntry(time_entry_frame, width=80)
            end_entry.pack(side="left", padx=5, pady=2)
            self.deletion_details["start_entry"] = start_entry
            self.deletion_details["end_entry"] = end_entry

    # --- Manage Buttons ---
        # Unpack the main back button first if it's currently packed
        if hasattr(self, 'back_to_main_button_s3') and self.back_to_main_button_s3.winfo_ismapped():
            self.back_to_main_button_s3.pack_forget()

        # Create the Apply button (don't pack yet)
        apply_button = self._create_styled_button(parent_frame, text="Apply This Deletion",
                                                  command=self.apply_deletion_action, width=200)

        # Re-pack the main Cancel/Back button FIRST at the very bottom
        if hasattr(self, 'back_to_main_button_s3'):
            self.back_to_main_button_s3.pack(side="bottom", pady=(5, 10))  # Add some padding below

        # Pack the Apply button SECOND (it will appear *above* the Back button)
        apply_button.pack(side="bottom", pady=(10, 5))  # Add padding above and below

    # --- Deletion Action Handlers (Refactored) ---
    def apply_deletion_action(self):
        """Main dispatcher for applying the selected deletion action."""
        selected_action = self.deletion_action_var.get()
        selected_sheet_key = self.selected_sheet_var_s3.get()

        if not selected_action or not selected_sheet_key or not self.working_data_dict:
            messagebox.showwarning("Action Error", "No action or sheet selected, or data is missing.")
            return
        if not self.current_analysis_results: # Should be set by start_modifying_sheet
            messagebox.showerror("Internal Error", "Sheet analysis results are missing.")
            return

        df_original = self.working_data_dict[selected_sheet_key]
        original_rows, original_cols = df_original.shape
        df_modified = df_original.copy() # Work on a copy initially

        log_params = {}; indices_to_drop = set(); columns_to_drop = set()
        action_type_for_log = None
        success = False

        try:
            if selected_action == self.ACTION_DELETE_TOPIC_ALL:
                action_type_for_log = self.LOG_ACTION_DELETE_TOPIC_ALL
                success, indices_to_drop, log_params = self._handle_delete_topic(df_modified, action_type_for_log)
            elif selected_action == self.ACTION_DELETE_TOPIC_CHANNEL:
                action_type_for_log = self.LOG_ACTION_DELETE_TOPIC_CHANNEL
                success, indices_to_drop, log_params = self._handle_delete_topic(df_modified, action_type_for_log)
            elif selected_action == self.ACTION_DELETE_CHANNEL:
                action_type_for_log = self.LOG_ACTION_DELETE_CHANNEL
                success, indices_to_drop, log_params = self._handle_delete_channel(df_modified)
            elif selected_action == self.ACTION_KEEP_ENTITIES:
                action_type_for_log = self.LOG_ACTION_KEEP_ENTITIES
                success, indices_to_drop, log_params = self._handle_keep_entities(df_modified)
            elif selected_action == self.ACTION_KEEP_TIME_RANGE:
                action_type_for_log = self.LOG_ACTION_KEEP_TIME_RANGE
                success, columns_to_drop, log_params = self._handle_keep_time_range(df_modified)

            if not success: # Handler indicated an input error or no action needed
                 return # Message shown by handler

            # --- Perform Actual Drops ---
            made_changes = False
            if indices_to_drop:
                valid_indices = sorted([idx for idx in indices_to_drop if idx in df_modified.index], reverse=True)
                if valid_indices:
                    print(f"Dropping {len(valid_indices)} rows...")
                    df_modified = df_modified.drop(index=valid_indices)
                    made_changes = True
            if columns_to_drop: # columns_to_drop now contains column NAMES/OBJECTS
                valid_cols_to_drop = [col for col in columns_to_drop if col in df_modified.columns]
                if valid_cols_to_drop:
                    print(f"Dropping {len(valid_cols_to_drop)} columns...")
                    df_modified = df_modified.drop(columns=valid_cols_to_drop)
                    made_changes = True

            # --- Update State if Changes Made ---
            if made_changes:
                df_modified = df_modified.reset_index(drop=True)
                self.working_data_dict[selected_sheet_key] = df_modified # Update main dict

                # Log the action using ActionLogger
                if action_type_for_log:
                    self.action_logger.log_deletion(sheet=selected_sheet_key, action_type=action_type_for_log, **log_params)

                self.update_change_log_display() # Refresh log display in left pane
                messagebox.showinfo("Success",
                                    f"Operation applied to '{selected_sheet_key}'.\nRows changed: {original_rows} -> {len(df_modified)}\nCols changed: {original_cols} -> {len(df_modified.columns)}")
                # Re-analyze and clear pane for next action on this sheet
                self.start_modifying_sheet()

            else: # No indices or columns to drop were generated by the handler
                messagebox.showinfo("No Change", "Operation resulted in no changes to the data.", parent=self.right_pane)
                # Optionally clear inputs or keep them: self.clear_right_pane()
                # Let's keep inputs visible if no change, maybe user wants to tweak.

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during deletion: {e}", parent=self.right_pane)
            print(f"Error during deletion apply: {e}")

    # --- Specific Deletion Logic Handlers ---
    def _handle_delete_topic(self, df_mod, action_type):
        indices_to_drop = set(); log_params = {}
        analysis = self.current_analysis_results

        topic_input = self.deletion_details["topic_entry"].get().strip()
        if not topic_input: topic_input = self.deletion_details["topic_combobox"].get()
        if not topic_input or topic_input == "-- Select Topic --":
             messagebox.showwarning("Input Error", "Please select or enter a topic.", parent=self.right_pane); return False, indices_to_drop, log_params
        canonical_topic = self.get_canonical_topic(topic_input, analysis["topics"])
        if canonical_topic is None:
             messagebox.showerror("Input Error", f"Topic '{topic_input}' not recognized.", parent=self.right_pane); return False, indices_to_drop, log_params
        print(f"Action targets canonical topic: '{canonical_topic}'")
        log_params["topic"] = canonical_topic

        cat_source = {}; target_category = None
        if action_type == self.LOG_ACTION_DELETE_TOPIC_ALL:
             cat_source = analysis["categories"] if analysis["categories"] else {None: (None, None)}
        else: # Delete Topic Channel
             target_category = self.deletion_details["category_combobox"].get()
             if target_category == "-- Select Channel --":
                  messagebox.showwarning("Input Error", "Please select a channel.", parent=self.right_pane); return False, indices_to_drop, log_params
             if target_category not in analysis["categories"]:
                  messagebox.showerror("Error", "Selected channel not found.", parent=self.right_pane); return False, indices_to_drop, log_params
             cat_source = {target_category: analysis["categories"][target_category]}
             log_params["channel"] = target_category

        found_count = 0
        for cat_name, (cat_start, cat_end) in cat_source.items():
            start_label, end_label = find_topic_block(df_mod, canonical_topic, cat_start, cat_end)
            if start_label is not None:
                found_count += 1
                try:
                    start_loc = df_mod.index.get_loc(start_label)
                    end_loc = len(df_mod) if isinstance(end_label, (int, np.integer)) and end_label > df_mod.index[-1] else df_mod.index.get_loc(end_label)
                    indices_to_drop.update(list(df_mod.iloc[start_loc:end_loc].index))
                except KeyError: pass # Ignore if index gone

        if found_count == 0:
             messagebox.showinfo("Info", f"Topic '{canonical_topic}' not found for deletion in the specified scope.", parent=self.right_pane)
             return False, indices_to_drop, log_params # Indicate no action needed

        return True, indices_to_drop, log_params

    def _handle_delete_channel(self, df_mod):
        indices_to_drop = set(); log_params = {}
        analysis = self.current_analysis_results
        checkbox_vars = self.deletion_details.get("channel_checkbox_vars", {})
        if not checkbox_vars:
            messagebox.showerror("Error", "Could not find channel selection controls.", parent=self.right_pane); return False, indices_to_drop, log_params
        channels_to_delete = {name for name, var in checkbox_vars.items() if var.get() == "on"}
        if not channels_to_delete:
            messagebox.showinfo("Info", "No channels selected for deletion.", parent=self.right_pane); return False, indices_to_drop, log_params
        confirm = messagebox.askyesno("Confirm Delete", f"Delete {len(channels_to_delete)} selected channels?", parent=self.right_pane)
        if not confirm: return False, indices_to_drop, log_params

        log_params = {"channels": list(channels_to_delete)} # Log intended channels
        deleted_count = 0
        for target_category in channels_to_delete:
            if target_category in analysis["categories"]:
                cat_start, cat_end = analysis["categories"][target_category]
                try:
                    start_loc = df_mod.index.get_loc(cat_start)
                    end_loc = len(df_mod) if isinstance(cat_end, (int, np.integer)) and cat_end > df_mod.index[-1] else df_mod.index.get_loc(cat_end)
                    indices_to_drop.update(list(df_mod.iloc[start_loc:end_loc].index))
                    deleted_count += 1
                except KeyError: print(f"Warning: Indices for channel '{target_category}' not found (already deleted?).")
            else: print(f"Warning: Channel '{target_category}' selected but not found in analysis results.")

        if deleted_count == 0: # Logged intent, but no rows found
             messagebox.showinfo("Info", "No deletable rows found for the selected channels.", parent=self.right_pane)
             return False, indices_to_drop, log_params

        return True, indices_to_drop, log_params

    def _handle_keep_entities(self, df_mod):
        indices_to_drop = set(); log_params = {}
        analysis = self.current_analysis_results
        entity_term = analysis["entity_term"]
        available_entities_orig = analysis["entities"]
        checkbox_vars = self.deletion_details.get("entity_checkbox_vars", {})
        if not checkbox_vars:
             messagebox.showerror("Error", f"Could not find {entity_term} selection controls.", parent=self.right_pane); return False, indices_to_drop, log_params

        entities_to_keep_orig = {name for name, var in checkbox_vars.items() if var.get() == "on"}
        log_params = {"entities_to_keep": list(entities_to_keep_orig)} # Log kept entities

        if not entities_to_keep_orig:
            confirm = messagebox.askyesno("Confirm Delete All", f"Delete ALL {entity_term} data?", parent=self.right_pane)
            if not confirm: return False, indices_to_drop, log_params

        entities_to_delete_orig = set(available_entities_orig) - entities_to_keep_orig
        if not entities_to_delete_orig:
            messagebox.showinfo("Info", f"All available {entity_term}s are selected to keep. No deletion.", parent=self.right_pane); return False, indices_to_drop, log_params

        print(f"Keeping: {len(entities_to_keep_orig)} {entity_term}s")
        print(f"Attempting to delete: {len(entities_to_delete_orig)} {entity_term}s")

        # Re-find entities on the potentially modified dataframe before dropping
        current_entity_map = find_entity_rows(df_mod)
        rows_for_deleted = set().union(*(current_entity_map[name_orig] for name_orig in entities_to_delete_orig if name_orig in current_entity_map))
        valid_indices_found = [idx for idx in rows_for_deleted if idx in df_mod.index]

        if not valid_indices_found:
             print(f"Note: No rows found corresponding to the {entity_term}s marked for deletion (already removed?).")
             messagebox.showinfo("Info", "No rows found for the entities marked for deletion.", parent=self.right_pane)
             return False, indices_to_drop, log_params # No action needed

        indices_to_drop.update(valid_indices_found)
        return True, indices_to_drop, log_params

    def _handle_keep_time_range(self, df_mod):
        columns_to_drop = set(); log_params = {}
        analysis = self.current_analysis_results
        start_m = self.deletion_details["start_entry"].get().strip().upper()
        end_m = self.deletion_details["end_entry"].get().strip().upper()
        start_dt = parse_month_code(start_m); end_dt = parse_month_code(end_m)

        if not start_dt or not end_dt or start_dt > end_dt:
             messagebox.showerror("Input Error", "Invalid date range or format (use MMMYY).", parent=self.right_pane); return False, columns_to_drop, log_params
        log_params = {"start_month": start_m, "end_month": end_m}

        # Use current time map from analysis results (should reflect current df state if re-analyzed)
        # If not re-analyzing often, might need: current_time_map = find_time_columns(df_mod)
        current_time_map = analysis["time_map"]
        if not current_time_map:
             messagebox.showerror("Error", "Could not find time columns in current sheet state.", parent=self.right_pane); return False, columns_to_drop, log_params

        cols_to_keep_indices = {idx for month, idx in current_time_map.items() if (m_dt := parse_month_code(month)) and start_dt <= m_dt <= end_dt}
        if not cols_to_keep_indices:
             messagebox.showinfo("Info", "No columns matched the specified time range.", parent=self.right_pane); return False, columns_to_drop, log_params

        all_current_time_cols_indices = set(current_time_map.values())
        cols_to_drop_indices = all_current_time_cols_indices - cols_to_keep_indices
        if not cols_to_drop_indices:
             messagebox.showinfo("Info", "All time columns are within the specified range.", parent=self.right_pane); return False, columns_to_drop, log_params

        current_cols = df_mod.columns
        # Store actual column NAMES/OBJECTS to drop
        columns_to_drop.update([current_cols[idx] for idx in cols_to_drop_indices if idx < len(current_cols)])
        if not columns_to_drop: # Should not happen if cols_to_drop_indices is not empty, but safety check
             return False, columns_to_drop, log_params

        return True, columns_to_drop, log_params


    # --- Stage 4: Organize & Filter Data ---
    def create_stage4_frame(self, parent_container):
        frame = ctk.CTkFrame(parent_container)
        frame.grid_columnconfigure(0, weight=1); frame.grid_columnconfigure(1, weight=3) # Filter vs Preview ratio
        frame.grid_rowconfigure(1, weight=1) # Allow preview to expand

        ctk.CTkLabel(frame, text="Step 4: Organize & Filter Combined Data", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 10), sticky="ew")

        # --- Left Pane (Filters) ---
        filter_pane = ctk.CTkFrame(frame)
        filter_pane.grid(row=1, column=0, padx=(10, 5), pady=10, sticky="nsew")
        filter_pane.grid_columnconfigure(0, weight=1)
        row_num = 0

        ctk.CTkLabel(filter_pane, text="Active Filters:", font=ctk.CTkFont(weight="bold")).grid(row=row_num, column=0, padx=5, pady=(5,0), sticky="w"); row_num+=1
        self.active_filters_label_s4 = ctk.CTkLabel(filter_pane, text="(None)", anchor="nw", justify="left", wraplength=250) # Wrap long filter lists
        self.active_filters_label_s4.grid(row=row_num, column=0, padx=5, pady=2, sticky="new"); row_num+=1

        # Scrollable frame for filter controls
        filter_scroll_frame = ctk.CTkScrollableFrame(filter_pane, label_text="Apply Filters")
        filter_scroll_frame.grid(row=row_num, column=0, padx=5, pady=5, sticky="nsew"); row_num+=1
        filter_pane.grid_rowconfigure(row_num-1, weight=1) # Make scroll frame expandable
        filter_scroll_frame.grid_columnconfigure(0, weight=1)

        self.filter_widgets = {} # Re-initialize here

        # -- Level Filter --
        level_frame = ctk.CTkFrame(filter_scroll_frame, fg_color="transparent")
        level_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(level_frame, text="Level:").pack(side="left", padx=5)
        self.filter_widgets["level_var"] = ctk.StringVar(value="ALL")
        self.level_filter_frame = level_frame
        # Radiobuttons will be populated in update_stage4_display

        # -- Other Filters (using popups) --
        ctk.CTkButton(filter_scroll_frame, text="Select Categories...", command=self.open_multi_select_category, width=150).pack(fill="x", pady=5, padx=10)
        self.filter_widgets["categories"] = set()
        ctk.CTkButton(filter_scroll_frame, text="Select Topics...", command=self.open_multi_select_topic, width=150).pack(fill="x", pady=5, padx=10)
        self.filter_widgets["topics"] = set()
        ctk.CTkButton(filter_scroll_frame, text="Select Entities...", command=self.open_multi_select_entity, width=150).pack(fill="x", pady=5, padx=10)
        self.filter_widgets["entities"] = set()

        # -- Time Filter --
        time_frame = ctk.CTkFrame(filter_scroll_frame, fg_color="transparent")
        time_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(time_frame, text="Time Range:").pack(side="left", padx=5)
        start_entry = ctk.CTkEntry(time_frame, placeholder_text="Start (MMMYY)", width=100)
        start_entry.pack(side="left", padx=3)
        end_entry = ctk.CTkEntry(time_frame, placeholder_text="End (MMMYY)", width=100)
        end_entry.pack(side="left", padx=3)
        self.filter_widgets["time_start"] = start_entry
        self.filter_widgets["time_end"] = end_entry

        # --- Filter Action Buttons ---
        filter_action_frame = ctk.CTkFrame(filter_pane) # Below scroll frame
        filter_action_frame.grid(row=row_num, column=0, padx=5, pady=5, sticky="ew"); row_num+=1
        filter_action_frame.grid_columnconfigure((0, 1, 2), weight=1) # Flexible columns

        apply_filt_btn = self._create_styled_button(filter_action_frame, text="Apply Filters", command=self.apply_filters_s4)
        apply_filt_btn.grid(row=0, column=0, padx=5, pady=5)
        reset_filt_btn = ctk.CTkButton(filter_action_frame, text="Reset Filters", command=self.reset_filters_s4, width=BUTTON_DEFAULT_WIDTH,
                                      fg_color=("gray50", "gray30"), hover_color=("gray40", "gray20")) # Standard style for reset
        reset_filt_btn.grid(row=0, column=1, padx=5, pady=5)
        export_preset_btn = ctk.CTkButton(filter_action_frame, text="Export Preset...", command=self.export_preset, width=BUTTON_DEFAULT_WIDTH) # Standard style
        export_preset_btn.grid(row=0, column=2, padx=5, pady=5)

        # --- Right Pane (Preview) ---
        preview_pane = ctk.CTkFrame(frame)
        preview_pane.grid(row=1, column=1, padx=(5, 10), pady=10, sticky="nsew")
        preview_pane.grid_rowconfigure(1, weight=1); preview_pane.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(preview_pane, text="Data Preview (Filtered):", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.preview_textbox_s4 = ctk.CTkTextbox(preview_pane, state="disabled", wrap="none") # No wrap for tables
        self.preview_textbox_s4.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")

        # --- Bottom Buttons ---
        bottom_frame_s4 = ctk.CTkFrame(frame, fg_color="transparent")
        bottom_frame_s4.grid(row=2, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="e") # Align right
        self.proceed_to_save_button_s4 = self._create_styled_button(bottom_frame_s4, text="Proceed to Save", command=self.go_to_stage5)
        self.proceed_to_save_button_s4.pack(pady=5, padx=5)

        return frame

    def update_stage4_display(self):
         """Called when Stage 4 is shown. Populates filters based on master_tidy_df and resets view."""
         print("Updating Stage 4 display...")
         if not hasattr(self, 'master_tidy_df') or self.master_tidy_df is None:
             messagebox.showerror("Error", "Master data not available for Stage 4.")
             self.show_frame("Stage1") # Go back if data is missing
             return

         # --- Reset filters visually and internally ---
         self.reset_filters_s4() # This clears active_filters, filter_widgets values, and updates preview/labels

         # --- Repopulate Level Filter Radiobuttons ---
         if "level_var" in self.filter_widgets and hasattr(self,
                                                           'level_filter_frame') and self.level_filter_frame:  # Check if frame exists
             level_var = self.filter_widgets["level_var"]
             level_rb_parent_frame = self.level_filter_frame  # <-- USE THE STORED REFERENCE

             # Clear old buttons except the label
             for widget in level_rb_parent_frame.winfo_children():
                 if isinstance(widget, ctk.CTkRadioButton): widget.destroy()

             # Add new buttons based on current master data
             levels = ["ALL"] + sorted(self.master_tidy_df[self.FILTER_KEY_LEVEL].unique())
             for level in levels:
                 rb = ctk.CTkRadioButton(level_rb_parent_frame, text=level, variable=level_var, value=level)
                 rb.pack(side="left", padx=3)
             level_var.set("ALL")  # Ensure default is selected
         else:
             # This handles the case where the frame wasn't found or doesn't exist
             print("Warning: Level filter frame reference not found. Cannot update Level radio buttons.")

         print("Stage 4 Display Updated.")


    def update_preview_s4(self):
         """Updates the preview textbox with current filtered data."""
         if not hasattr(self, 'preview_textbox_s4'): return
         self.preview_textbox_s4.configure(state="normal")
         self.preview_textbox_s4.delete("1.0", "end")
         if self.filtered_df is not None and not self.filtered_df.empty:
             # Show sample, convert to string with limited rows
             preview_text = self.filtered_df.head(50).to_string() # Adjust row limit as needed
             self.preview_textbox_s4.insert("1.0", preview_text)
         elif self.filtered_df is not None and self.filtered_df.empty:
             self.preview_textbox_s4.insert("1.0", "(No data matching current filters)")
         else:
             self.preview_textbox_s4.insert("1.0", "(Data not loaded or filtered yet)")
         self.preview_textbox_s4.configure(state="disabled")

    def update_active_filters_label_s4(self):
         """Updates the label showing active filters."""
         if not hasattr(self, 'active_filters_label_s4'): return
         if not self.active_filters:
              self.active_filters_label_s4.configure(text="(None)")
              return

         filter_lines = []
         for key, value in self.active_filters.items():
             display_value = value
             max_items_to_show = 5
             if isinstance(value, (list, set)):
                 sorted_list = sorted(list(value))
                 display_value = ", ".join(map(str, sorted_list[:max_items_to_show]))
                 if len(sorted_list) > max_items_to_show: display_value += f"... ({len(sorted_list)} total)"
             elif value is None or value == "": continue # Skip empty filters
             filter_lines.append(f"- {key}: {display_value}")

         self.active_filters_label_s4.configure(text="\n".join(filter_lines) if filter_lines else "(None)")


    def apply_filters_s4(self):
        """Applies all selected filters to the master tidy dataframe."""
        print("Applying filters...")
        if not hasattr(self, 'master_tidy_df') or self.master_tidy_df is None:
             messagebox.showwarning("Filter Error", "Master data is not available.")
             return

        temp_df = self.master_tidy_df.copy()
        current_active_filters = {} # Build active filters for this run

        try:
            # 1. Level Filter
            level_filter = self.filter_widgets["level_var"].get()
            if level_filter != "ALL":
                temp_df = temp_df[temp_df[self.FILTER_KEY_LEVEL] == level_filter]
                current_active_filters[self.FILTER_KEY_LEVEL] = level_filter

            # 2. Category Filter
            if self.filter_widgets.get("categories"):
                temp_df = temp_df[temp_df[self.FILTER_KEY_CATEGORY].isin(self.filter_widgets["categories"])]
                current_active_filters[self.FILTER_KEY_CATEGORY] = self.filter_widgets["categories"]

            # 3. Topic Filter
            if self.filter_widgets.get("topics"):
                temp_df = temp_df[temp_df[self.FILTER_KEY_TOPIC].isin(self.filter_widgets["topics"])]
                current_active_filters[self.FILTER_KEY_TOPIC] = self.filter_widgets["topics"]

            # 4. Entity Filter (Case-insensitive)
            if self.filter_widgets.get("entities"):
                selected_entities_lower = {str(e).lower() for e in self.filter_widgets["entities"]}
                temp_df = temp_df[temp_df[self.FILTER_KEY_ENTITY].astype(str).str.lower().isin(selected_entities_lower)]
                current_active_filters[self.FILTER_KEY_ENTITY] = self.filter_widgets["entities"] # Store original case

            # 5. Time Filter
            start_m = self.filter_widgets["time_start"].get().strip().upper()
            end_m = self.filter_widgets["time_end"].get().strip().upper()
            start_dt = parse_month_code(start_m)
            end_dt = parse_month_code(end_m)
            time_range_str = None
            if start_dt and end_dt:
                if start_dt <= end_dt:
                    temp_df = temp_df[(temp_df['Month_dt'] >= start_dt) & (temp_df['Month_dt'] <= end_dt)]
                    time_range_str = f"{start_m} to {end_m}"
                else: messagebox.showwarning("Filter Warning", "Start month is after end month. Time filter ignored.", parent=self.stage_frames["Stage4"])
            elif start_dt:
                temp_df = temp_df[temp_df['Month_dt'] >= start_dt]
                time_range_str = f"From {start_m}"
            elif end_dt:
                temp_df = temp_df[temp_df['Month_dt'] <= end_dt]
                time_range_str = f"Until {end_m}"
            if time_range_str: current_active_filters[self.FILTER_KEY_TIME] = time_range_str

            # Update the main filtered df and active filters dict
            self.filtered_df = temp_df
            self.active_filters = current_active_filters

            # Store applied filters in the logger
            if hasattr(self, 'action_logger'):
                self.action_logger.set_filter_settings(self.active_filters)

            print(f"Filtering complete. {len(self.filtered_df)} rows remaining.")
            self.update_preview_s4()
            self.update_active_filters_label_s4()

        except Exception as e:
             messagebox.showerror("Filter Error", f"An error occurred while applying filters:\n{e}", parent=self.stage_frames["Stage4"])
             print(f"Error applying filters: {e}")


    def reset_filters_s4(self):
        """Resets all filters to default and updates the view."""
        print("Resetting filters...")
        if not hasattr(self, 'master_tidy_df') or self.master_tidy_df is None: return
        self.filtered_df = self.master_tidy_df.copy()
        self.active_filters = {}

        # Clear stored filters in logger
        if hasattr(self, 'action_logger'):
            self.action_logger.set_filter_settings({})

        # Reset GUI filter widgets to defaults
        if "level_var" in self.filter_widgets: self.filter_widgets["level_var"].set("ALL")
        self.filter_widgets["categories"] = set()
        self.filter_widgets["topics"] = set()
        self.filter_widgets["entities"] = set()
        if "time_start" in self.filter_widgets: self.filter_widgets["time_start"].delete(0, "end")
        if "time_end" in self.filter_widgets: self.filter_widgets["time_end"].delete(0, "end")

        # Update displays
        self.update_preview_s4()
        self.update_active_filters_label_s4()
        print("Filters reset.")

    # --- Multi-Select Popups ---
    def open_multi_select_popup(self, title, items, selected_set_key):
        """Opens a Toplevel window for multi-selection with search."""
        if not hasattr(self, 'filter_widgets'): # Ensure filter_widgets exists
             print("Error: filter_widgets not initialized.")
             return
        # Get the currently selected *original* items for pre-selection
        currently_selected_orig_items = self.filter_widgets.get(selected_set_key, set())

        popup = Toplevel(self); popup.title(title); popup.geometry("400x450")
        popup.transient(self); popup.grab_set() # Make modal

        search_var = ctk.StringVar()
        search_entry = ctk.CTkEntry(popup, textvariable=search_var, placeholder_text="Search...")
        search_entry.pack(fill="x", padx=5, pady=(5, 0))

        list_frame = ctk.CTkFrame(popup); list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        listbox = Listbox(list_frame, selectmode=MULTIPLE, exportselection=False, width=50, height=15,
                          # Set background/foreground based on current theme
                          bg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkFrame"]["fg_color"]),
                          # Pass the tuple
                          fg=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkLabel"]["text_color"]),
                          # Pass the tuple
                          selectbackground=self._apply_appearance_mode(ctk.ThemeManager.theme["CTkButton"]["fg_color"]),
                          # Pass the tuple
                          selectforeground=self._apply_appearance_mode(
                              ctk.ThemeManager.theme["CTkButton"]["text_color"]),  # Pass the tuple
                          borderwidth=0, highlightthickness=0 # Match CTk look
                          )
        scrollbar = ctk.CTkScrollbar(list_frame, command=listbox.yview)
        listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y"); listbox.pack(side="left", fill="both", expand=True)

        # items might be simple strings or "[Level] Entity"
        all_items_formatted = sorted(list(items))

        def get_original_from_formatted(formatted_string):
            # Handles "[Level] Entity" format, returns original Entity or the string itself
            match = re.match(r"\[.*?\]\s*(.*)", formatted_string)
            return match.group(1) if match else formatted_string

        def populate_list(items_to_show):
            listbox.delete(0, END)
            for i, formatted_item in enumerate(items_to_show):
                listbox.insert(END, formatted_item)
                original_item = get_original_from_formatted(formatted_item)
                if original_item in currently_selected_orig_items:
                    listbox.selection_set(i)

        populate_list(all_items_formatted)

        def filter_listbox(*args):
            search_term = search_var.get().lower()
            filtered_items = [item for item in all_items_formatted if search_term in item.lower()] if search_term else all_items_formatted
            populate_list(filtered_items)

        search_var.trace_add("write", filter_listbox)

        button_frame = ctk.CTkFrame(popup); button_frame.pack(fill="x", padx=5, pady=(0, 5))
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        def select_all(): listbox.selection_set(0, END)
        def clear_all(): listbox.selection_clear(0, END)

        def on_ok():
            selected_indices = listbox.curselection()
            current_displayed_items = listbox.get(0, END) # Get items currently *in* the listbox

            # Store the *original* item names by parsing the selection
            selected_original_items = {get_original_from_formatted(current_displayed_items[i]) for i in selected_indices}

            # Update the main filter_widgets dictionary
            self.filter_widgets[selected_set_key] = selected_original_items

            print(f"Selected {len(selected_original_items)} items for '{selected_set_key}'")
            popup.destroy()
            self.update_active_filters_label_s4() # Update label immediately

        # Use styled button for OK
        ok_button = self._create_styled_button(button_frame, text="OK", command=on_ok, width=80)
        ok_button.grid(row=0, column=0, padx=5, pady=5)
        # Standard buttons for Select/Clear All
        select_all_button = ctk.CTkButton(button_frame, text="Select All", command=select_all, width=80); select_all_button.grid(row=0, column=1, padx=5, pady=5)
        clear_all_button = ctk.CTkButton(button_frame, text="Clear All", command=clear_all, width=80); clear_all_button.grid(row=0, column=2, padx=5, pady=5)

    def open_multi_select_category(self):
        if hasattr(self, 'master_tidy_df') and self.master_tidy_df is not None and not self.master_tidy_df.empty:
            items = sorted(self.master_tidy_df[self.FILTER_KEY_CATEGORY].unique())
            self.open_multi_select_popup("Select Data Categories", items, "categories") # Use constant key "categories"
        else: messagebox.showerror("Error", "Master data not ready.", parent=self.stage_frames["Stage4"])

    def open_multi_select_topic(self):
        if hasattr(self, 'master_tidy_df') and self.master_tidy_df is not None and not self.master_tidy_df.empty:
            items = sorted(self.master_tidy_df[self.FILTER_KEY_TOPIC].unique())
            self.open_multi_select_popup("Select Data Topics", items, "topics") # Use constant key "topics"
        else: messagebox.showerror("Error", "Master data not ready.", parent=self.stage_frames["Stage4"])

    def open_multi_select_entity(self):
        if not hasattr(self, 'master_tidy_df') or self.master_tidy_df is None or self.master_tidy_df.empty:
            messagebox.showerror("Error", "Master data not ready.", parent=self.stage_frames["Stage4"])
            return
        try:
            entity_data = self.master_tidy_df[[self.FILTER_KEY_LEVEL, self.FILTER_KEY_ENTITY]].drop_duplicates()
            entity_data = entity_data.sort_values(by=[self.FILTER_KEY_LEVEL, self.FILTER_KEY_ENTITY])
            formatted_items = [f"[{row[self.FILTER_KEY_LEVEL]}] {row[self.FILTER_KEY_ENTITY]}" for index, row in entity_data.iterrows()]
            self.open_multi_select_popup("Select Entities", formatted_items, "entities") # Use constant key "entities"
        except Exception as e:
            messagebox.showerror("Error", f"Could not prepare entity list: {e}", parent=self.stage_frames["Stage4"])
            print(f"Error preparing entity list: {e}")

    # --- Preset Export ---
    def export_preset(self):
        """Exports the current deletion steps and filter settings to a JSON file."""
        if not hasattr(self, 'action_logger'):
             messagebox.showerror("Error", "Action Logger not available.", parent=self.stage_frames["Stage4"])
             return

        preset_data = self.action_logger.get_preset_data()
        if not preset_data.get("deletion_steps") and not preset_data.get("filter_settings"):
             messagebox.showinfo("Export Preset", "No deletion steps or filter settings logged to create a preset.", parent=self.stage_frames["Stage4"])
             return

        file_path = filedialog.asksaveasfilename(
            title="Save Preset File", defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=OUTPUT_DIR, parent=self.stage_frames["Stage4"]
        )
        if not file_path: return # User cancelled

        try:
            with open(file_path, 'w') as f: json.dump(preset_data, f, indent=4)
            messagebox.showinfo("Export Preset", f"Preset saved successfully to:\n{file_path}", parent=self.stage_frames["Stage4"])
            print(f"Preset saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to save preset file:\n{e}", parent=self.stage_frames["Stage4"])
            print(f"Error saving preset: {e}")


    # --- Stage 5: Format & Save Output ---
    def create_stage5_frame(self, parent_container):
        frame = ctk.CTkFrame(parent_container)
        frame.grid_columnconfigure(1, weight=1) # Make entry/labels expand

        ctk.CTkLabel(frame, text="Step 5: Format & Save Output", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 15), sticky="ew")

        # Info Section
        info_frame = ctk.CTkFrame(frame, fg_color="transparent"); info_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        info_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(info_frame, text="Final Data:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=5, sticky="w")
        self.final_data_info_label = ctk.CTkLabel(info_frame, text="(Info unavailable)")
        self.final_data_info_label.grid(row=0, column=1, padx=5, sticky="w")
        ctk.CTkLabel(info_frame, text="Applied Filters:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=5, sticky="nw")
        self.final_filters_label = ctk.CTkLabel(info_frame, text="(None)", justify="left", anchor="nw", wraplength=500) # Wrap long filter lists
        self.final_filters_label.grid(row=1, column=1, padx=5, sticky="new")
        ctk.CTkLabel(info_frame, text="Output Format:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=5, pady=(5,0), sticky="w")
        self.output_format_label = ctk.CTkLabel(info_frame, text="(Will be determined)")
        self.output_format_label.grid(row=2, column=1, padx=5, pady=(5,0), sticky="w")

        # Save Location Section
        save_frame = ctk.CTkFrame(frame); save_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=15, sticky="ew")
        save_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(save_frame, text="Save Location:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.save_path_entry = ctk.CTkEntry(save_frame, placeholder_text="Click Browse to select location and filename", state="readonly", width=400)
        self.save_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.save_browse_button = self._create_styled_button(save_frame, text="Browse...", command=self._browse_save_location, width=100)
        self.save_browse_button.grid(row=0, column=2, padx=5, pady=5)

        # Save Action
        action_frame = ctk.CTkFrame(frame, fg_color="transparent"); action_frame.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")
        action_frame.grid_columnconfigure(0, weight=1)
        self.save_button = self._create_styled_button(action_frame, text="Save Organized Data", command=self._save_final_data_threaded, state="disabled", width=200) # Use threaded save
        self.save_button.pack(pady=5)

        # Progress Bar and Status
        self.progress_bar_s5 = ttk.Progressbar(frame, mode='indeterminate', length=300)
        self.progress_bar_s5.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        self.progress_bar_s5.grid_remove() # Hide initially
        self.status_label_s5 = ctk.CTkLabel(frame, text="")
        self.status_label_s5.grid(row=5, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

        # Restart Button
        restart_button = ctk.CTkButton(frame, text="Process Another File (Restart)", command=self._restart_app,
                                       fg_color=("gray50", "gray30"), hover_color=("gray40", "gray20"), width=200)
        restart_button.grid(row=6, column=0, columnspan=3, padx=10, pady=(10,5)) # Centered

        return frame

    def _update_stage5_display(self):
        """Populates labels when Stage 5 is shown. Assumes multi-sheet wide output."""
        if not hasattr(self, 'filtered_df') or self.filtered_df is None:
            messagebox.showerror("Error", "No filtered data available for saving.",
                                 parent=self.stage_frames.get("Stage5"))
            self.show_frame("Stage4")  # Go back
            return

        # Update Info Label
        rows, cols = self.filtered_df.shape
        data_cols = cols - 1 if 'Month_dt' in self.filtered_df.columns else cols  # Exclude Month_dt col
        self.final_data_info_label.configure(text=f"{rows} total rows across levels")

        # Update Filters Label (same logic as before)
        if not self.active_filters:
            self.final_filters_label.configure(text="(None Applied)")
        else:
            filter_lines = []
            for key, value in self.active_filters.items():
                display_value = value
                max_items = 5
                if isinstance(value, (list, set)):
                    val_list = sorted(list(value))
                    display_value = ", ".join(map(str, val_list[:max_items]))
                    if len(val_list) > max_items: display_value += f"... ({len(val_list)} total)"
                elif not value:
                    continue  # Skip empty values
                filter_lines.append(f"- {key}: {display_value}")
            self.final_filters_label.configure(text="\n".join(filter_lines) if filter_lines else "(None Applied)")

        # Update Output Format Label (Now static message, unless no data)
        if self.filtered_df.empty:
            self.output_format_label.configure(text="N/A (No data)")
        else:
            self.output_format_label.configure(text="Multiple Wide Sheets (per Level)")  # <-- Key change here

        # Reset save path and button state (same logic as before)
        self.save_path_entry.configure(state="normal")
        self.save_path_entry.delete(0, "end")
        self.save_path_entry.configure(state="readonly")
        self.save_button.configure(state="disabled")
        self.status_label_s5.configure(text="")
        self.progress_bar_s5.grid_remove()

    def _generate_suggested_filename(self):
        """Creates a descriptive filename based on format and filters."""
        base = "Organized_Data"
        format_desc = self.final_output_format.split(" ")[0] # "Wide" or "Long" or "NA"
        filter_parts = []
        if self.active_filters:
             # Use constants for keys
             if (lvl := self.active_filters.get(self.FILTER_KEY_LEVEL)) and lvl != 'ALL': filter_parts.append(f"Lvl_{lvl}")
             if cats := self.active_filters.get(self.FILTER_KEY_CATEGORY): filter_parts.append(f"Cat{len(cats)}")
             if tops := self.active_filters.get(self.FILTER_KEY_TOPIC): filter_parts.append(f"Top{len(tops)}")
             if ents := self.active_filters.get(self.FILTER_KEY_ENTITY): filter_parts.append(f"Ent{len(ents)}")
             if time_rng := self.active_filters.get(self.FILTER_KEY_TIME):
                 time_str = re.sub(r'[^\w\d-]+', '_', time_rng) # Sanitize time range
                 filter_parts.append(f"Time_{time_str}")

        filter_desc = "_".join(filter_parts) if filter_parts else "Unfiltered"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M") # Shorter timestamp
        clean_filter_desc = re.sub(r'[\\/*?:"<>|]', '', filter_desc)[:50] # Sanitize and limit length

        return f"{base}_{format_desc}_{clean_filter_desc}_{timestamp}.xlsx"

    def _browse_save_location(self):
        """Opens dialog to choose save location and filename."""
        suggested_name = self._generate_suggested_filename()
        filepath = filedialog.asksaveasfilename(
            title="Save Organized Data As...", initialfile=suggested_name,
            defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialdir=OUTPUT_DIR, parent=self.stage_frames.get("Stage5")
        )
        if filepath:
            self.save_path_entry.configure(state="normal")
            self.save_path_entry.delete(0, "end")
            self.save_path_entry.insert(0, filepath)
            self.save_path_entry.configure(state="readonly")
            self.save_button.configure(state="normal")
            self.status_label_s5.configure(text="")
        else:
            self.save_button.configure(state="disabled")

    def _save_final_data_threaded(self):
         """Initiates the save process in a separate thread to avoid GUI freeze."""
         output_path = self.save_path_entry.get()
         if not output_path:
             messagebox.showwarning("Save Error", "Please select a save location first.", parent=self.stage_frames.get("Stage5"))
             return
         if self.filtered_df is None or self.filtered_df.empty:
              messagebox.showwarning("Save Error", "No data available to save.", parent=self.stage_frames.get("Stage5"))
              return

         # --- Disable buttons, show progress ---
         self._set_stage5_controls_state("disabled")
         self.status_label_s5.configure(text="Preparing data and saving...", text_color="gray")
         self.progress_bar_s5.grid()
         self.progress_bar_s5.start(10)
         self.update_idletasks()

         # --- Start background thread ---
         save_thread = threading.Thread(target=self._save_final_data_logic, args=(output_path,), daemon=True)
         save_thread.start()

    def _set_stage5_controls_state(self, state):
        """Enable/disable relevant controls in Stage 5."""
        if hasattr(self, 'save_button'): self.save_button.configure(state=state)
        if hasattr(self, 'save_browse_button'): self.save_browse_button.configure(state=state)
        # Potentially disable restart button during save?
        # if hasattr(self, 'restart_button'): self.restart_button.configure(state=state)

    def _save_final_data_logic(self, output_path):
        """Saves data by creating a pivoted sheet for each Level."""
        saved_sheets_count = 0
        failed_levels = []
        try:
            print(f"Preparing multi-sheet wide output...")
            if not hasattr(self, 'filtered_df') or self.filtered_df is None or self.filtered_df.empty:
                raise ValueError("Filtered data is missing or empty.")

            unique_levels = sorted(self.filtered_df[self.FILTER_KEY_LEVEL].unique())
            print(f"Found levels to process: {unique_levels}")

            # --- Use ExcelWriter for multiple sheets ---
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for level in unique_levels:
                    print(f"  Processing Level: {level}")
                    # Filter data for the current level
                    level_df = self.filtered_df[self.filtered_df[self.FILTER_KEY_LEVEL] == level].copy()

                    if level_df.empty:
                        print(f"    Skipping Level '{level}' as it has no data after filtering.")
                        continue

                    # --- Attempt to pivot this level's data ---
                    pivoted_df = self._pivot_data_for_save(level_df)  # Pass the subset DataFrame

                    if pivoted_df is not None:
                        # --- Save the pivoted sheet ---
                        # Use the helper to sanitize the level name for the sheet
                        sanitized_sheet_name = self._sanitize_sheet_name(str(level))
                        print(f"    Saving pivoted data to sheet: '{sanitized_sheet_name}' ({len(pivoted_df)} rows)")
                        pivoted_df.to_excel(writer, sheet_name=sanitized_sheet_name, index=False)
                        saved_sheets_count += 1
                    else:
                        # --- Handle pivot failure for this level ---
                        print(
                            f"    *Warning*: Pivoting failed for Level '{level}'. This level will not be included in the output.")
                        failed_levels.append(str(level))
                        # Optional: Add fallback to save as long format here if desired

            # --- Final Success Feedback (Scheduled for main thread) ---
            if saved_sheets_count > 0:
                # Pass the counts and failed levels to the feedback function
                self.after(0, self._save_success_feedback, output_path, saved_sheets_count, failed_levels)
            else:
                # If no sheets were saved at all (e.g., all levels failed to pivot)
                raise ValueError("No sheets could be successfully pivoted and saved.")

        except Exception as e:
            # --- Error Feedback (Scheduled for main thread) ---
            print(f"Error during multi-sheet saving logic: {e}")
            self.after(0, self._save_error_feedback, e)

    def _pivot_data_for_save(self, df_to_pivot):  # <-- Takes DataFrame as argument now
        """Attempts to pivot the GIVEN DataFrame. Returns pivoted DataFrame or None on failure."""
        print("    Attempting to pivot data for current level...")
        try:
            # --- Input validation ---
            # Check if essential columns for pivoting exist in the passed DataFrame
            required_pivot_cols = [self.FILTER_KEY_CATEGORY, self.FILTER_KEY_ENTITY, 'Month', self.FILTER_KEY_TOPIC,
                                   'Value']
            if not all(col in df_to_pivot.columns for col in required_pivot_cols):
                print("    *Error*: Missing required columns for pivoting in the provided dataframe.")
                return None

            # Define pivot index columns
            pivot_index_cols = [self.FILTER_KEY_CATEGORY, self.FILTER_KEY_ENTITY, 'Month']

            # Sort data before pivoting for consistent output
            # Use 'Month_dt' if available, otherwise fall back to string sort on 'Month'
            sort_cols = [self.FILTER_KEY_CATEGORY, self.FILTER_KEY_ENTITY]
            if 'Month_dt' in df_to_pivot.columns:
                sort_cols.append('Month_dt')
            else:
                print("    *Warning*: 'Month_dt' column not found for sorting. Sorting by 'Month' string.")
                sort_cols.append('Month')

            # Handle potential errors during sorting
            try:
                # Drop duplicates before pivoting if necessary (can happen with edge cases)
                # Consider dropping Month_dt before pivot if it causes issues,
                # but sorting by it first is generally preferred.
                # Example: df_to_pivot.drop_duplicates(subset=pivot_index_cols + [self.FILTER_KEY_TOPIC], keep='last', inplace=True)
                sorted_long_df = df_to_pivot.sort_values(by=sort_cols)
            except KeyError as sort_e:
                print(f"    *Error* during sorting before pivot: {sort_e}. Cannot proceed.")
                return None

            # Pivot the sorted data
            wide_df = pd.pivot_table(sorted_long_df,
                                     index=pivot_index_cols,
                                     columns=self.FILTER_KEY_TOPIC,
                                     values='Value',
                                     aggfunc='first')  # Use aggfunc='first' or 'mean' if duplicates might exist after sorting
            wide_df = wide_df.reset_index().rename_axis(columns=None)

            # Reorder columns: Cat, Topics..., Entity, Month
            id_cols_final = pivot_index_cols
            # Get topic columns correctly, ensuring they exist in the pivoted frame
            topic_cols_final = sorted([col for col in wide_df.columns if
                                       col not in id_cols_final and col in df_to_pivot[self.FILTER_KEY_TOPIC].unique()])

            # Construct final column order carefully
            final_col_order = []
            if id_cols_final and id_cols_final[0] in wide_df.columns:  # Add category first if exists
                final_col_order.append(id_cols_final[0])
            final_col_order.extend(topic_cols_final)  # Add topics
            if len(id_cols_final) > 1:  # Add remaining index cols (Entity, Month)
                final_col_order.extend([col for col in id_cols_final[1:] if col in wide_df.columns])

            # Ensure all requested columns actually exist (redundant check, but safe)
            final_col_order = [col for col in final_col_order if col in wide_df.columns]

            print("    Pivoting successful for this level.")
            return wide_df[final_col_order]

        except Exception as pivot_e:
            print(f"    *Error*: Pivoting failed for this level: {pivot_e}")
            # Avoid GUI popups from threads if possible
            # messagebox.showwarning("Pivot Warning", f"Could not pivot data for level being processed:\n{pivot_e}", parent=self.stage_frames.get("Stage5"))
            return None  # Signal failure

    def _save_success_feedback(self, output_path, sheet_count, failed_levels):
        """Updates GUI after successful multi-sheet save."""
        self.progress_bar_s5.stop()
        self.progress_bar_s5.grid_remove()
        success_msg = f"Data saved successfully to {sheet_count} sheet(s)!"

        # Check if any levels failed
        if failed_levels:
            # Add warning about failed levels
            failed_str = ", ".join(failed_levels)
            success_msg += f"\nWarning: Failed to pivot/save levels: {failed_str}"
            self.status_label_s5.configure(text=success_msg, text_color="orange")  # Use orange for partial success
            # Show a warning messagebox
            messagebox.showwarning("Save Complete (with issues)",
                                   f"File saved to:\n{output_path}\n\n{sheet_count} sheet(s) created.\nFailed to process levels: {failed_str}",
                                   parent=self.stage_frames.get("Stage5"))
        else:
            # All levels succeeded
            self.status_label_s5.configure(text=success_msg, text_color="green")
            # Show an info messagebox
            messagebox.showinfo("Save Complete",
                                f"File saved successfully to:\n{output_path}\n\n{sheet_count} sheet(s) created.",
                                parent=self.stage_frames.get("Stage5"))

        # Keep buttons disabled after save, regardless of partial failure
        self._set_stage5_controls_state("disabled")


    def _save_error_feedback(self, error):
        """Updates GUI after save error (called from main thread)."""
        self.progress_bar_s5.stop()
        self.progress_bar_s5.grid_remove()
        self.status_label_s5.configure(text=f"Error saving file: {error}", text_color="red")
        messagebox.showerror("Save Error", f"An error occurred while saving:\n{error}", parent=self.stage_frames.get("Stage5"))
        # Re-enable buttons on error to allow retry or different path
        self._set_stage5_controls_state("normal")
        # But disable save button if path is now invalid somehow (though unlikely)
        if not self.save_path_entry.get():
             self.save_button.configure(state="disabled")


    # --- Application Restart ---
    def _restart_app(self):
        """Resets application state to start over."""
        print("Restarting application workflow...")
        # Reset state variables
        self.input_file_path = None
        self.loaded_data_dict = None
        self.working_data_dict = None
        self.action_logger.reset()
        self.master_tidy_df = None
        self.filtered_df = None
        self.active_filters = {}
        self.filter_widgets = {}
        self.current_analysis_results = None # Clear Stage 3 analysis
        self.final_output_format = "Long" # Reset default

        # Clear Stage 1 path entry and status
        if hasattr(self, 'file_path_entry'):
            self.file_path_entry.configure(state="normal")
            self.file_path_entry.delete(0, "end")
            self.file_path_entry.configure(state="readonly")
        if hasattr(self, 'load_button'): self.load_button.configure(state="disabled")
        if hasattr(self, 'status_label_s1'): self.status_label_s1.configure(text="")

        # Clear Stage 2 preset path
        if hasattr(self, 'preset_path_entry'):
             self.preset_path_entry.configure(state="normal")
             self.preset_path_entry.delete(0, "end")
             self.preset_path_entry.configure(state="disabled")
        if hasattr(self, 'preset_browse_button'): self.preset_browse_button.configure(state="disabled")
        if hasattr(self, 'preset_apply_button'): self.preset_apply_button.configure(state="disabled")


        # Clear Stage 3 (if exists)
        if "Stage3" in self.stage_frames:
            self.clear_right_pane()
            if hasattr(self, 'change_log_textbox'): self.update_change_log_display() # Show empty log
            if hasattr(self, 'sheet_list_widget_frame'): # Clear sheet list
                for widget in self.sheet_list_widget_frame.winfo_children(): widget.destroy()
                self.selected_sheet_var_s3.set(None)

        # Clear Stage 4 (if exists)
        if "Stage4" in self.stage_frames:
             if hasattr(self, 'preview_textbox_s4'): self.update_preview_s4() # Show empty preview
             if hasattr(self, 'active_filters_label_s4'): self.update_active_filters_label_s4() # Show no filters
             # Reset filter widget values (already done in self.filter_widgets = {})

        # Clear Stage 5 display elements
        if "Stage5" in self.stage_frames:
            if hasattr(self, 'save_path_entry'):
                 self.save_path_entry.configure(state="normal"); self.save_path_entry.delete(0, "end"); self.save_path_entry.configure(state="readonly")
            if hasattr(self, 'save_button'): self.save_button.configure(state="disabled")
            if hasattr(self, 'status_label_s5'): self.status_label_s5.configure(text="")
            if hasattr(self, 'final_data_info_label'): self.final_data_info_label.configure(text="(Info unavailable)")
            if hasattr(self, 'final_filters_label'): self.final_filters_label.configure(text="(None)")
            if hasattr(self, 'output_format_label'): self.output_format_label.configure(text="(Will be determined)")
            if hasattr(self, 'progress_bar_s5'): self.progress_bar_s5.grid_remove()

        # Go back to Stage 1
        self.show_frame("Stage1")

    # --- Transitions ---
    def go_to_stage3(self):
        """Transition to Deletion stage."""
        print("Transitioning to Stage 3 (Deletion)...")
        if "Stage3" not in self.stage_frames:
            self._create_and_register_frame("Stage3", self.create_stage3_frame)
        # update_stage3_display is called by show_frame
        self.show_frame("Stage3")

    def go_to_stage4(self):
        """Parse all sheets, combine, and transition to Filtering stage."""
        print("Preparing Data for Organization & Filtering (Stage 4)...")
        if not self.working_data_dict:
             messagebox.showerror("Error", "No data loaded or available to organize."); return

        # --- Parse sheets ---
        all_parsed_data = []
        print("Parsing sheets into tidy format...")
        for canonical_sheet_name, raw_df in self.working_data_dict.items():
            if raw_df.empty: print(f"  Skipping empty sheet: '{canonical_sheet_name}'"); continue
            entity_t = ENTITY_TYPE_MAP.get(canonical_sheet_name, 'Unknown')
            # Pass CATEGORY_ABBR_MAP explicitly
            parsed_df = parse_sheet_to_tidy(canonical_sheet_name, raw_df, entity_t, CATEGORY_ABBR_MAP)
            if parsed_df is not None: all_parsed_data.append(parsed_df)
            else: print(f"  *Warning*: Failed parsing sheet: '{canonical_sheet_name}'")

        if not all_parsed_data:
            messagebox.showerror("Error", "No data could be parsed successfully from any sheet."); return

        # --- Combine Data ---
        print("Combining parsed data...")
        try:
            self.master_tidy_df = pd.concat(all_parsed_data, ignore_index=True)
            self.master_tidy_df['Month_dt'] = self.master_tidy_df['Month'].apply(parse_month_code)
            self.filtered_df = self.master_tidy_df.copy() # Initialize filtered view
            print(f"Master Tidy DataFrame created: {self.master_tidy_df.shape}")

            # --- Create and Transition ---
            print("Transitioning to Stage 4 (Filtering)...")
            if "Stage4" not in self.stage_frames:
                 self._create_and_register_frame("Stage4", self.create_stage4_frame)
            # update_stage4_display called by show_frame
            self.show_frame("Stage4")

        except Exception as e:
             messagebox.showerror("Error", f"Failed to combine or process parsed data: {e}")
             print(f"Error combining data or creating Stage 4: {e}")


    def go_to_stage5(self):
        """Transition to Save stage."""
        print("Transitioning to Stage 5 (Save)...");
        # Final filter state should already be logged by apply_filters_s4, reset_filters_s4, or _load_and_apply_preset

        if "Stage5" not in self.stage_frames:
             self._create_and_register_frame("Stage5", self.create_stage5_frame)

        # _update_stage5_display called by show_frame
        self.show_frame("Stage5")


# --- Run the App ---
if __name__ == "__main__":
    app = App()
    app.mainloop()
