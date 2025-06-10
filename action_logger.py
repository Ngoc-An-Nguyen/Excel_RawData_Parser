# File: action_logger.py
import json
from datetime import datetime
import copy # To ensure filter settings are copied

class ActionLogger:
    """
    Handles logging of user actions (deletions, filters)
    and generating preset data structures.
    """
    def __init__(self):
        self.reset()

    def reset(self):
        """Resets the log and filter settings."""
        self.deletion_steps = []      # Stores structured deletion actions (list of dicts)
        self.filter_settings = {}     # Stores the final applied filter config (dict)
        print("Action logger reset.")

    def log_deletion(self, sheet, action_type, **kwargs):
        """
        Logs a specific deletion action with its parameters.

        Args:
            sheet (str): The canonical name of the sheet being modified.
            action_type (str): A unique identifier for the action
                               (e.g., 'delete_topic_all', 'keep_entities').
            **kwargs: Keyword arguments specific to the action
                      (e.g., topic='...', entities_to_keep=[...], etc.).
        """
        action = {
            "sheet": sheet,
            "action_type": action_type,
            "parameters": {}
        }

        # --- Populate parameters based on action_type and kwargs ---
        # Convert sets to lists for JSON compatibility where needed
        if action_type in ["delete_topic_all", "delete_topic_channel"]:
            action["parameters"]["topic"] = kwargs.get("topic")
            if action_type == "delete_topic_channel":
                action["parameters"]["channel"] = kwargs.get("channel")
        elif action_type == "delete_channel":
            # Ensure we get a list, even if None/empty set was passed
            channels = kwargs.get("channels", [])
            action["parameters"]["channels"] = list(channels) if channels else []
        elif action_type == "keep_entities":
            entities = kwargs.get("entities_to_keep", [])
            action["parameters"]["entities_to_keep"] = list(entities) if entities else []
        elif action_type == "keep_time_range":
            action["parameters"]["start_month"] = kwargs.get("start_month")
            action["parameters"]["end_month"] = kwargs.get("end_month")
        else:
            print(f"Warning: Unknown action_type '{action_type}' for logging.")
            return # Don't log unknown actions

        # Basic validation (optional but recommended)
        if not action["parameters"] or all(v is None for v in action["parameters"].values()):
             print(f"Warning: No valid parameters provided for action '{action_type}' on sheet '{sheet}'. Skipping log.")
             return

        self.deletion_steps.append(action)
        print(f"Action Logger: Logged {action_type} for sheet '{sheet}'")

    def get_deletion_log_readable(self):
        """Generates a human-readable string representation of the deletion log."""
        if not self.deletion_steps:
            return "(No deletion steps logged)"

        readable_lines = []
        for i, step in enumerate(self.deletion_steps):
            sheet = step.get("sheet", "Unknown Sheet")
            action = step.get("action_type", "Unknown Action")
            params = step.get("parameters", {})
            details = "No details"

            try: # Add try-except for safer formatting
                if action == "delete_topic_all":
                    details = f"Delete Topic '{params.get('topic', '?')}' everywhere"
                elif action == "delete_topic_channel":
                     details = f"Delete Topic '{params.get('topic', '?')}' from Channel '{params.get('channel', '?')}'"
                elif action == "delete_channel":
                     ch_list = params.get('channels', [])
                     details = f"Delete Channels: {', '.join(ch_list) if ch_list else 'None specified'}"
                elif action == "keep_entities":
                     kept_list = params.get('entities_to_keep', [])
                     # Show limited list for readability
                     kept_str = ", ".join(map(str, kept_list[:3])) + ('...' if len(kept_list) > 3 else '')
                     details = f"Keep only Entities: {kept_str if kept_list else 'None specified (Delete All)'}"
                elif action == "keep_time_range":
                     start = params.get('start_month','?')
                     end = params.get('end_month','?')
                     details = f"Keep Time Range: {start} - {end}"
                else:
                    details = f"Unknown action '{action}' with params {params}"

                readable_lines.append(f"{i+1}. [{sheet}] {details}")
            except Exception as e:
                 readable_lines.append(f"{i+1}. [{sheet}] Error formatting log entry: {e}")

        return "\n".join(readable_lines)

    def get_deletion_steps_structured(self):
        """Returns the raw list of deletion step dictionaries."""
        return self.deletion_steps

    def set_filter_settings(self, filters_dict):
        """
        Stores the final filter settings applied in Stage 4.

        Args:
            filters_dict (dict): The dictionary representing active filters
                                (e.g., self.active_filters from main_gui).
        """
        # Store a copy to prevent issues if the original dict is modified later
        self.filter_settings = copy.deepcopy(filters_dict)
        print(f"Action Logger: Stored filter settings: {self.filter_settings}")

    def get_filter_settings(self):
        """Returns the stored filter settings dictionary."""
        return self.filter_settings

    def get_preset_data(self):
        """
        Generates the complete data structure needed for saving a preset file.
        """
        # Convert any remaining sets in filter_settings to lists for JSON
        json_safe_filters = {}
        for key, value in self.filter_settings.items():
            if isinstance(value, set):
                json_safe_filters[key] = sorted(list(value)) # Sort for consistency
            else:
                json_safe_filters[key] = value

        preset_data = {
            "description": f"Preset generated by Excel Data Processor on {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "preset_format_version": "1.0", # Good practice
            "deletion_steps": self.deletion_steps, # Already list of dicts
            "filter_settings": json_safe_filters  # Use the JSON-safe version
        }
        return preset_data