import os
import json
import logging

logger = logging.getLogger(__name__)


class StateManager:
    """
    Manages persistent application state in JSON format so user edits
    (custom loads, edited agency info, added vehicles, custom invoice numbers)
    never revert to default Excel data on app restart or file reload.
    """

    def __init__(self, state_file_path):
        self.state_file_path = os.path.abspath(state_file_path)
        self._ensure_dir()

    def _ensure_dir(self):
        folder = os.path.dirname(self.state_file_path)
        if folder and not os.path.exists(folder):
            os.makedirs(folder, exist_ok=True)

    def load_state(self) -> dict:
        """Loads and returns saved state dict, or empty dict if not found/corrupted."""
        if not os.path.exists(self.state_file_path):
            return {}
        try:
            with open(self.state_file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
            return {}
        except Exception as e:
            logger.warning(f"Failed to load app state from {self.state_file_path}: {e}")
            return {}

    def save_state(self, state_dict: dict) -> bool:
        """Atomically saves the state dictionary to disk."""
        if not isinstance(state_dict, dict):
            return False
        try:
            self._ensure_dir()
            tmp_path = f"{self.state_file_path}.tmp"
            with open(tmp_path, "w", encoding="utf-8") as f:
                json.dump(state_dict, f, indent=2, ensure_ascii=False)
            os.replace(tmp_path, self.state_file_path)
            return True
        except Exception as e:
            logger.error(f"Failed to save app state to {self.state_file_path}: {e}")
            if os.path.exists(f"{self.state_file_path}.tmp"):
                try:
                    os.remove(f"{self.state_file_path}.tmp")
                except Exception:
                    pass
            return False

    def clear_state(self) -> bool:
        """Removes the saved state file (e.g. for reset to Excel defaults)."""
        try:
            if os.path.exists(self.state_file_path):
                os.remove(self.state_file_path)
            return True
        except Exception as e:
            logger.error(f"Failed to clear app state: {e}")
            return False
