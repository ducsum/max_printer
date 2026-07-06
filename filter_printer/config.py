import os
import json
import shutil
from typing import Any, Dict

class AppConfig:
    def __init__(self, config_file: str):
        self.config_file = config_file
        self.data: Dict[str, Any] = self._load()

    def _load(self) -> Dict[str, Any]:
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                try:
                    # Backup hỏng
                    shutil.copy2(self.config_file, self.config_file + ".bak")
                except Exception:
                    pass
        return {}

    def save(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=4)
        except Exception:
            pass

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def set(self, key: str, value: Any):
        self.data[key] = value
        self.save()
