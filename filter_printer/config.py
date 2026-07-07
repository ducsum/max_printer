import os
import json
import shutil
from typing import Any, Dict
import logging

from constants import DEFAULT_CONFIG

logger = logging.getLogger(__name__)

class AppConfig:
    def __init__(self, config_file: str):
        self.config_file = config_file
        self.data: Dict[str, Any] = self._load()

    def _load(self) -> Dict[str, Any]:
        data = DEFAULT_CONFIG.copy()
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    loaded_data = json.load(f)
                    data.update(loaded_data)
            except Exception as e:
                logger.error(f"Lỗi đọc config.json: {e}. Tiến hành backup và dùng default.")
                try:
                    shutil.copy2(self.config_file, self.config_file + ".bak")
                except Exception as backup_e:
                    logger.error(f"Lỗi khi backup config: {backup_e}")
        return data

    def save(self):
        tmp_file = self.config_file + ".tmp"
        try:
            with open(tmp_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=4)
            os.replace(tmp_file, self.config_file)
        except Exception as e:
            logger.error(f"Lỗi khi lưu config: {e}")
            if os.path.exists(tmp_file):
                try:
                    os.remove(tmp_file)
                except Exception:
                    pass

    def get(self, key: str, default: Any = None) -> Any:
        return self.data.get(key, default)

    def set(self, key: str, value: Any):
        self.data[key] = value
        self.save()
