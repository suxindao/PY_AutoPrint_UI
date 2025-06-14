import json
import os
from pathlib import Path
from typing import Dict, Any
from .path_utils import get_app_path, ensure_directory_exists

DEFAULT_CONFIG = {
    "source_dir": "",
    "monthly_printer_name": "",
    "default_paper_size": 132,
    "default_paper_zoom": 75,
    "delay_seconds": 5,
    "enable_wait_prompt": True,
    "wait_prompt_sleep": 30,
    "bw_print": True,
    "duplex_print": False,
}


class ConfigManager:
    def __init__(self):
        self.config_dir = ensure_directory_exists(get_app_path("config"))
        self.config_path = os.path.join(self.config_dir, "settings.json")
        self.config = DEFAULT_CONFIG.copy()
        self.load_config()

    def load_config(self) -> bool:
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config.update(json.load(f))
                return True
        except Exception as e:
            print(f"加载配置文件失败: {e}")
        return False

    def save_config(self) -> bool:
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
        return False

    def get(self, key: str, default=None) -> Any:
        return self.config.get(key, default)

    def set(self, key: str, value: Any) -> None:
        self.config[key] = value

    def get_all(self) -> Dict[str, Any]:
        return self.config.copy()

    def update_config(self, new_config: Dict[str, Any]) -> None:
        self.config.update(new_config)
