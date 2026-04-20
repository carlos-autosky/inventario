
from __future__ import annotations
import json
from pathlib import Path
from .config import CONFIG_FILENAME


def get_config_path() -> Path:
    return Path.home() / CONFIG_FILENAME


def load_config() -> dict:
    path = get_config_path()
    if not path.exists():
        return {
            "warehouse_mode": "Todas",
            "selected_warehouses": [],
            "excluded_skus": "",
            "last_cutoff": "",
        }
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {
            "warehouse_mode": "Todas",
            "selected_warehouses": [],
            "excluded_skus": "",
            "last_cutoff": "",
        }


def save_config(data: dict) -> None:
    path = get_config_path()
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
