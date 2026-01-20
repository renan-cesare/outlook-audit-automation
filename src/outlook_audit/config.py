import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict


@dataclass(frozen=True)
class AppConfig:
    raw: Dict[str, Any]


def load_config(path: Path) -> AppConfig:
    data = json.loads(path.read_text(encoding="utf-8"))
    return AppConfig(raw=data)


def get(cfg: AppConfig, *keys: str, default=None):
    cur = cfg.raw
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur
