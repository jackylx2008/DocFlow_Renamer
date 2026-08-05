from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from .constants import DATA_FILE_NAME, SCHEMA_VERSION
from .file_utils import atomic_replace_text


def empty_dataset(root: Path) -> dict[str, Any]:
    now = datetime.now().astimezone().isoformat(timespec="seconds")
    return {
        "schema_version": SCHEMA_VERSION,
        "dataset_revision": 0,
        "data_root": str(root.resolve()),
        "created_at": now,
        "updated_at": now,
        "applications": [],
        "unmatched_files": [],
        "runs": [],
        "changes": [],
    }


class JsonRepository:
    def __init__(self, root: Path) -> None:
        self.root = root.resolve()
        self.path = self.root / DATA_FILE_NAME

    def load(self) -> dict[str, Any]:
        if not self.path.is_file():
            return empty_dataset(self.root)
        data = json.loads(self.path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError(f"JSON 根节点必须是对象: {self.path}")
        if data.get("schema_version") != SCHEMA_VERSION:
            raise ValueError(
                f"不支持的数据版本: {data.get('schema_version')}，当前版本: {SCHEMA_VERSION}"
            )
        return data

    def save(self, data: dict[str, Any]) -> Path:
        data["schema_version"] = SCHEMA_VERSION
        data["data_root"] = str(self.root)
        data["updated_at"] = datetime.now().astimezone().isoformat(
            timespec="seconds"
        )
        atomic_replace_text(
            self.path,
            json.dumps(data, ensure_ascii=False, indent=2, sort_keys=False),
        )
        return self.path
