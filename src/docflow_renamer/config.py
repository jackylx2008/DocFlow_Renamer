from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from . import legacy


@dataclass(frozen=True)
class AppConfig:
    repo_root: Path
    data_root: Path

    @classmethod
    def resolve(
        cls, repo_root: Path, explicit_data_root: Path | None = None
    ) -> "AppConfig":
        data_root = (
            explicit_data_root.resolve()
            if explicit_data_root
            else legacy.resolve_input_dir(repo_root).resolve()
        )
        return cls(repo_root=repo_root.resolve(), data_root=data_root)
