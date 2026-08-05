from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from logging_config import setup_logger

from .config_loader import AppConfig


@dataclass(frozen=True)
class ProjectContext:
    config: AppConfig
    entry_name: str
    run_id: str
    log_path: Path


def bootstrap_context(
    entry_file: str | Path,
    explicit_data_root: Path | None = None,
) -> ProjectContext:
    entry_path = Path(entry_file).resolve()
    repo_root = entry_path.parent
    config = AppConfig.resolve(repo_root, explicit_data_root)
    log_path = setup_logger(
        log_level=config.log_level,
        log_dir=config.log_dir,
        entry_name=entry_path.stem,
    )
    return ProjectContext(
        config=config,
        entry_name=entry_path.stem,
        run_id=datetime.now().strftime("%Y%m%d_%H%M%S"),
        log_path=log_path,
    )
