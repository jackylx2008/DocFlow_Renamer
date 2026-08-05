"""项目统一日志配置。"""

from __future__ import annotations

import logging
import re
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path


LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"
DATE_FORMAT = "%H:%M:%S"


def setup_logger(
    *,
    log_level: str = "INFO",
    log_dir: Path | None = None,
    entry_name: str | None = None,
) -> Path:
    root_dir = Path(__file__).resolve().parent
    target_dir = (log_dir or root_dir / "logs").resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    raw_name = entry_name or Path(sys.argv[0]).stem or "application"
    safe_name = re.sub(r'[\\/:*?"<>|]+', "_", raw_name)
    log_path = target_dir / f"{safe_name}.log"

    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    rotating_file = RotatingFileHandler(
        log_path,
        maxBytes=10 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    rotating_file.setFormatter(formatter)

    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    root_logger.addHandler(console)
    root_logger.addHandler(rotating_file)
    return log_path


def get_logger(name: str) -> logging.Logger:
    return logging.getLogger(name)
