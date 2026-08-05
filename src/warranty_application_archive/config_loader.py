from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


ENV_PATTERN = re.compile(r"\$\{([A-Z0-9_]+)(?::-([^}]*))?\}")


def _read_env_file(path: Path) -> dict[str, str]:
    if not path.is_file():
        return {}
    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def _environment(repo_root: Path) -> dict[str, str]:
    values = dict(os.environ)
    for key, value in _read_env_file(repo_root / "common.env").items():
        values.setdefault(key, value)
    if not values.get("CLOUDSTATION_ROOT"):
        platform_key = {
            "win32": "CLOUDSTATION_ROOT_WINDOWS",
            "darwin": "CLOUDSTATION_ROOT_MACOS",
        }.get(sys.platform, "CLOUDSTATION_ROOT_LINUX")
        if values.get(platform_key):
            values["CLOUDSTATION_ROOT"] = values[platform_key]
    return values


def _expand(value: Any, environment: dict[str, str]) -> Any:
    if isinstance(value, dict):
        return {key: _expand(item, environment) for key, item in value.items()}
    if isinstance(value, list):
        return [_expand(item, environment) for item in value]
    if not isinstance(value, str):
        return value

    def replace(match: re.Match[str]) -> str:
        name, default = match.groups()
        return environment.get(name, default or "")

    expanded = value
    for _ in range(5):
        next_value = ENV_PATTERN.sub(replace, expanded)
        if next_value == expanded:
            break
        expanded = next_value
    return expanded


def load_config(repo_root: Path) -> dict[str, Any]:
    config_path = repo_root / "config.yaml"
    raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    if not isinstance(raw, dict):
        raise ValueError(f"配置文件格式不正确: {config_path}")
    return _expand(raw, _environment(repo_root))


def _resolve_path(repo_root: Path, value: str, fallback: str) -> Path:
    path = Path(value or fallback).expanduser()
    if not path.is_absolute():
        path = repo_root / path
    return path.resolve()


@dataclass(frozen=True)
class AppConfig:
    repo_root: Path
    data_root: Path
    output_dir: Path
    log_dir: Path
    log_level: str
    values: dict[str, Any]

    @classmethod
    def resolve(
        cls, repo_root: Path, explicit_data_root: Path | None = None
    ) -> "AppConfig":
        resolved_root = repo_root.resolve()
        values = load_config(resolved_root)
        app = values.get("app") or {}
        if not isinstance(app, dict):
            raise ValueError("config.yaml 的 app 节点必须是对象")
        data_root = (
            explicit_data_root.expanduser().resolve()
            if explicit_data_root
            else _resolve_path(
                resolved_root,
                str(app.get("input_path") or ""),
                "input",
            )
        )
        return cls(
            repo_root=resolved_root,
            data_root=data_root,
            output_dir=_resolve_path(
                resolved_root,
                str(app.get("output_dir") or ""),
                "output",
            ),
            log_dir=_resolve_path(
                resolved_root,
                str(app.get("log_dir") or ""),
                "logs",
            ),
            log_level=str(app.get("log_level") or "INFO").upper(),
            values=values,
        )
