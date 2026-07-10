from __future__ import annotations

import argparse
from datetime import date, datetime
import logging
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from docflow_renamer import setup_logging


APPLICATION_DOCX_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2})_(.+)_质保作业申请单\.docx$", re.IGNORECASE
)
WORKER_LIST_NAME_RE = re.compile(r"(?:人员|工人)名单")
TARGET_WORKER_LIST_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}_.+_质保作业申请单_工人名单\.(?:jpe?g|png)$",
    re.IGNORECASE,
)
IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png"}
WORKER_LIST_SUFFIX = "_工人名单"
LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class PlannedCopy:
    source: Path
    target: Path
    docx: Path
    modified_date: date


def load_env_file(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}

    values: dict[str, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def load_yaml_config(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError(f"配置文件格式不正确: {path}")
    return data


def resolve_setting(
    raw_value: str | None, env_values: dict[str, str], fallback: str
) -> str:
    if not raw_value:
        return fallback
    match = re.fullmatch(r"\$\{([A-Z0-9_]+):-([^}]*)\}", str(raw_value).strip())
    if not match:
        return str(raw_value)
    env_key, default_value = match.groups()
    return env_values.get(env_key, default_value)


def resolve_input_dir(repo_root: Path) -> Path:
    env_values = load_env_file(repo_root / "common.env")
    yaml_config = load_yaml_config(repo_root / "config.yaml")
    raw_input = env_values.get("INPUT_PATH") or resolve_setting(
        yaml_config.get("input_path"), env_values, "input"
    )
    input_dir = Path(raw_input)
    if not input_dir.is_absolute():
        input_dir = repo_root / input_dir
    return input_dir


def file_modified_date(path: Path) -> date:
    return datetime.fromtimestamp(path.stat().st_mtime).date()


def file_sort_key(path: Path) -> tuple[float, str]:
    return (path.stat().st_mtime, path.name.lower())


def collect_source_images(input_dir: Path) -> dict[date, list[Path]]:
    sources_by_date: dict[date, list[Path]] = {}
    for path in sorted(input_dir.iterdir(), key=lambda item: item.name.lower()):
        if not path.is_file() or path.suffix.lower() not in IMAGE_SUFFIXES:
            continue
        if not WORKER_LIST_NAME_RE.search(path.stem):
            continue
        if TARGET_WORKER_LIST_RE.fullmatch(path.name):
            LOGGER.info("跳过已符合目标命名规则的工人名单图片: %s", path.name)
            continue
        sources_by_date.setdefault(file_modified_date(path), []).append(path)
    for source_paths in sources_by_date.values():
        source_paths.sort(key=file_sort_key)
    return sources_by_date


def collect_application_docx(input_dir: Path) -> dict[date, list[Path]]:
    docx_by_date: dict[date, list[Path]] = {}
    for path in sorted(input_dir.glob("*.docx")):
        match = APPLICATION_DOCX_RE.fullmatch(path.name)
        if not match or not path.is_file():
            continue
        docx_by_date.setdefault(file_modified_date(path), []).append(path)
    for docx_paths in docx_by_date.values():
        docx_paths.sort(key=file_sort_key)
    return docx_by_date


def build_target_path(source: Path, docx_path: Path) -> Path:
    return docx_path.with_name(f"{docx_path.stem}{WORKER_LIST_SUFFIX}{source.suffix.lower()}")


def build_planned_copies(input_dir: Path) -> list[PlannedCopy]:
    sources_by_date = collect_source_images(input_dir)
    docx_by_date = collect_application_docx(input_dir)
    plans: list[PlannedCopy] = []

    for modified_date, source_files in sorted(sources_by_date.items()):
        docx_files = docx_by_date.get(modified_date, [])
        if not docx_files:
            LOGGER.warning(
                "未找到修改日期同一天的申请单，无法匹配: %s | 修改日期: %s",
                "；".join(source.name for source in source_files),
                modified_date.isoformat(),
            )
            continue
        if len(source_files) > 1:
            LOGGER.warning(
                "修改日期 %s 存在多张名单图片，仅使用第一张复制给同日期申请单: %s",
                modified_date.isoformat(),
                source_files[0].name,
            )
            LOGGER.warning(
                "跳过同日期其余名单图片: %s",
                "；".join(source.name for source in source_files[1:]),
            )
        source = source_files[0]
        for docx_path in docx_files:
            target = build_target_path(source, docx_path)
            plans.append(
                PlannedCopy(
                    source=source,
                    target=target,
                    docx=docx_path,
                    modified_date=modified_date,
                )
            )
    return plans


def copy_worker_list_images(plans: list[PlannedCopy]) -> tuple[int, int]:
    if not plans:
        LOGGER.info("未匹配到需要复制生成的人员名单图片")
        return 0, 0

    LOGGER.info("匹配到 %s 条复制生成关系，开始处理", len(plans))
    copied_count = 0
    skipped_count = 0
    for index, plan in enumerate(plans, start=1):
        if plan.source.resolve() == plan.target.resolve():
            skipped_count += 1
            LOGGER.info(
                "[%s/%s] 跳过，已是目标文件名: %s | 对应申请单: %s",
                index,
                len(plans),
                plan.source.name,
                plan.docx.name,
            )
            continue
        if plan.target.exists():
            skipped_count += 1
            LOGGER.info(
                "[%s/%s] 跳过，目标已存在: %s -> %s | 对应申请单: %s",
                index,
                len(plans),
                plan.source.name,
                plan.target.name,
                plan.docx.name,
            )
            continue

        shutil.copy2(plan.source, plan.target)
        copied_count += 1
        LOGGER.info(
            "[%s/%s] 已复制生成: %s -> %s | 修改日期: %s | 对应申请单: %s",
            index,
            len(plans),
            plan.source.name,
            plan.target.name,
            plan.modified_date.isoformat(),
            plan.docx.name,
        )
    LOGGER.info("复制生成完成: 新生成 %s 个，跳过 %s 个", copied_count, skipped_count)
    delete_completed_source_images(plans)
    return copied_count, skipped_count


def delete_completed_source_images(plans: list[PlannedCopy]) -> int:
    targets_by_source: dict[Path, list[Path]] = {}
    for plan in plans:
        targets_by_source.setdefault(plan.source, []).append(plan.target)

    deleted_count = 0
    for source, targets in targets_by_source.items():
        if not source.exists():
            continue
        source_path = source.resolve()
        target_paths = {target.resolve() for target in targets}
        if source_path in target_paths:
            LOGGER.info("保留原始名单图片，源文件也是目标文件: %s", source.name)
            continue
        missing_targets = [target for target in targets if not target.exists()]
        if missing_targets:
            LOGGER.warning(
                "保留原始名单图片，仍有目标未生成: %s | 缺失目标: %s",
                source.name,
                "；".join(target.name for target in missing_targets),
            )
            continue

        source.unlink()
        deleted_count += 1
        LOGGER.info("已删除原始名单图片: %s", source.name)

    LOGGER.info("原始名单图片清理完成: 删除 %s 个", deleted_count)
    return deleted_count


def main() -> int:
    repo_root = Path(__file__).resolve().parent
    log_path = setup_logging(repo_root)
    parser = argparse.ArgumentParser(
        description="将包含人员名单或工人名单的图片按修改日期复制生成为对应申请单工人名单图片"
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        help="输入目录，未指定时优先读取 common.env 中的 INPUT_PATH",
    )
    args = parser.parse_args()

    input_dir = args.input_dir.resolve() if args.input_dir else resolve_input_dir(repo_root)
    if not input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在: {input_dir}")

    LOGGER.info("终端输出日志: %s", log_path)
    LOGGER.info("输入目录: %s", input_dir)
    LOGGER.info("扫描源图片规则: 文件名包含“人员名单”或“工人名单”的 jpg/jpeg/png 图片")
    LOGGER.info("匹配规则: 图片修改日期与质保作业申请单修改日期为同一天")
    LOGGER.info("同一天一张名单图片可复制生成多个申请单对应的工人名单图片")
    LOGGER.info("目标命名规则: YYYY-MM-DD_施工内容_质保作业申请单_工人名单.源扩展名")
    plans = build_planned_copies(input_dir)
    copy_worker_list_images(plans)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
