from __future__ import annotations

import argparse
import logging
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


DATE_WORKER_LIST_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2})_(?:人员|工人)名单\.jpe?g$", re.IGNORECASE
)
APPLICATION_DOCX_RE = re.compile(
    r"^(\d{4}-\d{2}-\d{2})_(.+)_质保作业申请单\.docx$", re.IGNORECASE
)
WORKER_LIST_SUFFIX = "_工人名单.jpg"
LOGGER = logging.getLogger(__name__)


@dataclass(frozen=True)
class PlannedCopy:
    source: Path
    target: Path
    docx: Path


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


def collect_source_images(input_dir: Path) -> dict[str, Path]:
    sources: dict[str, Path] = {}
    for path in sorted(input_dir.glob("*.jp*g")):
        match = DATE_WORKER_LIST_RE.fullmatch(path.name)
        if not match or not path.is_file():
            continue
        work_date = match.group(1)
        if work_date in sources:
            LOGGER.warning("同日期存在多个人员名单图片，仅使用第一个: %s", sources[work_date].name)
            LOGGER.warning("跳过重复人员名单图片: %s", path.name)
            continue
        sources[work_date] = path
    return sources


def collect_application_docx(input_dir: Path) -> dict[str, list[Path]]:
    docx_by_date: dict[str, list[Path]] = {}
    for path in sorted(input_dir.glob("*.docx")):
        match = APPLICATION_DOCX_RE.fullmatch(path.name)
        if not match or not path.is_file():
            continue
        docx_by_date.setdefault(match.group(1), []).append(path)
    return docx_by_date


def build_planned_copies(input_dir: Path) -> list[PlannedCopy]:
    sources = collect_source_images(input_dir)
    docx_by_date = collect_application_docx(input_dir)
    plans: list[PlannedCopy] = []

    for work_date, source in sorted(sources.items()):
        docx_files = docx_by_date.get(work_date, [])
        if not docx_files:
            LOGGER.warning("未找到同日期申请单，无法匹配: %s", source.name)
            continue
        for docx_path in docx_files:
            target = docx_path.with_name(f"{docx_path.stem}{WORKER_LIST_SUFFIX}")
            plans.append(PlannedCopy(source=source, target=target, docx=docx_path))
    return plans


def copy_worker_list_images(plans: list[PlannedCopy]) -> tuple[int, int]:
    if not plans:
        LOGGER.info("未匹配到需要复制重命名的人员名单图片")
        return 0, 0

    LOGGER.info("匹配到 %s 条复制重命名关系，开始复制", len(plans))
    copied_count = 0
    skipped_count = 0
    for index, plan in enumerate(plans, start=1):
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
            "[%s/%s] 已复制: %s -> %s | 对应申请单: %s",
            index,
            len(plans),
            plan.source.name,
            plan.target.name,
            plan.docx.name,
        )
    LOGGER.info("复制完成: 新复制 %s 个，跳过 %s 个", copied_count, skipped_count)
    return copied_count, skipped_count


def main() -> int:
    repo_root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="将日期人员名单图片复制重命名为同日期申请单工人名单图片"
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        help="输入目录，未指定时优先读取 common.env 中的 INPUT_PATH",
    )
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    input_dir = args.input_dir.resolve() if args.input_dir else resolve_input_dir(repo_root)
    if not input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在: {input_dir}")

    LOGGER.info("输入目录: %s", input_dir)
    LOGGER.info("扫描源图片规则: YYYY-MM-DD_人员名单.jpg 或 YYYY-MM-DD_工人名单.jpg")
    LOGGER.info("目标命名规则: YYYY-MM-DD_施工内容_质保作业申请单_工人名单.jpg")
    plans = build_planned_copies(input_dir)
    copy_worker_list_images(plans)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
