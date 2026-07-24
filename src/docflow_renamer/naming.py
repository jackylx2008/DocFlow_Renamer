from __future__ import annotations

import re
from pathlib import Path

from .constants import (
    APPLICATION_SUFFIX,
    CONFINED_SPACE_ROLE,
    HIGH_ALTITUDE_ROLE,
    IMAGE_SUFFIXES,
    SIGNED_APPLICATION_ROLE,
    SPECIAL_WORK_ROLE,
    WORKER_LIST_ROLE,
)


APPLICATION_STEM_RE = re.compile(
    rf"^(?P<date>\d{{4}}-\d{{2}}-\d{{2}})_(?P<content>.+){re.escape(APPLICATION_SUFFIX)}$"
)
APPROVAL_PDF_RE = re.compile(
    r"^工程类-主体质保施工_编号：(?P<number>\d{12})\.pdf$", re.IGNORECASE
)


def parse_application_stem(stem: str) -> tuple[str, str] | None:
    match = APPLICATION_STEM_RE.fullmatch(stem)
    if not match:
        return None
    return match.group("date"), match.group("content")


def case_directory_name(word_path: Path) -> str:
    if not parse_application_stem(word_path.stem):
        raise ValueError(f"申请单文件名不符合规范: {word_path.name}")
    return word_path.stem


def application_prefix(word_path: Path) -> str:
    parsed = parse_application_stem(word_path.stem)
    if not parsed:
        raise ValueError(f"申请单文件名不符合规范: {word_path.name}")
    application_date, content = parsed
    return f"{application_date}_{content}"


def material_role(path: Path, word_stem: str) -> str | None:
    if not path.stem.startswith(word_stem):
        return None
    suffix = path.stem[len(word_stem) :].lstrip("_")
    if not suffix and path.suffix.lower() in IMAGE_SUFFIXES:
        return SIGNED_APPLICATION_ROLE
    if "工人名单" in suffix or "人员名单" in suffix:
        return WORKER_LIST_ROLE
    if "有限空间" in suffix:
        return CONFINED_SPACE_ROLE
    if "高处" in suffix or "高空" in suffix:
        return HIGH_ALTITUDE_ROLE
    if suffix:
        return SPECIAL_WORK_ROLE
    return None


def material_file_name(
    word_stem: str, role: str, index: int, suffix: str
) -> str:
    labels = {
        SIGNED_APPLICATION_ROLE: "手签",
        WORKER_LIST_ROLE: "施工人员名单",
        CONFINED_SPACE_ROLE: "有限空间申请",
        HIGH_ALTITUDE_ROLE: "高处作业申请",
        SPECIAL_WORK_ROLE: "专项作业材料",
    }
    label = labels[role]
    return f"{word_stem}_{label}_{index:02d}{suffix.lower()}"


def approval_application_no(file_name: str) -> str:
    match = APPROVAL_PDF_RE.fullmatch(file_name)
    return match.group("number") if match else ""
