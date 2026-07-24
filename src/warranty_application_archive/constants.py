from __future__ import annotations

from pathlib import Path


SCHEMA_VERSION = 1
DATA_FILE_NAME = "质保作业申请数据.json"
EXCEL_FILE_NAME = "质保作业申请汇总.xlsx"
APPROVAL_REVIEW_DATA_FILE_NAME = "审批PDF匹配审核.json"
APPROVAL_REVIEW_EXCEL_FILE_NAME = "审批PDF匹配审核.xlsx"
APPROVAL_REVIEW_SCHEMA_VERSION = 1
TEMPLATE_FILE_NAME = "01 安全生产及消防安全协议（建工）.pdf"

INBOX_DIR_NAME = "_inbox"
TEMPLATES_DIR_NAME = "_templates"
CASES_DIR_NAME = "_cases"
INTERNAL_DIR_NAME = ".docflow"
QUARANTINE_DIR_NAME = "quarantine"
LEGACY_DIR_NAME = "legacy"

APPLICATION_SUFFIX = "_质保作业申请单"
SIGNED_APPLICATION_ROLE = "signed_application"
WORKER_LIST_ROLE = "worker_list"
SAFETY_AGREEMENT_ROLE = "safety_agreement"
CONFINED_SPACE_ROLE = "confined_space"
HIGH_ALTITUDE_ROLE = "high_altitude"
SPECIAL_WORK_ROLE = "special_work"
APPROVAL_PDF_ROLE = "approval_pdf"
WORD_ROLE = "word"

IMAGE_SUFFIXES = {".jpg", ".jpeg", ".png"}
PDF_SUFFIX = ".pdf"
WORD_SUFFIX = ".docx"


def inbox_dir(root: Path) -> Path:
    return root / INBOX_DIR_NAME


def templates_dir(root: Path) -> Path:
    return root / TEMPLATES_DIR_NAME


def cases_dir(root: Path) -> Path:
    return root / CASES_DIR_NAME


def internal_dir(root: Path) -> Path:
    return root / INTERNAL_DIR_NAME
