from __future__ import annotations

import logging
import shutil
import uuid
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from . import legacy
from .constants import (
    APPROVAL_PDF_ROLE,
    CASES_DIR_NAME,
    CONFINED_SPACE_ROLE,
    DATA_FILE_NAME,
    EXCEL_FILE_NAME,
    HIGH_ALTITUDE_ROLE,
    IMAGE_SUFFIXES,
    INBOX_DIR_NAME,
    SAFETY_AGREEMENT_ROLE,
    SIGNED_APPLICATION_ROLE,
    SPECIAL_WORK_ROLE,
    TEMPLATE_FILE_NAME,
    TEMPLATES_DIR_NAME,
    WORD_ROLE,
    WORKER_LIST_ROLE,
)
from .file_utils import ensure_within, relative_posix, sha256_file
from .naming import (
    application_prefix,
    approval_application_no,
    case_directory_name,
    material_file_name,
    material_role,
)
from .repository import empty_dataset


LOGGER = logging.getLogger(__name__)
CASE_NAMESPACE = uuid.UUID("76cf298a-84bf-45df-992c-34749139b7c4")
MIGRATABLE_SUFFIXES = {".docx", ".pdf", *IMAGE_SUFFIXES}


@dataclass(frozen=True)
class FileOperation:
    action: str
    source: str
    target: str
    sha256: str
    case_id: str = ""
    role: str = ""


@dataclass
class MigrationPlan:
    data_root: str
    created_at: str
    applications: list[dict[str, Any]]
    unmatched_files: list[dict[str, Any]]
    operations: list[FileOperation]
    warnings: list[str]

    def to_dict(self) -> dict[str, Any]:
        return {
            "data_root": self.data_root,
            "created_at": self.created_at,
            "applications": self.applications,
            "unmatched_files": self.unmatched_files,
            "operations": [asdict(operation) for operation in self.operations],
            "warnings": self.warnings,
        }


def file_record(
    source: Path,
    target: Path,
    root: Path,
    role: str,
    fingerprint: str | None = None,
    derived: bool = False,
) -> dict[str, Any]:
    resolved_fingerprint = fingerprint or sha256_file(source)
    relative_target = relative_posix(target, root)
    return {
        "file_id": str(
            uuid.uuid5(
                CASE_NAMESPACE,
                f"{resolved_fingerprint}:{relative_target}",
            )
        ),
        "role": role,
        "path": relative_target,
        "original_name": source.name,
        "current_name": target.name,
        "sha256": resolved_fingerprint,
        "size": source.stat().st_size,
        "derived": derived,
        "exists": True,
    }


def _legacy_pdf_matches(
    root: Path, parsed_cases: list[tuple[Path, dict[str, Any]]]
) -> dict[str, list[str]]:
    excel_path = root / EXCEL_FILE_NAME
    if not excel_path.is_file():
        return {}
    cache = legacy.load_existing_pdf_match_cache(excel_path)
    result: dict[str, list[str]] = {}
    for word_path, parsed in parsed_cases:
        key = legacy.pdf_match_cache_key(
            parsed["施工区域"],
            parsed["施工内容"],
            parsed["施工开始时间"],
            parsed["施工结束时间"],
        )
        names = legacy.split_pdf_names(cache.get(key, ""))
        if names:
            result[word_path.name] = names
    return result


def _required_roles(parsed: dict[str, Any]) -> list[str]:
    required = [
        WORD_ROLE,
        SIGNED_APPLICATION_ROLE,
        WORKER_LIST_ROLE,
        SAFETY_AGREEMENT_ROLE,
    ]
    dangerous_work = str(parsed.get("危险作业") or "")
    if "有限空间" in dangerous_work:
        required.append(CONFINED_SPACE_ROLE)
    if "高处" in dangerous_work or "高空" in dangerous_work:
        required.append(HIGH_ALTITUDE_ROLE)
    return required


def _case_status(
    approval_files: list[dict[str, Any]],
    missing_roles: list[str],
) -> str:
    if approval_files:
        return "approved"
    if missing_roles:
        return "materials_incomplete"
    return "materials_ready"


def build_migration_plan(root: Path) -> MigrationPlan:
    root = root.resolve()
    if (root / DATA_FILE_NAME).exists():
        raise FileExistsError(f"数据 JSON 已存在，不能重复执行旧目录迁移: {root / DATA_FILE_NAME}")

    word_files = sorted(
        path
        for path in root.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".docx"
        and not path.name.startswith("~$")
        and legacy.TARGET_NAME_RE.fullmatch(path.name)
    )
    if not word_files:
        raise FileNotFoundError(f"没有找到可迁移的规范命名 Word 申请单: {root}")

    template_source = root / TEMPLATE_FILE_NAME
    if not template_source.is_file():
        raise FileNotFoundError(f"缺少安全协议模板: {template_source}")
    template_target = root / TEMPLATES_DIR_NAME / TEMPLATE_FILE_NAME
    template_hash = sha256_file(template_source)

    parsed_cases = [(word_path, legacy.parse_document(word_path)) for word_path in word_files]
    pdf_matches = _legacy_pdf_matches(root, parsed_cases)
    operations: list[FileOperation] = []
    legacy_excel = root / EXCEL_FILE_NAME
    if legacy_excel.is_file():
        operations.append(
            FileOperation(
                action="move",
                source=str(legacy_excel),
                target=str(
                    root
                    / ".docflow"
                    / "legacy"
                    / "质保作业申请汇总_legacy.xlsx"
                ),
                sha256=sha256_file(legacy_excel),
                role="legacy_excel",
            )
        )
    operations.append(
        FileOperation(
            action="move",
            source=str(template_source),
            target=str(template_target),
            sha256=template_hash,
            role="template",
        )
    )
    applications: list[dict[str, Any]] = []
    claimed_sources: set[Path] = {template_source.resolve()}
    warnings: list[str] = []

    for word_path, parsed in parsed_cases:
        case_name = case_directory_name(word_path)
        case_id = str(uuid.uuid5(CASE_NAMESPACE, case_name))
        case_dir = root / CASES_DIR_NAME / case_name
        word_target = case_dir / word_path.name
        word_hash = sha256_file(word_path)
        operations.append(
            FileOperation(
                "move",
                str(word_path),
                str(word_target),
                word_hash,
                case_id,
                WORD_ROLE,
            )
        )
        claimed_sources.add(word_path.resolve())

        materials: dict[str, list[dict[str, Any]]] = {
            WORD_ROLE: [
                file_record(
                    word_path,
                    word_target,
                    root,
                    WORD_ROLE,
                    fingerprint=word_hash,
                )
            ],
            SIGNED_APPLICATION_ROLE: [],
            WORKER_LIST_ROLE: [],
            SAFETY_AGREEMENT_ROLE: [],
            CONFINED_SPACE_ROLE: [],
            HIGH_ALTITUDE_ROLE: [],
            SPECIAL_WORK_ROLE: [],
        }

        related_files: dict[str, list[Path]] = {}
        for candidate in sorted(root.iterdir(), key=lambda item: item.name.lower()):
            if not candidate.is_file() or candidate.resolve() in claimed_sources:
                continue
            role = material_role(candidate, word_path.stem)
            if role:
                related_files.setdefault(role, []).append(candidate)

        for role, source_files in related_files.items():
            for index, source in enumerate(source_files, start=1):
                target_name = material_file_name(
                    word_path.stem,
                    role,
                    index,
                    source.suffix,
                )
                target = case_dir / target_name
                fingerprint = sha256_file(source)
                operations.append(
                    FileOperation(
                        "move",
                        str(source),
                        str(target),
                        fingerprint,
                        case_id,
                        role,
                    )
                )
                claimed_sources.add(source.resolve())
                materials[role].append(
                    file_record(
                        source,
                        target,
                        root,
                        role,
                        fingerprint=fingerprint,
                    )
                )

        agreement_target = (
            case_dir
            / f"{application_prefix(word_path)}_{TEMPLATE_FILE_NAME}"
        )
        operations.append(
            FileOperation(
                "copy",
                str(template_target),
                str(agreement_target),
                template_hash,
                case_id,
                SAFETY_AGREEMENT_ROLE,
            )
        )
        materials[SAFETY_AGREEMENT_ROLE].append(
            {
                "file_id": str(
                    uuid.uuid5(
                        CASE_NAMESPACE,
                        f"{case_id}:{SAFETY_AGREEMENT_ROLE}:{template_hash}",
                    )
                ),
                "role": SAFETY_AGREEMENT_ROLE,
                "path": relative_posix(agreement_target, root),
                "original_name": TEMPLATE_FILE_NAME,
                "current_name": agreement_target.name,
                "sha256": template_hash,
                "size": template_source.stat().st_size,
                "derived": True,
                "template_path": relative_posix(template_target, root),
                "exists": True,
            }
        )

        approval_files: list[dict[str, Any]] = []
        for pdf_name in pdf_matches.get(word_path.name, []):
            source = root / pdf_name
            if not source.is_file():
                warnings.append(f"旧 Excel 匹配的审批 PDF 不存在: {word_path.name} -> {pdf_name}")
                continue
            target = case_dir / source.name
            fingerprint = sha256_file(source)
            operations.append(
                FileOperation(
                    "move",
                    str(source),
                    str(target),
                    fingerprint,
                    case_id,
                    APPROVAL_PDF_ROLE,
                )
            )
            claimed_sources.add(source.resolve())
            approval_files.append(
                file_record(
                    source,
                    target,
                    root,
                    APPROVAL_PDF_ROLE,
                    fingerprint=fingerprint,
                )
            )

        required_roles = _required_roles(parsed)
        missing_roles = [
            role for role in required_roles if not materials.get(role)
        ]
        application_data = {
            "项目名称": parsed["项目名称"],
            "质保单位": parsed["质保单位"],
            "分包单位": parsed["分包单位"],
            "质保负责人": parsed["质保负责人"],
            "质保负责人联系电话": parsed["质保负责人联系电话"],
            "施工区域": parsed["施工区域"],
            "施工开始时间": parsed["施工开始时间"],
            "施工结束时间": parsed["施工结束时间"],
            "时长天": parsed["时长天"],
            "施工内容": parsed["施工内容"],
            "施工负责人": parsed["施工负责人"],
            "施工负责人联系电话": parsed["施工负责人联系电话"],
            "影响改动消防设备设施": parsed["影响改动消防设备设施"],
            "影响堵塞应急疏散通道": parsed["影响堵塞应急疏散通道"],
            "危险作业": parsed["危险作业"],
        }
        applications.append(
            {
                "case_id": case_id,
                "case_name": case_name,
                "case_directory": relative_posix(case_dir, root),
                "status": _case_status(approval_files, missing_roles),
                "application": application_data,
                "required_material_types": required_roles,
                "missing_material_types": missing_roles,
                "materials": materials,
                "approval": {
                    "status": "approved" if approval_files else "not_received",
                    "application_no": (
                        approval_application_no(approval_files[0]["current_name"])
                        if approval_files
                        else ""
                    ),
                    "pdfs": approval_files,
                    "match_source": "legacy_excel" if approval_files else "",
                },
                "history": [],
            }
        )

    unmatched_files: list[dict[str, Any]] = []
    inbox_dir = root / INBOX_DIR_NAME
    excluded_names = {
        EXCEL_FILE_NAME,
        DATA_FILE_NAME,
    }
    for source in sorted(root.iterdir(), key=lambda item: item.name.lower()):
        if (
            not source.is_file()
            or source.resolve() in claimed_sources
            or source.name in excluded_names
            or source.name.startswith("~$")
            or source.suffix.lower() not in MIGRATABLE_SUFFIXES
        ):
            continue
        target = inbox_dir / source.name
        fingerprint = sha256_file(source)
        role = (
            APPROVAL_PDF_ROLE
            if source.suffix.lower() == ".pdf"
            else "unclassified"
        )
        operations.append(
            FileOperation(
                "move",
                str(source),
                str(target),
                fingerprint,
                role=role,
            )
        )
        claimed_sources.add(source.resolve())
        unmatched_files.append(
            file_record(
                source,
                target,
                root,
                role,
                fingerprint=fingerprint,
            )
        )

    return MigrationPlan(
        data_root=str(root),
        created_at=datetime.now().astimezone().isoformat(timespec="seconds"),
        applications=applications,
        unmatched_files=unmatched_files,
        operations=operations,
        warnings=warnings,
    )


def _inventory_signature(root: Path) -> dict[str, tuple[int, str]]:
    return {
        path.relative_to(root).as_posix(): (path.stat().st_size, sha256_file(path))
        for path in root.rglob("*")
        if path.is_file() and not path.name.startswith("~$")
    }


def verify_backup(primary: Path, backup: Path) -> None:
    primary = primary.resolve()
    backup = backup.resolve()
    if not backup.is_dir():
        raise FileNotFoundError(f"备份目录不存在: {backup}")
    primary_inventory = _inventory_signature(primary)
    backup_inventory = _inventory_signature(backup)
    if primary_inventory != backup_inventory:
        primary_only = sorted(primary_inventory.keys() - backup_inventory.keys())
        backup_only = sorted(backup_inventory.keys() - primary_inventory.keys())
        changed = sorted(
            name
            for name in primary_inventory.keys() & backup_inventory.keys()
            if primary_inventory[name] != backup_inventory[name]
        )
        raise RuntimeError(
            "正式目录与备份不一致，拒绝迁移。"
            f" 正式目录独有={primary_only[:5]}，备份独有={backup_only[:5]}，"
            f"内容不同={changed[:5]}"
        )


def apply_migration_plan(plan: MigrationPlan) -> dict[str, Any]:
    root = Path(plan.data_root).resolve()
    completed: list[dict[str, Any]] = []
    for operation in plan.operations:
        source = ensure_within(Path(operation.source), root)
        target = ensure_within(Path(operation.target), root)
        target.parent.mkdir(parents=True, exist_ok=True)
        if target.exists():
            if sha256_file(target) == operation.sha256:
                completed.append({**asdict(operation), "result": "already_exists"})
                continue
            raise FileExistsError(f"迁移目标已存在且内容不同: {target}")
        if not source.is_file():
            raise FileNotFoundError(f"迁移源文件不存在: {source}")
        if sha256_file(source) != operation.sha256:
            raise RuntimeError(f"迁移前文件内容发生变化，已停止: {source}")
        if operation.action == "move":
            shutil.move(str(source), str(target))
        elif operation.action == "copy":
            shutil.copy2(source, target)
        else:
            raise ValueError(f"未知迁移动作: {operation.action}")
        if not target.is_file() or sha256_file(target) != operation.sha256:
            raise RuntimeError(f"迁移后校验失败: {target}")
        completed.append({**asdict(operation), "result": "completed"})

    dataset = empty_dataset(root)
    dataset["dataset_revision"] = 1
    dataset["applications"] = plan.applications
    dataset["unmatched_files"] = plan.unmatched_files
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    dataset["runs"].append(
        {
            "run_id": run_id,
            "type": "legacy_migration",
            "status": "completed",
            "started_at": plan.created_at,
            "completed_at": datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
            "application_count": len(plan.applications),
            "operation_count": len(completed),
            "warning_count": len(plan.warnings),
        }
    )
    dataset["changes"] = completed
    return dataset
