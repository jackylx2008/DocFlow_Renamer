from __future__ import annotations

import logging
import re
import shutil
import uuid
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Callable

from ..modules import legacy
from ..modules.constants import (
    APPROVAL_PDF_ROLE,
    CASES_DIR_NAME,
    CONFINED_SPACE_ROLE,
    HIGH_ALTITUDE_ROLE,
    IMAGE_SUFFIXES,
    INPUT_DIR_NAME,
    INTERNAL_DIR_NAME,
    PDF_SUFFIX,
    QUARANTINE_DIR_NAME,
    SAFETY_AGREEMENT_ROLE,
    SIGNED_APPLICATION_ROLE,
    SPECIAL_WORK_ROLE,
    TEMPLATE_FILE_NAME,
    TEMPLATES_DIR_NAME,
    WORD_ROLE,
    WORKER_LIST_ROLE,
)
from ..modules.file_utils import ensure_within, relative_posix, sha256_file
from .migration_flow import CASE_NAMESPACE, _required_roles, file_record
from ..modules.naming import (
    application_prefix,
    approval_application_no,
    material_file_name,
    material_role,
)
from ..modules.recognition import RecognitionService


LOGGER = logging.getLogger(__name__)
INPUT_ROUTE_VERSION = "input-router-v1"
APPLICATION_MATERIAL_ROUTE = "application_material"
APPROVAL_PDF_ROUTE = "approval_pdf"
IMAGE_CONTENT_SIMILARITY_MIN = 0.60
IMAGE_CONTENT_SIMILARITY_MARGIN = 0.15


def _change(
    action: str,
    source: Path,
    target: Path,
    role: str,
    case_id: str = "",
    result: str = "completed",
) -> dict[str, Any]:
    return {
        "action": action,
        "source": str(source),
        "target": str(target),
        "role": role,
        "case_id": case_id,
        "result": result,
        "at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }


def _refresh_status(application: dict[str, Any]) -> None:
    if application.get("status") == "terminated":
        return
    materials = application.get("materials") or {}
    required = application.get("required_material_types") or []
    missing = [role for role in required if not materials.get(role)]
    application["missing_material_types"] = missing
    approval = application.get("approval") or {}
    if approval.get("pdfs"):
        application["status"] = "approved"
        approval["status"] = "approved"
    elif missing:
        application["status"] = "materials_incomplete"
    else:
        application["status"] = "materials_ready"


def _business_data(parsed: dict[str, Any]) -> dict[str, Any]:
    keys = [
        "项目名称",
        "质保单位",
        "分包单位",
        "质保负责人",
        "质保负责人联系电话",
        "施工区域",
        "施工开始时间",
        "施工结束时间",
        "时长天",
        "施工内容",
        "施工负责人",
        "施工负责人联系电话",
        "影响改动消防设备设施",
        "影响堵塞应急疏散通道",
        "危险作业",
    ]
    return {key: parsed.get(key, "") for key in keys}


def _application_duplicate_key(
    business: dict[str, Any],
) -> tuple[str, str, str, str] | None:
    start = legacy.normalize_date_for_pdf_match(
        str(business.get("施工开始时间") or "")
    )
    end = legacy.normalize_date_for_pdf_match(
        str(business.get("施工结束时间") or "")
    )
    content = legacy.normalize_match_text(
        str(business.get("施工内容") or "")
    )
    area = legacy.normalize_match_text(
        str(business.get("施工区域") or "")
    )
    if not all((start, end, content, area)):
        return None
    return start, end, content, area


def _find_duplicate_application(
    parsed: dict[str, Any],
    applications: list[dict[str, Any]],
    root: Path | None = None,
) -> dict[str, Any] | None:
    matches = _matching_duplicate_applications(parsed, applications)
    if not matches:
        return None
    selected = max(
        matches,
        key=lambda application: _application_completeness_score(
            application, root
        ),
    )
    if len(matches) > 1:
        LOGGER.warning(
            "发现 %d 条相同质保申请记录，按资料完整度保留: %s",
            len(matches),
            selected.get("case_name") or "",
        )
    return selected


def _application_completeness_score(
    application: dict[str, Any],
    root: Path | None = None,
) -> tuple[int, int, int, int, int, int, int]:
    materials = application.get("materials") or {}
    required = application.get("required_material_types") or []
    approval = application.get("approval") or {}
    valid_materials = {
        role: [
            file_item
            for file_item in (role_files or [])
            if _application_file_exists(application, file_item, root)
        ]
        for role, role_files in materials.items()
    }
    approval_files = [
        file_item
        for file_item in (approval.get("pdfs") or [])
        if _application_file_exists(application, file_item, root)
    ]
    required_present = sum(
        1 for role in required if valid_materials.get(role)
    )
    filled_roles = sum(
        1 for role_files in valid_materials.values() if role_files
    )
    file_keys = {
        str(file_item.get("sha256") or file_item.get("path") or "")
        for role_files in valid_materials.values()
        for file_item in (role_files or [])
        if isinstance(file_item, dict)
    }
    file_keys.update(
        str(file_item.get("sha256") or file_item.get("path") or "")
        for file_item in approval_files
        if isinstance(file_item, dict)
    )
    file_keys.discard("")
    missing_count = sum(
        1 for role in required if not valid_materials.get(role)
    )
    case_name = str(application.get("case_name") or "")
    _prefix, separator, suffix = case_name.rpartition("_")
    original_name = int(
        not (separator and len(suffix) == 2 and suffix.isdigit())
    )
    approved = int(
        bool(approval_files) or application.get("status") == "approved"
    )
    return (
        required_present,
        approved,
        len(approval_files),
        filled_roles,
        len(file_keys),
        -missing_count,
        original_name,
    )


def _application_file_exists(
    application: dict[str, Any],
    file_item: Any,
    root: Path | None,
) -> bool:
    if not isinstance(file_item, dict):
        return False
    if root is None:
        return True
    relative_path = str(file_item.get("path") or "")
    relative_case_dir = str(application.get("case_directory") or "")
    if not relative_path or not relative_case_dir:
        return False
    try:
        path = ensure_within(root / relative_path, root)
        case_dir = ensure_within(root / relative_case_dir, root)
        path.relative_to(case_dir)
    except (ValueError, OSError):
        return False
    return path.is_file()


def _matching_duplicate_applications(
    parsed: dict[str, Any],
    applications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    incoming_key = _application_duplicate_key(parsed)
    if incoming_key is None:
        return []
    return [
        application
        for application in applications
        if isinstance(application.get("application"), dict)
        and _application_duplicate_key(application["application"])
        == incoming_key
    ]


def _quarantine_less_complete_duplicate_cases(
    dataset: dict[str, Any],
    parsed: dict[str, Any],
    retained: dict[str, Any],
    root: Path,
    changes: list[dict[str, Any]],
) -> int:
    applications = dataset.setdefault("applications", [])
    duplicates = [
        application
        for application in _matching_duplicate_applications(
            parsed, applications
        )
        if application is not retained
    ]
    if not duplicates:
        return 0
    quarantine_dir = (
        root
        / INTERNAL_DIR_NAME
        / QUARANTINE_DIR_NAME
        / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        / "duplicate_cases"
    )
    removed_ids = {id(application) for application in duplicates}
    for application in duplicates:
        _merge_duplicate_case_files(
            application,
            retained,
            root,
            changes,
        )
        relative_case_dir = str(application.get("case_directory") or "")
        cases_root = root / CASES_DIR_NAME
        source = (
            ensure_within(root / relative_case_dir, cases_root)
            if relative_case_dir
            else cases_root / "_missing_case_directory"
        )
        target = quarantine_dir / (
            source.name
            if relative_case_dir
            else str(application.get("case_id") or "unknown_case")
        )
        result = "case_directory_missing"
        if relative_case_dir and source.is_dir():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(source), str(target))
            result = "completed"
        changes.append(
            _change(
                "quarantine_less_complete_duplicate_case",
                source,
                target,
                "application_case",
                str(application.get("case_id") or ""),
                result,
            )
        )
        LOGGER.warning(
            "重复案卷资料较少，已从汇总移除: 移除=%s；保留=%s；"
            "案卷目录处理=%s",
            application.get("case_name") or "",
            retained.get("case_name") or "",
            target if result == "completed" else "原案卷目录不存在",
        )
    applications[:] = [
        application
        for application in applications
        if id(application) not in removed_ids
    ]
    return len(duplicates)


def _merge_duplicate_case_files(
    source_application: dict[str, Any],
    retained: dict[str, Any],
    root: Path,
    changes: list[dict[str, Any]],
) -> None:
    retained_materials = retained.setdefault("materials", {})
    for role, retained_files in list(retained_materials.items()):
        retained_materials[role] = [
            file_item
            for file_item in (retained_files or [])
            if _application_file_exists(retained, file_item, root)
        ]
    retained_case_dir = ensure_within(
        root / str(retained["case_directory"]),
        root / CASES_DIR_NAME,
    )
    for role, role_files in (
        source_application.get("materials") or {}
    ).items():
        for file_item in role_files or []:
            if not _application_file_exists(
                source_application, file_item, root
            ):
                continue
            source = ensure_within(
                root / str(file_item["path"]),
                root / CASES_DIR_NAME,
            )
            fingerprint = sha256_file(source)
            if _application_contains_hash(retained, fingerprint):
                _quarantine_duplicate_application_sources(
                    [source],
                    root,
                    changes,
                    str(retained.get("case_id") or ""),
                )
                continue
            target = _unique_inbox_target(retained_case_dir, source)
            _move_verified(source, target, root)
            retained_materials.setdefault(role, []).append(
                file_record(
                    target,
                    target,
                    root,
                    role,
                    fingerprint=fingerprint,
                )
            )
            changes.append(
                _change(
                    "merge_duplicate_case_material",
                    source,
                    target,
                    role,
                    str(retained.get("case_id") or ""),
                )
            )
    retained_approval = retained.setdefault("approval", {})
    retained_approval_files = [
        file_item
        for file_item in (retained_approval.get("pdfs") or [])
        if _application_file_exists(retained, file_item, root)
    ]
    retained_approval["pdfs"] = retained_approval_files
    known_approval_hashes = {
        str(file_item.get("sha256") or "")
        for file_item in retained_approval_files
    }
    source_approval = source_application.get("approval") or {}
    for file_item in source_approval.get("pdfs") or []:
        if not _application_file_exists(
            source_application, file_item, root
        ):
            continue
        source = ensure_within(
            root / str(file_item["path"]),
            root / CASES_DIR_NAME,
        )
        fingerprint = sha256_file(source)
        if fingerprint in known_approval_hashes:
            _quarantine_duplicate_application_sources(
                [source],
                root,
                changes,
                str(retained.get("case_id") or ""),
            )
            continue
        target = _unique_inbox_target(retained_case_dir, source)
        _move_verified(source, target, root)
        retained_approval_files.append(
            file_record(
                target,
                target,
                root,
                APPROVAL_PDF_ROLE,
                fingerprint=fingerprint,
            )
        )
        known_approval_hashes.add(fingerprint)
        changes.append(
            _change(
                "merge_duplicate_case_approval",
                source,
                target,
                APPROVAL_PDF_ROLE,
                str(retained.get("case_id") or ""),
            )
        )
    if retained_approval_files:
        for field in ("application_no", "match_source"):
            if not retained_approval.get(field) and source_approval.get(
                field
            ):
                retained_approval[field] = source_approval[field]
        retained_approval["status"] = "approved"
    _refresh_status(retained)


def deduplicate_applications(
    dataset: dict[str, Any],
    root: Path,
    checkpoint: Callable[[dict[str, Any]], None] | None = None,
) -> int:
    root = root.resolve()
    applications = dataset.setdefault("applications", [])
    changes = dataset.setdefault("changes", [])
    duplicate_keys = {
        key
        for application in applications
        if isinstance(application.get("application"), dict)
        if (key := _application_duplicate_key(application["application"]))
        is not None
        if sum(
            1
            for candidate in applications
            if isinstance(candidate.get("application"), dict)
            and _application_duplicate_key(candidate["application"]) == key
        )
        > 1
    }
    removed = 0
    for duplicate_key in duplicate_keys:
        matches = [
            application
            for application in applications
            if isinstance(application.get("application"), dict)
            and _application_duplicate_key(application["application"])
            == duplicate_key
        ]
        if len(matches) < 2:
            continue
        retained = max(
            matches,
            key=lambda application: _application_completeness_score(
                application, root
            ),
        )
        business = retained.get("application") or {}
        LOGGER.warning(
            "发现 %d 条相同质保申请记录，按资料完整度保留: %s",
            len(matches),
            retained.get("case_name") or "",
        )
        removed += _quarantine_less_complete_duplicate_cases(
            dataset,
            business,
            retained,
            root,
            changes,
        )
        if checkpoint:
            checkpoint(dataset)
    return removed


def _next_case_name(
    base_name: str, applications: list[dict[str, Any]]
) -> str:
    existing = {
        str(application.get("case_name") or "")
        for application in applications
    }
    if base_name not in existing:
        return base_name
    index = 2
    while f"{base_name}_{index:02d}" in existing:
        index += 1
    return f"{base_name}_{index:02d}"


def _move_verified(source: Path, target: Path, root: Path) -> None:
    source = ensure_within(source, root)
    target = ensure_within(target, root)
    fingerprint = sha256_file(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        if sha256_file(target) == fingerprint:
            return
        raise FileExistsError(f"目标文件已存在且内容不同: {target}")
    shutil.move(str(source), str(target))
    if sha256_file(target) != fingerprint:
        raise RuntimeError(f"文件移动后内容校验失败: {target}")


def _archive_material(
    application: dict[str, Any],
    source: Path,
    role: str,
    root: Path,
    changes: list[dict[str, Any]],
) -> Path:
    materials = application.setdefault("materials", {})
    role_files = materials.setdefault(role, [])
    index = len(role_files) + 1
    target = (
        root
        / Path(str(application["case_directory"]))
        / material_file_name(
            str(application["case_name"]),
            role,
            index,
            source.suffix,
        )
    )
    fingerprint = sha256_file(source)
    _move_verified(source, target, root)
    role_files.append(
        file_record(
            target,
            target,
            root,
            role,
            fingerprint=fingerprint,
        )
    )
    changes.append(
        _change(
            "move",
            source,
            target,
            role,
            str(application["case_id"]),
        )
    )
    return target


def _classify_recognized_image(text: str) -> str | None:
    normalized = legacy.normalize_match_text(text)
    if (
        "工人名单" in normalized
        or "人员名单" in normalized
        or all(label in normalized for label in ("姓名", "性别", "电话"))
    ):
        return WORKER_LIST_ROLE
    if "质保作业申请单" in normalized or "质保申请单" in normalized:
        return SIGNED_APPLICATION_ROLE
    if (
        "有限空间作业申请" in normalized
        or _work_option_is_checked(text, "有限空间作业")
    ):
        return CONFINED_SPACE_ROLE
    if (
        "高处作业申请" in normalized
        or "高空作业申请" in normalized
        or _work_option_is_checked(text, "5米以上高处作业")
        or _work_option_is_checked(text, "5米以上高空作业")
    ):
        return HIGH_ALTITUDE_ROLE
    if any(
        phrase in normalized
        for phrase in (
            "动火作业申请",
            "危大工程专项方案",
            "配电室接电申请",
        )
    ) or any(
        _work_option_is_checked(text, option)
        for option in ("动火作业", "危大工程", "配电室接电")
    ):
        return SPECIAL_WORK_ROLE
    return None


def _match_recognized_image_application(
    text: str,
    applications: list[dict[str, Any]],
    role: str,
) -> tuple[dict[str, Any] | None, str]:
    """Find one existing case from strong form fields in recognized image text."""
    normalized_text = legacy.normalize_match_text(text)
    image_content = legacy.extract_pdf_construction_content(normalized_text)
    if not image_content:
        return None, "OCR 未提取到施工内容字段"

    ranked: list[tuple[float, bool, dict[str, Any]]] = []
    for application in applications:
        if application.get("status") == "terminated":
            continue
        business = application.get("application") or {}
        area_key = legacy.normalize_pdf_ocr_match_text(
            str(business.get("施工区域") or "")
        )
        if not area_key or area_key not in legacy.normalize_pdf_ocr_match_text(
            normalized_text
        ):
            continue
        if not legacy.pdf_construction_dates_match(
            normalized_text,
            str(business.get("施工开始时间") or ""),
            str(business.get("施工结束时间") or ""),
        ):
            continue

        content_key = legacy.normalize_match_text(
            str(business.get("施工内容") or "")
        )
        if not content_key:
            continue
        exact_content = (
            content_key in image_content or image_content in content_key
        )
        similarity = (
            1.0
            if exact_content
            else SequenceMatcher(None, content_key, image_content).ratio()
        )
        role_missing = not bool(
            ((application.get("materials") or {}).get(role) or [])
        )
        ranked.append((similarity, role_missing, application))

    if not ranked:
        return None, "没有同时命中施工区域和起止日期的已有案卷"

    exact_matches = [item for item in ranked if item[0] == 1.0]
    if len(exact_matches) == 1:
        return exact_matches[0][2], "施工区域、起止日期和施工内容严格命中"
    if len(exact_matches) > 1:
        missing_exact = [item for item in exact_matches if item[1]]
        if len(missing_exact) == 1:
            return (
                missing_exact[0][2],
                "多个案卷字段严格命中，仅该案卷缺少当前材料类型",
            )
        return None, f"施工内容严格命中多个已有案卷: {len(exact_matches)}"

    ranked.sort(
        key=lambda item: (
            item[0],
            item[1],
            str(item[2].get("case_name") or ""),
        ),
        reverse=True,
    )
    best_similarity, best_role_missing, best_application = ranked[0]
    runner_up_similarity = ranked[1][0] if len(ranked) > 1 else 0.0
    if (
        best_similarity >= IMAGE_CONTENT_SIMILARITY_MIN
        and best_similarity - runner_up_similarity
        >= IMAGE_CONTENT_SIMILARITY_MARGIN
    ):
        missing_note = "，且缺少当前材料类型" if best_role_missing else ""
        return (
            best_application,
            "施工区域和起止日期命中；"
            f"施工内容相似度 {best_similarity:.0%}，"
            f"领先次选 {best_similarity - runner_up_similarity:.0%}"
            f"{missing_note}",
        )
    return (
        None,
        "施工区域和起止日期命中，但施工内容无法唯一确认："
        f"最高相似度 {best_similarity:.0%}，"
        f"次高 {runner_up_similarity:.0%}",
    )


def _work_option_is_checked(text: str, option: str) -> bool:
    normalized = legacy.normalize_match_text(text)
    checked_markers = "☑☒✓✔√■●◆⊠▣▩◉x×"
    escaped_option = re.escape(legacy.normalize_match_text(option))
    separators = r"[：:;；,、\[\](){}【】□☐○◇]{0,3}"
    patterns = (
        (
            rf"{escaped_option}{separators}"
            rf"[{re.escape(checked_markers)}]"
        ),
        rf"{escaped_option}[：:]?(?:已勾选|是)",
    )
    return any(re.search(pattern, normalized) for pattern in patterns)


def _is_approval_pdf(path: Path, text: str) -> tuple[bool, str]:
    if approval_application_no(path.name):
        return True, "文件名符合审批 PDF 编号格式"
    normalized_name = legacy.normalize_match_text(path.stem)
    if "工程类主体质保施工" in normalized_name:
        return True, "文件名包含工程类主体质保施工"
    normalized_text = legacy.normalize_match_text(text)
    if "工程类主体质保施工" in normalized_text:
        return True, "PDF 内容包含工程类主体质保施工"
    if (
        "主体质保施工" in normalized_text
        and (
            "申请编号" in normalized_text
            or "审批编号" in normalized_text
        )
    ):
        return True, "PDF 内容包含主体质保施工及审批编号"
    return False, "未发现审批 PDF 标题或编号特征"


def _unique_inbox_target(inbox: Path, source: Path) -> Path:
    target = inbox / source.name
    if not target.exists():
        return target
    index = 2
    while True:
        candidate = inbox / f"{source.stem}_{index:02d}{source.suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def _quarantine_input_duplicate(
    source: Path,
    root: Path,
    changes: list[dict[str, Any]],
) -> Path:
    quarantine = (
        root
        / INTERNAL_DIR_NAME
        / QUARANTINE_DIR_NAME
        / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        / source.name
    )
    quarantine.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(source), str(quarantine))
    changes.append(
        _change(
            "quarantine_input_duplicate",
            source,
            quarantine,
            "input_file",
        )
    )
    return quarantine


def _quarantine_duplicate_application_sources(
    sources: list[Path],
    root: Path,
    changes: list[dict[str, Any]],
    case_id: str,
) -> list[Path]:
    quarantine_dir = (
        root
        / INTERNAL_DIR_NAME
        / QUARANTINE_DIR_NAME
        / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    )
    quarantined: list[Path] = []
    for source in sources:
        source = ensure_within(source, root)
        target = quarantine_dir / source.name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(source), str(target))
        changes.append(
            _change(
                "quarantine_duplicate_application",
                source,
                target,
                WORD_ROLE if source.suffix.lower() == ".docx" else "input_file",
                case_id,
            )
        )
        quarantined.append(target)
    return quarantined


def _application_contains_hash(
    application: dict[str, Any],
    fingerprint: str,
) -> bool:
    return any(
        str(file_item.get("sha256") or "") == fingerprint
        for role_files in (application.get("materials") or {}).values()
        for file_item in role_files or []
        if isinstance(file_item, dict)
    )


def _archive_material_once(
    application: dict[str, Any],
    source: Path,
    role: str,
    root: Path,
    changes: list[dict[str, Any]],
) -> bool:
    fingerprint = sha256_file(source)
    if _application_contains_hash(application, fingerprint):
        quarantined = _quarantine_duplicate_application_sources(
            [source],
            root,
            changes,
            str(application.get("case_id") or ""),
        )
        LOGGER.warning(
            "申请材料内容重复，未再次归档: 文件=%s；案卷=%s；"
            "SHA-256=%s；重复文件已移至=%s",
            source.name,
            application.get("case_name") or "",
            fingerprint,
            quarantined[0],
        )
        return False
    _archive_material(application, source, role, root, changes)
    return True


def _applications_missing_worker_list(
    applications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    return [
        application
        for application in applications
        if application.get("status") != "terminated"
        and not (
            ((application.get("materials") or {}).get(WORKER_LIST_ROLE) or [])
        )
    ]


def _unique_applications(
    applications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    unique: list[dict[str, Any]] = []
    seen: set[str] = set()
    for application in applications:
        key = str(
            application.get("case_id")
            or application.get("case_name")
            or id(application)
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(application)
    return unique


def _archive_named_worker_list_batch(
    images: list[Path],
    batch_word_files: set[Path],
    batch_applications: list[dict[str, Any]],
    root: Path,
    changes: list[dict[str, Any]],
) -> int:
    """Assign filename-marked worker lists only when the batch is one-to-one."""
    if not images:
        return 0

    applications = _unique_applications(batch_applications)
    if (
        len(images) != len(batch_word_files)
        or len(applications) != len(batch_word_files)
    ):
        raise ValueError(
            "本批次工人名单与质保单无法一一对应，已停止文件保存和 JSON "
            "入库："
            f"工人名单={len(images)}；质保单={len(batch_word_files)}；"
            f"可入库案卷={len(applications)}"
        )
        return 0

    sources = sorted(images, key=lambda path: path.name)
    targets = sorted(
        applications,
        key=lambda application: str(application.get("case_name") or ""),
    )
    archived = 0
    LOGGER.info(
        "本批次文件名标记的工人名单与质保单数量一致，按稳定顺序一一入库: %d 份",
        len(sources),
    )
    for source, application in zip(sources, targets):
        if _archive_material_once(
            application,
            source,
            WORKER_LIST_ROLE,
            root,
            changes,
        ):
            archived += 1
            _refresh_status(application)
    return archived


def _archive_identical_worker_list_groups(
    images: list[Path],
    applications: list[dict[str, Any]],
    batch_applications: list[dict[str, Any]],
    root: Path,
    changes: list[dict[str, Any]],
) -> int:
    """Assign identical worker-list copies one-to-one to a provable case set."""
    groups: dict[str, list[Path]] = {}
    for image in images:
        groups.setdefault(sha256_file(image), []).append(image)

    archived = 0
    for fingerprint, group in sorted(groups.items()):
        sources = sorted(group, key=lambda path: path.name)
        batch_candidates = _applications_missing_worker_list(
            _unique_applications(batch_applications)
        )
        candidates: list[dict[str, Any]] = []
        evidence = ""
        if len(batch_candidates) == len(sources):
            candidates = batch_candidates
            evidence = "同批 Word 案卷数与相同名单图片数一致"
        else:
            modified_dates = {
                datetime.fromtimestamp(source.stat().st_mtime)
                .date()
                .isoformat()
                for source in sources
            }
            if len(modified_dates) == 1:
                modified_date = next(iter(modified_dates))
                date_candidates = [
                    application
                    for application in _applications_missing_worker_list(
                        applications
                    )
                    if str(
                        (application.get("application") or {}).get(
                            "施工开始时间"
                        )
                        or ""
                    )
                    == modified_date
                ]
                if len(date_candidates) == len(sources):
                    candidates = date_candidates
                    evidence = (
                        f"图片修改日期 {modified_date} 对应的缺名单案卷数"
                        "与图片数一致"
                    )

        if not candidates:
            LOGGER.warning(
                "相同工人名单无法一对一关联案卷，保留在 _inbox: "
                "SHA-256=%s；图片数=%d；同批缺名单案卷数=%d",
                fingerprint,
                len(sources),
                len(batch_candidates),
            )
            continue

        ordered_candidates = sorted(
            candidates,
            key=lambda application: str(application.get("case_name") or ""),
        )
        LOGGER.info(
            "相同工人名单按批次一对一入库: SHA-256=%s；图片数=%d；"
            "案卷数=%d（%s）",
            fingerprint,
            len(sources),
            len(ordered_candidates),
            evidence,
        )
        for source, application in zip(sources, ordered_candidates):
            if _archive_material_once(
                application,
                source,
                WORKER_LIST_ROLE,
                root,
                changes,
            ):
                archived += 1
                _refresh_status(application)
    return archived


def reclassify_historical_materials(
    dataset: dict[str, Any],
    root: Path,
    checkpoint: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, int]:
    root = root.resolve()
    cache = dataset.get("recognition_cache") or {}
    changes = dataset.setdefault("changes", [])
    summary = {
        "records_checked": 0,
        "records_reclassified": 0,
        "missing_references_removed": 0,
        "files_moved": 0,
        "recognition_cache_missing": 0,
    }
    special_roles = (
        CONFINED_SPACE_ROLE,
        HIGH_ALTITUDE_ROLE,
        SPECIAL_WORK_ROLE,
    )
    for application in dataset.get("applications") or []:
        application_changed = False
        materials = application.setdefault("materials", {})
        for current_role in special_roles:
            role_files = materials.setdefault(current_role, [])
            for file_item in list(role_files):
                if not isinstance(file_item, dict):
                    continue
                summary["records_checked"] += 1
                fingerprint = str(file_item.get("sha256") or "")
                recognition = cache.get(fingerprint) or {}
                text = str(recognition.get("text") or "")
                if not text:
                    summary["recognition_cache_missing"] += 1
                    continue
                resolved_role = _classify_recognized_image(text)
                if resolved_role is None or resolved_role == current_role:
                    continue

                relative = str(file_item.get("path") or "")
                try:
                    source = ensure_within(root / Path(relative), root)
                except ValueError:
                    LOGGER.warning(
                        "历史材料路径越界，跳过重判: 案卷=%s；路径=%s",
                        application.get("case_name") or "",
                        relative,
                    )
                    continue
                role_files.remove(file_item)
                target = source
                result = "source_missing_record_removed"
                if source.is_file():
                    before_changes = len(changes)
                    archived = _archive_material_once(
                        application,
                        source,
                        resolved_role,
                        root,
                        changes,
                    )
                    if archived:
                        target_item = materials[resolved_role][-1]
                        target = ensure_within(
                            root / Path(str(target_item["path"])),
                            root,
                        )
                        summary["files_moved"] += 1
                        result = "file_reclassified"
                    elif len(changes) > before_changes:
                        target = Path(str(changes[-1]["target"]))
                        result = "duplicate_file_quarantined"
                else:
                    summary["missing_references_removed"] += 1

                changes.append(
                    _change(
                        "reclassify_historical_material",
                        source,
                        target,
                        f"{current_role}->{resolved_role}",
                        str(application.get("case_id") or ""),
                        result,
                    )
                )
                summary["records_reclassified"] += 1
                application_changed = True
                LOGGER.warning(
                    "历史申请材料重新分类: 案卷=%s；原类型=%s；"
                    "新类型=%s；处理结果=%s；文件=%s",
                    application.get("case_name") or "",
                    ROLE_LABELS_FOR_LOG.get(current_role, current_role),
                    ROLE_LABELS_FOR_LOG.get(resolved_role, resolved_role),
                    RESULT_LABELS_FOR_LOG.get(result, result),
                    source.name,
                )
        if application_changed:
            _refresh_status(application)
            if checkpoint:
                checkpoint(dataset)
    return summary


ROLE_LABELS_FOR_LOG = {
    SIGNED_APPLICATION_ROLE: "手签申请单",
    WORKER_LIST_ROLE: "施工人员名单",
    CONFINED_SPACE_ROLE: "有限空间申请",
    HIGH_ALTITUDE_ROLE: "高处作业申请",
    SPECIAL_WORK_ROLE: "专项作业材料",
}

RESULT_LABELS_FOR_LOG = {
    "source_missing_record_removed": "源文件不存在，已清理旧记录",
    "file_reclassified": "文件已重新分类并规范命名",
    "duplicate_file_quarantined": "重复文件已移入隔离区",
}


def _input_batch_files(
    dataset: dict[str, Any],
    root: Path,
    input_batch_id: str,
) -> set[Path]:
    if not input_batch_id:
        return set()
    paths: set[Path] = set()
    for item in dataset.get("input_routes") or []:
        if (
            item.get("input_batch_id") != input_batch_id
            or item.get("kind") != APPLICATION_MATERIAL_ROUTE
        ):
            continue
        relative_path = str(item.get("processing_path") or "")
        if not relative_path and item.get("action") == "routed":
            relative_path = str(item.get("path") or "")
        if not relative_path and item.get("action") == "quarantine_duplicate":
            source_name = Path(str(item.get("source_path") or "")).name
            relative_path = str(Path("_inbox") / source_name)
        try:
            path = ensure_within(
                root / Path(relative_path),
                root,
            )
        except ValueError:
            continue
        if path.is_file():
            paths.add(path.resolve())
    return paths


def route_input_files(
    dataset: dict[str, Any],
    root: Path,
    repo_root: Path,
    checkpoint: Callable[[dict[str, Any]], None] | None = None,
    input_batch_id: str = "",
) -> dict[str, int]:
    """Classify the public _input drop zone before business workflows run."""
    root = root.resolve()
    input_root = root / INPUT_DIR_NAME
    inbox = root / "_inbox"
    input_root.mkdir(parents=True, exist_ok=True)
    sources = sorted(
        path
        for path in input_root.iterdir()
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.lower()
        in {*IMAGE_SUFFIXES, ".docx", PDF_SUFFIX}
    )
    summary = {
        "input_files_routed": 0,
        "approval_pdfs_routed": 0,
        "application_files_routed": 0,
        "input_duplicates_quarantined": 0,
    }
    if not sources:
        return summary

    input_word_count = sum(
        path.suffix.lower() == ".docx" for path in sources
    )
    named_worker_list_count = sum(
        path.suffix.lower() in IMAGE_SUFFIXES and "工人名单" in path.stem
        for path in sources
    )
    if named_worker_list_count != input_word_count:
        message = (
            "_input 中工人名单与质保单数量不一致，已停止文件保存和 JSON "
            f"入库：工人名单={named_worker_list_count}；"
            f"质保单={input_word_count}"
        )
        LOGGER.error(message)
        raise ValueError(message)

    inbox.mkdir(parents=True, exist_ok=True)

    resolved_batch_id = input_batch_id or str(uuid.uuid4())
    routes = dataset.setdefault("input_routes", [])
    changes = dataset.setdefault("changes", [])
    with RecognitionService(dataset, repo_root) as recognition:
        for source in sources:
            fingerprint = sha256_file(source)
            route_kind = APPLICATION_MATERIAL_ROUTE
            reason = "Word 或图片进入申请材料流程"
            recognition_method = ""
            if source.suffix.lower() == PDF_SUFFIX:
                LOGGER.info("开始判断 _input PDF 类型: %s", source.name)
                is_approval, reason = _is_approval_pdf(source, "")
                if is_approval:
                    recognition_method = "filename"
                else:
                    try:
                        text = recognition.pdf_text(source)
                    except Exception as exc:
                        LOGGER.warning(
                            "_input PDF 类型识别失败，保留在 _input: %s (%s)",
                            source.name,
                            exc,
                        )
                        continue
                    if checkpoint:
                        checkpoint(dataset)
                    is_approval, reason = _is_approval_pdf(
                        source,
                        text,
                    )
                    recognition_method = str(
                        (
                            (dataset.get("recognition_cache") or {})
                            .get(fingerprint, {})
                            .get("method", "unknown")
                        )
                    )
                route_kind = (
                    APPROVAL_PDF_ROUTE
                    if is_approval
                    else APPLICATION_MATERIAL_ROUTE
                )
                LOGGER.info(
                    "_input PDF 分类完成: %s -> %s (%s)",
                    source.name,
                    (
                        "审批 PDF 流程"
                        if is_approval
                        else "申请材料流程"
                    ),
                    reason,
                )

            existing_target = inbox / source.name
            if (
                existing_target.is_file()
                and sha256_file(existing_target) == fingerprint
            ):
                target = _quarantine_input_duplicate(
                    source,
                    root,
                    changes,
                )
                summary["input_duplicates_quarantined"] += 1
                action = "quarantine_duplicate"
                processing_target = existing_target
            else:
                target = _unique_inbox_target(inbox, source)
                _move_verified(source, target, root)
                changes.append(
                    _change(
                        "route_input",
                        source,
                        target,
                        route_kind,
                    )
                )
                summary["input_files_routed"] += 1
                action = "routed"
                processing_target = target

            routes.append(
                {
                    "route_version": INPUT_ROUTE_VERSION,
                    "sha256": fingerprint,
                    "source_path": relative_posix(
                        input_root / source.name,
                        root,
                    ),
                    "path": relative_posix(target, root),
                    "processing_path": relative_posix(
                        processing_target,
                        root,
                    ),
                    "kind": route_kind,
                    "reason": reason,
                    "recognition_method": recognition_method,
                    "action": action,
                    "input_batch_id": resolved_batch_id,
                    "routed_at": datetime.now().astimezone().isoformat(
                        timespec="seconds"
                    ),
                }
            )
            if route_kind == APPROVAL_PDF_ROUTE:
                summary["approval_pdfs_routed"] += 1
            else:
                summary["application_files_routed"] += 1
            if checkpoint:
                checkpoint(dataset)
    return summary


def intake_applications(
    dataset: dict[str, Any],
    root: Path,
    repo_root: Path,
    checkpoint: Callable[[dict[str, Any]], None] | None = None,
    input_batch_id: str = "",
    intake_stats: dict[str, int] | None = None,
) -> int:
    root = root.resolve()
    inbox = root / "_inbox"
    if not inbox.is_dir():
        return 0
    word_files = sorted(
        path
        for path in inbox.glob("*.docx")
        if path.is_file() and not path.name.startswith("~$")
    )

    applications = dataset.setdefault("applications", [])
    changes = dataset.setdefault("changes", [])
    template = root / TEMPLATES_DIR_NAME / TEMPLATE_FILE_NAME
    template_hash = ""
    if word_files:
        if not template.is_file():
            raise FileNotFoundError(f"缺少安全协议模板: {template}")
        template_hash = sha256_file(template)
    created: list[dict[str, Any]] = []
    claimed: set[Path] = set()
    batch_files = _input_batch_files(dataset, root, input_batch_id)
    batch_word_files = {
        path
        for path in batch_files
        if path.suffix.lower() == ".docx"
    }
    single_batch_word = (
        next(iter(batch_word_files))
        if len(batch_word_files) == 1
        else None
    )
    batch_application: dict[str, Any] | None = None
    batch_applications: list[dict[str, Any]] = []
    named_batch_worker_lists = sorted(
        path
        for path in batch_files
        if path.is_file()
        and path.suffix.lower() in IMAGE_SUFFIXES
        and "工人名单" in path.stem
    )
    named_worker_list_paths = {
        path.resolve() for path in named_batch_worker_lists
    }

    for source_word in word_files:
        parsed = legacy.parse_document(source_word)
        canonical_word_name = legacy.build_target_name(
            str(parsed["施工开始时间"]),
            str(parsed["施工内容"]),
        )
        base_case_name = Path(canonical_word_name).stem
        duplicate = _find_duplicate_application(
            parsed, applications, root
        )
        if duplicate is not None:
            removed_duplicates = (
                _quarantine_less_complete_duplicate_cases(
                    dataset,
                    parsed,
                    duplicate,
                    root,
                    changes,
                )
            )
            quarantined = _quarantine_duplicate_application_sources(
                [source_word],
                root,
                changes,
                str(duplicate.get("case_id") or ""),
            )
            claimed.add(source_word.resolve())
            if (
                single_batch_word is not None
                and source_word.resolve() == single_batch_word
            ):
                batch_application = duplicate
            if source_word.resolve() in batch_word_files:
                batch_applications.append(duplicate)
            duplicate_business = duplicate.get("application") or {}
            LOGGER.warning(
                "检测到重复质保申请，已跳过写入 JSON: 文件=%s；"
                "已有案卷=%s；施工开始时间=%s；施工结束时间=%s；"
                "施工内容=%s；施工区域=%s；重复文件已移至=%s",
                source_word.name,
                duplicate.get("case_name") or "",
                duplicate_business.get("施工开始时间") or "",
                duplicate_business.get("施工结束时间") or "",
                duplicate_business.get("施工内容") or "",
                duplicate_business.get("施工区域") or "",
                quarantined[0].parent,
            )
            if removed_duplicates and checkpoint:
                checkpoint(dataset)
            direct_candidates = sorted(
                path
                for path in inbox.iterdir()
                if path.is_file()
                and path.resolve() not in claimed
                and path.resolve() not in named_worker_list_paths
                and path.suffix.lower()
                in {*IMAGE_SUFFIXES, ".pdf"}
            )
            for candidate in direct_candidates:
                role = material_role(candidate, source_word.stem)
                if role is None:
                    role = material_role(candidate, base_case_name)
                if role is None:
                    continue
                _archive_material_once(
                    duplicate,
                    candidate,
                    role,
                    root,
                    changes,
                )
                claimed.add(candidate.resolve())
            _refresh_status(duplicate)
            if checkpoint:
                checkpoint(dataset)
            continue
        case_name = _next_case_name(base_case_name, applications)
        case_id = str(uuid.uuid5(CASE_NAMESPACE, case_name))
        case_dir = root / CASES_DIR_NAME / case_name
        target_word = case_dir / f"{case_name}.docx"
        word_hash = sha256_file(source_word)
        _move_verified(source_word, target_word, root)
        claimed.add(source_word.resolve())
        changes.append(
            _change(
                "move",
                source_word,
                target_word,
                WORD_ROLE,
                case_id,
            )
        )
        materials: dict[str, list[dict[str, Any]]] = {
            WORD_ROLE: [
                file_record(
                    target_word,
                    target_word,
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
        application = {
            "case_id": case_id,
            "case_name": case_name,
            "case_directory": relative_posix(case_dir, root),
            "status": "materials_incomplete",
            "application": _business_data(parsed),
            "required_material_types": _required_roles(parsed),
            "missing_material_types": [],
            "materials": materials,
            "approval": {
                "status": "not_received",
                "application_no": "",
                "pdfs": [],
                "match_source": "",
            },
            "history": [],
        }
        applications.append(application)
        created.append(application)
        if (
            single_batch_word is not None
            and source_word.resolve() == single_batch_word
        ):
            batch_application = application
        if source_word.resolve() in batch_word_files:
            batch_applications.append(application)

        direct_candidates = sorted(
            path
            for path in inbox.iterdir()
            if path.is_file()
            and path.resolve() not in claimed
            and path.resolve() not in named_worker_list_paths
            and path.suffix.lower()
            in {*IMAGE_SUFFIXES, ".pdf"}
        )
        for candidate in direct_candidates:
            role = material_role(candidate, source_word.stem)
            if role is None:
                role = material_role(candidate, base_case_name)
            if role is None:
                continue
            _archive_material_once(
                application,
                candidate,
                role,
                root,
                changes,
            )
            claimed.add(candidate.resolve())

        agreement_target = (
            case_dir
            / f"{application_prefix(Path(canonical_word_name))}_{TEMPLATE_FILE_NAME}"
        )
        agreement_target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template, agreement_target)
        if sha256_file(agreement_target) != template_hash:
            raise RuntimeError(f"安全协议复制校验失败: {agreement_target}")
        materials[SAFETY_AGREEMENT_ROLE].append(
            file_record(
                agreement_target,
                agreement_target,
                root,
                SAFETY_AGREEMENT_ROLE,
                fingerprint=template_hash,
                derived=True,
            )
        )
        changes.append(
            _change(
                "copy",
                template,
                agreement_target,
                SAFETY_AGREEMENT_ROLE,
                case_id,
            )
        )
        _refresh_status(application)

    named_worker_lists_archived = _archive_named_worker_list_batch(
        named_batch_worker_lists,
        batch_word_files,
        batch_applications,
        root,
        changes,
    )
    if named_worker_lists_archived:
        claimed.update(named_worker_list_paths)

    remaining_images = sorted(
        path
        for path in inbox.iterdir()
        if path.is_file()
        and path.suffix.lower() in IMAGE_SUFFIXES
        and path.resolve() not in claimed
    )
    pending_worker_lists: list[Path] = []
    with RecognitionService(dataset, repo_root) as recognition:
        for image in remaining_images:
            if image.resolve() in named_worker_list_paths:
                continue
            try:
                text = recognition.image_text(image)
            except Exception as exc:
                LOGGER.warning("新申请材料图片识别失败: %s (%s)", image.name, exc)
                continue
            if checkpoint:
                checkpoint(dataset)
            role = _classify_recognized_image(text)
            if role is None:
                LOGGER.warning(
                    "无法判断申请材料图片类型，保留在 _inbox: %s",
                    image.name,
                )
                continue
            if role == WORKER_LIST_ROLE and not (
                batch_application is not None
                and image.resolve() in batch_files
            ):
                pending_worker_lists.append(image)
                continue
            if (
                batch_application is not None
                and image.resolve() in batch_files
            ):
                candidates = [batch_application]
                LOGGER.info(
                    "按 _input 同批投放策略关联申请材料: %s -> %s",
                    image.name,
                    batch_application.get("case_name") or "",
                )
            else:
                normalized_text = legacy.normalize_match_text(text)
                candidates = [
                    application
                    for application in created
                    if legacy.normalize_match_text(
                        str(
                            (application.get("application") or {}).get(
                                "施工内容"
                            )
                            or ""
                        )
                    )
                    in normalized_text
                ]
                if len(candidates) != 1:
                    matched_application, match_reason = (
                        _match_recognized_image_application(
                            text,
                            applications,
                            role,
                        )
                    )
                    candidates = (
                        [matched_application] if matched_application else []
                    )
                    if matched_application is not None:
                        LOGGER.info(
                            "按图片识别字段关联申请材料: %s -> %s (%s)",
                            image.name,
                            matched_application.get("case_name") or "",
                            match_reason,
                        )
                    else:
                        LOGGER.warning(
                            "图片识别字段未能唯一关联案卷: %s (%s)",
                            image.name,
                            match_reason,
                        )
            if len(candidates) != 1:
                LOGGER.warning(
                    "申请材料图片无法唯一关联案卷，保留在 _inbox: %s",
                    image.name,
                )
                continue
            _archive_material_once(
                candidates[0],
                image,
                role,
                root,
                changes,
            )
            _refresh_status(candidates[0])
    worker_lists_archived = named_worker_lists_archived
    worker_lists_archived += _archive_identical_worker_list_groups(
        pending_worker_lists,
        applications,
        batch_applications,
        root,
        changes,
    )
    if worker_lists_archived and checkpoint:
        checkpoint(dataset)
    if intake_stats is not None:
        intake_stats["worker_lists_ingested"] = worker_lists_archived
    return len(created)


def _approval_candidates(
    path: Path,
    text: str,
    applications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for application in applications:
        if application.get("status") == "terminated":
            continue
        business = application.get("application") or {}
        matched = legacy.find_matching_pdf_paths(
            str(business.get("施工区域") or ""),
            str(business.get("施工内容") or ""),
            {path: text},
            str(business.get("施工开始时间") or ""),
            str(business.get("施工结束时间") or ""),
        )
        if matched:
            candidates.append(application)
    return candidates


def _resolve_unmatched_path(root: Path, item: dict[str, Any]) -> Path:
    return ensure_within(root / Path(str(item.get("path") or "")), root)


def archive_reviewed_approval_pdf(
    dataset: dict[str, Any],
    root: Path,
    item: dict[str, Any],
    application: dict[str, Any],
    review_id: str,
    review_note: str = "",
) -> Path:
    """Archive one approval PDF after an explicit human confirmation."""
    root = root.resolve()
    path = _resolve_unmatched_path(root, item)
    if not path.is_file():
        raise FileNotFoundError(f"待审核审批 PDF 不存在: {path}")
    expected_hash = str(item.get("sha256") or "")
    actual_hash = sha256_file(path)
    if expected_hash and expected_hash != actual_hash:
        raise RuntimeError(f"待审核审批 PDF 哈希已变化: {path.name}")

    approval = application.setdefault("approval", {})
    existing_files = list(approval.get("pdfs") or [])
    if existing_files and not any(
        str(existing.get("sha256") or "") == actual_hash
        for existing in existing_files
    ):
        raise RuntimeError(
            f"案卷已有其他审批 PDF，不能再次人工匹配: "
            f"{application.get('case_name')}"
        )

    text = str(
        (dataset.get("recognition_cache") or {})
        .get(actual_hash, {})
        .get("text", "")
    )
    application_no = (
        approval_application_no(path.name)
        or legacy.extract_pdf_rename_application_no(text)
        or legacy.extract_pdf_application_no_from_name(path.name)
    )
    target_name = (
        f"{legacy.PDF_TARGET_NAME_PREFIX}{application_no}.pdf"
        if application_no
        else path.name
    )
    target = (
        root / Path(str(application["case_directory"])) / target_name
    )
    target.parent.mkdir(parents=True, exist_ok=True)
    changes = dataset.setdefault("changes", [])
    if target.exists():
        if sha256_file(target) != actual_hash:
            raise FileExistsError(
                f"审批 PDF 目标文件已存在且内容不同: {target}"
            )
        quarantine = (
            root
            / INTERNAL_DIR_NAME
            / QUARANTINE_DIR_NAME
            / datetime.now().strftime("%Y%m%d_%H%M%S")
            / path.name
        )
        quarantine.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(path), str(quarantine))
        changes.append(
            _change(
                "quarantine_duplicate",
                path,
                quarantine,
                APPROVAL_PDF_ROLE,
                str(application["case_id"]),
            )
        )
    else:
        _move_verified(path, target, root)
        changes.append(
            _change(
                "move_after_human_review",
                path,
                target,
                APPROVAL_PDF_ROLE,
                str(application["case_id"]),
            )
        )

    approval_files = approval.setdefault("pdfs", [])
    if not any(
        str(existing.get("sha256") or "") == actual_hash
        for existing in approval_files
    ):
        approval_files.append(
            file_record(
                target,
                target,
                root,
                APPROVAL_PDF_ROLE,
                fingerprint=actual_hash,
            )
        )
    reviewed_at = datetime.now().astimezone().isoformat(timespec="seconds")
    approval.update(
        {
            "application_no": application_no,
            "match_source": "human_review",
            "status": "approved",
            "review_id": review_id,
            "review_note": review_note,
            "reviewed_at": reviewed_at,
        }
    )
    application.setdefault("history", []).append(
        {
            "action": "approval_pdf_human_confirmed",
            "review_id": review_id,
            "pdf_sha256": actual_hash,
            "note": review_note,
            "at": reviewed_at,
        }
    )
    _refresh_status(application)
    dataset["unmatched_files"] = [
        unmatched
        for unmatched in dataset.get("unmatched_files") or []
        if not (
            str(unmatched.get("sha256") or "") == actual_hash
            or str(unmatched.get("path") or "")
            == str(item.get("path") or "")
        )
    ]
    return target


def ingest_approval_pdfs(
    dataset: dict[str, Any],
    root: Path,
    repo_root: Path,
    checkpoint: Callable[[dict[str, Any]], None] | None = None,
) -> int:
    root = root.resolve()
    applications = list(dataset.get("applications") or [])
    unmatched = list(dataset.get("unmatched_files") or [])
    known_paths = {str(item.get("path") or "") for item in unmatched}
    application_material_pdf_hashes = {
        str(item.get("sha256") or "")
        for item in dataset.get("input_routes") or []
        if item.get("kind") == APPLICATION_MATERIAL_ROUTE
    }
    inbox = root / "_inbox"
    if inbox.is_dir():
        for path in sorted(inbox.glob("*.pdf")):
            if sha256_file(path) in application_material_pdf_hashes:
                LOGGER.info(
                    "跳过已分类为申请材料的 PDF: %s",
                    path.name,
                )
                continue
            relative = relative_posix(path, root)
            if relative in known_paths:
                continue
            unmatched.append(
                file_record(path, path, root, APPROVAL_PDF_ROLE)
            )

    retained: list[dict[str, Any]] = []
    ingested = 0
    changes = dataset.setdefault("changes", [])
    with RecognitionService(dataset, repo_root) as recognition:
        for item in unmatched:
            if item.get("role") != APPROVAL_PDF_ROLE:
                retained.append(item)
                continue
            path = _resolve_unmatched_path(root, item)
            if not path.is_file():
                item["exists"] = False
                item["review_reason"] = "文件不存在"
                retained.append(item)
                continue
            try:
                text = recognition.pdf_text(path)
            except Exception as exc:
                item["review_reason"] = f"识别失败: {exc}"
                retained.append(item)
                LOGGER.warning("审批 PDF 识别失败: %s (%s)", path.name, exc)
                continue
            if checkpoint:
                checkpoint(dataset)
            candidates = _approval_candidates(path, text, applications)
            if len(candidates) != 1:
                item["review_reason"] = (
                    "未匹配到案卷"
                    if not candidates
                    else f"匹配到多个案卷: {len(candidates)}"
                )
                item["candidate_case_ids"] = [
                    candidate["case_id"] for candidate in candidates
                ]
                retained.append(item)
                continue

            application = candidates[0]
            application_no = (
                approval_application_no(path.name)
                or legacy.extract_pdf_rename_application_no(text)
                or legacy.extract_pdf_application_no_from_name(path.name)
            )
            target_name = (
                f"{legacy.PDF_TARGET_NAME_PREFIX}{application_no}.pdf"
                if application_no
                else path.name
            )
            target = (
                root
                / Path(str(application["case_directory"]))
                / target_name
            )
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                if sha256_file(target) != sha256_file(path):
                    item["review_reason"] = f"目标文件已存在且内容不同: {target.name}"
                    retained.append(item)
                    continue
                quarantine = (
                    root
                    / INTERNAL_DIR_NAME
                    / QUARANTINE_DIR_NAME
                    / datetime.now().strftime("%Y%m%d_%H%M%S")
                    / path.name
                )
                quarantine.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(path), str(quarantine))
                changes.append(
                    _change(
                        "quarantine_duplicate",
                        path,
                        quarantine,
                        APPROVAL_PDF_ROLE,
                        application["case_id"],
                    )
                )
            else:
                shutil.move(str(path), str(target))
                changes.append(
                    _change(
                        "move",
                        path,
                        target,
                        APPROVAL_PDF_ROLE,
                        application["case_id"],
                    )
                )
            approval = application.setdefault("approval", {})
            approval_files = approval.setdefault("pdfs", [])
            if not any(
                item.get("path") == relative_posix(target, root)
                and item.get("sha256") == sha256_file(target)
                for item in approval_files
            ):
                approval_files.append(
                    file_record(
                        target,
                        target,
                        root,
                        APPROVAL_PDF_ROLE,
                    )
                )
            approval["application_no"] = application_no
            approval["match_source"] = "content_recognition"
            approval["status"] = "approved"
            _refresh_status(application)
            ingested += 1
            if checkpoint:
                checkpoint(dataset)

    dataset["unmatched_files"] = retained
    return ingested


def ingest_worker_lists(
    dataset: dict[str, Any],
    root: Path,
    *,
    input_batch_id: str = "",
) -> int:
    root = root.resolve()
    inbox = root / "_inbox"
    if not inbox.is_dir():
        return 0
    current_batch_files = _input_batch_files(dataset, root, input_batch_id)
    sources = sorted(
        path
        for path in inbox.iterdir()
        if path.is_file()
        and path.suffix.lower() in {*IMAGE_SUFFIXES, PDF_SUFFIX}
        and ("工人名单" in path.stem or "人员名单" in path.stem)
        and path.resolve() not in current_batch_files
    )
    ingested = 0
    changes = dataset.setdefault("changes", [])
    for source in sources:
        direct = [
            application
            for application in dataset.get("applications") or []
            if source.stem.startswith(str(application.get("case_name") or ""))
        ]
        candidates = direct
        if not candidates:
            modified_date = datetime.fromtimestamp(
                source.stat().st_mtime
            ).date().isoformat()
            candidates = [
                application
                for application in dataset.get("applications") or []
                if str(
                    (application.get("application") or {}).get(
                        "施工开始时间"
                    )
                    or ""
                )
                == modified_date
            ]
        if not candidates:
            continue

        for application in candidates:
            existing = (application.get("materials") or {}).setdefault(
                WORKER_LIST_ROLE, []
            )
            index = len(existing) + 1
            target = (
                root
                / Path(str(application["case_directory"]))
                / f"{application['case_name']}_施工人员名单_{index:02d}{source.suffix.lower()}"
            )
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists():
                if sha256_file(target) != sha256_file(source):
                    LOGGER.warning("施工人员名单目标冲突，已跳过: %s", target)
                    continue
            else:
                shutil.copy2(source, target)
            existing.append(
                file_record(
                    source,
                    target,
                    root,
                    WORKER_LIST_ROLE,
                    fingerprint=sha256_file(source),
                    derived=len(candidates) > 1,
                )
            )
            changes.append(
                _change(
                    "copy",
                    source,
                    target,
                    WORKER_LIST_ROLE,
                    application["case_id"],
                )
            )
            _refresh_status(application)
            ingested += 1

        quarantine = (
            root
            / INTERNAL_DIR_NAME
            / QUARANTINE_DIR_NAME
            / datetime.now().strftime("%Y%m%d_%H%M%S")
            / source.name
        )
        quarantine.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(source), str(quarantine))
        changes.append(
            _change(
                "quarantine_source",
                source,
                quarantine,
                WORKER_LIST_ROLE,
            )
        )
    return ingested


def append_run(
    dataset: dict[str, Any],
    run_type: str,
    summary: dict[str, Any],
) -> None:
    dataset.setdefault("runs", []).append(
        {
            "run_id": datetime.now().strftime("%Y%m%d_%H%M%S"),
            "type": run_type,
            "status": "completed",
            "completed_at": datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
            **summary,
        }
    )
