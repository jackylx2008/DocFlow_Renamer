from __future__ import annotations

import logging
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from . import legacy
from .constants import (
    APPROVAL_PDF_ROLE,
    CASES_DIR_NAME,
    CONFINED_SPACE_ROLE,
    HIGH_ALTITUDE_ROLE,
    IMAGE_SUFFIXES,
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
from .file_utils import ensure_within, relative_posix, sha256_file
from .migration import CASE_NAMESPACE, _required_roles, file_record
from .naming import (
    application_prefix,
    approval_application_no,
    material_file_name,
    material_role,
)
from .recognition import RecognitionService


LOGGER = logging.getLogger(__name__)


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


def _classify_recognized_image(text: str) -> str:
    normalized = legacy.normalize_match_text(text)
    if "工人名单" in normalized or "人员名单" in normalized:
        return WORKER_LIST_ROLE
    if "有限空间" in normalized:
        return CONFINED_SPACE_ROLE
    if "高处作业" in normalized or "高空作业" in normalized:
        return HIGH_ALTITUDE_ROLE
    return SIGNED_APPLICATION_ROLE


def intake_applications(
    dataset: dict[str, Any],
    root: Path,
    repo_root: Path,
    checkpoint: Callable[[dict[str, Any]], None] | None = None,
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
    if not word_files:
        return 0

    applications = dataset.setdefault("applications", [])
    changes = dataset.setdefault("changes", [])
    template = root / TEMPLATES_DIR_NAME / TEMPLATE_FILE_NAME
    if not template.is_file():
        raise FileNotFoundError(f"缺少安全协议模板: {template}")
    template_hash = sha256_file(template)
    created: list[dict[str, Any]] = []
    claimed: set[Path] = set()

    for source_word in word_files:
        parsed = legacy.parse_document(source_word)
        canonical_word_name = legacy.build_target_name(
            str(parsed["施工开始时间"]),
            str(parsed["施工内容"]),
        )
        base_case_name = Path(canonical_word_name).stem
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

        direct_candidates = sorted(
            path
            for path in inbox.iterdir()
            if path.is_file()
            and path.resolve() not in claimed
            and path.suffix.lower()
            in {*IMAGE_SUFFIXES, ".pdf"}
        )
        for candidate in direct_candidates:
            role = material_role(candidate, source_word.stem)
            if role is None:
                role = material_role(candidate, base_case_name)
            if role is None:
                continue
            _archive_material(application, candidate, role, root, changes)
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

    remaining_images = sorted(
        path
        for path in inbox.iterdir()
        if path.is_file()
        and path.suffix.lower() in IMAGE_SUFFIXES
        and path.resolve() not in claimed
        and "工人名单" not in path.stem
        and "人员名单" not in path.stem
    )
    with RecognitionService(dataset, repo_root) as recognition:
        for image in remaining_images:
            try:
                text = recognition.image_text(image)
                if checkpoint:
                    checkpoint(dataset)
            except Exception as exc:
                LOGGER.warning("新申请材料图片识别失败: %s (%s)", image.name, exc)
                continue
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
                in text
            ]
            if len(candidates) != 1:
                continue
            role = _classify_recognized_image(text)
            _archive_material(
                candidates[0],
                image,
                role,
                root,
                changes,
            )
            _refresh_status(candidates[0])
    return len(created)


def _approval_candidates(
    path: Path,
    text: str,
    applications: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for application in applications:
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
    inbox = root / "_inbox"
    if inbox.is_dir():
        for path in sorted(inbox.glob("*.pdf")):
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
                if checkpoint:
                    checkpoint(dataset)
            except Exception as exc:
                item["review_reason"] = f"识别失败: {exc}"
                retained.append(item)
                LOGGER.warning("审批 PDF 识别失败: %s (%s)", path.name, exc)
                continue
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


def ingest_worker_lists(dataset: dict[str, Any], root: Path) -> int:
    root = root.resolve()
    inbox = root / "_inbox"
    if not inbox.is_dir():
        return 0
    sources = sorted(
        path
        for path in inbox.iterdir()
        if path.is_file()
        and path.suffix.lower() in {*IMAGE_SUFFIXES, PDF_SUFFIX}
        and ("工人名单" in path.stem or "人员名单" in path.stem)
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
