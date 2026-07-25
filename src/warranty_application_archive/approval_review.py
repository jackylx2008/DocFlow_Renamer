from __future__ import annotations

import json
import re
import shutil
import uuid
from datetime import date, datetime
from pathlib import Path
from typing import Any

from . import legacy
from .constants import (
    APPROVAL_PDF_ROLE,
    APPROVAL_REVIEW_DATA_FILE_NAME,
    APPROVAL_REVIEW_SCHEMA_VERSION,
    INBOX_DIR_NAME,
    INTERNAL_DIR_NAME,
    LEGACY_APPROVAL_REVIEW_DATA_FILE_NAME,
    LEGACY_DIR_NAME,
    TRASH_DIR_NAME,
)
from .file_utils import atomic_replace_text, ensure_within, sha256_file
from .migration import CASE_NAMESPACE
from .recognition import RecognitionService
from .workflows import archive_reviewed_approval_pdf


DECISION_OPTIONS = ("待审核", "确认匹配", "排除", "移至_trash")
MATCH_RULE_VERSION = "strict-non-date-relaxed-date-v1"
NON_DATE_CONFIDENCE = 80
CASE_STATUS_LABELS = {
    "materials_incomplete": "材料待补充",
    "materials_ready": "材料齐全，待审批PDF",
    "approval_pdf_unmatched": "审批PDF待确认",
    "approved": "审批完成",
    "needs_review": "待人工确认",
    "terminated": "终止",
}
CASE_NAME_RE = re.compile(
    r"^(?P<date>\d{4}-\d{2}-\d{2})_(?P<content>.+?)_质保作业申请单"
)


def _now() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _empty_review(root: Path) -> dict[str, Any]:
    return {
        "schema_version": APPROVAL_REVIEW_SCHEMA_VERSION,
        "data_root": str(root.resolve()),
        "source_dataset_revision": 0,
        "generated_at": _now(),
        "match_rule_version": MATCH_RULE_VERSION,
        "pending_reviews": [],
        "unresolved_pdfs": [],
        "decisions": [],
    }


class ApprovalReviewRepository:
    def __init__(self, root: Path) -> None:
        self.root = root.resolve()
        self.path = self.root / APPROVAL_REVIEW_DATA_FILE_NAME
        self.legacy_path = (
            self.root / LEGACY_APPROVAL_REVIEW_DATA_FILE_NAME
        )

    def load(self) -> dict[str, Any]:
        source = self.path if self.path.is_file() else self.legacy_path
        if not source.is_file():
            return _empty_review(self.root)
        data = json.loads(source.read_text(encoding="utf-8"))
        schema_version = int(data.get("schema_version") or 0)
        if schema_version not in {1, APPROVAL_REVIEW_SCHEMA_VERSION}:
            raise ValueError(
                "不支持的审批 PDF 审核数据版本: "
                f"{data.get('schema_version')}"
            )
        data["schema_version"] = APPROVAL_REVIEW_SCHEMA_VERSION
        return data

    def save(self, data: dict[str, Any]) -> Path:
        data["schema_version"] = APPROVAL_REVIEW_SCHEMA_VERSION
        data["data_root"] = str(self.root)
        data["updated_at"] = _now()
        atomic_replace_text(
            self.path,
            json.dumps(data, ensure_ascii=False, indent=2),
        )
        _archive_legacy_artifact(self.root, self.legacy_path)
        return self.path


def _archive_legacy_artifact(root: Path, path: Path) -> None:
    if not path.is_file():
        return
    archive_dir = root / INTERNAL_DIR_NAME / LEGACY_DIR_NAME
    archive_dir.mkdir(parents=True, exist_ok=True)
    target = archive_dir / path.name
    if target.exists():
        target = archive_dir / (
            f"{path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            f"{path.suffix}"
        )
    shutil.move(str(path), str(target))


def _date_from_text(value: str) -> str:
    return legacy.normalize_date_for_pdf_match(value)


def _case_values(application: dict[str, Any]) -> dict[str, str]:
    business = application.get("application") or {}
    case_name = str(application.get("case_name") or "")
    name_match = CASE_NAME_RE.match(case_name)
    return {
        "area": str(business.get("施工区域") or ""),
        "start": (
            _date_from_text(str(business.get("施工开始时间") or ""))
            or (name_match.group("date") if name_match else "")
        ),
        "end": _date_from_text(str(business.get("施工结束时间") or "")),
        "content": (
            str(business.get("施工内容") or "")
            or (name_match.group("content") if name_match else "")
        ),
    }


def _between(text: str, start: str, ends: tuple[str, ...]) -> str:
    start_index = text.find(start)
    if start_index < 0:
        return ""
    value_start = start_index + len(start)
    value_end = len(text)
    for end in ends:
        index = text.find(end, value_start)
        if index >= 0:
            value_end = min(value_end, index)
    return text[value_start:value_end].strip("~:：-—")


def _pdf_values(text: str) -> dict[str, str]:
    normalized = legacy.normalize_match_text(text)
    construction_segment = _between(
        normalized,
        "施工开始时间",
        ("时长", "施工内容"),
    )
    construction_dates = [
        legacy.date_match_groups_to_iso(match.groups())
        for match in legacy.PDF_DATE_TOKEN_RE.finditer(construction_segment)
    ]
    construction_dates = [value for value in construction_dates if value]
    start = construction_dates[0] if construction_dates else ""
    end = construction_dates[1] if len(construction_dates) > 1 else ""
    if not start:
        match = legacy.PDF_START_DATE_RE.search(normalized)
        start = (
            legacy.date_match_groups_to_iso(match.groups()) if match else ""
        )
    if not end:
        match = legacy.PDF_END_DATE_RE.search(normalized)
        end = legacy.date_match_groups_to_iso(match.groups()) if match else ""
    return {
        "area": _between(
            normalized,
            "施工区域",
            ("施工开始时间", "施工开始日期"),
        ),
        "start": start,
        "end": end,
        "content": _between(
            normalized,
            "施工内容",
            (
                "施工负责人",
                "影响改动消防",
                "影响堵塞",
                "危险作业",
                "附件",
                "审批记录",
            ),
        ),
    }


def _date_delta(first: str, second: str) -> int | None:
    try:
        return abs((date.fromisoformat(first) - date.fromisoformat(second)).days)
    except (TypeError, ValueError):
        return None


def _strict_non_date_match(
    pdf_text: str,
    case_values: dict[str, str],
) -> tuple[bool, bool]:
    normalized_pdf = legacy.normalize_match_text(pdf_text)
    area_key = legacy.normalize_match_text(case_values["area"])
    content_key = legacy.normalize_match_text(case_values["content"])
    if not area_key or not content_key:
        return False, False

    content_keys = [content_key]
    if content_key.startswith(area_key):
        shortened = content_key[len(area_key) :]
        if shortened:
            content_keys.append(shortened)
    return (
        area_key in normalized_pdf,
        any(key in normalized_pdf for key in content_keys),
    )


def _date_points(delta: int | None, maximum: int) -> int:
    if delta is None:
        return 0
    if delta == 0:
        return maximum
    if delta <= 3:
        return maximum - 2
    if delta <= 7:
        return max(0, maximum - 4)
    if delta <= 14:
        return max(0, maximum - 7)
    if delta <= 31:
        return max(0, maximum - 10)
    return 0


def _candidate_confidence(
    pdf_values: dict[str, str],
    case_values: dict[str, str],
) -> tuple[int, list[str]]:
    evidence: list[str] = []
    start_delta = _date_delta(case_values["start"], pdf_values["start"])
    end_delta = _date_delta(case_values["end"], pdf_values["end"])
    confidence = (
        NON_DATE_CONFIDENCE
        + _date_points(start_delta, 12)
        + _date_points(end_delta, 8)
    )
    evidence.extend(["施工内容严格命中", "施工区域严格命中"])
    if start_delta is None:
        evidence.append("开始日期缺失，不参与淘汰")
    else:
        evidence.append(f"开始日期相差{start_delta}天")
    if end_delta is None:
        evidence.append("结束日期缺失，不参与淘汰")
    else:
        evidence.append(f"结束日期相差{end_delta}天")
    return min(100, confidence), evidence


def _confidence_level(confidence: int) -> str:
    if confidence >= 95:
        return "高"
    if confidence >= 85:
        return "中"
    return "低"


def _selection_confidence(
    candidate_confidence: int,
    runner_up_confidence: int | None,
) -> int:
    if runner_up_confidence is None:
        return candidate_confidence
    gap = max(0, candidate_confidence - runner_up_confidence)
    ambiguity_limit = 60 + min(40, gap * 3)
    return min(candidate_confidence, ambiguity_limit)


def _review_id(pdf_hash: str, case_id: str) -> str:
    return str(
        uuid.uuid5(
            CASE_NAMESPACE,
            f"approval-review:{pdf_hash}:{case_id}",
        )
    )


def _unresolved_review_id(pdf_hash: str) -> str:
    return str(
        uuid.uuid5(
            CASE_NAMESPACE,
            f"approval-review-unresolved:{pdf_hash}",
        )
    )


def _direct_inbox_pdf(root: Path, item: dict[str, Any]) -> Path | None:
    try:
        path = ensure_within(
            root / Path(str(item.get("path") or "")),
            root,
        )
    except ValueError:
        return None
    inbox = (root / INBOX_DIR_NAME).resolve()
    if path.parent != inbox or path.suffix.lower() != ".pdf":
        return None
    return path


def build_approval_review(
    dataset: dict[str, Any],
    root: Path,
    repo_root: Path,
    existing: dict[str, Any] | None = None,
) -> dict[str, Any]:
    root = root.resolve()
    previous = existing or _empty_review(root)
    decisions = list(previous.get("decisions") or [])
    previous_pending = {
        str(item.get("review_id") or ""): item
        for item in previous.get("pending_reviews") or []
    }
    previous_unresolved = {
        str((item.get("pdf") or {}).get("sha256") or ""): item
        for item in previous.get("unresolved_pdfs") or []
    }
    excluded_ids = {
        str(item.get("review_id") or "")
        for item in decisions
        if item.get("decision") == "排除"
    }
    applications = [
        application
        for application in dataset.get("applications") or []
        if (
            application.get("status") != "terminated"
            and not (application.get("approval") or {}).get("pdfs")
        )
    ]
    pending: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []
    seen_pdf_hashes: set[str] = set()
    with RecognitionService(dataset, repo_root) as recognition:
        for pdf_item in dataset.get("unmatched_files") or []:
            if pdf_item.get("role") != APPROVAL_PDF_ROLE:
                continue
            pdf_path = _direct_inbox_pdf(root, pdf_item)
            if pdf_path is None or not pdf_path.is_file():
                continue
            pdf_hash = sha256_file(pdf_path)
            if pdf_hash in seen_pdf_hashes:
                continue
            seen_pdf_hashes.add(pdf_hash)
            text = recognition.pdf_text(pdf_path)
            values = _pdf_values(text)
            application_no = (
                legacy.extract_pdf_application_no_from_name(pdf_path.name)
                or legacy.extract_pdf_rename_application_no(text)
            )
            recognition_method = str(
                (
                    (dataset.get("recognition_cache") or {})
                    .get(pdf_hash, {})
                    .get("method", "unknown")
                )
            )
            pdf_payload = {
                "path": str(pdf_item.get("path") or ""),
                "file_name": pdf_path.name,
                "sha256": pdf_hash,
                "application_no": application_no,
                "recognition_method": recognition_method,
                **values,
            }
            ranked: list[
                tuple[
                    int,
                    str,
                    dict[str, Any],
                    dict[str, str],
                    list[str],
                    str,
                ]
            ] = []
            any_area_match = False
            any_content_match = False
            strict_matches = 0
            for application in applications:
                case_values = _case_values(application)
                area_match, content_match = _strict_non_date_match(
                    text,
                    case_values,
                )
                any_area_match = any_area_match or area_match
                any_content_match = any_content_match or content_match
                if not (area_match and content_match):
                    continue
                strict_matches += 1
                confidence, evidence = _candidate_confidence(
                    values,
                    case_values,
                )
                case_id = str(application.get("case_id") or "")
                review_id = _review_id(pdf_hash, case_id)
                if review_id in excluded_ids:
                    continue
                ranked.append(
                    (
                        confidence,
                        str(application.get("case_name") or ""),
                        application,
                        case_values,
                        evidence,
                        review_id,
                    )
                )
            ranked.sort(key=lambda item: (-item[0], item[1]))
            if not ranked:
                if strict_matches:
                    reason = "所有严格候选均已被人工排除"
                elif not applications:
                    reason = "没有可匹配的未审批案卷"
                elif not any_content_match:
                    reason = "施工内容没有严格文本命中"
                elif not any_area_match:
                    reason = "施工区域没有严格文本命中"
                else:
                    reason = "施工区域和施工内容未在同一案卷同时严格命中"
                unresolved.append(
                    {
                        "review_id": _unresolved_review_id(pdf_hash),
                        "decision": str(
                            previous_unresolved.get(pdf_hash, {}).get(
                                "decision", "待审核"
                            )
                        ),
                        "review_note": str(
                            previous_unresolved.get(pdf_hash, {}).get(
                                "review_note", ""
                            )
                        ),
                        "status": "无严格候选",
                        "reason": reason,
                        "pdf": pdf_payload,
                    }
                )
                continue

            (
                candidate_confidence,
                _case_name,
                application,
                case_values,
                evidence,
                review_id,
            ) = ranked[0]
            runner_up_confidence = ranked[1][0] if len(ranked) > 1 else None
            confidence_gap = (
                candidate_confidence - runner_up_confidence
                if runner_up_confidence is not None
                else None
            )
            confidence = _selection_confidence(
                candidate_confidence,
                runner_up_confidence,
            )
            case_id = str(application.get("case_id") or "")
            ambiguity_evidence = (
                [
                    f"候选自身匹配度{candidate_confidence}%",
                    f"次高匹配度{runner_up_confidence}%",
                    f"领先{confidence_gap}个百分点",
                    f"综合选择置信度{confidence}%",
                ]
                if runner_up_confidence is not None
                else [f"唯一严格候选，置信度{confidence}%"]
            )
            pending.append(
                {
                    "review_id": review_id,
                    "decision": str(
                        previous_pending.get(review_id, {}).get(
                            "decision", "待审核"
                        )
                    ),
                    "review_note": str(
                        previous_pending.get(review_id, {}).get(
                            "review_note", ""
                        )
                    ),
                    "confidence": confidence,
                    "candidate_match_confidence": candidate_confidence,
                    "confidence_level": _confidence_level(confidence),
                    "strict_candidate_count": len(ranked),
                    "runner_up_confidence": runner_up_confidence,
                    "confidence_gap": confidence_gap,
                    "matching_evidence": "；".join(
                        [
                            *evidence,
                            *ambiguity_evidence,
                            f"严格候选{len(ranked)}个，仅展示最高项",
                        ]
                    ),
                    "pdf": pdf_payload,
                    "case": {
                        "case_id": case_id,
                        "case_name": str(application.get("case_name") or ""),
                        "case_directory": str(
                            application.get("case_directory") or ""
                        ),
                        "status": str(application.get("status") or ""),
                        **case_values,
                    },
                }
            )
    return {
        "schema_version": APPROVAL_REVIEW_SCHEMA_VERSION,
        "data_root": str(root),
        "source_dataset_revision": int(
            dataset.get("dataset_revision") or 0
        ),
        "generated_at": _now(),
        "match_rule_version": MATCH_RULE_VERSION,
        "pending_reviews": pending,
        "unresolved_pdfs": unresolved,
        "decisions": decisions,
    }


def import_json_decisions(
    review: dict[str, Any],
) -> list[dict[str, Any]]:
    imported: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    review_items = [
        *(review.get("pending_reviews") or []),
        *(review.get("unresolved_pdfs") or []),
    ]
    unresolved_ids = {
        str(item.get("review_id") or "")
        for item in review.get("unresolved_pdfs") or []
    }
    for item in review_items:
        review_id = str(item.get("review_id") or "").strip()
        if not review_id:
            raise ValueError("审核 JSON 中存在缺少审核ID的候选")
        if review_id in seen_ids:
            raise ValueError(f"审核 JSON 中审核ID重复: {review_id}")
        seen_ids.add(review_id)
        decision = str(item.get("decision") or "待审核").strip()
        if decision not in DECISION_OPTIONS:
            raise ValueError(
                f"审核 JSON 中审核结果无效: {review_id}={decision}"
            )
        if (
            review_id in unresolved_ids
            and decision not in {"待审核", "移至_trash"}
        ):
            raise ValueError(
                f"无严格候选 PDF 只允许待审核或移至_trash: {review_id}"
            )
        if decision != "待审核":
            imported.append(
                {
                    **item,
                    "decision": decision,
                    "review_note": str(
                        item.get("review_note") or ""
                    ).strip(),
                }
            )
    return imported


def apply_review_decisions(
    dataset: dict[str, Any],
    review: dict[str, Any],
    decisions: list[dict[str, Any]],
    root: Path,
) -> dict[str, int]:
    confirmed = [
        item for item in decisions if item.get("decision") == "确认匹配"
    ]
    trashed = [
        item for item in decisions if item.get("decision") == "移至_trash"
    ]
    counts: dict[str, int] = {}
    selected_cases: set[str] = set()
    for item in confirmed:
        pdf_hash = str((item.get("pdf") or {}).get("sha256") or "")
        counts[pdf_hash] = counts.get(pdf_hash, 0) + 1
        case_id = str((item.get("case") or {}).get("case_id") or "")
        if case_id in selected_cases:
            raise ValueError(f"同一案卷被确认了多个审批 PDF: {case_id}")
        selected_cases.add(case_id)
    duplicates = [pdf_hash for pdf_hash, count in counts.items() if count > 1]
    if duplicates:
        raise ValueError(
            "同一审批 PDF 只能确认一个案卷: " + "，".join(duplicates)
        )

    applications = {
        str(item.get("case_id") or ""): item
        for item in dataset.get("applications") or []
        if item.get("status") != "terminated"
    }
    unmatched = {
        str(item.get("sha256") or ""): item
        for item in dataset.get("unmatched_files") or []
        if (
            item.get("role") == APPROVAL_PDF_ROLE
            and _direct_inbox_pdf(root.resolve(), item) is not None
        )
    }
    for item in [*confirmed, *trashed]:
        pdf_hash = str((item.get("pdf") or {}).get("sha256") or "")
        case_id = str((item.get("case") or {}).get("case_id") or "")
        if pdf_hash not in unmatched:
            raise ValueError(f"审批 PDF 已不在待匹配列表: {pdf_hash}")
        if item.get("decision") == "确认匹配" and case_id not in applications:
            raise ValueError(f"案卷已不存在或已终止: {case_id}")
        path = ensure_within(
            root / Path(str(unmatched[pdf_hash].get("path") or "")),
            root,
        )
        if not path.is_file() or sha256_file(path) != pdf_hash:
            raise RuntimeError(f"审批 PDF 文件不存在或内容已变化: {path}")

    decided_at = _now()
    decision_history = review.setdefault("decisions", [])
    existing_history = {
        (
            str(item.get("review_id") or ""),
            str(item.get("decision") or ""),
        )
        for item in decision_history
    }
    for item in decisions:
        pdf = item.get("pdf") or {}
        case = item.get("case") or {}
        history_item = {
            "review_id": item.get("review_id", ""),
            "decision": item.get("decision", ""),
            "review_note": item.get("review_note", ""),
            "pdf_file_name": pdf.get("file_name", ""),
            "pdf_sha256": pdf.get("sha256", ""),
            "case_name": case.get("case_name", ""),
            "case_id": case.get("case_id", ""),
            "decided_at": decided_at,
        }
        history_key = (
            str(history_item["review_id"]),
            str(history_item["decision"]),
        )
        if history_key not in existing_history:
            decision_history.append(history_item)
            existing_history.add(history_key)

    applied_hashes: set[str] = set()
    for item in confirmed:
        pdf_hash = str((item.get("pdf") or {}).get("sha256") or "")
        case_id = str((item.get("case") or {}).get("case_id") or "")
        archive_reviewed_approval_pdf(
            dataset,
            root,
            unmatched[pdf_hash],
            applications[case_id],
            str(item.get("review_id") or ""),
            str(item.get("review_note") or ""),
        )
        applied_hashes.add(pdf_hash)

    trashed_hashes: set[str] = set()
    for item in trashed:
        pdf_hash = str((item.get("pdf") or {}).get("sha256") or "")
        unmatched_item = unmatched[pdf_hash]
        source = ensure_within(
            root / Path(str(unmatched_item.get("path") or "")),
            root,
        )
        trash_root = ensure_within(root / TRASH_DIR_NAME, root)
        trash_root.mkdir(parents=True, exist_ok=True)
        target = trash_root / source.name
        index = 2
        while target.exists():
            target = (
                trash_root
                / f"{source.stem}_{index:02d}{source.suffix}"
            )
            index += 1
        shutil.move(str(source), str(target))
        if sha256_file(target) != pdf_hash:
            raise RuntimeError(f"PDF 移入 _trash 后哈希校验失败: {target}")
        dataset.setdefault("changes", []).append(
            {
                "action": "move_to_trash_after_human_review",
                "source": str(source),
                "target": str(target),
                "role": APPROVAL_PDF_ROLE,
                "case_id": "",
                "result": "completed",
                "review_id": str(item.get("review_id") or ""),
                "at": decided_at,
            }
        )
        for history_item in reversed(decision_history):
            if (
                history_item.get("review_id") == item.get("review_id")
                and history_item.get("decision") == "移至_trash"
            ):
                history_item["result_path"] = str(
                    target.relative_to(root).as_posix()
                )
                break
        trashed_hashes.add(pdf_hash)

    if trashed_hashes:
        dataset["unmatched_files"] = [
            item
            for item in dataset.get("unmatched_files") or []
            if str(item.get("sha256") or "") not in trashed_hashes
        ]

    excluded_ids = {
        str(item.get("review_id") or "")
        for item in decisions
        if item.get("decision") in {"排除", "移至_trash"}
    }
    review["pending_reviews"] = [
        item
        for item in review.get("pending_reviews") or []
        if (
            str((item.get("pdf") or {}).get("sha256") or "")
            not in {*applied_hashes, *trashed_hashes}
            and str(item.get("review_id") or "") not in excluded_ids
        )
    ]
    review["last_applied_at"] = decided_at
    return {
        "confirmed": len(confirmed),
        "trashed": len(trashed),
    }
