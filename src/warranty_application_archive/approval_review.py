from __future__ import annotations

import json
import os
import re
import uuid
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from . import legacy
from .constants import (
    APPROVAL_PDF_ROLE,
    APPROVAL_REVIEW_DATA_FILE_NAME,
    APPROVAL_REVIEW_EXCEL_FILE_NAME,
    APPROVAL_REVIEW_SCHEMA_VERSION,
    INBOX_DIR_NAME,
)
from .file_utils import atomic_replace_text, ensure_within, sha256_file
from .migration import CASE_NAMESPACE
from .recognition import RecognitionService
from .workflows import archive_reviewed_approval_pdf


PENDING_HEADERS = [
    "审核结果",
    "人工备注",
    "候选评分",
    "匹配依据",
    "审批PDF文件",
    "审批编号",
    "审批施工区域",
    "审批施工开始",
    "审批施工结束",
    "审批施工内容",
    "候选案卷",
    "案卷状态",
    "案卷施工区域",
    "案卷施工开始",
    "案卷施工结束",
    "案卷施工内容",
    "PDF链接",
    "案卷目录",
    "审核ID",
    "PDF SHA-256",
    "案卷ID",
]
DECISION_OPTIONS = ("待审核", "确认匹配", "排除")
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
        "pending_reviews": [],
        "decisions": [],
    }


class ApprovalReviewRepository:
    def __init__(self, root: Path) -> None:
        self.root = root.resolve()
        self.path = self.root / APPROVAL_REVIEW_DATA_FILE_NAME

    def load(self) -> dict[str, Any]:
        if not self.path.is_file():
            return _empty_review(self.root)
        data = json.loads(self.path.read_text(encoding="utf-8"))
        if data.get("schema_version") != APPROVAL_REVIEW_SCHEMA_VERSION:
            raise ValueError(
                "不支持的审批 PDF 审核数据版本: "
                f"{data.get('schema_version')}"
            )
        return data

    def save(self, data: dict[str, Any]) -> Path:
        data["schema_version"] = APPROVAL_REVIEW_SCHEMA_VERSION
        data["data_root"] = str(self.root)
        data["updated_at"] = _now()
        atomic_replace_text(
            self.path,
            json.dumps(data, ensure_ascii=False, indent=2),
        )
        return self.path


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


def _candidate_score(
    pdf_text: str,
    pdf_values: dict[str, str],
    case_values: dict[str, str],
) -> tuple[int, list[str]]:
    score = 0
    evidence: list[str] = []
    normalized_pdf = legacy.normalize_match_text(pdf_text)
    case_content = legacy.normalize_match_text(case_values["content"])
    pdf_content = legacy.normalize_match_text(pdf_values["content"])
    if case_content and case_content in normalized_pdf:
        score += 60
        evidence.append("施工内容文字命中")
    elif case_content and pdf_content:
        ratio = SequenceMatcher(None, case_content, pdf_content).ratio()
        if ratio >= 0.45:
            points = round(ratio * 40)
            score += points
            evidence.append(f"施工内容相似{ratio:.0%}")

    case_area = legacy.normalize_match_text(case_values["area"])
    pdf_area = legacy.normalize_match_text(pdf_values["area"])
    if case_area and case_area in normalized_pdf:
        score += 15
        evidence.append("施工区域命中")
    elif case_area and pdf_area:
        ratio = SequenceMatcher(None, case_area, pdf_area).ratio()
        if ratio >= 0.6:
            score += 8
            evidence.append(f"施工区域相似{ratio:.0%}")

    start_delta = _date_delta(case_values["start"], pdf_values["start"])
    if start_delta is not None:
        if start_delta == 0:
            score += 25
        elif start_delta <= 3:
            score += 20
        elif start_delta <= 7:
            score += 12
        elif start_delta <= 14:
            score += 5
        evidence.append(f"开始日期相差{start_delta}天")

    end_delta = _date_delta(case_values["end"], pdf_values["end"])
    if end_delta is not None:
        if end_delta == 0:
            score += 15
        elif end_delta <= 3:
            score += 10
        elif end_delta <= 7:
            score += 5
        evidence.append(f"结束日期相差{end_delta}天")
    return score, evidence


def _review_id(pdf_hash: str, case_id: str) -> str:
    return str(
        uuid.uuid5(
            CASE_NAMESPACE,
            f"approval-review:{pdf_hash}:{case_id}",
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
    excluded_ids = {
        str(item.get("review_id") or "")
        for item in decisions
        if item.get("decision") == "排除"
    }
    applications = [
        application
        for application in dataset.get("applications") or []
        if not (application.get("approval") or {}).get("pdfs")
    ]
    pending: list[dict[str, Any]] = []
    with RecognitionService(dataset, repo_root) as recognition:
        for pdf_item in dataset.get("unmatched_files") or []:
            if pdf_item.get("role") != APPROVAL_PDF_ROLE:
                continue
            pdf_path = _direct_inbox_pdf(root, pdf_item)
            if pdf_path is None or not pdf_path.is_file():
                continue
            pdf_hash = sha256_file(pdf_path)
            text = recognition.pdf_text(pdf_path)
            values = _pdf_values(text)
            application_no = (
                legacy.extract_pdf_application_no_from_name(pdf_path.name)
                or legacy.extract_pdf_rename_application_no(text)
            )
            ranked: list[tuple[int, str, dict[str, Any], dict[str, str], list[str]]] = []
            for application in applications:
                case_values = _case_values(application)
                score, evidence = _candidate_score(text, values, case_values)
                ranked.append(
                    (
                        score,
                        str(application.get("case_name") or ""),
                        application,
                        case_values,
                        evidence,
                    )
                )
            ranked.sort(key=lambda item: (-item[0], item[1]))
            for score, _case_name, application, case_values, evidence in ranked:
                case_id = str(application.get("case_id") or "")
                review_id = _review_id(pdf_hash, case_id)
                if review_id in excluded_ids:
                    continue
                pending.append(
                    {
                        "review_id": review_id,
                        "decision": "待审核",
                        "review_note": "",
                        "candidate_score": score,
                        "matching_evidence": "；".join(evidence) or "无自动命中依据",
                        "pdf": {
                            "path": str(pdf_item.get("path") or ""),
                            "file_name": pdf_path.name,
                            "sha256": pdf_hash,
                            "application_no": application_no,
                            **values,
                        },
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
        "pending_reviews": pending,
        "decisions": decisions,
    }


def _pending_row(item: dict[str, Any]) -> list[Any]:
    pdf = item.get("pdf") or {}
    case = item.get("case") or {}
    return [
        item.get("decision", "待审核"),
        item.get("review_note", ""),
        item.get("candidate_score", 0),
        item.get("matching_evidence", ""),
        pdf.get("file_name", ""),
        pdf.get("application_no", ""),
        pdf.get("area", ""),
        pdf.get("start", ""),
        pdf.get("end", ""),
        pdf.get("content", ""),
        case.get("case_name", ""),
        case.get("status", ""),
        case.get("area", ""),
        case.get("start", ""),
        case.get("end", ""),
        case.get("content", ""),
        "打开审批PDF",
        "打开案卷目录",
        item.get("review_id", ""),
        pdf.get("sha256", ""),
        case.get("case_id", ""),
    ]


def _format_table_sheet(sheet: Any, rows: int, columns: int) -> None:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = (
        f"A1:{get_column_letter(columns)}{max(2, rows + 1)}"
    )
    header_fill = PatternFill("solid", fgColor="44546A")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 28


def export_approval_review_excel(
    review: dict[str, Any], root: Path
) -> Path:
    root = root.resolve()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "待审核"
    sheet.append(PENDING_HEADERS)
    pending = list(review.get("pending_reviews") or [])
    for item in pending:
        sheet.append(_pending_row(item))
    _format_table_sheet(sheet, len(pending), len(PENDING_HEADERS))
    if pending:
        table = Table(
            displayName="ApprovalReviewCandidates",
            ref=f"A1:U{len(pending) + 1}",
        )
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showFirstColumn=False,
            showLastColumn=False,
            showColumnStripes=False,
        )
        sheet.add_table(table)
        validation = DataValidation(
            type="list",
            formula1='"待审核,确认匹配,排除"',
            allow_blank=False,
        )
        validation.error = "请选择：待审核、确认匹配或排除"
        validation.errorTitle = "审核结果无效"
        sheet.add_data_validation(validation)
        validation.add(f"A2:A{len(pending) + 1}")
        sheet.conditional_formatting.add(
            f"A2:A{len(pending) + 1}",
            FormulaRule(
                formula=['$A2="确认匹配"'],
                fill=PatternFill("solid", fgColor="C6EFCE"),
            ),
        )
        sheet.conditional_formatting.add(
            f"A2:A{len(pending) + 1}",
            FormulaRule(
                formula=['$A2="排除"'],
                fill=PatternFill("solid", fgColor="FFC7CE"),
            ),
        )
    editable_fill = PatternFill("solid", fgColor="FFF2CC")
    for row_index in range(2, len(pending) + 2):
        for column_index in (1, 2):
            sheet.cell(row=row_index, column=column_index).fill = editable_fill
            sheet.cell(row=row_index, column=column_index).protection = Protection(
                locked=False
            )
        item = pending[row_index - 2]
        pdf_path = root / Path(str((item.get("pdf") or {}).get("path") or ""))
        case_path = root / Path(
            str((item.get("case") or {}).get("case_directory") or "")
        )
        sheet.cell(row=row_index, column=17).hyperlink = str(pdf_path.resolve())
        sheet.cell(row=row_index, column=18).hyperlink = str(case_path.resolve())
        for column_index in (17, 18):
            sheet.cell(row=row_index, column=column_index).font = Font(
                color="0563C1", underline="single"
            )
    widths = [
        12, 28, 10, 28, 38, 16, 22, 13, 13, 32, 42,
        16, 22, 13, 13, 32, 16, 16, 38, 68, 38,
    ]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for column_index in (19, 20, 21):
        sheet.column_dimensions[get_column_letter(column_index)].hidden = True

    history = workbook.create_sheet("已处理决定")
    history_headers = [
        "处理时间",
        "审核结果",
        "人工备注",
        "审批PDF文件",
        "PDF SHA-256",
        "候选案卷",
        "案卷ID",
        "审核ID",
    ]
    history.append(history_headers)
    for item in review.get("decisions") or []:
        history.append(
            [
                item.get("decided_at", ""),
                item.get("decision", ""),
                item.get("review_note", ""),
                item.get("pdf_file_name", ""),
                item.get("pdf_sha256", ""),
                item.get("case_name", ""),
                item.get("case_id", ""),
                item.get("review_id", ""),
            ]
        )
    _format_table_sheet(
        history,
        len(review.get("decisions") or []),
        len(history_headers),
    )
    for column, width in zip(
        "ABCDEFGH", (22, 12, 32, 42, 68, 42, 38, 38)
    ):
        history.column_dimensions[column].width = width

    notes = workbook.create_sheet("说明")
    notes.sheet_view.showGridLines = False
    instructions = [
        ["审批 PDF 人工匹配审核", ""],
        ["1", "同一个审批 PDF 会列出多条候选案卷，并按候选评分从高到低排列。"],
        ["2", "只修改黄色的“审核结果”和“人工备注”列。"],
        ["3", "每个审批 PDF 最多只能有一行选择“确认匹配”；不可能的候选可选“排除”。"],
        [
            "4",
            "保存并关闭 Excel 后，运行："
            "python warranty_application_archive.py apply-approval-review",
        ],
        ["5", "命令先保存审核决定，再更新正式 JSON，并从正式 JSON 重建正式汇总 Excel。"],
        ["正式数据版本", review.get("source_dataset_revision", 0)],
        ["生成时间", review.get("generated_at", "")],
        ["待审核候选关系", len(pending)],
    ]
    for row in instructions:
        notes.append(row)
    notes.merge_cells("A1:B1")
    notes["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    notes["A1"].fill = PatternFill("solid", fgColor="44546A")
    notes["A1"].alignment = Alignment(horizontal="center")
    notes.column_dimensions["A"].width = 18
    notes.column_dimensions["B"].width = 100
    for row in notes.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    output = root / APPROVAL_REVIEW_EXCEL_FILE_NAME
    temporary = root / f".{APPROVAL_REVIEW_EXCEL_FILE_NAME}.tmp.xlsx"
    workbook.save(temporary)
    try:
        os.replace(temporary, output)
    except PermissionError as exc:
        temporary.unlink(missing_ok=True)
        raise PermissionError(
            f"审核 Excel 正在使用，请先关闭后重试: {output}"
        ) from exc
    return output


def import_excel_decisions(
    review: dict[str, Any], root: Path
) -> list[dict[str, Any]]:
    path = root.resolve() / APPROVAL_REVIEW_EXCEL_FILE_NAME
    if not path.is_file():
        raise FileNotFoundError(f"审批 PDF 审核 Excel 不存在: {path}")
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        if "待审核" not in workbook.sheetnames:
            raise ValueError("审核 Excel 缺少“待审核”工作表")
        sheet = workbook["待审核"]
        headers = {
            str(cell.value or "").strip(): cell.column
            for cell in sheet[1]
        }
        missing = [header for header in PENDING_HEADERS if header not in headers]
        if missing:
            raise ValueError(f"审核 Excel 缺少列: {', '.join(missing)}")
        expected = {
            str(item.get("review_id") or ""): item
            for item in review.get("pending_reviews") or []
        }
        imported: list[dict[str, Any]] = []
        seen_ids: set[str] = set()
        for row_index in range(2, sheet.max_row + 1):
            review_id = str(
                sheet.cell(row_index, headers["审核ID"]).value or ""
            ).strip()
            if not review_id:
                continue
            if review_id in seen_ids:
                raise ValueError(f"审核 Excel 中审核ID重复: {review_id}")
            seen_ids.add(review_id)
            item = expected.get(review_id)
            if item is None:
                raise ValueError(
                    f"审核 Excel 含有过期或未知审核ID: {review_id}"
                )
            decision = str(
                sheet.cell(row_index, headers["审核结果"]).value or ""
            ).strip()
            if decision not in DECISION_OPTIONS:
                raise ValueError(
                    f"第 {row_index} 行审核结果无效: {decision}"
                )
            pdf_hash = str(
                sheet.cell(row_index, headers["PDF SHA-256"]).value or ""
            ).strip()
            case_id = str(
                sheet.cell(row_index, headers["案卷ID"]).value or ""
            ).strip()
            if pdf_hash != str((item.get("pdf") or {}).get("sha256") or ""):
                raise ValueError(f"第 {row_index} 行 PDF SHA-256 被修改")
            if case_id != str((item.get("case") or {}).get("case_id") or ""):
                raise ValueError(f"第 {row_index} 行案卷ID被修改")
            if decision != "待审核":
                imported.append(
                    {
                        **item,
                        "decision": decision,
                        "review_note": str(
                            sheet.cell(
                                row_index, headers["人工备注"]
                            ).value
                            or ""
                        ).strip(),
                    }
                )
        return imported
    finally:
        workbook.close()


def apply_review_decisions(
    dataset: dict[str, Any],
    review: dict[str, Any],
    decisions: list[dict[str, Any]],
    root: Path,
) -> int:
    confirmed = [
        item for item in decisions if item.get("decision") == "确认匹配"
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
    }
    unmatched = {
        str(item.get("sha256") or ""): item
        for item in dataset.get("unmatched_files") or []
        if (
            item.get("role") == APPROVAL_PDF_ROLE
            and _direct_inbox_pdf(root.resolve(), item) is not None
        )
    }
    for item in confirmed:
        pdf_hash = str((item.get("pdf") or {}).get("sha256") or "")
        case_id = str((item.get("case") or {}).get("case_id") or "")
        if pdf_hash not in unmatched:
            raise ValueError(f"审批 PDF 已不在待匹配列表: {pdf_hash}")
        if case_id not in applications:
            raise ValueError(f"案卷已不存在: {case_id}")
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

    excluded_ids = {
        str(item.get("review_id") or "")
        for item in decisions
        if item.get("decision") == "排除"
    }
    review["pending_reviews"] = [
        item
        for item in review.get("pending_reviews") or []
        if (
            str((item.get("pdf") or {}).get("sha256") or "")
            not in applied_hashes
            and str(item.get("review_id") or "") not in excluded_ids
        )
    ]
    review["last_applied_at"] = decided_at
    return len(confirmed)
