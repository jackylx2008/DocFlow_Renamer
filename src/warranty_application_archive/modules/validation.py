from __future__ import annotations

import json
from collections import Counter
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Iterable

from .constants import SUMMARY_HTML_FILE_NAME
from .file_utils import ensure_within, sha256_file


EXPECTED_SUMMARY_SHEETS = [
    "申请汇总",
    "待补材料",
    "待审批PDF",
    "已完成",
]


def _all_file_records(application: dict[str, Any]) -> Iterable[dict[str, Any]]:
    for files in (application.get("materials") or {}).values():
        yield from files or []
    yield from (application.get("approval") or {}).get("pdfs") or []


def validate_dataset(
    data: dict[str, Any],
    root: Path,
    verify_hashes: bool = True,
) -> dict[str, Any]:
    root = root.resolve()
    errors: list[str] = []
    warnings: list[str] = []
    case_ids: set[str] = set()
    file_ids: set[str] = set()
    file_count = 0
    statuses: Counter[str] = Counter()

    for application in data.get("applications") or []:
        case_id = str(application.get("case_id") or "")
        if not case_id:
            errors.append("存在缺少 case_id 的案卷")
        elif case_id in case_ids:
            errors.append(f"重复 case_id: {case_id}")
        case_ids.add(case_id)
        statuses[str(application.get("status") or "unknown")] += 1

        relative_case_dir = str(application.get("case_directory") or "")
        try:
            case_dir = ensure_within(root / Path(relative_case_dir), root)
        except ValueError:
            errors.append(f"案卷路径越界: {relative_case_dir}")
            continue
        if not case_dir.is_dir():
            errors.append(f"案卷目录不存在: {relative_case_dir}")

        for item in _all_file_records(application):
            file_count += 1
            file_id = str(item.get("file_id") or "")
            if not file_id:
                errors.append(f"文件记录缺少 file_id: {item.get('path')}")
            elif file_id in file_ids:
                errors.append(f"重复 file_id: {file_id}")
            file_ids.add(file_id)
            _validate_file_record(
                item,
                root,
                errors,
                warnings,
                verify_hashes,
            )

    for item in data.get("unmatched_files") or []:
        file_count += 1
        _validate_file_record(
            item,
            root,
            errors,
            warnings,
            verify_hashes,
        )

    return {
        "valid": not errors,
        "applications": len(data.get("applications") or []),
        "files": file_count,
        "statuses": dict(statuses),
        "errors": errors,
        "warnings": warnings,
    }


def _validate_file_record(
    item: dict[str, Any],
    root: Path,
    errors: list[str],
    warnings: list[str],
    verify_hashes: bool,
) -> None:
    relative_path = str(item.get("path") or "")
    try:
        path = ensure_within(root / Path(relative_path), root)
    except ValueError:
        errors.append(f"文件路径越界: {relative_path}")
        return
    if not path.is_file():
        errors.append(f"文件不存在: {relative_path}")
        return
    expected_size = item.get("size")
    if isinstance(expected_size, int) and path.stat().st_size != expected_size:
        errors.append(f"文件大小不一致: {relative_path}")
    if verify_hashes and item.get("sha256"):
        if sha256_file(path) != item["sha256"]:
            errors.append(f"文件哈希不一致: {relative_path}")
    elif not item.get("sha256"):
        warnings.append(f"文件缺少 SHA-256: {relative_path}")


class _SummaryDataParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_summary_data = False
        self.parts: list[str] = []

    def handle_starttag(
        self,
        tag: str,
        attrs: list[tuple[str, str | None]],
    ) -> None:
        attributes = dict(attrs)
        if tag == "script" and attributes.get("id") == "summaryData":
            self.in_summary_data = True

    def handle_endtag(self, tag: str) -> None:
        if tag == "script" and self.in_summary_data:
            self.in_summary_data = False

    def handle_data(self, data: str) -> None:
        if self.in_summary_data:
            self.parts.append(data)


def validate_summary_html(
    root: Path,
    expected_applications: int,
) -> dict[str, Any]:
    root = root.resolve()
    html_path = root / SUMMARY_HTML_FILE_NAME
    if not html_path.is_file():
        return {
            "valid": False,
            "path": str(html_path),
            "errors": ["汇总 HTML 文件不存在"],
        }
    errors: list[str] = []
    try:
        parser = _SummaryDataParser()
        parser.feed(html_path.read_text(encoding="utf-8"))
        if not parser.parts:
            raise ValueError("缺少 summaryData 数据块")
        view = json.loads("".join(parser.parts))
        sheets = list(view.get("sheets") or [])
        sheet_names = [str(item.get("title") or "") for item in sheets]
        if sheet_names != EXPECTED_SUMMARY_SHEETS:
            errors.append(
                f"HTML 页签不符合预期: {sheet_names}"
            )
        summary = next(
            (
                item
                for item in sheets
                if item.get("title") == "申请汇总"
            ),
            None,
        )
        if summary is None:
            errors.append("HTML 缺少申请汇总页签")
        else:
            actual_rows = len(summary.get("rows") or [])
            if actual_rows != expected_applications:
                errors.append(
                    f"申请汇总行数错误: {actual_rows}，预期 {expected_applications}"
                )
        for sheet in sheets:
            for row_index, row in enumerate(
                sheet.get("rows") or [],
                start=1,
            ):
                for cell in row:
                    for link in cell.get("links") or []:
                        relative = str(link.get("path") or "")
                        if not relative:
                            continue
                        try:
                            target = ensure_within(root / relative, root)
                        except ValueError:
                            errors.append(
                                f"HTML 链接路径越界: {relative}"
                            )
                            continue
                        if not target.exists():
                            errors.append(
                                "HTML 链接目标不存在: "
                                f"{sheet.get('title')} 第 {row_index} 行"
                                f" -> {relative}"
                            )
    except (ValueError, json.JSONDecodeError) as exc:
        errors.append(f"汇总 HTML 数据无效: {exc}")
    return {
        "valid": not errors,
        "path": str(html_path),
        "sheets": EXPECTED_SUMMARY_SHEETS,
        "errors": errors,
    }
