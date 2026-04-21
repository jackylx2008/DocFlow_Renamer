from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import yaml
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": WORD_NS}
CHECKED_SYMBOL = "0052"
TARGET_NAME_RE = re.compile(
    r"^\d{4}-\d{2}-\d{2}_.+_质保作业申请单\.docx$", re.IGNORECASE
)
DATE_RANGE_RE = re.compile(
    r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日\s*[~～\-至到]+\s*(?:(\d{4})年\s*)?(\d{1,2})月\s*(\d{1,2})日"
)
SINGLE_DATE_RE = re.compile(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日")
INVALID_FILENAME_CHARS_RE = re.compile(r'[\\/:*?"<>|]+')
WHITESPACE_RE = re.compile(r"\s+")
WORK_TYPES = ["动火作业", "有限空间作业", "5米以上高处作业", "危大工程", "配电室接电"]


@dataclass
class Record:
    项目名称: str
    质保单位: str
    分包单位: str
    质保负责人: str
    质保负责人联系电话: str
    施工区域: str
    施工开始时间: str
    施工结束时间: str
    时长天: int | None
    施工内容: str
    施工负责人: str
    施工负责人联系电话: str
    影响改动消防设备设施: str
    影响堵塞应急疏散通道: str
    危险作业: str
    原文件名: str
    新文件名: str
    文件路径: str
    申请单文件链接: str
    附件目录: str
    附件目录链接: str
    处理状态: str


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


def normalize_whitespace(value: str) -> str:
    return WHITESPACE_RE.sub(" ", value or "").strip()


def sanitize_filename_part(value: str) -> str:
    cleaned = INVALID_FILENAME_CHARS_RE.sub("-", value.strip())
    cleaned = normalize_whitespace(cleaned)
    return cleaned.strip(". ")


def read_first_form_text(docx_path: Path) -> str:
    document = Document(docx_path)
    if not document.tables:
        raise ValueError("Word 文档中未找到表格")
    table = document.tables[0]
    if not table.rows or not table.rows[0].cells:
        raise ValueError("Word 表格结构不符合预期")
    return table.rows[0].cells[0].text


def read_first_form_runs(docx_path: Path) -> list[dict[str, str]]:
    with ZipFile(docx_path) as archive:
        xml_bytes = archive.read("word/document.xml")

    root = ET.fromstring(xml_bytes)
    table = root.find(".//w:tbl", NS)
    if table is None:
        raise ValueError("document.xml 中未找到表格")

    row = table.find("./w:tr", NS)
    if row is None:
        raise ValueError("document.xml 中未找到第一行")

    cell = row.find("./w:tc", NS)
    if cell is None:
        raise ValueError("document.xml 中未找到第一列")

    runs: list[dict[str, str]] = []
    for paragraph in cell.findall(".//w:p", NS):
        for run in paragraph.findall("./w:r", NS):
            text_nodes = run.findall("./w:t", NS)
            text = "".join(node.text or "" for node in text_nodes)
            if text:
                runs.append({"type": "text", "value": text})

            sym = run.find("./w:sym", NS)
            if sym is not None:
                char_value = sym.attrib.get(f"{{{WORD_NS}}}char", "").upper()
                if char_value:
                    runs.append({"type": "sym", "value": char_value})
    return runs


def capture_between(text: str, start_label: str, end_label: str) -> str:
    pattern = re.compile(
        rf"{re.escape(start_label)}\s*(.*?)\s*(?={re.escape(end_label)})",
        re.DOTALL,
    )
    match = pattern.search(text)
    return normalize_whitespace(match.group(1)) if match else ""


def parse_date_range(raw_value: str) -> tuple[str, str, int | None]:
    raw_value = normalize_whitespace(raw_value)
    match = DATE_RANGE_RE.search(raw_value)
    if match:
        start_year, start_month, start_day, end_year, end_month, end_day = (
            match.groups()
        )
        start_date = date(int(start_year), int(start_month), int(start_day))
        end_date = date(int(end_year or start_year), int(end_month), int(end_day))
        duration = (end_date - start_date).days + 1
        return start_date.isoformat(), end_date.isoformat(), duration

    match = SINGLE_DATE_RE.search(raw_value)
    if not match:
        return "", "", None

    year, month, day = match.groups()
    only_date = date(int(year), int(month), int(day)).isoformat()
    return only_date, only_date, 1


def extract_checkbox_value(runs: list[dict[str, str]], label: str) -> str:
    for index, item in enumerate(runs):
        if item["type"] != "text" or label not in item["value"]:
            continue

        symbol_values: list[str] = []
        for candidate in runs[index + 1 :]:
            if candidate["type"] == "text" and any(
                marker in candidate["value"]
                for marker in (
                    "一、",
                    "二、",
                    "三、",
                    "会投工程部意见：",
                    "分公司工程部意见：",
                )
            ):
                break
            if candidate["type"] == "sym":
                symbol_values.append(candidate["value"])
                if len(symbol_values) == 2:
                    break

        if len(symbol_values) < 2:
            return ""
        if symbol_values[0] == CHECKED_SYMBOL:
            return "是"
        if symbol_values[1] == CHECKED_SYMBOL:
            return "否"
        return ""
    return ""


def next_symbol_value(runs: list[dict[str, str]], start_index: int) -> str:
    for candidate in runs[start_index + 1 :]:
        if candidate["type"] == "sym":
            return candidate["value"]
        if candidate["type"] == "text" and any(
            option in candidate["value"] for option in WORK_TYPES
        ):
            return ""
    return ""


def extract_dangerous_work(runs: list[dict[str, str]]) -> str:
    selected: list[str] = []
    cursor = 0

    for option in WORK_TYPES:
        while cursor < len(runs):
            item = runs[cursor]
            cursor += 1
            if item["type"] != "text" or option not in item["value"]:
                continue
            if next_symbol_value(runs, cursor - 1) == CHECKED_SYMBOL:
                selected.append(option)
            break

    return "、".join(selected)


def detect_attachment_dir(docx_path: Path) -> Path | None:
    candidates = [
        docx_path.with_name(f"{docx_path.stem}_附件"),
        docx_path.with_name(f"{docx_path.stem}附件"),
        docx_path.with_name(docx_path.stem),
    ]
    for candidate in candidates:
        if candidate.is_dir():
            return candidate
    return None


def build_target_name(start_date: str, work_content: str) -> str:
    if not start_date or not work_content:
        raise ValueError("缺少施工开始时间或施工内容，无法生成新文件名")
    return f"{start_date}_{sanitize_filename_part(work_content)}_质保作业申请单.docx"


def parse_document(docx_path: Path) -> dict[str, Any]:
    form_text = normalize_whitespace(read_first_form_text(docx_path))
    runs = read_first_form_runs(docx_path)

    date_range_raw = capture_between(form_text, "施工日期：", "施工内容：")
    start_date, end_date, duration_days = parse_date_range(date_range_raw)
    constructor_section = capture_between(
        form_text, "施工内容：", "一、影响、改动消防设备设施"
    )
    phone_match = re.search(r"联系电话：\s*(\d+)", constructor_section)

    return {
        "项目名称": capture_between(form_text, "项目名称：", "质保单位（盖章）："),
        "质保单位": capture_between(form_text, "质保单位（盖章）：", "分包单位："),
        "分包单位": capture_between(form_text, "分包单位：", "质保负责人："),
        "质保负责人": capture_between(form_text, "质保负责人：", "联系电话："),
        "质保负责人联系电话": capture_between(form_text, "联系电话：", "施工区域："),
        "施工区域": capture_between(form_text, "施工区域：", "施工日期："),
        "施工开始时间": start_date,
        "施工结束时间": end_date,
        "时长天": duration_days,
        "施工内容": capture_between(form_text, "施工内容：", "施工负责人："),
        "施工负责人": capture_between(form_text, "施工负责人：", "联系电话："),
        "施工负责人联系电话": phone_match.group(1) if phone_match else "",
        "影响改动消防设备设施": extract_checkbox_value(
            runs, "一、影响、改动消防设备设施"
        ),
        "影响堵塞应急疏散通道": extract_checkbox_value(
            runs, "二、影响、堵塞应急疏散通道"
        ),
        "危险作业": extract_dangerous_work(runs),
    }


def export_excel(records: list[Record], output_path: Path) -> None:
    workbook = Workbook()
    summary_sheet = workbook.active
    summary_sheet.title = "汇总"

    headers = [
        "项目名称",
        "质保单位",
        "分包单位",
        "质保负责人",
        "质保负责人联系电话",
        "施工区域",
        "施工开始时间",
        "施工结束时间",
        "时长（天）",
        "施工内容",
        "施工负责人",
        "施工负责人联系电话",
        "影响、改动消防设备设施",
        "影响、堵塞应急疏散通道",
        "危险作业",
        "申请单文件",
        "附件目录",
        "原文件名",
        "新文件名",
        "文件路径",
        "处理状态",
    ]
    summary_sheet.append(headers)

    for record in records:
        summary_sheet.append(
            [
                record.项目名称,
                record.质保单位,
                record.分包单位,
                record.质保负责人,
                record.质保负责人联系电话,
                record.施工区域,
                record.施工开始时间,
                record.施工结束时间,
                record.时长天,
                record.施工内容,
                record.施工负责人,
                record.施工负责人联系电话,
                record.影响改动消防设备设施,
                record.影响堵塞应急疏散通道,
                record.危险作业,
                "打开文件" if record.申请单文件链接 else "",
                "打开目录" if record.附件目录链接 else "",
                record.原文件名,
                record.新文件名,
                record.文件路径,
                record.处理状态,
            ]
        )

    apply_summary_styles(summary_sheet)
    fill_summary_hyperlinks(summary_sheet, records)

    note_sheet = workbook.create_sheet("说明")
    build_note_sheet(note_sheet, records, output_path.name)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        workbook.save(output_path)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_path = output_path.with_name(
            f"{output_path.stem}_{timestamp}{output_path.suffix}"
        )
        workbook.save(fallback_path)


def apply_summary_styles(summary_sheet: Any) -> None:
    thin_black = Side(style="thin", color="000000")
    border = Border(top=thin_black, bottom=thin_black, left=thin_black, right=thin_black)
    header_fill = PatternFill(fill_type="solid", fgColor="5B6F84")
    band_fill = PatternFill(fill_type="solid", fgColor="D9E5F3")
    plain_fill = PatternFill(fill_type=None)
    header_font = Font(bold=True, color="FFFFFF")
    hyperlink_font = Font(color="0563C1", underline="single")
    centered = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrapped = Alignment(vertical="center", wrap_text=True)

    column_widths = {
        1: 16,
        2: 16,
        3: 16,
        4: 12,
        5: 16,
        6: 20,
        7: 12,
        8: 12,
        9: 10,
        10: 24,
        11: 12,
        12: 16,
        13: 16,
        14: 18,
        15: 24,
        16: 12,
        17: 12,
        18: 28,
        19: 32,
        20: 40,
        21: 12,
    }
    for col_idx, width in column_widths.items():
        summary_sheet.column_dimensions[get_column_letter(col_idx)].width = width

    summary_sheet.freeze_panes = "A2"
    summary_sheet.auto_filter.ref = f"A1:{get_column_letter(summary_sheet.max_column)}{summary_sheet.max_row}"
    summary_sheet.sheet_view.showGridLines = False
    summary_sheet.row_dimensions[1].height = 24

    for cell in summary_sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = centered

    for row_idx, row in enumerate(
        summary_sheet.iter_rows(
        min_row=2,
        max_row=summary_sheet.max_row,
        min_col=1,
        max_col=summary_sheet.max_column,
        ),
        start=2,
    ):
        if not any(cell.value not in (None, "") for cell in row):
            continue
        fill = band_fill if row_idx % 2 == 0 else plain_fill
        for cell in row:
            cell.alignment = wrapped
            cell.border = border
            cell.fill = fill
        row[15].font = hyperlink_font
        row[16].font = hyperlink_font

    for row_idx in range(2, summary_sheet.max_row + 1):
        summary_sheet[f"G{row_idx}"].number_format = "yyyy-mm-dd"
        summary_sheet[f"H{row_idx}"].number_format = "yyyy-mm-dd"
        summary_sheet[f"I{row_idx}"].number_format = "0"
        summary_sheet.row_dimensions[row_idx].height = 42


def fill_summary_hyperlinks(summary_sheet: Any, records: list[Record]) -> None:
    for row_idx, record in enumerate(records, start=2):
        if record.申请单文件链接:
            summary_sheet[f"P{row_idx}"].hyperlink = record.申请单文件链接
        if record.附件目录链接:
            summary_sheet[f"Q{row_idx}"].hyperlink = record.附件目录链接


def build_note_sheet(
    note_sheet: Any, records: list[Record], workbook_name: str
) -> None:
    rows = [
        ["项目", "说明"],
        ["生成时间", date.today().isoformat()],
        ["文档总数", len(records)],
        ["重命名数量", sum(1 for record in records if record.处理状态 == "已重命名")],
        ["汇总文件", workbook_name],
        ["说明", "申请单文件和附件目录列可直接点击打开本地文件或目录。"],
    ]
    for row in rows:
        note_sheet.append(row)

    thin_black = Side(style="thin", color="000000")
    border = Border(top=thin_black, bottom=thin_black, left=thin_black, right=thin_black)
    header_fill = PatternFill(fill_type="solid", fgColor="5B6F84")
    header_font = Font(bold=True, color="000000")
    wrapped = Alignment(vertical="center", wrap_text=True)

    note_sheet.column_dimensions["A"].width = 18
    note_sheet.column_dimensions["B"].width = 40

    for cell in note_sheet[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = border
        cell.alignment = wrapped

    for row in note_sheet.iter_rows(
        min_row=2, max_row=note_sheet.max_row, min_col=1, max_col=2
    ):
        for cell in row:
            cell.border = border
            cell.alignment = wrapped


def collect_docx_files(input_dir: Path) -> list[Path]:
    return sorted(path for path in input_dir.glob("*.docx") if path.is_file())


def process_documents(input_dir: Path) -> list[Record]:
    records: list[Record] = []

    for docx_path in collect_docx_files(input_dir):
        original_name = docx_path.name
        parsed = parse_document(docx_path)
        target_name = build_target_name(parsed["施工开始时间"], parsed["施工内容"])

        current_path = docx_path
        status = "已符合命名规则"
        if not TARGET_NAME_RE.fullmatch(docx_path.name):
            target_path = docx_path.with_name(target_name)
            if target_path.exists() and target_path.resolve() != docx_path.resolve():
                raise FileExistsError(f"目标文件已存在，无法重命名: {target_path}")
            if target_path.name != docx_path.name:
                docx_path.rename(target_path)
                current_path = target_path
                status = "已重命名"

        attachment_dir = detect_attachment_dir(current_path)
        records.append(
            Record(
                项目名称=parsed["项目名称"],
                质保单位=parsed["质保单位"],
                分包单位=parsed["分包单位"],
                质保负责人=parsed["质保负责人"],
                质保负责人联系电话=parsed["质保负责人联系电话"],
                施工区域=parsed["施工区域"],
                施工开始时间=parsed["施工开始时间"],
                施工结束时间=parsed["施工结束时间"],
                时长天=parsed["时长天"],
                施工内容=parsed["施工内容"],
                施工负责人=parsed["施工负责人"],
                施工负责人联系电话=parsed["施工负责人联系电话"],
                影响改动消防设备设施=parsed["影响改动消防设备设施"],
                影响堵塞应急疏散通道=parsed["影响堵塞应急疏散通道"],
                危险作业=parsed["危险作业"],
                原文件名=original_name,
                新文件名=current_path.name,
                文件路径=str(current_path),
                申请单文件链接=str(current_path),
                附件目录=str(attachment_dir) if attachment_dir else "",
                附件目录链接=str(attachment_dir) if attachment_dir else "",
                处理状态=status,
            )
        )

    return records


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


def main() -> int:
    parser = argparse.ArgumentParser(
        description="重命名质保作业申请单并导出 Excel 汇总"
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        help="输入目录，未指定时优先读取 common.env 中的 INPUT_PATH",
    )
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parent
    input_dir = (
        args.input_dir.resolve() if args.input_dir else resolve_input_dir(repo_root)
    )

    if not input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在: {input_dir}")

    records = process_documents(input_dir)
    output_path = input_dir / "质保作业申请汇总.xlsx"
    export_excel(records, output_path)

    summary = {
        "input_dir": str(input_dir),
        "total_docs": len(records),
        "renamed_docs": sum(1 for record in records if record.处理状态 == "已重命名"),
        "excel_path": str(output_path),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
