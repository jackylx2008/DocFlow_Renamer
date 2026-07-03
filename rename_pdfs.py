from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path

import fitz

from docflow_renamer import (
    LlamaCppClient,
    LlamaCppConfig,
    resolve_input_dir,
    setup_logging,
)


LOGGER = logging.getLogger(__name__)

SOURCE_DIGIT_NAME_RE = re.compile(r"^\d+\.pdf$", re.IGNORECASE)
SOURCE_APPLICATION_NAME_RE = re.compile(
    r"^申请\s*编号\s*[:：]?\s*[0-9０-９][0-9０-９\s\-—–－]{10,30}[0-9０-９]\.pdf$",
    re.IGNORECASE,
)
APPLICATION_NO_RE = re.compile(r"申请\s*编号\s*[:：]?\s*([0-9０-９][0-9０-９\s\-—–－]{10,30}[0-9０-９])")
TWELVE_DIGIT_RE = re.compile(r"\d{12}")
TARGET_NAME_TEMPLATE = "工程类-主体质保施工_编号：{application_no}.pdf"
PDF_APPLICATION_NO_PROMPT = (
    "请识别这页 PDF 截图里的文字，并提取“申请编号”后面的 12 位数字编码。"
    "只输出这 12 位数字；如果本页没有申请编号，只输出空字符串。"
)


@dataclass
class RenameResult:
    source: Path
    target: Path | None
    application_no: str
    status: str


def normalize_digits(value: str) -> str:
    return value.translate(str.maketrans("０１２３４５６７８９", "0123456789"))


def extract_application_no(text: str) -> str:
    normalized = normalize_digits(text or "")
    match = APPLICATION_NO_RE.search(normalized)
    if match:
        number_match = TWELVE_DIGIT_RE.search(re.sub(r"\D", "", match.group(1)))
        if number_match:
            return number_match.group(0)

    compact_text = re.sub(r"\s+", "", normalized)
    fallback_match = re.search(r"申请编号[:：]?(\d{12})", compact_text)
    return fallback_match.group(1) if fallback_match else ""


def extract_application_no_from_filename(pdf_path: Path) -> str:
    if TWELVE_DIGIT_RE.fullmatch(pdf_path.stem):
        return pdf_path.stem
    if SOURCE_APPLICATION_NAME_RE.fullmatch(pdf_path.name):
        return extract_application_no(pdf_path.stem)
    return ""


def extract_application_no_with_ai(pdf_path: Path, client: LlamaCppClient, page_limit: int) -> str:
    document = fitz.open(pdf_path)
    try:
        page_count = min(len(document), page_limit)
        for page_index in range(page_count):
            page = document[page_index]
            pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            text = client.extract_image_bytes_text(
                pixmap.tobytes("png"),
                "image/png",
                PDF_APPLICATION_NO_PROMPT,
                max_tokens=128,
            )
            application_no = extract_application_no(text)
            if not application_no:
                direct_match = TWELVE_DIGIT_RE.search(normalize_digits(text))
                application_no = direct_match.group(0) if direct_match else ""
            if application_no:
                return application_no
    finally:
        document.close()
    return ""


def collect_source_pdfs(input_dir: Path) -> list[Path]:
    return sorted(
        path
        for path in input_dir.iterdir()
        if path.is_file()
        and (
            SOURCE_DIGIT_NAME_RE.fullmatch(path.name)
            or SOURCE_APPLICATION_NAME_RE.fullmatch(path.name)
        )
    )


def build_target_path(input_dir: Path, application_no: str) -> Path:
    return input_dir / TARGET_NAME_TEMPLATE.format(application_no=application_no)


def rename_pdf(pdf_path: Path, input_dir: Path, client: LlamaCppClient | None, page_limit: int) -> RenameResult:
    application_no = extract_application_no_from_filename(pdf_path)
    if not application_no:
        if client is None:
            raise RuntimeError("需要本地 AI 识别，但服务未初始化")
        application_no = extract_application_no_with_ai(pdf_path, client, page_limit)
    if not application_no:
        return RenameResult(pdf_path, None, "", "未识别到申请编号")

    target_path = build_target_path(input_dir, application_no)
    if pdf_path.resolve() == target_path.resolve():
        return RenameResult(pdf_path, target_path, application_no, "已是目标文件名")
    if target_path.exists():
        return RenameResult(pdf_path, target_path, application_no, "目标文件已存在，跳过")

    pdf_path.rename(target_path)
    return RenameResult(pdf_path, target_path, application_no, "已重命名")


def process_pdfs(input_dir: Path, repo_root: Path, page_limit: int) -> list[RenameResult]:
    pdf_files = collect_source_pdfs(input_dir)
    LOGGER.info("发现 %s 个待处理 PDF", len(pdf_files))
    if not pdf_files:
        return []

    client: LlamaCppClient | None = None

    def ensure_client() -> LlamaCppClient:
        nonlocal client
        if client is None:
            client = LlamaCppClient(LlamaCppConfig.from_repo(repo_root), repo_root)
            LOGGER.info("初始化本地 AI 服务")
            client.ensure_server()
            client.assert_model_available()
        return client

    try:
        results: list[RenameResult] = []
        for index, pdf_path in enumerate(pdf_files, start=1):
            LOGGER.info("处理 PDF %s/%s: %s", index, len(pdf_files), pdf_path.name)
            try:
                pdf_client = None
                if not extract_application_no_from_filename(pdf_path):
                    pdf_client = ensure_client()
                result = rename_pdf(pdf_path, input_dir, pdf_client, page_limit)
            except Exception as exc:
                LOGGER.warning("处理失败: %s (%s)", pdf_path, exc)
                result = RenameResult(pdf_path, None, "", f"处理失败: {exc}")
            results.append(result)
            target_name = result.target.name if result.target else ""
            LOGGER.info(
                "%s -> %s | %s | %s",
                result.source.name,
                target_name,
                result.application_no or "未提取",
                result.status,
            )
        return results
    finally:
        if client is not None:
            client.shutdown_server()


def main() -> int:
    repo_root = Path(__file__).resolve().parent
    log_path = setup_logging(repo_root)
    parser = argparse.ArgumentParser(
        description="识别 INPUT_PATH 下 数字.pdf 或 申请编号.pdf 的申请编号并重命名"
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        help="输入目录，未指定时优先读取 common.env 中的 INPUT_PATH",
    )
    parser.add_argument(
        "--page-limit",
        type=int,
        default=2,
        help="每个 PDF 最多识别前几页，默认 2",
    )
    args = parser.parse_args()

    LOGGER.info("终端输出日志: %s", log_path)
    input_dir = args.input_dir.resolve() if args.input_dir else resolve_input_dir(repo_root)
    if not input_dir.exists():
        raise FileNotFoundError(f"输入目录不存在: {input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"输入路径不是目录: {input_dir}")
    if args.page_limit < 1:
        raise ValueError("--page-limit 必须大于等于 1")

    LOGGER.info("输入目录: %s", input_dir)
    results = process_pdfs(input_dir, repo_root, args.page_limit)
    summary = {
        "processed": len(results),
        "renamed": sum(1 for result in results if result.status == "已重命名"),
        "skipped": sum(1 for result in results if result.status != "已重命名"),
    }
    LOGGER.info("处理汇总: %s", summary)
    return 0


if __name__ == "__main__":
    sys.exit(main())
