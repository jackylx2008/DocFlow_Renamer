import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from warranty_application_archive.modules import legacy as archive_workflow
from warranty_application_archive.modules.recognition import RecognitionService


class PdfTextCacheTest(unittest.TestCase):
    def test_content_fingerprint_changes_with_pdf_content(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            pdf_path = Path(temporary_dir) / "sample.pdf"
            pdf_path.write_bytes(b"first")
            first_fingerprint = archive_workflow.pdf_content_fingerprint(pdf_path)

            pdf_path.write_bytes(b"second")
            second_fingerprint = archive_workflow.pdf_content_fingerprint(pdf_path)

        self.assertNotEqual(first_fingerprint, second_fingerprint)

    def test_build_index_reuses_persisted_text_without_reading_pdf_again(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            input_dir = Path(temporary_dir)
            pdf_path = input_dir / "sample.pdf"
            pdf_path.write_bytes(b"same pdf content")
            expected_text = archive_workflow.normalize_match_text("施工区域 冷却塔")

            with (
                patch.object(
                    archive_workflow,
                    "read_pdf_plain_text",
                    return_value="施工区域 冷却塔",
                ),
                patch.object(archive_workflow, "count_cjk_chars", return_value=20),
            ):
                first_index = archive_workflow.build_pdf_text_index(input_dir)

            self.assertEqual(first_index[pdf_path], expected_text)
            self.assertTrue(
                (input_dir / archive_workflow.PDF_TEXT_CACHE_NAME).is_file()
            )

            with patch.object(
                archive_workflow,
                "read_pdf_plain_text",
                side_effect=AssertionError("cache miss caused PDF to be read again"),
            ):
                second_index = archive_workflow.build_pdf_text_index(input_dir)

            self.assertEqual(second_index[pdf_path], expected_text)

    def test_empty_recognition_result_is_still_a_cache_hit(self) -> None:
        entries = {
            "fingerprint": {
                "text": "",
                "method": "ocr",
            }
        }

        text, method = archive_workflow.get_cached_pdf_text(
            entries, "fingerprint"
        )

        self.assertEqual(text, "")
        self.assertEqual(method, "ocr")

    def test_pdf_rename_reuses_cache_without_starting_ai(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            input_dir = Path(temporary_dir)
            pdf_path = input_dir / "incoming.pdf"
            pdf_path.write_bytes(b"cached pdf")
            fingerprint = archive_workflow.pdf_content_fingerprint(pdf_path)
            cache_path = input_dir / archive_workflow.PDF_TEXT_CACHE_NAME
            archive_workflow.cache_pdf_text(
                cache_path,
                {},
                fingerprint,
                pdf_path,
                archive_workflow.normalize_match_text(
                    "工程类-主体质保施工 申请编号：202607240001"
                ),
                "ocr",
            )

            with patch.object(
                archive_workflow,
                "read_pdf_ai_ocr_text",
                side_effect=AssertionError("cached PDF caused AI OCR"),
            ):
                renamed_count = (
                    archive_workflow.rename_subject_warranty_pdfs_by_local_ai(
                        input_dir,
                        Path(archive_workflow.__file__).resolve().parent,
                    )
                )

            target_path = input_dir / (
                f"{archive_workflow.PDF_TARGET_NAME_PREFIX}202607240001.pdf"
            )
            self.assertEqual(renamed_count, 1)
            self.assertTrue(target_path.is_file())
            self.assertFalse(pdf_path.exists())

    def test_recognition_logs_local_ai_ocr_progress(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            root = Path(temporary_dir)
            pdf_path = root / "扫描审批单.pdf"
            pdf_path.write_bytes(b"scanned pdf")
            client = Mock()

            with (
                patch.object(
                    archive_workflow,
                    "read_pdf_plain_text",
                    return_value="",
                ),
                patch.object(
                    archive_workflow,
                    "read_pdf_ai_ocr_text",
                    return_value="工程类-主体质保施工 申请编号 202607240001",
                ),
                patch.object(
                    RecognitionService,
                    "_ensure_client",
                    return_value=client,
                ),
                self.assertLogs(
                    "warranty_application_archive.modules.recognition",
                    level="INFO",
                ) as captured,
            ):
                text = RecognitionService({}, root).pdf_text(pdf_path)

            messages = "\n".join(captured.output)
            self.assertIn("PDF 内嵌文本不足，准备调用本地 AI OCR", messages)
            self.assertIn("本地 AI OCR 开始识别 PDF", messages)
            self.assertIn("本地 AI OCR 识别完成", messages)
            self.assertIn("申请编号", text)


if __name__ == "__main__":
    unittest.main()
