import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import docflow_renamer


class PdfTextCacheTest(unittest.TestCase):
    def test_content_fingerprint_changes_with_pdf_content(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            pdf_path = Path(temporary_dir) / "sample.pdf"
            pdf_path.write_bytes(b"first")
            first_fingerprint = docflow_renamer.pdf_content_fingerprint(pdf_path)

            pdf_path.write_bytes(b"second")
            second_fingerprint = docflow_renamer.pdf_content_fingerprint(pdf_path)

        self.assertNotEqual(first_fingerprint, second_fingerprint)

    def test_build_index_reuses_persisted_text_without_reading_pdf_again(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            input_dir = Path(temporary_dir)
            pdf_path = input_dir / "sample.pdf"
            pdf_path.write_bytes(b"same pdf content")
            expected_text = docflow_renamer.normalize_match_text("施工区域 冷却塔")

            with (
                patch.object(
                    docflow_renamer,
                    "read_pdf_plain_text",
                    return_value="施工区域 冷却塔",
                ),
                patch.object(docflow_renamer, "count_cjk_chars", return_value=20),
            ):
                first_index = docflow_renamer.build_pdf_text_index(input_dir)

            self.assertEqual(first_index[pdf_path], expected_text)
            self.assertTrue(
                (input_dir / docflow_renamer.PDF_TEXT_CACHE_NAME).is_file()
            )

            with patch.object(
                docflow_renamer,
                "read_pdf_plain_text",
                side_effect=AssertionError("cache miss caused PDF to be read again"),
            ):
                second_index = docflow_renamer.build_pdf_text_index(input_dir)

            self.assertEqual(second_index[pdf_path], expected_text)

    def test_empty_recognition_result_is_still_a_cache_hit(self) -> None:
        entries = {
            "fingerprint": {
                "text": "",
                "method": "ocr",
            }
        }

        text, method = docflow_renamer.get_cached_pdf_text(
            entries, "fingerprint"
        )

        self.assertEqual(text, "")
        self.assertEqual(method, "ocr")

    def test_pdf_rename_reuses_cache_without_starting_ai(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            input_dir = Path(temporary_dir)
            pdf_path = input_dir / "incoming.pdf"
            pdf_path.write_bytes(b"cached pdf")
            fingerprint = docflow_renamer.pdf_content_fingerprint(pdf_path)
            cache_path = input_dir / docflow_renamer.PDF_TEXT_CACHE_NAME
            docflow_renamer.cache_pdf_text(
                cache_path,
                {},
                fingerprint,
                pdf_path,
                docflow_renamer.normalize_match_text(
                    "工程类-主体质保施工 申请编号：202607240001"
                ),
                "ocr",
            )

            with patch.object(
                docflow_renamer,
                "read_pdf_ai_ocr_text",
                side_effect=AssertionError("cached PDF caused AI OCR"),
            ):
                renamed_count = (
                    docflow_renamer.rename_subject_warranty_pdfs_by_local_ai(
                        input_dir,
                        Path(docflow_renamer.__file__).resolve().parent,
                    )
                )

            target_path = input_dir / (
                f"{docflow_renamer.PDF_TARGET_NAME_PREFIX}202607240001.pdf"
            )
            self.assertEqual(renamed_count, 1)
            self.assertTrue(target_path.is_file())
            self.assertFalse(pdf_path.exists())


if __name__ == "__main__":
    unittest.main()
