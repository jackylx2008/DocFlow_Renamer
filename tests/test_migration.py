import os
import shutil
import tempfile
import unittest
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from src.warranty_application_archive.approval_review import (
    ApprovalReviewRepository,
    apply_review_decisions,
    build_approval_review,
    import_json_decisions,
)
from src.warranty_application_archive.approval_review_web import (
    _file_drop_clipboard_data,
    export_approval_review_html,
    save_and_apply_review_payload,
    save_review_payload,
)
from src.warranty_application_archive.constants import (
    APPROVAL_REVIEW_DATA_FILE_NAME,
    APPROVAL_REVIEW_HTML_FILE_NAME,
    APPROVAL_REVIEW_LAUNCHER_FILE_NAME,
    RETIRED_APPROVAL_REVIEW_EXCEL_FILE_NAME,
    DATA_FILE_NAME,
    LEGACY_SUMMARY_EXCEL_FILE_NAME,
    SUMMARY_LAUNCHER_FILE_NAME,
    TEMPLATE_FILE_NAME,
)
from src.warranty_application_archive.file_utils import sha256_file
from src.warranty_application_archive.migration import (
    apply_migration_plan,
    build_migration_plan,
    file_record,
    verify_backup,
)
from src.warranty_application_archive.repository import JsonRepository
from src.warranty_application_archive.summary_html import (
    build_summary_view,
    export_summary_html,
)
from src.warranty_application_archive.validation import validate_summary_html
from src.warranty_application_archive.workflows import (
    intake_applications,
    ingest_approval_pdfs,
    ingest_worker_lists,
    route_input_files,
)


PARSED_APPLICATION = {
    "项目名称": "测试项目",
    "质保单位": "测试单位",
    "分包单位": "",
    "质保负责人": "负责人",
    "质保负责人联系电话": "13800000000",
    "施工区域": "冷却塔",
    "施工开始时间": "2026-07-24",
    "施工结束时间": "2026-07-24",
    "时长天": 1,
    "施工内容": "维修冷塔",
    "施工负责人": "施工负责人",
    "施工负责人联系电话": "13900000000",
    "影响改动消防设备设施": "否",
    "影响堵塞应急疏散通道": "否",
    "危险作业": "",
}


class MigrationTest(unittest.TestCase):
    def _migrated_fixture(
        self, base: Path
    ) -> tuple[Path, dict[str, object], str]:
        primary = base / "质保作业申请单"
        primary.mkdir()
        stem = "2026-07-24_维修冷塔_质保作业申请单"
        (primary / f"{stem}.docx").write_bytes(b"word")
        (primary / f"{stem}.jpg").write_bytes(b"signed")
        (primary / f"{stem}_工人名单.jpg").write_bytes(b"workers")
        (primary / TEMPLATE_FILE_NAME).write_bytes(b"agreement")
        with patch(
            "src.warranty_application_archive.migration.legacy.parse_document",
            return_value=PARSED_APPLICATION,
        ):
            plan = build_migration_plan(primary)
        return primary, apply_migration_plan(plan), stem

    def test_migration_builds_case_directory_and_json(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary = base / "质保作业申请单"
            backup = base / "质保作业申请单_backup"
            primary.mkdir()
            stem = "2026-07-24_维修冷塔_质保作业申请单"
            (primary / f"{stem}.docx").write_bytes(b"word")
            (primary / f"{stem}.jpg").write_bytes(b"signed")
            (primary / f"{stem}_工人名单.jpg").write_bytes(b"workers")
            (primary / TEMPLATE_FILE_NAME).write_bytes(b"agreement")
            shutil.copytree(primary, backup)

            verify_backup(primary, backup)
            with patch(
                "src.warranty_application_archive.migration.legacy.parse_document",
                return_value=PARSED_APPLICATION,
            ):
                plan = build_migration_plan(primary)

            self.assertEqual(len(plan.applications), 1)
            self.assertEqual(plan.applications[0]["status"], "materials_ready")
            self.assertEqual(plan.applications[0]["missing_material_types"], [])

            dataset = apply_migration_plan(plan)
            repository = JsonRepository(primary)
            repository.save(dataset)
            retired_excel = primary / LEGACY_SUMMARY_EXCEL_FILE_NAME
            retired_excel.write_bytes(b"retired summary")
            html_path = export_summary_html(dataset, primary)

            case_dir = primary / "_cases" / stem
            self.assertTrue((case_dir / f"{stem}.docx").is_file())
            self.assertTrue(
                (case_dir / f"{stem}_手签_01.jpg").is_file()
            )
            self.assertTrue(
                (case_dir / f"{stem}_施工人员名单_01.jpg").is_file()
            )
            self.assertTrue(
                (
                    case_dir
                    / "2026-07-24_维修冷塔_01 安全生产及消防安全协议（建工）.pdf"
                ).is_file()
            )
            self.assertTrue((primary / DATA_FILE_NAME).is_file())
            self.assertTrue(html_path.is_file())
            report = validate_summary_html(primary, 1)
            self.assertTrue(report["valid"], report["errors"])
            self.assertEqual(
                report["sheets"],
                [
                    "申请汇总",
                    "待补材料",
                    "待审批PDF",
                    "已完成",
                ],
            )
            html = html_path.read_text(encoding="utf-8")
            self.assertIn("材料齐全，待审批PDF", html)
            self.assertIn("维修冷塔", html)
            self.assertIn(f"{stem}.docx", html)
            self.assertNotIn('"title": "说明"', html)
            self.assertNotIn('"title": "本次变更"', html)
            self.assertIn("width: min(2520px, calc(100% - 28px))", html)
            self.assertIn("table-layout: fixed", html)
            self.assertIn("overflow-x: hidden", html)
            self.assertNotIn("min-width: 116px", html)
            self.assertIn("复制文件（可直接粘贴）", html)
            self.assertIn("/api/copy-file", html)
            summary_launcher = primary / SUMMARY_LAUNCHER_FILE_NAME
            self.assertTrue(summary_launcher.is_file())
            self.assertIn(
                "--page summary",
                summary_launcher.read_text(encoding="utf-8"),
            )
            self.assertTrue(dataset.get("changes"))
            self.assertFalse(retired_excel.exists())
            self.assertTrue(
                (
                    primary
                    / ".docflow"
                    / "legacy"
                    / LEGACY_SUMMARY_EXCEL_FILE_NAME
                ).is_file()
            )

    def test_backup_mismatch_blocks_migration(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary = base / "primary"
            backup = base / "backup"
            primary.mkdir()
            backup.mkdir()
            (primary / "file.txt").write_text("primary", encoding="utf-8")
            (backup / "file.txt").write_text("changed", encoding="utf-8")

            with self.assertRaises(RuntimeError):
                verify_backup(primary, backup)

    def test_terminated_case_is_gray_and_excluded_from_work_queues(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            application = dataset["applications"][0]
            application["status"] = "terminated"

            view = build_summary_view(dataset)
            sheets = {
                sheet["title"]: sheet
                for sheet in view["sheets"]
            }
            row = sheets["申请汇总"]["rows"][0]
            self.assertEqual(row[0]["text"], "终止")
            self.assertTrue(
                all(cell["tone"] == "terminated" for cell in row)
            )
            self.assertEqual(sheets["待补材料"]["rows"], [])
            self.assertEqual(sheets["待审批PDF"]["rows"], [])
            self.assertEqual(sheets["已完成"]["rows"], [])

            html_path = export_summary_html(dataset, primary)
            html = html_path.read_text(encoding="utf-8")
            self.assertIn("td.tone-terminated", html)
            self.assertIn('"text": "终止"', html)

    def test_worker_list_and_approval_pdf_subworkflows(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary, dataset, stem = self._migrated_fixture(base)
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)

            worker_source = inbox / "现场施工人员名单.jpg"
            worker_source.write_bytes(b"new workers")
            worker_timestamp = datetime(2026, 7, 24, 12, 0).timestamp()
            os.utime(
                worker_source,
                (worker_timestamp, worker_timestamp),
            )
            application = dataset["applications"][0]
            before_workers = len(
                application["materials"]["worker_list"]
            )
            worker_count = ingest_worker_lists(dataset, primary)
            self.assertEqual(worker_count, 1)
            self.assertEqual(
                len(application["materials"]["worker_list"]),
                before_workers + 1,
            )
            self.assertFalse(worker_source.exists())

            approval_source = (
                inbox / "工程类-主体质保施工_编号：202607240001.pdf"
            )
            approval_source.write_bytes(b"approval")
            dataset["unmatched_files"].append(
                file_record(
                    approval_source,
                    approval_source,
                    primary,
                    "approval_pdf",
                )
            )
            recognized_text = (
                "工程类-主体质保施工 申请编号：202607240001 "
                "施工区域：冷却塔 施工内容：维修冷塔 "
                "施工开始时间：2026年7月24日 "
                "施工结束时间：2026年7月24日"
            )
            with patch(
                "src.warranty_application_archive.workflows."
                "RecognitionService.pdf_text",
                return_value=recognized_text,
            ):
                approval_count = ingest_approval_pdfs(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(approval_count, 1)
            target = (
                primary
                / "_cases"
                / stem
                / "工程类-主体质保施工_编号：202607240001.pdf"
            )
            self.assertTrue(target.is_file())
            self.assertEqual(application["status"], "approved")
            self.assertEqual(
                application["approval"]["application_no"],
                "202607240001",
            )

    def test_input_router_separates_approval_and_application_files(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary, dataset, stem = self._migrated_fixture(base)
            input_root = primary / "_input"
            input_root.mkdir()
            approval_pdf = (
                input_root
                / "工程类-主体质保施工_编号：202607240088.pdf"
            )
            material_pdf = input_root / f"{stem}_有限空间申请.pdf"
            image = input_root / f"{stem}_补充图片.jpg"
            word = input_root / "新增质保申请.docx"
            approval_pdf.write_bytes(b"approval")
            material_pdf.write_bytes(b"confined space")
            image.write_bytes(b"image")
            word.write_bytes(b"word")

            def recognized_text(path: Path) -> str:
                if path.name == approval_pdf.name:
                    return (
                        "工程类-主体质保施工 申请编号：202607240088 "
                        "施工区域：冷却塔 施工内容：维修冷塔 "
                        "施工开始时间：2026年7月24日 "
                        "施工结束时间：2026年7月24日"
                    )
                return "有限空间作业申请 施工区域：冷却塔"

            with patch(
                "src.warranty_application_archive.workflows."
                "RecognitionService.pdf_text",
                side_effect=recognized_text,
            ):
                summary = route_input_files(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )
                approval_count = ingest_approval_pdfs(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(summary["input_files_routed"], 4)
            self.assertEqual(summary["approval_pdfs_routed"], 1)
            self.assertEqual(summary["application_files_routed"], 3)
            self.assertEqual(approval_count, 1)
            self.assertFalse(any(input_root.iterdir()))
            self.assertTrue((primary / "_inbox" / material_pdf.name).is_file())
            self.assertTrue((primary / "_inbox" / image.name).is_file())
            self.assertTrue((primary / "_inbox" / word.name).is_file())
            self.assertFalse(
                any(
                    item.get("path", "").endswith(material_pdf.name)
                    for item in dataset.get("unmatched_files") or []
                )
            )
            route_kinds = {
                item["sha256"]: item["kind"]
                for item in dataset["input_routes"]
            }
            approval_route = next(
                item
                for item in dataset["input_routes"]
                if item["kind"] == "approval_pdf"
            )
            self.assertEqual(
                approval_route["recognition_method"],
                "filename",
            )
            self.assertEqual(
                route_kinds[sha256_file(primary / "_inbox" / material_pdf.name)],
                "application_material",
            )
            archived_approval = (
                primary
                / "_cases"
                / stem
                / "工程类-主体质保施工_编号：202607240088.pdf"
            )
            self.assertTrue(archived_approval.is_file())

    def test_new_application_intake_creates_case(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary, dataset, _stem = self._migrated_fixture(base)
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            incoming_word = inbox / "待处理申请.docx"
            incoming_word.write_bytes(b"new word")
            incoming_image = inbox / "待处理申请.jpg"
            incoming_image.write_bytes(b"new signed")
            parsed = {
                **PARSED_APPLICATION,
                "施工开始时间": "2026-07-25",
                "施工结束时间": "2026-07-25",
                "施工内容": "保温修复",
            }

            with patch(
                "src.warranty_application_archive.workflows."
                "legacy.parse_document",
                return_value=parsed,
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 1)
            application = dataset["applications"][-1]
            self.assertEqual(
                application["case_name"],
                "2026-07-25_保温修复_质保作业申请单",
            )
            case_dir = (
                primary
                / "_cases"
                / "2026-07-25_保温修复_质保作业申请单"
            )
            self.assertTrue(
                (
                    case_dir
                    / "2026-07-25_保温修复_质保作业申请单.docx"
                ).is_file()
            )
            self.assertTrue(
                (
                    case_dir
                    / "2026-07-25_保温修复_质保作业申请单_手签_01.jpg"
                ).is_file()
            )

    def test_human_approval_review_updates_formal_data(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary, dataset, stem = self._migrated_fixture(base)
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            source = (
                inbox / "工程类-主体质保施工_编号：202607240099.pdf"
            )
            source.write_bytes(b"approval requiring human review")
            source_record = file_record(
                source,
                source,
                primary,
                "approval_pdf",
            )
            dataset["unmatched_files"].append(source_record)
            dataset.setdefault("recognition_cache", {})[
                sha256_file(source)
            ] = {
                "text": (
                    "工程类主体质保施工申请编号202607240099"
                    "施工区域冷却塔施工开始时间施工结束时间"
                    "2026~07~23至2026~07~25时长3天"
                    "施工内容维修冷塔施工负责人测试人员"
                ),
                "method": "test",
            }
            outside_inbox = primary / "历史未匹配审批.pdf"
            outside_inbox.write_bytes(b"outside inbox")
            outside_record = file_record(
                outside_inbox,
                outside_inbox,
                primary,
                "approval_pdf",
            )
            dataset["unmatched_files"].append(outside_record)
            dataset["recognition_cache"][sha256_file(outside_inbox)] = {
                "text": "施工区域冷却塔施工内容维修冷塔",
                "method": "test",
            }

            review = build_approval_review(
                dataset,
                primary,
                Path(__file__).resolve().parents[1],
            )
            self.assertEqual(len(review["pending_reviews"]), 1)
            candidate = review["pending_reviews"][0]
            self.assertEqual(candidate["case"]["case_name"], stem)
            self.assertEqual(candidate["pdf"]["start"], "2026-07-23")
            self.assertGreaterEqual(candidate["confidence"], 80)
            self.assertEqual(review["unresolved_pdfs"], [])

            html_path = export_approval_review_html(review, primary)
            decision_payload = {
                "source_dataset_revision": review[
                    "source_dataset_revision"
                ],
                "decisions": [
                    {
                        "review_id": candidate["review_id"],
                        "decision": "确认匹配",
                        "review_note": "人工核对施工内容与日期后确认",
                        "pdf_sha256": candidate["pdf"]["sha256"],
                        "case_id": candidate["case"]["case_id"],
                    }
                ],
            }
            tampered_payload = deepcopy(decision_payload)
            tampered_payload["decisions"][0]["pdf_sha256"] = "tampered"
            with self.assertRaisesRegex(ValueError, "SHA-256"):
                save_review_payload(
                    review,
                    tampered_payload,
                    int(dataset.get("dataset_revision") or 0),
                )
            with self.assertRaisesRegex(ValueError, "正式数据已更新"):
                save_review_payload(
                    review,
                    decision_payload,
                    int(dataset.get("dataset_revision") or 0) + 1,
                )
            JsonRepository(primary).save(dataset)
            ApprovalReviewRepository(primary).save(review)
            applied = save_and_apply_review_payload(
                primary,
                Path(__file__).resolve().parents[1],
                decision_payload,
            )
            self.assertEqual(
                applied["approval_pdfs_human_confirmed"],
                1,
            )
            self.assertEqual(
                applied["approval_pdfs_moved_to_trash"],
                0,
            )
            self.assertEqual(
                applied["review"]["pending_reviews"],
                [],
            )
            dataset = JsonRepository(primary).load()
            application = dataset["applications"][0]
            self.assertEqual(application["status"], "approved")
            self.assertEqual(
                application["approval"]["match_source"],
                "human_review",
            )
            self.assertEqual(len(dataset["unmatched_files"]), 1)
            self.assertEqual(
                dataset["unmatched_files"][0]["path"],
                "历史未匹配审批.pdf",
            )
            self.assertFalse(source.exists())
            self.assertTrue(
                (
                    primary
                    / "_cases"
                    / stem
                    / "工程类-主体质保施工_编号：202607240099.pdf"
                ).is_file()
            )
            self.assertTrue(html_path.is_file())
            self.assertTrue(
                (primary / "质保作业申请汇总.html").is_file()
            )

    def test_review_outputs_only_top_strict_candidate_per_inbox_pdf(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            base = Path(temporary_dir)
            primary, dataset, _stem = self._migrated_fixture(base)
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)

            first_application = dataset["applications"][0]
            first_application["application"]["施工开始时间"] = "2026-01-01"
            first_application["application"]["施工结束时间"] = "2026-01-02"
            second_application = deepcopy(first_application)
            second_application["case_id"] = "second-case"
            second_application["case_name"] = (
                "2026-08-18_维修冷塔_质保作业申请单"
            )
            second_application["case_directory"] = (
                "_cases/2026-08-18_维修冷塔_质保作业申请单"
            )
            second_application["application"]["施工开始时间"] = "2026-08-18"
            second_application["application"]["施工结束时间"] = "2026-08-20"
            (
                primary / second_application["case_directory"]
            ).mkdir(parents=True)
            dataset["applications"].append(second_application)

            matching_pdf = inbox / "工程类-主体质保施工_编号：202608190001.pdf"
            matching_pdf.write_bytes(b"top candidate")
            dataset["unmatched_files"].append(
                file_record(
                    matching_pdf,
                    matching_pdf,
                    primary,
                    "approval_pdf",
                )
            )
            dataset["unmatched_files"].append(
                deepcopy(dataset["unmatched_files"][-1])
            )
            matching_hash = sha256_file(matching_pdf)
            dataset.setdefault("recognition_cache", {})[matching_hash] = {
                "text": (
                    "工程类主体质保施工申请编号202608190001"
                    "施工区域冷却塔施工开始时间2026年8月19日"
                    "施工结束时间2026年8月20日"
                    "施工内容维修冷塔施工负责人测试人员"
                ),
                "method": "ocr",
            }

            unresolved_pdf = (
                inbox / "工程类-主体质保施工_编号：202608190002.pdf"
            )
            unresolved_pdf.write_bytes(b"no strict candidate")
            dataset["unmatched_files"].append(
                file_record(
                    unresolved_pdf,
                    unresolved_pdf,
                    primary,
                    "approval_pdf",
                )
            )
            unresolved_hash = sha256_file(unresolved_pdf)
            dataset["recognition_cache"][unresolved_hash] = {
                "text": (
                    "工程类主体质保施工申请编号202608190002"
                    "施工区域冷却塔施工开始时间2026年7月24日"
                    "施工结束时间2026年7月24日"
                    "施工内容更换UV灯管施工负责人测试人员"
                ),
                "method": "plain",
            }

            nested = inbox / "nested"
            nested.mkdir()
            nested_pdf = nested / "嵌套目录审批.pdf"
            nested_pdf.write_bytes(b"nested")
            dataset["unmatched_files"].append(
                file_record(
                    nested_pdf,
                    nested_pdf,
                    primary,
                    "approval_pdf",
                )
            )
            dataset["recognition_cache"][sha256_file(nested_pdf)] = {
                "text": "施工区域冷却塔施工内容维修冷塔",
                "method": "test",
            }

            review = build_approval_review(
                dataset,
                primary,
                Path(__file__).resolve().parents[1],
            )

            self.assertEqual(len(review["pending_reviews"]), 1)
            candidate = review["pending_reviews"][0]
            self.assertEqual(candidate["pdf"]["sha256"], matching_hash)
            self.assertEqual(
                candidate["case"]["case_id"],
                "second-case",
            )
            self.assertEqual(candidate["strict_candidate_count"], 2)
            self.assertEqual(candidate["confidence"], 98)
            self.assertEqual(candidate["runner_up_confidence"], 80)
            self.assertEqual(candidate["confidence_gap"], 18)
            self.assertEqual(len(review["unresolved_pdfs"]), 1)
            self.assertEqual(
                review["unresolved_pdfs"][0]["pdf"]["sha256"],
                unresolved_hash,
            )
            self.assertIn(
                "施工内容",
                review["unresolved_pdfs"][0]["reason"],
            )

            review_path = ApprovalReviewRepository(primary).save(review)
            self.assertEqual(
                review_path.name,
                APPROVAL_REVIEW_DATA_FILE_NAME,
            )
            retired_excel = (
                primary / RETIRED_APPROVAL_REVIEW_EXCEL_FILE_NAME
            )
            retired_excel.write_bytes(b"retired workbook")
            html_path = export_approval_review_html(review, primary)
            self.assertEqual(
                html_path.name,
                APPROVAL_REVIEW_HTML_FILE_NAME,
            )
            html = html_path.read_text(encoding="utf-8")
            self.assertIn("保存并执行审核结果", html)
            self.assertNotIn("showOpenFilePicker", html)
            self.assertIn("打开待人工审核匹配PDF.cmd", html)
            self.assertIn("移至 _trash", html)
            self.assertIn(matching_pdf.name, html)
            self.assertIn(unresolved_pdf.name, html)
            self.assertIn('"confidence": 98', html)
            self.assertIn("width: min(2200px, calc(100% - 32px))", html)
            self.assertIn(
                "html, body { max-width: 100%; overflow-x: hidden; }",
                html,
            )
            self.assertIn("overflow-wrap: anywhere", html)
            self.assertIn("复制文件（可直接粘贴）", html)
            self.assertIn("/api/copy-file", html)
            launcher = primary / APPROVAL_REVIEW_LAUNCHER_FILE_NAME
            self.assertTrue(launcher.is_file())
            self.assertIn(
                "approval-review-server",
                launcher.read_text(encoding="utf-8"),
            )

            self.assertFalse(retired_excel.exists())
            self.assertTrue(
                (
                    primary
                    / ".docflow"
                    / "legacy"
                    / RETIRED_APPROVAL_REVIEW_EXCEL_FILE_NAME
                ).is_file()
            )

            unresolved_item = review["unresolved_pdfs"][0]
            save_review_payload(
                review,
                {
                    "source_dataset_revision": review[
                        "source_dataset_revision"
                    ],
                    "decisions": [
                        {
                            "review_id": unresolved_item["review_id"],
                            "decision": "移至_trash",
                            "review_note": "人工确认不是有效审批单",
                            "pdf_sha256": unresolved_item["pdf"]["sha256"],
                            "case_id": "",
                        }
                    ],
                },
                int(dataset.get("dataset_revision") or 0),
            )
            trash_result = apply_review_decisions(
                dataset,
                review,
                import_json_decisions(review),
                primary,
            )
            self.assertEqual(trash_result["confirmed"], 0)
            self.assertEqual(trash_result["trashed"], 1)
            self.assertFalse(unresolved_pdf.exists())
            trashed_pdf = primary / "_trash" / unresolved_pdf.name
            self.assertTrue(trashed_pdf.is_file())
            self.assertEqual(
                sha256_file(trashed_pdf),
                unresolved_hash,
            )
            self.assertFalse(
                any(
                    item.get("sha256") == unresolved_hash
                    for item in dataset.get("unmatched_files") or []
                )
            )

    def test_windows_file_clipboard_payload_contains_absolute_path(
        self,
    ) -> None:
        path = Path("资料") / "审批单.pdf"
        payload = _file_drop_clipboard_data(path)

        self.assertEqual(payload[:4], (20).to_bytes(4, "little"))
        self.assertEqual(payload[16:20], (1).to_bytes(4, "little"))
        self.assertEqual(
            payload[20:].decode("utf-16le"),
            f"{path.resolve()}\0\0",
        )


if __name__ == "__main__":
    unittest.main()
