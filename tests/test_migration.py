import os
import shutil
import tempfile
import unittest
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from unittest.mock import Mock, patch

from warranty_application_archive.modules import legacy
from warranty_application_archive.flows.approval_review_flow import (
    ApprovalReviewRepository,
    apply_review_decisions,
    build_approval_review,
    import_json_decisions,
)
from warranty_application_archive.flows.approval_review_web_flow import (
    _file_drop_clipboard_data,
    copy_file_to_macos_clipboard,
    export_approval_review_html,
    open_path_with_default_application,
    save_and_apply_review_payload,
    save_review_payload,
)
from warranty_application_archive.modules.constants import (
    APPROVAL_REVIEW_DATA_FILE_NAME,
    APPROVAL_REVIEW_HTML_FILE_NAME,
    APPROVAL_REVIEW_LAUNCHER_FILE_NAME,
    APPROVAL_REVIEW_MACOS_LAUNCHER_FILE_NAME,
    CONFINED_SPACE_ROLE,
    RETIRED_APPROVAL_REVIEW_EXCEL_FILE_NAME,
    DATA_FILE_NAME,
    HIGH_ALTITUDE_ROLE,
    LEGACY_SUMMARY_EXCEL_FILE_NAME,
    SIGNED_APPLICATION_ROLE,
    SPECIAL_WORK_ROLE,
    SUMMARY_LAUNCHER_FILE_NAME,
    SUMMARY_MACOS_LAUNCHER_FILE_NAME,
    TEMPLATE_FILE_NAME,
    WORKER_LIST_ROLE,
)
from warranty_application_archive.modules.file_utils import sha256_file
from warranty_application_archive.flows.migration_flow import (
    apply_migration_plan,
    build_migration_plan,
    file_record,
    verify_backup,
)
from warranty_application_archive.modules.repository import JsonRepository
from warranty_application_archive.modules.summary_html import (
    build_summary_view,
    export_summary_html,
)
from warranty_application_archive.modules.validation import validate_summary_html
from warranty_application_archive.flows.archive_flow import (
    _classify_recognized_image,
    _find_duplicate_application,
    _work_option_is_checked,
    intake_applications,
    ingest_approval_pdfs,
    ingest_worker_lists,
    reclassify_historical_materials,
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
            "warranty_application_archive.flows.migration_flow.legacy.parse_document",
            return_value=PARSED_APPLICATION,
        ):
            plan = build_migration_plan(primary)
        return primary, apply_migration_plan(plan), stem

    def test_image_classification_respects_checked_work_options(
        self,
    ) -> None:
        unchecked_form = (
            "质保作业申请单 "
            "动火作业□ 有限空间作业□ 5米以上高处作业□ "
            "危大工程□ 配电室接电□"
        )
        self.assertEqual(
            _classify_recognized_image(unchecked_form),
            SIGNED_APPLICATION_ROLE,
        )
        self.assertFalse(
            _work_option_is_checked(
                "动火作业□ 有限空间作业□",
                "有限空间作业",
            )
        )
        self.assertTrue(
            _work_option_is_checked(
                "动火作业□ 有限空间作业☑",
                "有限空间作业",
            )
        )
        self.assertTrue(
            _work_option_is_checked(
                "动火作业□ 有限空间作业□√",
                "有限空间作业",
            )
        )
        self.assertFalse(
            _work_option_is_checked(
                "动火作业☑ 有限空间作业□",
                "有限空间作业",
            )
        )
        self.assertEqual(
            _classify_recognized_image("专项作业 有限空间作业☑"),
            CONFINED_SPACE_ROLE,
        )
        self.assertEqual(
            _classify_recognized_image("专项作业 5米以上高空作业☒"),
            HIGH_ALTITUDE_ROLE,
        )
        self.assertEqual(
            _classify_recognized_image("专项作业 动火作业：已勾选"),
            SPECIAL_WORK_ROLE,
        )

    def test_approval_content_match_uses_only_construction_field(
        self,
    ) -> None:
        approval = Path("approval.pdf")
        common = (
            "施工区域：冷却塔 "
            "施工开始时间：2026年7月24日 "
            "施工结束时间：2026年7月24日 "
        )
        contained = (
            common
            + "施工内容：维修冷塔（含设备调试及现场清理） "
            + "施工负责人：测试人员"
        )
        self.assertEqual(
            legacy.find_matching_pdf_paths(
                "冷却塔",
                "维修冷塔",
                {approval: contained},
                "2026-07-24",
                "2026-07-24",
            ),
            [approval],
        )

        content_only_in_approval_comment = (
            common
            + "施工内容：更换照明灯具 施工负责人：测试人员 "
            + "审批记录：备注中提到维修冷塔"
        )
        self.assertEqual(
            legacy.find_matching_pdf_paths(
                "冷却塔",
                "维修冷塔",
                {approval: content_only_in_approval_comment},
                "2026-07-24",
                "2026-07-24",
            ),
            [],
        )

    def test_approval_area_match_normalizes_ocr_location_code(
        self,
    ) -> None:
        approval = Path("approval.pdf")
        recognized = (
            "施工区域：i3M1层南区冷却塔 "
            "施工开始时间：2026年7月24日 "
            "施工结束时间：2026年7月24日 "
            "施工内容：维修冷塔及现场清理 "
            "施工负责人：测试人员"
        )
        self.assertEqual(
            legacy.find_matching_pdf_paths(
                "L3M1层南区冷却塔",
                "维修冷塔",
                {approval: recognized},
                "2026-07-24",
                "2026-07-24",
            ),
            [approval],
        )

    def test_reclassify_historical_materials_corrects_false_positive(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            application = dataset["applications"][0]
            case_directory = (
                primary / Path(str(application["case_directory"]))
            )
            misclassified = case_directory / "历史误判有限空间.png"
            misclassified.write_bytes(b"misclassified signed form")
            misclassified_hash = sha256_file(misclassified)
            materials = application["materials"]
            materials[CONFINED_SPACE_ROLE].append(
                file_record(
                    misclassified,
                    misclassified,
                    primary,
                    CONFINED_SPACE_ROLE,
                )
            )

            missing_source = case_directory / "已经不存在的误判图片.png"
            missing_source.write_bytes(b"missing signed form")
            missing_record = file_record(
                missing_source,
                missing_source,
                primary,
                CONFINED_SPACE_ROLE,
            )
            missing_hash = sha256_file(missing_source)
            missing_source.unlink()
            materials[CONFINED_SPACE_ROLE].append(missing_record)
            recognition_cache = dataset.setdefault(
                "recognition_cache",
                {},
            )
            unchecked_form = (
                "质保作业申请单 动火作业□ 有限空间作业□ "
                "5米以上高处作业□ 危大工程□ 配电室接电□"
            )
            recognition_cache[misclassified_hash] = {
                "text": unchecked_form,
                "method": "test",
            }
            recognition_cache[missing_hash] = {
                "text": unchecked_form,
                "method": "test",
            }

            summary = reclassify_historical_materials(
                dataset,
                primary,
            )

            self.assertEqual(summary["records_checked"], 2)
            self.assertEqual(summary["records_reclassified"], 2)
            self.assertEqual(summary["files_moved"], 1)
            self.assertEqual(summary["missing_references_removed"], 1)
            self.assertEqual(materials[CONFINED_SPACE_ROLE], [])
            signed_hashes = {
                item["sha256"]
                for item in materials[SIGNED_APPLICATION_ROLE]
            }
            self.assertIn(misclassified_hash, signed_hashes)
            self.assertNotIn(missing_hash, signed_hashes)
            self.assertFalse(misclassified.exists())
            self.assertTrue(
                any(
                    item["action"] == "reclassify_historical_material"
                    and item["role"]
                    == (
                        f"{CONFINED_SPACE_ROLE}->"
                        f"{SIGNED_APPLICATION_ROLE}"
                    )
                    for item in dataset["changes"]
                )
            )

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
                "warranty_application_archive.flows.migration_flow.legacy.parse_document",
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
            self.assertIn("质保负责人及\\n联系电话", html)
            self.assertIn("施工负责人及\\n联系电话", html)
            self.assertIn("质保单位和分包单位", html)
            self.assertIn("质保：测试单位", html)
            self.assertIn("负责人\\n13800000000", html)
            self.assertIn("施工负责人\\n13900000000", html)
            self.assertNotIn('"缺少材料", "Word申请单"', html)
            self.assertNotIn('"审批编号"', html)
            self.assertNotIn('"案卷目录"', html)
            self.assertNotIn('"案卷ID"', html)
            self.assertIn(f"{stem}.docx", html)
            self.assertNotIn('"title": "说明"', html)
            self.assertNotIn('"title": "本次变更"', html)
            self.assertIn("width: min(2520px, calc(100% - 28px))", html)
            self.assertIn("table-layout: fixed", html)
            self.assertIn("overflow-x: hidden", html)
            self.assertNotIn("min-width: 116px", html)
            self.assertIn("复制文件（可直接粘贴）", html)
            self.assertIn("右键分别复制姓名或电话", html)
            self.assertIn("复制姓名", html)
            self.assertIn("复制电话", html)
            self.assertIn("复制质保单位", html)
            self.assertIn("复制分包单位", html)
            self.assertIn("复制内容", html)
            self.assertIn("dataset.copyText", html)
            self.assertIn("dataset.copyName", html)
            self.assertIn("dataset.copyPhone", html)
            self.assertIn("/api/copy-file", html)
            self.assertIn("/api/open-path", html)
            self.assertIn("打开文件失败", html)
            summary_launcher = primary / SUMMARY_LAUNCHER_FILE_NAME
            self.assertTrue(summary_launcher.is_file())
            self.assertIn(
                "--page summary",
                summary_launcher.read_text(encoding="utf-8"),
            )
            summary_macos_launcher = (
                primary / SUMMARY_MACOS_LAUNCHER_FILE_NAME
            )
            self.assertTrue(summary_macos_launcher.is_file())
            summary_macos_script = summary_macos_launcher.read_text(
                encoding="utf-8"
            )
            self.assertIn(
                "#!/bin/zsh",
                summary_macos_script,
            )
            self.assertIn('SCRIPT_DIR="${0:A:h}"', summary_macos_script)
            self.assertIn(
                '"$HOME/anaconda3/bin/python"',
                summary_macos_script,
            )
            self.assertIn("DOCFLOW_PROJECT_ROOT", summary_macos_script)
            self.assertNotIn(str(primary), summary_macos_script)
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
            self.assertEqual(row[0]["text"], "终止；\n材料完整")
            self.assertTrue(
                all(cell["tone"] == "terminated" for cell in row)
            )
            self.assertEqual(sheets["待补材料"]["rows"], [])
            self.assertEqual(sheets["待审批PDF"]["rows"], [])
            self.assertEqual(sheets["已完成"]["rows"], [])

            html_path = export_summary_html(dataset, primary)
            html = html_path.read_text(encoding="utf-8")
            self.assertIn("td.tone-terminated", html)
            self.assertIn('"text": "终止；\\n材料完整"', html)

    def test_summary_business_fields_are_ordered_and_copyable(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            _primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            application = dataset["applications"][0]
            application["status"] = "materials_incomplete"
            application["missing_material_types"] = ["worker_list"]
            application["application"]["分包单位"] = "测试分包"

            view = build_summary_view(dataset)
            sheet = view["sheets"][0]
            headers = sheet["headers"]
            row = sheet["rows"][0]

            self.assertEqual(len(headers), 16)
            self.assertNotIn("缺少材料", headers)
            self.assertNotIn("审批编号", headers)
            self.assertNotIn("案卷目录", headers)
            self.assertNotIn("案卷ID", headers)
            self.assertEqual(
                headers[:10],
                [
                    "案卷状态及\n材料完整性",
                    "项目名称",
                    "质保单位和分包单位",
                    "质保负责人及\n联系电话",
                    "施工区域",
                    "施工开始时间",
                    "施工结束时间",
                    "施工内容",
                    "施工负责人及\n联系电话",
                    "危险作业及专项材料核对",
                ],
            )
            self.assertIn("施工人员名单", row[0]["text"])
            self.assertEqual(row[1]["copy_text"], "测试项目")
            self.assertEqual(
                row[2]["text"],
                "质保：测试单位；\n分包：测试分包",
            )
            self.assertEqual(
                row[2]["copy_warranty_unit"],
                "测试单位",
            )
            self.assertEqual(
                row[2]["copy_subcontract_unit"],
                "测试分包",
            )
            self.assertEqual(row[3]["text"], "负责人\n13800000000")
            self.assertTrue(row[3]["copyable"])
            self.assertEqual(row[3]["copy_name"], "负责人")
            self.assertEqual(row[3]["copy_phone"], "13800000000")
            self.assertEqual(row[4]["copy_text"], "冷却塔")
            self.assertEqual(row[7]["copy_text"], "维修冷塔")
            self.assertEqual(
                row[8]["text"],
                "施工负责人\n13900000000",
            )
            self.assertTrue(row[8]["copyable"])
            self.assertEqual(row[8]["copy_name"], "施工负责人")
            self.assertEqual(row[8]["copy_phone"], "13900000000")
            self.assertIn(
                "影响、改动消防设备设施：否",
                row[9]["text"],
            )
            self.assertIn(
                "影响、堵塞应急疏散通道：否",
                row[9]["text"],
            )
            self.assertIn("危险作业：无", row[9]["text"])
            self.assertIn(
                "专项材料核对：相符（无需专项作业材料）",
                row[9]["text"],
            )
            self.assertEqual(row[9]["tone"], "success")

            business = application["application"]
            business["影响改动消防设备设施"] = "是"
            business["危险作业"] = (
                "动火作业、有限空间作业、5米以上高空作业、"
                "危大工程、配电室接电"
            )
            danger_cell = build_summary_view(dataset)["sheets"][0][
                "rows"
            ][0][9]
            self.assertIn("有限空间申请", danger_cell["text"])
            self.assertIn("高处作业申请", danger_cell["text"])
            self.assertIn("专项作业材料", danger_cell["text"])
            self.assertEqual(danger_cell["tone"], "warning")

            materials = application["materials"]
            materials["confined_space"] = [{"path": "confined.pdf"}]
            materials["high_altitude"] = [{"path": "height.pdf"}]
            materials["special_work"] = [{"path": "special.pdf"}]
            matched_danger_cell = build_summary_view(dataset)[
                "sheets"
            ][0]["rows"][0][9]
            self.assertIn(
                "专项材料核对：相符",
                matched_danger_cell["text"],
            )
            self.assertEqual(matched_danger_cell["tone"], "success")

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
                "施工区域：冷却塔 施工内容：维修冷塔（含设备调试） "
                "施工开始时间：2026年7月24日 "
                "施工结束时间：2026年7月24日"
            )
            with patch(
                "warranty_application_archive.flows.archive_flow."
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
                "warranty_application_archive.flows.archive_flow."
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
                "warranty_application_archive.flows.archive_flow."
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

    def test_image_only_intake_matches_existing_case_with_ocr_content_error(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            application = dataset["applications"][0]
            materials = application["materials"]
            for file_item in materials[SIGNED_APPLICATION_ROLE]:
                signed_path = primary / file_item["path"]
                if signed_path.is_file():
                    signed_path.unlink()
            materials[SIGNED_APPLICATION_ROLE] = []

            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            incoming_image = inbox / "random-signed.png"
            incoming_image.write_bytes(b"later signed application")
            incoming_hash = sha256_file(incoming_image)
            recognized_text = (
                "质保作业申请单 "
                "施工区域：冷却塔 "
                "施工日期：2026年7月24日~2026年7月24日 "
                "施工内容：维修水塔 "
                "施工负责人：施工负责人 联系电话：13900000000"
            )

            with patch(
                "warranty_application_archive.flows.archive_flow."
                "RecognitionService.image_text",
                return_value=recognized_text,
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 0)
            self.assertFalse(incoming_image.exists())
            self.assertTrue(
                any(
                    item["sha256"] == incoming_hash
                    for item in materials[SIGNED_APPLICATION_ROLE]
                )
            )

    def test_image_only_intake_keeps_ambiguous_fuzzy_match_in_inbox(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            application = dataset["applications"][0]
            materials = application["materials"]
            for file_item in materials[SIGNED_APPLICATION_ROLE]:
                signed_path = primary / file_item["path"]
                if signed_path.is_file():
                    signed_path.unlink()
            materials[SIGNED_APPLICATION_ROLE] = []
            application["application"]["施工内容"] = "风机烧坏维修"

            competing = deepcopy(application)
            competing["case_id"] = "competing-case"
            competing["case_name"] = "2026-07-24_风机盘管清洗_质保作业申请单"
            competing["case_directory"] = (
                "_cases/2026-07-24_风机盘管清洗_质保作业申请单"
            )
            competing["application"]["施工内容"] = "风机盘管清洗"
            dataset["applications"].append(competing)

            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            incoming_image = inbox / "ambiguous-signed.png"
            incoming_image.write_bytes(b"ambiguous signed application")
            recognized_text = (
                "质保作业申请单 "
                "施工区域：冷却塔 "
                "施工日期：2026年7月24日~2026年7月24日 "
                "施工内容：风机盘管维修 "
                "施工负责人：施工负责人 联系电话：13900000000"
            )

            with patch(
                "warranty_application_archive.flows.archive_flow."
                "RecognitionService.image_text",
                return_value=recognized_text,
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 0)
            self.assertTrue(incoming_image.is_file())
            self.assertFalse(materials[SIGNED_APPLICATION_ROLE])

    def test_duplicate_application_is_quarantined_and_not_added(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            incoming_word = inbox / "重复申请.docx"
            incoming_word.write_bytes(b"duplicate word")
            original_count = len(dataset["applications"])

            with (
                patch(
                    "warranty_application_archive.flows.archive_flow."
                    "legacy.parse_document",
                    return_value=deepcopy(PARSED_APPLICATION),
                ),
                self.assertLogs(
                    "warranty_application_archive.flows.archive_flow",
                    level="WARNING",
                ) as captured,
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 0)
            self.assertEqual(len(dataset["applications"]), original_count)
            self.assertFalse(incoming_word.exists())
            quarantined = list(
                (primary / ".docflow" / "quarantine").rglob("*")
            )
            self.assertTrue(
                any(path.name == incoming_word.name for path in quarantined)
            )
            messages = "\n".join(captured.output)
            self.assertIn("检测到重复质保申请", messages)
            self.assertIn("已跳过写入 JSON", messages)
            self.assertIn(PARSED_APPLICATION["施工开始时间"], messages)
            self.assertTrue(
                any(
                    change.get("action")
                    == "quarantine_duplicate_application"
                    for change in dataset["changes"]
                )
            )

    def test_duplicate_application_retains_more_complete_case(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            original = dataset["applications"][0]
            more_complete = deepcopy(original)
            more_complete["case_id"] = "more-complete-case"
            more_complete["case_name"] = f"{stem}_02"
            more_complete["case_directory"] = (
                f"_cases/{stem}_02"
            )
            more_complete["status"] = "approved"
            more_complete["approval"] = {
                "status": "approved",
                "pdfs": [
                    {
                        "path": f"_cases/{stem}_02/approval.pdf",
                        "sha256": "approval-hash",
                    }
                ],
            }
            more_complete_dir = primary / more_complete["case_directory"]
            more_complete_dir.mkdir(parents=True)
            for role_files in more_complete["materials"].values():
                for file_item in role_files:
                    source_file = primary / file_item["path"]
                    copied_file = more_complete_dir / source_file.name
                    shutil.copy2(source_file, copied_file)
                    file_item["path"] = copied_file.relative_to(
                        primary
                    ).as_posix()
            (more_complete_dir / "approval.pdf").write_bytes(b"approval")
            dataset["applications"].append(more_complete)
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            incoming_word = inbox / "再次提交.docx"
            incoming_word.write_bytes(b"duplicate word")

            selected = _find_duplicate_application(
                PARSED_APPLICATION,
                dataset["applications"],
            )
            self.assertIs(selected, more_complete)

            with patch(
                "warranty_application_archive.flows.archive_flow."
                "legacy.parse_document",
                return_value=deepcopy(PARSED_APPLICATION),
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 0)
            self.assertEqual(dataset["applications"], [more_complete])
            self.assertTrue(more_complete_dir.is_dir())
            self.assertFalse(
                (primary / "_cases" / stem).exists()
            )
            quarantined_case_dirs = list(
                (
                    primary / ".docflow" / "quarantine"
                ).rglob(stem)
            )
            self.assertTrue(quarantined_case_dirs)
            self.assertTrue(
                any(
                    change.get("action")
                    == "quarantine_less_complete_duplicate_case"
                    for change in dataset["changes"]
                )
            )

    def test_one_input_batch_adds_random_named_images_to_duplicate_case(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            input_dir = primary / "_input"
            inbox = primary / "_inbox"
            input_dir.mkdir(exist_ok=True)
            inbox.mkdir(exist_ok=True)
            incoming_word = input_dir / "重复申请.docx"
            signed_image = input_dir / "random-signed.png"
            worker_image = input_dir / "random-workers.jpg"
            incoming_word.write_bytes(b"duplicate word")
            signed_image.write_bytes(b"new signed application")
            worker_image.write_bytes(b"new worker list")
            existing_worker_image = inbox / worker_image.name
            existing_worker_image.write_bytes(worker_image.read_bytes())
            signed_hash = sha256_file(signed_image)
            worker_hash = sha256_file(worker_image)
            batch_id = "test-input-batch"

            route_input_files(
                dataset,
                primary,
                Path(__file__).resolve().parents[1],
                input_batch_id=batch_id,
            )

            def recognized_text(path: Path) -> str:
                if path.suffix.lower() == ".png":
                    return "质保作业申请单 施工区域 冷却塔"
                return "姓名 性别 电话 张三 男 13800000000"

            with (
                patch(
                    "warranty_application_archive.flows.archive_flow."
                    "legacy.parse_document",
                    return_value=deepcopy(PARSED_APPLICATION),
                ),
                patch(
                    "warranty_application_archive.flows.archive_flow."
                    "RecognitionService.image_text",
                    side_effect=recognized_text,
                ),
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                    input_batch_id=batch_id,
                )

            self.assertEqual(count, 0)
            application = dataset["applications"][0]
            materials = application["materials"]
            self.assertTrue(
                any(
                    item["sha256"] == signed_hash
                    for item in materials["signed_application"]
                )
            )
            self.assertTrue(
                any(
                    item["sha256"] == worker_hash
                    for item in materials["worker_list"]
                )
            )
            self.assertFalse(any(input_dir.iterdir()))
            self.assertFalse(any(inbox.iterdir()))
            worker_route = next(
                item
                for item in dataset["input_routes"]
                if item["sha256"] == worker_hash
            )
            self.assertEqual(
                worker_route["action"],
                "quarantine_duplicate",
            )
            self.assertEqual(
                worker_route["processing_path"],
                "_inbox/random-workers.jpg",
            )

    def test_identical_worker_lists_are_distributed_one_per_existing_case(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            prototype = dataset["applications"][0]
            for file_item in prototype["materials"][WORKER_LIST_ROLE]:
                worker_path = primary / file_item["path"]
                if worker_path.is_file():
                    worker_path.unlink()
            prototype["materials"][WORKER_LIST_ROLE] = []
            prototype["application"]["施工开始时间"] = "2026-08-05"

            applications = []
            for index in range(4):
                application = deepcopy(prototype)
                application["case_id"] = f"worker-batch-{index}"
                application["case_name"] = (
                    f"2026-08-05_批次{index + 1}_质保作业申请单"
                )
                application["case_directory"] = (
                    f"_cases/{application['case_name']}"
                )
                (primary / application["case_directory"]).mkdir(
                    parents=True,
                    exist_ok=True,
                )
                applications.append(application)
            dataset["applications"] = applications

            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            modified_at = datetime(2026, 8, 5, 10, 38).timestamp()
            worker_bytes = b"same worker list for four cases"
            for index in range(4):
                image = inbox / f"random-worker-{index}.jpg"
                image.write_bytes(worker_bytes)
                os.utime(image, (modified_at, modified_at))
            worker_hash = sha256_file(inbox / "random-worker-0.jpg")

            with patch(
                "warranty_application_archive.flows.archive_flow."
                "RecognitionService.image_text",
                return_value="姓名 性别 电话 张三 男 13800000000",
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 0)
            self.assertFalse(any(inbox.iterdir()))
            for application in applications:
                worker_files = application["materials"][WORKER_LIST_ROLE]
                self.assertEqual(len(worker_files), 1)
                self.assertEqual(worker_files[0]["sha256"], worker_hash)

    def test_identical_worker_lists_stay_when_case_count_does_not_match(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            prototype = dataset["applications"][0]
            for file_item in prototype["materials"][WORKER_LIST_ROLE]:
                worker_path = primary / file_item["path"]
                if worker_path.is_file():
                    worker_path.unlink()
            prototype["materials"][WORKER_LIST_ROLE] = []
            prototype["application"]["施工开始时间"] = "2026-08-05"

            applications = []
            for index in range(3):
                application = deepcopy(prototype)
                application["case_id"] = f"ambiguous-worker-{index}"
                application["case_name"] = (
                    f"2026-08-05_候选{index + 1}_质保作业申请单"
                )
                application["case_directory"] = (
                    f"_cases/{application['case_name']}"
                )
                applications.append(application)
            dataset["applications"] = applications

            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            modified_at = datetime(2026, 8, 5, 10, 38).timestamp()
            for index in range(4):
                image = inbox / f"ambiguous-worker-{index}.jpg"
                image.write_bytes(b"same ambiguous worker list")
                os.utime(image, (modified_at, modified_at))

            with patch(
                "warranty_application_archive.flows.archive_flow."
                "RecognitionService.image_text",
                return_value="姓名 性别 电话 张三 男 13800000000",
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 0)
            self.assertEqual(len(list(inbox.iterdir())), 4)
            self.assertTrue(
                all(
                    not application["materials"][WORKER_LIST_ROLE]
                    for application in applications
                )
            )

    def test_same_content_with_different_end_date_is_not_duplicate(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            primary, dataset, _stem = self._migrated_fixture(
                Path(temporary_dir)
            )
            inbox = primary / "_inbox"
            inbox.mkdir(exist_ok=True)
            incoming_word = inbox / "日期不同申请.docx"
            incoming_word.write_bytes(b"different end date")
            parsed = {
                **PARSED_APPLICATION,
                "施工结束时间": "2026-07-25",
            }

            with patch(
                "warranty_application_archive.flows.archive_flow."
                "legacy.parse_document",
                return_value=parsed,
            ):
                count = intake_applications(
                    dataset,
                    primary,
                    Path(__file__).resolve().parents[1],
                )

            self.assertEqual(count, 1)
            self.assertEqual(
                dataset["applications"][-1]["application"]["施工结束时间"],
                "2026-07-25",
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
                    "施工内容维修冷塔及现场清理施工负责人测试人员"
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
                    "审批记录备注中提到维修冷塔"
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
            self.assertIn("Windows 使用 .cmd，macOS 使用 .command", html)
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
            self.assertIn("/api/open-path", html)
            self.assertIn("使用默认程序打开审批 PDF", html)
            launcher = primary / APPROVAL_REVIEW_LAUNCHER_FILE_NAME
            self.assertTrue(launcher.is_file())
            self.assertIn(
                "serve_archive_review.py",
                launcher.read_text(encoding="utf-8"),
            )
            self.assertIn(
                "--port 0",
                launcher.read_text(encoding="utf-8"),
            )
            macos_launcher = (
                primary / APPROVAL_REVIEW_MACOS_LAUNCHER_FILE_NAME
            )
            self.assertTrue(macos_launcher.is_file())
            macos_script = macos_launcher.read_text(encoding="utf-8")
            self.assertIn(
                "#!/bin/zsh",
                macos_script,
            )
            self.assertIn(
                '--input-dir "$SCRIPT_DIR" --port 0',
                macos_script,
            )
            self.assertIn("--port 0", macos_script)
            self.assertNotIn(str(primary), macos_script)

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

    def test_windows_opens_file_with_default_application(self) -> None:
        path = Path("资料") / "申请单.docx"
        with (
            patch(
                "warranty_application_archive.flows."
                "approval_review_web_flow.sys.platform",
                "win32",
            ),
            patch(
                "warranty_application_archive.flows."
                "approval_review_web_flow.os.startfile",
                create=True,
            ) as startfile,
        ):
            open_path_with_default_application(path)

        startfile.assert_called_once_with(str(path.resolve()))

    def test_macos_file_clipboard_uses_alias_list(self) -> None:
        path = Path("资料") / "审批单.pdf"
        completed = Mock(
            returncode=0,
            stderr="",
            stdout="",
        )
        with patch(
            "warranty_application_archive.flows.approval_review_web_flow."
            "subprocess.run",
            return_value=completed,
        ) as run:
            copy_file_to_macos_clipboard(path)

        command = run.call_args.args[0]
        self.assertEqual(command[0], "osascript")
        self.assertIn("set the clipboard to {targetItem}", command[2])
        self.assertEqual(command[-1], str(path.resolve()))


if __name__ == "__main__":
    unittest.main()
