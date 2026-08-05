from __future__ import annotations

import argparse
import json
import logging
import sys
import uuid
from pathlib import Path
from typing import Sequence

from logging_config import setup_logger

from .approval_review_flow import (
    ApprovalReviewRepository,
    apply_review_decisions,
    build_approval_review,
    import_json_decisions,
)
from .approval_review_web_flow import (
    export_approval_review_html,
    serve_approval_review,
)
from ..config_loader import AppConfig
from .migration_flow import (
    apply_migration_plan,
    build_migration_plan,
    verify_backup,
)
from ..modules.repository import JsonRepository
from ..modules.summary_html import export_summary_html
from ..modules.validation import validate_dataset, validate_summary_html
from .archive_flow import (
    append_run,
    deduplicate_applications,
    intake_applications,
    ingest_approval_pdfs,
    ingest_worker_lists,
    reclassify_historical_materials,
    route_input_files,
)


LOGGER = logging.getLogger(__name__)

STATUS_LABELS = {
    "materials_incomplete": "材料不完整",
    "approved": "已审批",
    "terminated": "已终止",
    "unknown": "状态未知",
}


def _repo_root() -> Path:
    current_directory = Path.cwd().resolve()
    if (current_directory / "common.env").is_file():
        return current_directory
    return Path(__file__).resolve().parents[3]


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="质保作业申请案卷归档与人工审查数据生成"
    )
    parser.add_argument("--input-dir", type=Path, help="资料根目录")
    subparsers = parser.add_subparsers(dest="command")

    migrate_parser = subparsers.add_parser(
        "migrate", help="将旧版平铺目录迁移为独立案卷目录"
    )
    migrate_parser.add_argument("--apply", action="store_true", help="执行迁移")
    migrate_parser.add_argument("--backup-dir", type=Path, help="迁移前核对的备份目录")
    migrate_parser.add_argument(
        "--plan-output",
        type=Path,
        help="迁移计划 JSON 输出路径",
    )

    subparsers.add_parser("status", help="显示当前 JSON 数据状态")
    subparsers.add_parser("validate", help="校验 JSON、案卷文件和汇总 HTML")
    subparsers.add_parser("export", help="从 JSON 重新生成汇总 HTML")
    subparsers.add_parser("applications", help="处理待入库质保申请材料")
    subparsers.add_parser("approval-pdfs", help="处理待入库审批 PDF")
    subparsers.add_parser(
        "approval-review",
        help="生成未匹配审批 PDF 的独立审核 JSON 和 HTML",
    )
    review_server_parser = subparsers.add_parser(
        "approval-review-server",
        help="启动可回写审核 JSON 的本地人工审核页面",
    )
    review_server_parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="监听地址（默认 127.0.0.1）",
    )
    review_server_parser.add_argument(
        "--port",
        type=int,
        default=8765,
        help="监听端口（默认 8765）",
    )
    review_server_parser.add_argument(
        "--no-open",
        action="store_true",
        help="启动后不自动打开浏览器",
    )
    review_server_parser.add_argument(
        "--page",
        choices=("review", "summary"),
        default="review",
        help="启动后打开的页面（默认人工审核页）",
    )
    subparsers.add_parser(
        "apply-approval-review",
        help="读取审核 JSON 中的人工决定并同步正式 JSON/HTML",
    )
    subparsers.add_parser("worker-lists", help="处理待入库施工人员名单")
    subparsers.add_parser(
        "reclassify-materials",
        help="按 OCR 方框勾选状态重新判断既往专项材料",
    )
    subparsers.add_parser("run", help="执行完整增量工作流")
    return parser


def _status(repository: JsonRepository) -> dict[str, object]:
    data = repository.load()
    applications = data.get("applications") or []
    statuses: dict[str, int] = {}
    for item in applications:
        status = str(item.get("status") or "unknown")
        statuses[status] = statuses.get(status, 0) + 1
    return {
        "data_file": str(repository.path),
        "dataset_revision": data.get("dataset_revision"),
        "applications": len(applications),
        "statuses": statuses,
        "unmatched_files": len(data.get("unmatched_files") or []),
    }


def _log_workflow_summary(
    status: dict[str, object],
    *,
    applications_ingested: int,
    approval_pdfs_ingested: int,
    worker_lists_ingested: int,
    duplicate_applications_removed: int,
    route_summary: dict[str, int],
    html_file: str,
    review_json: str,
    review_html: str,
) -> None:
    """Write the interactive run report as readable timestamped log lines."""
    LOGGER.info("增量归档工作流执行完成")
    LOGGER.info("正式数据文件: %s", status["data_file"])
    LOGGER.info(
        "数据版本: %s；申请记录: %s 条；待匹配文件: %s 个",
        status["dataset_revision"],
        status["applications"],
        status["unmatched_files"],
    )
    statuses = status.get("statuses")
    if isinstance(statuses, dict) and statuses:
        status_description = "，".join(
            f"{STATUS_LABELS.get(str(name), str(name))} {count} 条"
            for name, count in statuses.items()
        )
        LOGGER.info("申请状态统计: %s", status_description)
    LOGGER.info(
        "本次接收: 申请材料 %s 份；审批 PDF %s 份；施工人员名单 %s 份",
        applications_ingested,
        approval_pdfs_ingested,
        worker_lists_ingested,
    )
    if duplicate_applications_removed:
        LOGGER.info(
            "重复申请整理: 移除资料较少的重复案卷 %s 条",
            duplicate_applications_removed,
        )
    LOGGER.info(
        "_input 分流: 共 %s 个文件；审批 PDF %s 个；申请材料 %s 个；"
        "重复文件隔离 %s 个",
        route_summary["input_files_routed"],
        route_summary["approval_pdfs_routed"],
        route_summary["application_files_routed"],
        route_summary["input_duplicates_quarantined"],
    )
    LOGGER.info("申请汇总页面: %s", html_file)
    if review_json:
        LOGGER.info("待人工审核数据: %s", review_json)
    if review_html:
        LOGGER.info("待人工审核页面: %s", review_html)


def main(argv: Sequence[str] | None = None) -> int:
    repo_root = _repo_root()
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    command = args.command or "run"
    config = AppConfig.resolve(repo_root, args.input_dir)
    setup_logger(log_level=config.log_level, log_dir=config.log_dir)
    if not config.data_root.is_dir():
        raise NotADirectoryError(f"资料根目录不存在: {config.data_root}")
    repository = JsonRepository(config.data_root)
    review_repository = ApprovalReviewRepository(config.data_root)

    if command == "migrate":
        plan = build_migration_plan(config.data_root)
        plan_output = args.plan_output or (
            config.output_dir / "migration_plan.json"
        )
        _write_json(plan_output, plan.to_dict())
        LOGGER.info(
            "迁移计划: %s 个案卷，%s 个操作，%s 个警告",
            len(plan.applications),
            len(plan.operations),
            len(plan.warnings),
        )
        LOGGER.info("迁移计划已保存: %s", plan_output)
        if not args.apply:
            print(
                json.dumps(
                    {
                        "mode": "dry-run",
                        "applications": len(plan.applications),
                        "operations": len(plan.operations),
                        "warnings": plan.warnings,
                        "plan": str(plan_output),
                    },
                    ensure_ascii=False,
                    indent=2,
                )
            )
            return 0
        if not args.backup_dir:
            parser.error("执行迁移必须提供 --backup-dir")
        verify_backup(config.data_root, args.backup_dir)
        dataset = apply_migration_plan(plan)
        data_path = repository.save(dataset)
        LOGGER.info("迁移完成，JSON 已保存: %s", data_path)
        html_path = export_summary_html(dataset, config.data_root)
        LOGGER.info("汇总 HTML 已从 JSON 生成: %s", html_path)
        print(
            json.dumps(
                {
                    **_status(repository),
                    "html_file": str(html_path),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    if command == "status":
        print(json.dumps(_status(repository), ensure_ascii=False, indent=2))
        return 0

    if command == "validate":
        data = repository.load()
        data_report = validate_dataset(data, config.data_root)
        html_report = validate_summary_html(
            config.data_root,
            len(data.get("applications") or []),
        )
        report = {
            "data": data_report,
            "html": html_report,
        }
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if data_report["valid"] and html_report["valid"] else 1

    if command == "export":
        data = repository.load()
        output_path = export_summary_html(data, config.data_root)
        LOGGER.info("汇总 HTML 已从 JSON 生成: %s", output_path)
        print(
            json.dumps(
                {
                    **_status(repository),
                    "html_file": str(output_path),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    if command == "reclassify-materials":
        data = repository.load()
        if not repository.path.is_file():
            raise FileNotFoundError(
                "尚未建立正式 JSON 数据，请先执行 migrate"
            )
        summary = reclassify_historical_materials(
            data,
            config.data_root,
            checkpoint=repository.save,
        )
        if summary["records_reclassified"]:
            data["dataset_revision"] = int(
                data.get("dataset_revision") or 0
            ) + 1
        append_run(data, command, summary)
        repository.save(data)
        formal_html = export_summary_html(data, config.data_root)
        review = build_approval_review(
            data,
            config.data_root,
            repo_root,
            existing=review_repository.load(),
        )
        repository.save(data)
        review_path = review_repository.save(review)
        review_html = export_approval_review_html(
            review,
            config.data_root,
        )
        LOGGER.info("历史专项材料重新判断完成")
        LOGGER.info(
            "已检查 %s 条；重新分类 %s 条；移动文件 %s 个",
            summary["records_checked"],
            summary["records_reclassified"],
            summary["files_moved"],
        )
        LOGGER.info(
            "清理不存在的旧文件引用 %s 条；缺少 OCR 缓存 %s 条",
            summary["missing_references_removed"],
            summary["recognition_cache_missing"],
        )
        LOGGER.info("正式数据文件: %s", repository.path)
        LOGGER.info("申请汇总页面: %s", formal_html)
        LOGGER.info("待人工审核数据: %s", review_path)
        LOGGER.info("待人工审核页面: %s", review_html)
        return 0

    if command == "approval-review":
        data = repository.load()
        if not repository.path.is_file():
            raise FileNotFoundError(
                "尚未建立正式 JSON 数据，请先执行 migrate"
            )
        review = build_approval_review(
            data,
            config.data_root,
            repo_root,
            existing=review_repository.load(),
        )
        repository.save(data)
        review_path = review_repository.save(review)
        review_html = export_approval_review_html(
            review, config.data_root
        )
        print(
            json.dumps(
                {
                    **_status(repository),
                    "review_candidates": len(
                        review.get("pending_reviews") or []
                    ),
                    "review_unresolved_pdfs": len(
                        review.get("unresolved_pdfs") or []
                    ),
                    "review_json": str(review_path),
                    "review_html": str(review_html),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    if command == "approval-review-server":
        if args.page == "review":
            if not review_repository.path.is_file():
                raise FileNotFoundError(
                    "尚未生成审核 JSON，请先运行 approval-review"
                )
            export_approval_review_html(
                review_repository.load(),
                config.data_root,
            )
        else:
            if not repository.path.is_file():
                raise FileNotFoundError("尚未建立正式 JSON 数据")
            export_summary_html(repository.load(), config.data_root)
        serve_approval_review(
            config.data_root,
            host=args.host,
            port=args.port,
            open_browser=not args.no_open,
            initial_page=args.page,
        )
        return 0

    if command == "apply-approval-review":
        data = repository.load()
        review = review_repository.load()
        source_revision = int(
            review.get("source_dataset_revision") or 0
        )
        current_revision = int(data.get("dataset_revision") or 0)
        if source_revision != current_revision:
            raise RuntimeError(
                "审核文件对应的正式数据版本已过期，请先运行 "
                "approval-review 重新生成审核文件"
            )
        decisions = import_json_decisions(review)
        apply_result = apply_review_decisions(
            data,
            review,
            decisions,
            config.data_root,
        )
        if apply_result["confirmed"] or apply_result["trashed"]:
            data["dataset_revision"] = current_revision + 1
        append_run(
            data,
            command,
            {
                "review_decisions_imported": len(decisions),
                "approval_pdfs_human_confirmed": apply_result["confirmed"],
                "approval_pdfs_moved_to_trash": apply_result["trashed"],
            },
        )
        repository.save(data)
        formal_html = export_summary_html(data, config.data_root)
        refreshed_review = build_approval_review(
            data,
            config.data_root,
            repo_root,
            existing=review,
        )
        review_path = review_repository.save(refreshed_review)
        review_html = export_approval_review_html(
            refreshed_review, config.data_root
        )
        print(
            json.dumps(
                {
                    **_status(repository),
                    "review_decisions_imported": len(decisions),
                    "approval_pdfs_human_confirmed": apply_result["confirmed"],
                    "approval_pdfs_moved_to_trash": apply_result["trashed"],
                    "formal_html": str(formal_html),
                    "review_json": str(review_path),
                    "review_html": str(review_html),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    if command in {
        "applications",
        "approval-pdfs",
        "worker-lists",
        "run",
    }:
        data = repository.load()
        if not repository.path.is_file():
            raise FileNotFoundError(
                "尚未建立 JSON 数据，请先执行 migrate 并确认迁移计划"
            )
        approval_count = 0
        worker_list_count = 0
        application_count = 0
        intake_stats: dict[str, int] = {}
        duplicate_application_count = deduplicate_applications(
            data,
            config.data_root,
            checkpoint=repository.save,
        )
        input_batch_id = str(uuid.uuid4())
        route_summary = route_input_files(
            data,
            config.data_root,
            repo_root,
            checkpoint=repository.save,
            input_batch_id=input_batch_id,
        )
        if command in {"applications", "run"}:
            application_count = intake_applications(
                data,
                config.data_root,
                repo_root,
                checkpoint=repository.save,
                input_batch_id=input_batch_id,
                intake_stats=intake_stats,
            )
            worker_list_count += intake_stats.get(
                "worker_lists_ingested", 0
            )
        if command in {"approval-pdfs", "run"}:
            approval_count = ingest_approval_pdfs(
                data,
                config.data_root,
                repo_root,
                checkpoint=repository.save,
            )
        if command in {"worker-lists", "run"}:
            worker_list_count += ingest_worker_lists(data, config.data_root)
        changed = (
            application_count
            + approval_count
            + worker_list_count
            + route_summary["input_files_routed"]
            + route_summary["input_duplicates_quarantined"]
            + duplicate_application_count
        )
        if changed:
            data["dataset_revision"] = int(
                data.get("dataset_revision") or 0
            ) + 1
        append_run(
            data,
            command,
            {
                "applications_ingested": application_count,
                "approval_pdfs_ingested": approval_count,
                "worker_lists_ingested": worker_list_count,
                "duplicate_applications_removed": (
                    duplicate_application_count
                ),
                **route_summary,
            },
        )
        repository.save(data)
        html_path = export_summary_html(data, config.data_root)
        review_path = ""
        review_html = ""
        if command in {"approval-pdfs", "run"}:
            review = build_approval_review(
                data,
                config.data_root,
                repo_root,
                existing=review_repository.load(),
            )
            repository.save(data)
            review_path = str(review_repository.save(review))
            review_html = str(
                export_approval_review_html(review, config.data_root)
            )
        _log_workflow_summary(
            _status(repository),
            applications_ingested=application_count,
            approval_pdfs_ingested=approval_count,
            worker_lists_ingested=worker_list_count,
            duplicate_applications_removed=duplicate_application_count,
            route_summary=route_summary,
            html_file=str(html_path),
            review_json=review_path,
            review_html=review_html,
        )
        return 0

    raise NotImplementedError(
        f"{command} 工作流将在数据迁移基础验证完成后启用"
    )


if __name__ == "__main__":
    sys.exit(main())
