from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Sequence

from . import legacy
from .approval_review import (
    ApprovalReviewRepository,
    apply_review_decisions,
    build_approval_review,
    export_approval_review_excel,
    import_excel_decisions,
)
from .config import AppConfig
from .excel_export import export_excel
from .migration import (
    apply_migration_plan,
    build_migration_plan,
    verify_backup,
)
from .repository import JsonRepository
from .validation import validate_dataset, validate_excel
from .workflows import (
    append_run,
    intake_applications,
    ingest_approval_pdfs,
    ingest_worker_lists,
)


LOGGER = logging.getLogger(__name__)


def _repo_root() -> Path:
    current_directory = Path.cwd().resolve()
    if (current_directory / "common.env").is_file():
        return current_directory
    return Path(__file__).resolve().parents[2]


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
    subparsers.add_parser("validate", help="校验 JSON、案卷文件和 Excel")
    subparsers.add_parser("export", help="从 JSON 重新生成 Excel")
    subparsers.add_parser("applications", help="处理待入库质保申请材料")
    subparsers.add_parser("approval-pdfs", help="处理待入库审批 PDF")
    subparsers.add_parser(
        "approval-review",
        help="生成未匹配审批 PDF 的独立审核 JSON 和 Excel",
    )
    subparsers.add_parser(
        "apply-approval-review",
        help="读取人工填写的审核 Excel 并同步正式 JSON/Excel",
    )
    subparsers.add_parser("worker-lists", help="处理待入库施工人员名单")
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


def main(argv: Sequence[str] | None = None) -> int:
    repo_root = _repo_root()
    legacy.setup_logging(repo_root)
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    command = args.command or "run"
    config = AppConfig.resolve(repo_root, args.input_dir)
    if not config.data_root.is_dir():
        raise NotADirectoryError(f"资料根目录不存在: {config.data_root}")
    repository = JsonRepository(config.data_root)
    review_repository = ApprovalReviewRepository(config.data_root)

    if command == "migrate":
        plan = build_migration_plan(config.data_root)
        plan_output = args.plan_output or (
            repo_root / "output" / "migration_plan.json"
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
        excel_path = export_excel(dataset, config.data_root)
        LOGGER.info("Excel 已从 JSON 生成: %s", excel_path)
        print(
            json.dumps(
                {
                    **_status(repository),
                    "excel_file": str(excel_path),
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
        excel_report = validate_excel(
            config.data_root,
            len(data.get("applications") or []),
        )
        report = {
            "data": data_report,
            "excel": excel_report,
        }
        print(json.dumps(report, ensure_ascii=False, indent=2))
        return 0 if data_report["valid"] and excel_report["valid"] else 1

    if command == "export":
        data = repository.load()
        output_path = export_excel(data, config.data_root)
        LOGGER.info("Excel 已从 JSON 生成: %s", output_path)
        print(
            json.dumps(
                {
                    **_status(repository),
                    "excel_file": str(output_path),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
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
        review_excel = export_approval_review_excel(
            review, config.data_root
        )
        print(
            json.dumps(
                {
                    **_status(repository),
                    "review_candidates": len(
                        review.get("pending_reviews") or []
                    ),
                    "review_json": str(review_path),
                    "review_excel": str(review_excel),
                },
                ensure_ascii=False,
                indent=2,
            )
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
        decisions = import_excel_decisions(review, config.data_root)
        applied = apply_review_decisions(
            data,
            review,
            decisions,
            config.data_root,
        )
        if applied:
            data["dataset_revision"] = current_revision + 1
        append_run(
            data,
            command,
            {
                "review_decisions_imported": len(decisions),
                "approval_pdfs_human_confirmed": applied,
            },
        )
        repository.save(data)
        formal_excel = export_excel(data, config.data_root)
        refreshed_review = build_approval_review(
            data,
            config.data_root,
            repo_root,
            existing=review,
        )
        review_path = review_repository.save(refreshed_review)
        review_excel = export_approval_review_excel(
            refreshed_review, config.data_root
        )
        print(
            json.dumps(
                {
                    **_status(repository),
                    "review_decisions_imported": len(decisions),
                    "approval_pdfs_human_confirmed": applied,
                    "formal_excel": str(formal_excel),
                    "review_json": str(review_path),
                    "review_excel": str(review_excel),
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
        if command in {"applications", "run"}:
            application_count = intake_applications(
                data,
                config.data_root,
                repo_root,
                checkpoint=repository.save,
            )
        if command in {"approval-pdfs", "run"}:
            approval_count = ingest_approval_pdfs(
                data,
                config.data_root,
                repo_root,
                checkpoint=repository.save,
            )
        if command in {"worker-lists", "run"}:
            worker_list_count = ingest_worker_lists(data, config.data_root)
        changed = application_count + approval_count + worker_list_count
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
            },
        )
        repository.save(data)
        excel_path = export_excel(data, config.data_root)
        review_path = ""
        review_excel = ""
        if command in {"approval-pdfs", "run"}:
            review = build_approval_review(
                data,
                config.data_root,
                repo_root,
                existing=review_repository.load(),
            )
            repository.save(data)
            review_path = str(review_repository.save(review))
            review_excel = str(
                export_approval_review_excel(review, config.data_root)
            )
        print(
            json.dumps(
                {
                    **_status(repository),
                    "applications_ingested": application_count,
                    "approval_pdfs_ingested": approval_count,
                    "worker_lists_ingested": worker_list_count,
                    "excel_file": str(excel_path),
                    "review_json": review_path,
                    "review_excel": review_excel,
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0

    raise NotImplementedError(
        f"{command} 工作流将在数据迁移基础验证完成后启用"
    )


if __name__ == "__main__":
    sys.exit(main())
