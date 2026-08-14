from __future__ import annotations

import json
import logging
import uuid
from pathlib import Path

from logging_config import setup_logger

from ..config_loader import AppConfig
from ..modules.repository import JsonRepository
from ..modules.summary_html import export_summary_html
from .approval_review_flow import (
    ApprovalReviewRepository,
    build_approval_review,
)
from .approval_review_web_flow import (
    export_approval_review_html,
    serve_approval_review,
)
from .archive_flow import (
    append_run,
    deduplicate_applications,
    ingest_approval_pdfs,
    ingest_worker_lists,
    intake_applications,
    route_input_files,
)
from .migration_flow import (
    apply_migration_plan,
    build_migration_plan,
    verify_backup,
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


def _open_runtime(
    input_dir: Path | None,
) -> tuple[
    Path,
    AppConfig,
    JsonRepository,
    ApprovalReviewRepository,
]:
    repo_root = _repo_root()
    config = AppConfig.resolve(repo_root, input_dir)
    setup_logger(log_level=config.log_level, log_dir=config.log_dir)
    if not config.data_root.is_dir():
        raise NotADirectoryError(f"资料根目录不存在: {config.data_root}")
    return (
        repo_root,
        config,
        JsonRepository(config.data_root),
        ApprovalReviewRepository(config.data_root),
    )


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


def run_archive(input_dir: Path | None = None) -> int:
    repo_root, config, repository, review_repository = _open_runtime(
        input_dir
    )
    data = repository.load()
    if not repository.path.is_file():
        raise FileNotFoundError(
            "尚未建立 JSON 数据，请先执行 migrate_archive.py 并确认迁移计划"
        )

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
    intake_stats: dict[str, int] = {}
    application_count = intake_applications(
        data,
        config.data_root,
        repo_root,
        checkpoint=repository.save,
        input_batch_id=input_batch_id,
        intake_stats=intake_stats,
    )
    worker_list_count = intake_stats.get("worker_lists_ingested", 0)
    approval_count = ingest_approval_pdfs(
        data,
        config.data_root,
        repo_root,
        checkpoint=repository.save,
    )
    worker_list_count += ingest_worker_lists(
        data,
        config.data_root,
        input_batch_id=input_batch_id,
    )
    changed = (
        application_count
        + approval_count
        + worker_list_count
        + route_summary["input_files_routed"]
        + route_summary["input_duplicates_quarantined"]
        + duplicate_application_count
    )
    if changed:
        data["dataset_revision"] = int(data.get("dataset_revision") or 0) + 1
    append_run(
        data,
        "run",
        {
            "applications_ingested": application_count,
            "approval_pdfs_ingested": approval_count,
            "worker_lists_ingested": worker_list_count,
            "duplicate_applications_removed": duplicate_application_count,
            **route_summary,
        },
    )
    repository.save(data)
    html_path = export_summary_html(data, config.data_root)
    review = build_approval_review(
        data,
        config.data_root,
        repo_root,
        existing=review_repository.load(),
    )
    repository.save(data)
    review_path = review_repository.save(review)
    review_html = export_approval_review_html(review, config.data_root)
    _log_workflow_summary(
        _status(repository),
        applications_ingested=application_count,
        approval_pdfs_ingested=approval_count,
        worker_lists_ingested=worker_list_count,
        duplicate_applications_removed=duplicate_application_count,
        route_summary=route_summary,
        html_file=str(html_path),
        review_json=str(review_path),
        review_html=str(review_html),
    )
    return 0


def migrate_archive(
    *,
    input_dir: Path | None = None,
    plan_output: Path | None = None,
    apply: bool = False,
    backup_dir: Path | None = None,
) -> int:
    _, config, repository, _ = _open_runtime(input_dir)
    plan = build_migration_plan(config.data_root)
    output = plan_output or config.output_dir / "migration_plan.json"
    _write_json(output, plan.to_dict())
    LOGGER.info(
        "迁移计划: %s 个案卷，%s 个操作，%s 个警告",
        len(plan.applications),
        len(plan.operations),
        len(plan.warnings),
    )
    LOGGER.info("迁移计划已保存: %s", output)
    if not apply:
        print(
            json.dumps(
                {
                    "mode": "dry-run",
                    "applications": len(plan.applications),
                    "operations": len(plan.operations),
                    "warnings": plan.warnings,
                    "plan": str(output),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    if backup_dir is None:
        raise ValueError("执行迁移必须提供 --backup-dir")
    verify_backup(config.data_root, backup_dir)
    dataset = apply_migration_plan(plan)
    data_path = repository.save(dataset)
    LOGGER.info("迁移完成，JSON 已保存: %s", data_path)
    html_path = export_summary_html(dataset, config.data_root)
    LOGGER.info("汇总 HTML 已从 JSON 生成: %s", html_path)
    print(
        json.dumps(
            {**_status(repository), "html_file": str(html_path)},
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


def serve_archive_pages(
    *,
    input_dir: Path | None = None,
    host: str = "127.0.0.1",
    port: int = 8765,
    page: str = "review",
    open_browser: bool = True,
) -> int:
    _, config, repository, review_repository = _open_runtime(input_dir)
    if page == "review":
        if not review_repository.path.is_file():
            raise FileNotFoundError(
                "尚未生成审核 JSON，请先运行 run_archive.py"
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
        host=host,
        port=port,
        open_browser=open_browser,
        initial_page=page,
    )
    return 0
