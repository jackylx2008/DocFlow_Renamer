from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_root_python_files_are_only_operator_interfaces() -> None:
    root_python_files = {path.name for path in ROOT.glob("*.py")}
    assert root_python_files == {
        "logging_config.py",
        "migrate_archive.py",
        "run_archive.py",
        "serve_archive_review.py",
    }


def test_business_flows_live_under_package_flows() -> None:
    flows_dir = ROOT / "src" / "warranty_application_archive" / "flows"
    assert {
        "approval_review_flow.py",
        "approval_review_web_flow.py",
        "archive_flow.py",
        "application_flow.py",
        "migration_flow.py",
    }.issubset({path.name for path in flows_dir.glob("*.py")})
