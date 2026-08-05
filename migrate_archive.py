"""旧版平铺资料迁移入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir、--plan-output、--apply；执行迁移时必须提供 --backup-dir。
示例：python migrate_archive.py --plan-output output/migration_plan.json
输出：先生成迁移计划；显式指定 --apply 后才会移动文件并建立正式 JSON。
"""

from __future__ import annotations

import argparse
from pathlib import Path

from entry_bootstrap import run_command


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input-dir", type=Path, help="资料根目录")
    parser.add_argument("--plan-output", type=Path, help="迁移计划 JSON 路径")
    parser.add_argument("--apply", action="store_true", help="执行迁移")
    parser.add_argument("--backup-dir", type=Path, help="迁移前核对的备份目录")
    args = parser.parse_args()
    forwarded: list[str] = []
    if args.input_dir:
        forwarded.extend(["--input-dir", str(args.input_dir)])
    forwarded.append("migrate")
    if args.plan_output:
        forwarded.extend(["--plan-output", str(args.plan_output)])
    if args.apply:
        forwarded.append("--apply")
    if args.backup_dir:
        forwarded.extend(["--backup-dir", str(args.backup_dir)])
    return run_command(forwarded)


if __name__ == "__main__":
    raise SystemExit(main())
