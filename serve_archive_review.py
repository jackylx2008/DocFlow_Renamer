"""本地归档页面服务入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir、--host、--port、--page 和 --no-open。
示例：python serve_archive_review.py --page summary --port 0
输出：启动本地 HTTP 服务，并按配置打开汇总页或审批审核页。
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
    parser.add_argument("--host", default="127.0.0.1", help="监听地址")
    parser.add_argument("--port", type=int, default=8765, help="监听端口")
    parser.add_argument(
        "--page",
        choices=("review", "summary"),
        default="review",
        help="初始页面",
    )
    parser.add_argument("--no-open", action="store_true", help="不自动打开浏览器")
    args = parser.parse_args()
    forwarded: list[str] = []
    if args.input_dir:
        forwarded.extend(["--input-dir", str(args.input_dir)])
    forwarded.extend(
        [
            "approval-review-server",
            "--host",
            args.host,
            "--port",
            str(args.port),
            "--page",
            args.page,
        ]
    )
    if args.no_open:
        forwarded.append("--no-open")
    return run_command(forwarded)


if __name__ == "__main__":
    raise SystemExit(main())
