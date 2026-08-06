"""本地归档页面服务入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir、--host、--port、--page 和 --no-open。
示例：python serve_archive_review.py --page summary --port 0
输出：启动本地 HTTP 服务，并按配置打开汇总页或审批审核页。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from warranty_application_archive.flows.application_flow import (  # noqa: E402
    serve_archive_pages,
)


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
    return serve_archive_pages(
        input_dir=args.input_dir,
        host=args.host,
        port=args.port,
        page=args.page,
        open_browser=not args.no_open,
    )


if __name__ == "__main__":
    raise SystemExit(main())
