"""完整增量归档入口。

配置文件：读取 config.yaml，并用本机 common.env 覆盖路径和模型配置。
可选参数：--input-dir 可临时指定资料根目录。
示例：python run_archive.py
输出：更新正式 JSON、案卷目录、汇总 HTML、审核页面和 logs/run_archive.log。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from warranty_application_archive.flows.application_flow import (  # noqa: E402
    run_archive,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--input-dir", type=Path, help="资料根目录")
    args = parser.parse_args()
    return run_archive(args.input_dir)


if __name__ == "__main__":
    raise SystemExit(main())
