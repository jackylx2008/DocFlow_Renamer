"""完整增量归档入口。

配置文件：读取 config.yaml，并用本机 common.env 覆盖路径和模型配置。
可选参数：--input-dir 可临时指定资料根目录。
示例：python run_archive.py
输出：更新正式 JSON、案卷目录、汇总 HTML、审核页面和 logs/run_archive.log。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("run", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
