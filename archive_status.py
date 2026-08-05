"""归档状态查询入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python archive_status.py
输出：在控制台打印正式数据版本、案卷数量和状态统计 JSON。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("status", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
