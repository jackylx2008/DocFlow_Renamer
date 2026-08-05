"""汇总页面重新生成入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python export_archive.py
输出：从正式 JSON 重新生成质保作业申请汇总 HTML。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("export", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
