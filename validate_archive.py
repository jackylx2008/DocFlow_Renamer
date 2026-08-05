"""正式归档数据校验入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python validate_archive.py
输出：打印 JSON、案卷文件和汇总 HTML 的结构化校验报告。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("validate", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
