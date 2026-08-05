"""审批 PDF 入库入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python archive_approval_pdfs.py
输出：匹配审批 PDF，并刷新正式数据及待人工审核页面。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("approval-pdfs", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
