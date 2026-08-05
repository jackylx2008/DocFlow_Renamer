"""审批 PDF 人工审核结果回写入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python apply_approval_review.py
输出：验证并回写审核决定，随后刷新正式 JSON 和相关 HTML。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("apply-approval-review", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
