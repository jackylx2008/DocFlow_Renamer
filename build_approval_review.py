"""审批 PDF 人工审核材料生成入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python build_approval_review.py
输出：生成待人工审核匹配 PDF 的 JSON、HTML 和启动器。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("approval-review", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
