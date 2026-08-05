"""申请材料入库入口。

配置文件：读取 config.yaml 和本机 common.env。
可选参数：--input-dir 可临时指定资料根目录。
示例：python archive_applications.py
输出：归档 Word、图片及专项材料，并刷新正式 JSON 和汇总 HTML。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("applications", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
