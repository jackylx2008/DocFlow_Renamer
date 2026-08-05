"""历史专项材料重新分类入口。

配置文件：读取 config.yaml 和本机 common.env，并复用已保存的 OCR 缓存。
可选参数：--input-dir 可临时指定资料根目录。
示例：python reclassify_materials.py
输出：更新材料分类、正式 JSON、汇总 HTML 和审核页面。
"""

from __future__ import annotations

from entry_bootstrap import run_simple_entry


def main() -> int:
    return run_simple_entry("reclassify-materials", __doc__ or "")


if __name__ == "__main__":
    raise SystemExit(main())
