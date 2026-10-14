# 更新记录

## 2026-10-14

- 恢复审批 PDF 的机器识别与自动筛选：识别带有中间描述的质保施工审批标题，并结合申请编号和最终审批状态判断文件类型。唯一且证据充分的匹配直接归档；无法确定的文件继续进入人工审核。汇总 HTML 根据正式 JSON 显示归档结果。
- 保留 HTML 中手工选择审批 PDF 的入口，用于自动匹配未成功的情况；它不替代自动识别流程。
- `run_archive.py` 增加 Windows 图形模式。窗口可更新归档、打开汇总和人工审核页面，退出时关闭本地服务；源码默认仍运行命令行归档。
- 新增 `assets/docflow_archive.png` 和 `assets/docflow_archive.ico`，以及 `build/build_run_archive_exe.ps1` 和 `requirements-build.txt`。打包程序为 `build/delivery/质保作业申请归档.exe`，构建脚本位于 `build` 目录。
- 更新项目根目录定位方式，使嵌套在 `build/delivery` 中的 EXE 能读取项目配置和本机 `common.env`。

## 使用和验证

将审批 PDF 放入资料目录的 `_input` 第一层，运行 `python run_archive.py` 或双击 EXE。自动匹配成功的文件应出现在对应案卷和正式汇总 HTML；未匹配的文件应出现在待人工审核页面。构建命令及配置说明见 [README](../README.md)。

本次验证：Python 测试 46 项通过；EXE 的归档与页面自检通过；移动到 `build/delivery` 后配置检查通过。
