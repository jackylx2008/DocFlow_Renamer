# 更新记录

## 2026-11-03（入口整理）

- 根目录统一为 `main.py` 一个 Python 启动入口；归档、旧目录迁移和审批审核的操作入口由本地 HTML 工作台承接。
- 将日志配置移入 `src/warranty_application_archive/`，服务启动器和生成的 Windows/macOS 页面启动脚本统一调用 `main.py`。
- 没有正式 JSON 时，`python main.py` 打开旧目录迁移页面；用户可以先预览迁移计划，再填写备份路径并确认执行。
- 更新项目布局、构建入口和 README。
- Windows 启动 HTML 时使用独立 Chrome 窗口，并将本次页面窗口移动到主显示器工作区后最大化；未安装 Chrome 时回退到默认浏览器。
- HTML 文件输入支持一次拖入多个文件及分批累加；批量上传并发执行，单个文件失败不会阻断其他文件，失败项会保留在页面队列并显示原因。

## 2026-11-03

- 将本机 HTML 汇总页扩展为文件输入工作台，支持选择或拖入 Word、PDF 和常见图片文件，并通过本机服务接收至 `_input`。
- 增加“识别并归档”操作，复用现有 AI/OCR 识别与归档编排；无法唯一匹配的审批 PDF 继续进入人工审核页面，结果保存后刷新正式 JSON 和汇总 HTML。
- 输入服务限制在 localhost，单文件上限 50 MB；拒绝不支持的文件类型和覆盖同名文件。
- 更新 README 与架构说明，明确 HTML 输入、识别、人工审核和汇总展示流程。页面沿用本项目现有蓝灰色卡片与表格样式，并参考 B24DesignArchiver 的信息布局与筛选交互。

## 2026-10-14

- 恢复审批 PDF 的机器识别与自动筛选：识别带有中间描述的质保施工审批标题，并结合申请编号和最终审批状态判断文件类型。唯一且证据充分的匹配直接归档；无法确定的文件继续进入人工审核。汇总 HTML 根据正式 JSON 显示归档结果。
- 保留 HTML 中手工选择审批 PDF 的入口，用于自动匹配未成功的情况；它不替代自动识别流程。
- `run_archive.py` 增加 Windows 图形模式。窗口可更新归档、打开汇总和人工审核页面，退出时关闭本地服务；源码默认仍运行命令行归档。
- 新增 `assets/docflow_archive.png` 和 `assets/docflow_archive.ico`，以及 `build/build_run_archive_exe.ps1` 和 `requirements-build.txt`。打包程序为 `build/delivery/质保作业申请归档.exe`，构建脚本位于 `build` 目录。
- 更新项目根目录定位方式，使嵌套在 `build/delivery` 中的 EXE 能读取项目配置和本机 `common.env`。

## 使用和验证

将审批 PDF 放入资料目录的 `_input` 第一层，运行 `python run_archive.py` 或双击 EXE。自动匹配成功的文件应出现在对应案卷和正式汇总 HTML；未匹配的文件应出现在待人工审核页面。构建命令及配置说明见 [README](../README.md)。

本次验证：Python 测试 46 项通过；EXE 的归档与页面自检通过；移动到 `build/delivery` 后配置检查通过。
