# DocFlow Renamer

质保作业申请案卷归档工具。项目以一份质保申请为一个案卷，完成申报材料入库、统一命名、审批 PDF 回传、JSON 归档和 Excel 人工审查视图生成。

## 核心原则

- `质保作业申请数据.json` 是唯一事实数据源。
- `质保作业申请汇总.xlsx` 只用于人工调阅，由 JSON 单向生成。
- 未自动匹配的审批 PDF 使用独立的 `审批PDF匹配审核.json/.xlsx` 闭环处理。
- 人工只编辑审核 Excel 的结果和备注；程序校验后先留存审核决定，再同步正式 JSON/Excel。
- Word 申请单是案卷锚点，每份申请建立独立目录。
- AI 客户端和识别缓存由所有支流程共享，同一内容不会重复识别。
- 文件移动、复制和隔离均记录在 JSON 的 `changes` 中。
- 不直接删除重复文件；重复件和已处理的公共源文件进入 `.docflow/quarantine`。

## 资料目录

```text
质保作业申请单/
├─ _inbox/                  # 人工放入的新申请材料和飞书审批 PDF
├─ _templates/
│  └─ 01 安全生产及消防安全协议（建工）.pdf
├─ _cases/
│  └─ 2026-07-24_维修冷塔_质保作业申请单/
│     ├─ 2026-07-24_维修冷塔_质保作业申请单.docx
│     ├─ 2026-07-24_维修冷塔_质保作业申请单_手签_01.jpg
│     ├─ 2026-07-24_维修冷塔_质保作业申请单_施工人员名单_01.jpg
│     ├─ 2026-07-24_维修冷塔_质保作业申请单_有限空间申请_01.pdf
│     ├─ 2026-07-24_维修冷塔_01 安全生产及消防安全协议（建工）.pdf
│     └─ 工程类-主体质保施工_编号：202607240001.pdf
├─ .docflow/
│  ├─ legacy/               # 迁移前的旧版汇总表
│  └─ quarantine/           # 可恢复的重复件和已处理公共源文件
├─ 质保作业申请数据.json
├─ 质保作业申请汇总.xlsx
├─ 审批PDF匹配审核.json
└─ 审批PDF匹配审核.xlsx
```

## 两阶段工作流

### 1. 申报材料入库

1. 人工将 Word、手签图片、施工人员名单和专项作业材料放入 `_inbox`。
2. 系统解析 Word 中的日期、施工内容、区域和危险作业。
3. 创建独立案卷目录并统一文件名。
4. 关联手签图片、人员名单和有限空间/高处作业材料。
5. 从 `_templates` 复制每份申请必需的安全协议。
6. 将案卷字段、文件路径、SHA-256、识别结果和缺失材料写入 JSON。
7. 根据 JSON 更新 Excel，供人工审查和填报飞书。

### 2. 审批 PDF 回传

1. 人工将飞书审批通过生成的 PDF 放入 `_inbox`。
2. 系统优先读取 PDF 文本，必要时调用本地多模态模型 OCR。
3. 按施工区域、内容、开始时间和结束时间匹配已有案卷。
4. 提取申请编号，规范命名并移动到对应案卷目录。
5. 更新 JSON 中的审批信息和案卷状态。
6. 根据 JSON 更新 Excel。

自动匹配要求唯一且证据完整。未自动匹配的 PDF 不使用宽松规则直接归档，而是进入独立人工审核流程：

1. 系统为每个未匹配 PDF 和每个尚未归档审批单的案卷建立候选关系。
2. 施工内容、施工区域及起止日期差异转成候选评分和可读匹配依据。
3. 候选关系写入 `审批PDF匹配审核.json`，并生成 `审批PDF匹配审核.xlsx`。
4. 人工在 Excel 中将唯一正确的一行改为 `确认匹配`，错误候选可标记为 `排除`。
5. 程序校验同一 PDF 只能确认一个案卷、案卷不能重复绑定、隐藏 ID 和 SHA-256 未被修改。
6. 审核决定写回审核 JSON，PDF 归档到案卷，正式 JSON 更新，再由正式 JSON 重建正式汇总 Excel。

## 案卷状态

- `materials_incomplete`：缺少必需申报材料。
- `materials_ready`：材料齐全，等待人工填报或回传审批 PDF。
- `approved`：审批 PDF 已匹配并归档。
- `needs_review`：识别、匹配或文件状态需要人工确认。

已审批案卷仍保留 `missing_material_types`，用于发现历史材料缺失。

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

默认从 `common.env` 读取资料目录：

```env
INPUT_PATH=D:\path\to\质保作业申请单
```

也可以使用 `--input-dir` 指定：

```powershell
python docflow_renamer.py --input-dir "D:\path\to\质保作业申请单" status
```

## 旧目录迁移

旧版平铺目录必须先演练。演练只生成计划，不移动资料：

```powershell
python docflow_renamer.py migrate
```

确认计划后，执行迁移时必须提供内容一致的备份目录：

```powershell
python docflow_renamer.py migrate `
  --apply `
  --backup-dir "D:\path\to\质保作业申请单_backup"
```

执行前会逐文件比较正式目录与备份的大小和 SHA-256。任何差异都会阻止迁移。迁移时还会在每次移动前后再次校验哈希。

旧版 Excel 只在首次迁移时读取已有 PDF 对应关系，随后归档到 `.docflow/legacy`。迁移完成后 Excel 不再作为缓存或数据输入。

## 日常命令

执行完整增量流程：

```powershell
python docflow_renamer.py run
```

单独执行支流程：

```powershell
python docflow_renamer.py applications
python docflow_renamer.py worker-lists
python docflow_renamer.py approval-pdfs
```

`approval-pdfs` 会自动刷新未匹配 PDF 的独立审核 JSON/Excel。也可以单独重新生成：

```powershell
python docflow_renamer.py approval-review
```

人工填写并关闭 `审批PDF匹配审核.xlsx` 后，应用审核结果：

```powershell
python docflow_renamer.py apply-approval-review
```

如果正式 JSON 在审核 Excel 生成后发生过版本变化，应用命令会拒绝过期结果，须先重新生成审核文件。

仅从 JSON 重新生成 Excel：

```powershell
python docflow_renamer.py export
```

查看状态及进行完整校验：

```powershell
python docflow_renamer.py status
python docflow_renamer.py validate
```

`validate` 会检查：

- 案卷 ID 和文件 ID 是否重复。
- JSON 中的目录和文件是否存在。
- 文件大小和 SHA-256 是否一致。
- Excel 工作表结构和汇总行数是否与 JSON 一致。

## Excel 人工审查视图

工作簿包含：

- `申请汇总`
- `待补材料`
- `待审批PDF`
- `已完成`
- `本次变更`
- `说明`

正式 Excel 中的 Word、图片、协议、审批 PDF 和案卷目录均提供本地超链接。正式 Excel 不作为数据输入。

独立的 `审批PDF匹配审核.xlsx` 包含：

- `待审核`：一行代表一个“审批 PDF—候选案卷”关系，按评分降序排列；黄色列允许人工填写。
- `已处理决定`：从审核 JSON 输出的历史确认和排除记录。
- `说明`：操作步骤、正式数据版本和候选数量。

## 兼容入口

以下脚本仅保留为薄入口，不再包含业务代码：

```powershell
python rename_pdfs.py
python copy_worker_list_images.py
```

它们分别调用统一工作流中的审批 PDF 和人员名单支流程。

## 项目结构

```text
docflow_renamer.py                 # 主命令薄入口
rename_pdfs.py                     # 审批 PDF 兼容入口
copy_worker_list_images.py         # 人员名单兼容入口
src/docflow_renamer/
├─ cli.py                          # 命令编排
├─ config.py                       # 配置
├─ constants.py                    # 目录、状态和材料类型
├─ repository.py                   # JSON 唯一数据源
├─ naming.py                       # 案卷及材料命名
├─ recognition.py                  # 共享 AI/OCR 与缓存
├─ migration.py                    # 旧目录迁移计划与执行
├─ workflows.py                    # 申报、人员名单、审批 PDF 支流程
├─ approval_review.py              # 未匹配审批 PDF 人工审核与回写
├─ excel_export.py                 # JSON → Excel
├─ validation.py                   # 数据和成果校验
└─ legacy.py                       # 迁移期间保留的解析/AI兼容层
```

## 开发验证

```powershell
python -m compileall -q docflow_renamer.py rename_pdfs.py copy_worker_list_images.py src tests
python -m unittest discover -s tests -v
python -m flake8 src tests docflow_renamer.py rename_pdfs.py copy_worker_list_images.py
```
