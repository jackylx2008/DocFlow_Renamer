# Warranty Application Archive

质保作业申请案卷归档工具。项目以一份质保申请为一个案卷，完成申报材料入库、统一命名、审批 PDF 回传、JSON 归档和 Excel 人工审查视图生成。

## 核心原则

- `质保作业申请数据.json` 是唯一事实数据源。
- `质保作业申请汇总.xlsx` 只用于人工调阅，由 JSON 单向生成。
- 仅 `_inbox` 第一层中未自动匹配的审批 PDF 使用独立的
  `待人工审核匹配PDF.json/.xlsx` 闭环处理。
- 自动唯一匹配成功的审批 PDF 视为可信，直接归档，不进入人工审核。
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
├─ 待人工审核匹配PDF.json
└─ 待人工审核匹配PDF.xlsx
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

自动匹配要求唯一且证据完整；成功匹配的结果直接视为可信。只有 `_inbox` 第一层中仍未匹配的 PDF 才进入独立人工审核流程：

1. 施工区域和施工内容必须同时通过严格文本匹配；日期不淘汰候选，
   只根据相差天数降低候选自身匹配度。
2. 每个 PDF 只保留最高项；若存在第二名，再根据领先差值下调最终选择置信度。
3. 没有严格候选的 PDF 写入 `unresolved_pdfs`，并在 Excel 的
   `无严格候选` 工作表单独显示，不强行推荐案卷。
4. 结果写入 `待人工审核匹配PDF.json`，并生成
   `待人工审核匹配PDF.xlsx`。
5. 人工只在最高候选行选择 `确认匹配` 或 `排除`；排除后重新生成，
   程序会从剩余严格候选中选择下一名。
6. 程序校验案卷不能重复绑定、隐藏 ID 和 SHA-256 未被修改。
7. 审核决定写回审核 JSON，PDF 归档到案卷，正式 JSON 更新，
   再由正式 JSON 重建正式汇总 Excel。

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
python warranty_application_archive.py --input-dir "D:\path\to\质保作业申请单" status
```

## 旧目录迁移

旧版平铺目录必须先演练。演练只生成计划，不移动资料：

```powershell
python warranty_application_archive.py migrate
```

确认计划后，执行迁移时必须提供内容一致的备份目录：

```powershell
python warranty_application_archive.py migrate `
  --apply `
  --backup-dir "D:\path\to\质保作业申请单_backup"
```

执行前会逐文件比较正式目录与备份的大小和 SHA-256。任何差异都会阻止迁移。迁移时还会在每次移动前后再次校验哈希。

旧版 Excel 只在首次迁移时读取已有 PDF 对应关系，随后归档到 `.docflow/legacy`。迁移完成后 Excel 不再作为缓存或数据输入。

## 日常命令

执行完整增量流程：

```powershell
python warranty_application_archive.py run
```

单独执行支流程：

```powershell
python warranty_application_archive.py applications
python warranty_application_archive.py worker-lists
python warranty_application_archive.py approval-pdfs
```

`approval-pdfs` 会自动刷新未匹配 PDF 的独立审核 JSON/Excel。也可以单独重新生成：

```powershell
python warranty_application_archive.py approval-review
```

人工填写并关闭 `待人工审核匹配PDF.xlsx` 后，应用审核结果：

```powershell
python warranty_application_archive.py apply-approval-review
```

如果正式 JSON 在审核 Excel 生成后发生过版本变化，应用命令会拒绝过期结果，须先重新生成审核文件。

仅从 JSON 重新生成 Excel：

```powershell
python warranty_application_archive.py export
```

查看状态及进行完整校验：

```powershell
python warranty_application_archive.py status
python warranty_application_archive.py validate
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

独立的 `待人工审核匹配PDF.xlsx` 包含：

- `待审核`：每个审批 PDF 最多一行，只展示最高置信度严格候选；
  黄色列允许人工填写。
- `无严格候选`：区域和内容无法同时严格命中的 PDF，仅供人工查看，
  不能直接确认归档。
- `已处理决定`：从审核 JSON 输出的历史确认和排除记录及回写状态。
- `说明`：匹配规则、置信度含义、操作步骤和数据版本。

置信度规则：

- 施工区域和施工内容严格命中占 80%，属于候选准入条件。
- 开始、结束日期合计占 20%；日期相差超过 31 天也不会淘汰候选，
  但日期部分不再加分。
- 存在多个严格候选时，最终选择置信度还会根据第一名与第二名的
  领先差值下调；并列第一会显示为低置信度。
- PDF 没有足够的内嵌文字时，程序调用本地 AI OCR，并在审核产物中
  记录识别方式；已有 SHA-256 识别缓存会直接复用。

## 统一入口

项目根目录只保留一个 Python 入口：

```powershell
python warranty_application_archive.py <command>
```

审批 PDF、人员名单和申请材料均通过子命令进入统一工作流，不再维护独立兼容脚本。

## 项目结构

```text
warranty_application_archive.py   # 唯一命令入口
src/warranty_application_archive/
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
python -m compileall -q warranty_application_archive.py src tests
python -m unittest discover -s tests -v
python -m flake8 src tests warranty_application_archive.py
```
