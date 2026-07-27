# Warranty Application Archive

质保作业申请案卷归档工具。项目以一份质保申请为一个案卷，完成申报材料入库、统一命名、审批 PDF 回传、JSON 归档和 HTML 人工审查视图生成。

## 核心原则

- `质保作业申请数据.json` 是唯一事实数据源。
- `质保作业申请汇总.html` 只用于人工调阅，由 JSON 单向生成。
- 仅 `_inbox` 第一层中未自动匹配的审批 PDF 使用独立的
  `待人工审核匹配PDF.json/.html` 闭环处理。
- 自动唯一匹配成功的审批 PDF 视为可信，直接归档，不进入人工审核。
- 人工在本地 HTML 页面填写审核结果和备注；保存时校验并立即同步审核
  JSON、正式 JSON 和正式汇总 HTML。
- Word 申请单是案卷锚点，每份申请建立独立目录。
- AI 客户端和识别缓存由所有支流程共享，同一内容不会重复识别。
- 文件移动、复制和隔离均记录在 JSON 的 `changes` 中。
- 不直接删除重复文件；重复件和已处理的公共源文件进入 `.docflow/quarantine`。

## 资料目录

```text
质保作业申请单/
├─ _input/                  # 唯一人工投放入口
├─ _inbox/                  # 程序分类后的内部待处理区
├─ _trash/                  # 人工判定删除的 PDF，可恢复
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
├─ 质保作业申请汇总.html
├─ 待人工审核匹配PDF.json
├─ 待人工审核匹配PDF.html
├─ 打开待人工审核匹配PDF.cmd       # Windows
└─ 打开待人工审核匹配PDF.command   # macOS
```

## 两阶段工作流

### 1. 申报材料入库

1. 人工将 Word、手签图片、施工人员名单和专项作业材料放入 `_input`。
2. 程序将 Word、图片和非审批 PDF 分类为申请材料并转入内部 `_inbox`。
3. 系统解析 Word 中的日期、施工内容、区域和危险作业。
4. 创建独立案卷目录并统一文件名。
5. 关联手签图片、人员名单和有限空间/高处作业材料。
6. 从 `_templates` 复制每份申请必需的安全协议。
7. 将案卷字段、文件路径、SHA-256、识别结果和缺失材料写入 JSON。
8. 根据 JSON 更新汇总 HTML，供人工审查和填报飞书。

### 2. 审批 PDF 回传

1. 人工将飞书审批通过生成的 PDF 放入 `_input`。
2. 系统优先读取 PDF 文本，必要时调用本地多模态模型 OCR，判断是否为
   “工程类-主体质保施工”审批 PDF。
3. 审批 PDF 转入内部 `_inbox` 的审批归档流程；其他 PDF 标记为申请材料，
   不会进入审批匹配。
4. 按施工区域、内容、开始时间和结束时间匹配已有案卷。
5. 提取申请编号，规范命名并移动到对应案卷目录。
6. 更新 JSON 中的审批信息和案卷状态。
7. 根据 JSON 更新汇总 HTML。

自动匹配要求唯一且证据完整；成功匹配的结果直接视为可信。只有 `_inbox` 第一层中仍未匹配的 PDF 才进入独立人工审核流程：

1. 施工区域和施工内容必须同时通过严格文本匹配；日期不淘汰候选，
   只根据相差天数降低候选自身匹配度。
2. 每个 PDF 只保留最高项；若存在第二名，再根据领先差值下调最终选择置信度。
3. 没有严格候选的 PDF 写入 `unresolved_pdfs`，并在 HTML 的
   “无严格候选”页签单独显示，不强行推荐案卷。
4. 结果写入 `待人工审核匹配PDF.json`，并生成
   `待人工审核匹配PDF.html`。
5. 人工可选择 `确认匹配`、`排除`，或点击“移至 `_trash`”，再点击
   “保存并执行审核结果”。
6. 程序校验案卷不能重复绑定、隐藏 ID 和 SHA-256 未被修改。
7. 保存成功后立即执行：确认项归档到案卷，删除项移动到 `_trash`，
   正式 JSON、正式汇总 HTML 和审核 JSON/HTML 同步刷新；已处理条目
   从待审核列表消失。

## 案卷状态

- `materials_incomplete`：缺少必需申报材料。
- `materials_ready`：材料齐全，等待人工填报或回传审批 PDF。
- `approved`：审批 PDF 已匹配并归档。
- `terminated`：案卷已人工终止；保留全部数据与文件，但不再进入待补材料、
  待审批 PDF 或审批匹配流程，汇总页整行显示为灰色。
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

日常只需将文件放入 `_input` 第一层。程序不递归扫描子目录；支持 `.docx`、
`.jpg`、`.jpeg`、`.png` 和 `.pdf`。PDF 分类结果保存在正式 JSON 的
`input_routes` 中，识别文本按 SHA-256 缓存。

单独执行支流程：

```powershell
python warranty_application_archive.py applications
python warranty_application_archive.py worker-lists
python warranty_application_archive.py approval-pdfs
```

`approval-pdfs` 会自动刷新未匹配 PDF 的独立审核 JSON/HTML。也可以单独重新生成：

```powershell
python warranty_application_archive.py approval-review
```

Windows 双击 `打开待人工审核匹配PDF.cmd`，macOS 双击
`打开待人工审核匹配PDF.command` 启动审核。程序会自动在浏览器中打开
HTML 页面；“保存并执行审核结果”固定更新同目录的
`待人工审核匹配PDF.json`，不会询问保存位置，并立即完成确认归档、
移入 `_trash` 及正式数据刷新。直接双击 HTML 只能预览，不能写入本地文件。

也可以从命令行启动：

```powershell
python warranty_application_archive.py approval-review-server
```

正常使用网页时不再需要第二步应用命令。下面的命令只用于兼容处理旧版页面
已经保存、但尚未执行的审核决定：

```powershell
python warranty_application_archive.py apply-approval-review
```

如果正式 JSON 在审核页面生成后发生过版本变化，页面保存和应用命令都会
拒绝过期结果，须先重新生成审核文件。

仅从 JSON 重新生成汇总 HTML：

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
- HTML 页签结构、汇总行数和本地链接是否与 JSON 一致。

## HTML 人工审查视图

`质保作业申请汇总.html` 包含以下页签：

- `申请汇总`
- `待补材料`
- `待审批PDF`
- `已完成`

“说明”和“本次变更”页签不再生成；文件变更记录只保留在正式 JSON 的
`changes` 中。Word、图片、协议、审批 PDF 和案卷目录均提供本地链接；
页面支持页签切换、当前页签搜索、固定表头和打印。正式汇总 HTML 不作为
数据输入。

需要把页面中的文档复制到其他位置时，应先通过对应系统的独立启动器打开：

- Windows：`打开质保作业申请汇总.cmd`
- macOS：`打开质保作业申请汇总.command`

macOS 启动器不会保存 Windows 绝对路径。资料目录自动取 `.command` 文件
所在目录；项目目录会从资料目录逐级向上查找
`Python/Project/DocFlow_Renamer`，并兼容 `~/Clooustation`、
`~/CloudStation` 和 `~/Cloudstation`。Python 依次尝试
`~/anaconda3/bin/python`、`~/opt/anaconda3/bin/python`、
`/opt/anaconda3/bin/python` 和系统 `python3`。特殊安装位置可通过
`DOCFLOW_PROJECT_ROOT` 和 `DOCFLOW_PYTHON` 环境变量指定。

再右键文档链接并选择“复制文件（可直接粘贴）”。程序会把真实文件写入
系统文件剪贴板，随后可在 Windows 资源管理器或 macOS Finder 等位置直接
粘贴。直接双击 HTML 仍可打开文档，但受浏览器安全限制，不能复制为可粘贴
的文件。

原 `质保作业申请汇总.xlsx` 不再生成；首次生成正式汇总 HTML 时会将已有
文件移入 `.docflow/legacy`。

独立的 `待人工审核匹配PDF.html` 包含：

- `最高候选`：每个审批 PDF 最多一张卡片，对比 PDF 识别内容和推荐案卷，
  可填写审核结果及备注，也可标记移至 `_trash`。
- `无严格候选`：区域和内容无法同时严格命中的 PDF，可继续保留或标记移至
  `_trash`。
- `已处理记录`：从审核 JSON 展示历史确认、排除和移至 `_trash` 记录。
- 页面会显示置信度、严格候选数、匹配依据和正式数据版本。

人工审核页面使用 Windows 的 `打开待人工审核匹配PDF.cmd` 或 macOS 的
`打开待人工审核匹配PDF.command` 启动时，同样可右键文档链接并选择
“复制文件（可直接粘贴）”。

原 `待人工审核匹配PDF.xlsx` 不再生成；首次生成 HTML 时会将已有文件移入
`.docflow/legacy`，保留历史记录。

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
├─ approval_review.py              # 未匹配审批 PDF 候选、决定与正式回写
├─ approval_review_web.py          # 本地 HTML 审核页与审核 JSON 回写服务
├─ summary_html.py                 # JSON → 正式汇总 HTML
├─ validation.py                   # 数据和成果校验
└─ legacy.py                       # 迁移期间保留的解析/AI兼容层
```

## 开发验证

```powershell
python -m compileall -q warranty_application_archive.py src tests
python -m unittest discover -s tests -v
python -m flake8 src tests warranty_application_archive.py
```
