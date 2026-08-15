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

1. 人工将 Word、手签图片、施工人员名单和专项作业材料放入 `_input`；施工
   人员名单图片的文件名必须包含“工人名单”。
2. 程序在移动文件和写入 JSON 前，先比较本批次“工人名单”图片数与质保单
   Word 数。数量一致时才按文件名稳定排序一一对应，并继续将材料分类转入
   内部 `_inbox`；数量不一致时直接报错并终止整个批次，所有文件原样保留在
   `_input`，不执行文件保存、案卷入库或 JSON 写入。若本批 Word 全部命中
   已有申请且未投放工人名单，则视为“仅覆盖更新”批次，可直接进入更新流程。
3. 系统解析 Word 中的日期、施工内容、区域和危险作业。
4. 入库前先比较施工开始时间、施工结束时间、施工内容和施工区域；四项均
   相同即判定为已有申请，以新 Word 重新解析并覆盖该申请的 JSON 业务字段
   和案卷 Word。旧 Word 先移入 `.docflow/quarantine` 以便恢复；案卷目录
   不存在时自动创建并补入安全协议。如果历史数据中已有多条相同记录，按
   实际存在的必需资料、审批资料和附件数量选择更完整的一条，并把其他记录
   中的有效附件合并进去；完整度相同时优先保留不带 `_02` 等数字后缀的原
   案卷。其余案卷从汇总中移除，并整体移入隔离区以便恢复。
5. 一次放入 `_input` 的申请材料默认属于同一个申请批次；同批只有一份
   Word 时，随机文件名的图片也关联到该 Word 对应案卷。重复 Word 不新建
   案卷，但同批附件可以补充到已有原始案卷。图片后续单独补投时，系统会
   在已有未终止案卷中先严格匹配施工区域和起止日期，再按 OCR 施工内容的
   严格命中或唯一高相似度结果关联；候选不唯一时仍保留在 `_inbox`。
6. 图片 OCR 同时包含“姓名、性别、电话”时识别为施工人员名单；包含
   “质保申请单”或“质保作业申请单”时识别为手签申请单。有限空间、高处、
   动火、危大工程和配电室接电等专项材料，必须出现对应专用申请标题，或
   对应作业选项后的方框已明确勾选；申请单模板中的未勾选选项不再作为
   专项材料分类依据。文件名包含“工人名单”的同批图片不依赖 OCR 内容或
   图片哈希区分；通过步骤 2 的数量校验后，按案卷名稳定排序一对一入库。
7. 创建独立案卷目录并统一文件名。
8. 关联手签图片、人员名单和有限空间/高处作业材料。
9. 从 `_templates` 复制每份申请必需的安全协议。
10. 将案卷字段、文件路径、SHA-256、识别结果和缺失材料写入 JSON。
11. 根据 JSON 更新汇总 HTML，供人工审查和填报飞书。

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

1. 施工区域必须通过严格文本匹配（位置代码中紧邻数字的 OCR `i/l` 混淆会先
   归一化）；申请记录的施工内容经过规范化后，必须完整包含在审批 PDF 的
   “施工内容”字段中，审批 PDF 可以带有额外补充说明。日期不淘汰候选，只
   根据相差天数降低候选自身匹配度。
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

复制 `common.env.example` 为本机私有的 `common.env`，并配置资料目录：

```env
CLOUDSTATION_ROOT_WINDOWS=D:\CloudStaion
INPUT_PATH=${CLOUDSTATION_ROOT}/sample/project/input
```

也可以使用 `--input-dir` 指定：

```powershell
python run_archive.py --input-dir "sample/project/input"
```

## 旧目录迁移

旧版平铺目录必须先演练。演练只生成计划，不移动资料：

```powershell
python migrate_archive.py
```

确认计划后，执行迁移时必须提供内容一致的备份目录：

```powershell
python migrate_archive.py `
  --apply `
  --backup-dir "sample/project/input_backup"
```

执行前会逐文件比较正式目录与备份的大小和 SHA-256。任何差异都会阻止迁移。迁移时还会在每次移动前后再次校验哈希。

旧版 Excel 只在首次迁移时读取已有 PDF 对应关系，随后归档到 `.docflow/legacy`。迁移完成后 Excel 不再作为缓存或数据输入。

## 日常命令

执行完整增量流程：

```powershell
python run_archive.py
```

日常只需将文件放入 `_input` 第一层。程序不递归扫描子目录；支持 `.docx`、
`.jpg`、`.jpeg`、`.png` 和 `.pdf`。PDF 分类结果保存在正式 JSON 的
`input_routes` 中，识别文本按 SHA-256 缓存。

申请材料、施工人员名单、审批 PDF、人工审核数据和汇总 HTML 均由这个入口
按顺序处理。各支流程属于内部业务编排，统一位于
`src/warranty_application_archive/flows/`，不再为每个内部步骤在项目根目录
增加启动脚本。所有处理均写入正式 JSON 的 `changes` 和 `runs`，终端同时
显示带时间戳的中文处理日志。

Windows 双击 `打开待人工审核匹配PDF.cmd`，macOS 双击
`打开待人工审核匹配PDF.command` 启动审核。程序会自动在浏览器中打开
HTML 页面；“保存并执行审核结果”固定更新同目录的
`待人工审核匹配PDF.json`，不会询问保存位置，并立即完成确认归档、
移入 `_trash` 及正式数据刷新。直接双击 HTML 只能预览，不能写入本地文件。
启动器每次自动选择空闲的本地端口，避免程序升级后浏览器仍连接旧版后台。
如果页面明确提示后台服务仍是旧版本，请关闭之前打开的启动器终端窗口，
再重新双击启动器。

也可以从命令行启动：

```powershell
python serve_archive_review.py
```

正常使用网页时不需要第二个根目录命令；审核结果保存、正式 JSON 回写和页面
刷新由 `serve_archive_review.py` 对应的服务流程完成。如果正式 JSON 在审核
页面生成后发生过版本变化，页面会拒绝过期结果，须先重新执行完整归档流程。

## HTML 人工审查视图

`质保作业申请汇总.html` 包含以下页签：

- `申请汇总`
- `待补材料`
- `待审批PDF`
- `已完成`

“说明”和“本次变更”页签不再生成；文件变更记录只保留在正式 JSON 的
`changes` 中。Word、图片、协议和审批 PDF 提供本地链接；页面不显示
案卷目录、案卷 ID 和审批编号。“案卷状态”和“材料完整性”合并为第一列，
两项使用分号分隔并换行，缺少的材料直接列在该列中。
“质保单位和分包单位”合并显示在一列，使用分号分隔并换行；右键可分别
选择“复制质保单位”或“复制分包单位”。
“质保负责人及联系电话”和“施工负责人及联系电话”各占一列，右键单元格
可分别选择“复制姓名”或“复制电话”，姓名和电话在单元格内分两行显示。
这三个较长表头分别固定为两行显示，避免浏览器按列宽产生不规则换行。
页面支持页签切换、当前页签搜索、固定表头和打印。正式汇总 HTML 不作为
数据输入。

“项目名称”“施工区域”和“施工内容”单元格可通过右键菜单选择“复制内容”，
用于直接粘贴到线上填报流程。

“危险作业及专项材料核对”列分行显示“影响、改动消防设备设施”“影响、
堵塞应急疏散通道”和“危险作业”三个正式 JSON 字段，并根据申报内容核对
有限空间申请、高处作业申请及通用专项作业材料。缺少对应附件或危险作业
填写为“无”但存在专项材料时，页面以警示色显示核对结果；该提示不修改
正式 JSON 或案卷状态。

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
粘贴。通过 `.cmd` 或 `.command` 启动器打开页面后，左键点击文档链接会
直接调用系统默认程序打开真实文件，不再通过浏览器下载。直接双击 HTML
时仍受浏览器安全限制，不能调用本地服务或复制为可粘贴的文件。

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

- 施工区域严格命中，且审批 PDF 施工内容字段包含申请施工内容，占 80%，属于
  候选准入条件。
- 开始、结束日期合计占 20%；日期相差超过 31 天也不会淘汰候选，
  但日期部分不再加分。
- 存在多个严格候选时，最终选择置信度还会根据第一名与第二名的
  领先差值下调；并列第一会显示为低置信度。
- PDF 没有足够的内嵌文字时，程序调用本地 AI OCR，并在审核产物中
  记录识别方式；已有 SHA-256 识别缓存会直接复用。

## 独立入口

项目根目录按工作流提供独立 Python 入口：

```powershell
python run_archive.py
python migrate_archive.py
python serve_archive_review.py
```

根目录只保留以上三个实际操作入口。每个入口文件头部均说明用途、配置、
参数和输出；公共能力位于 `modules/`，所有业务步骤和完整流程编排均位于
`flows/`。不再保留统一子命令兼容入口，也不为内部支流程创建根目录脚本。

最常用的完整增量处理命令是：

```powershell
python run_archive.py
```

处理完成后的主要成果包括：

- `质保作业申请数据.json`：唯一正式数据源。
- `质保作业申请汇总.html`：2K 屏幕人工查阅及飞书填报辅助页面。
- `待人工审核匹配PDF.json/.html`：未自动匹配审批 PDF 的审核闭环。
- `_cases/`：规范命名后的独立案卷目录。
- `.docflow/quarantine/`：可恢复的重复文件和重复案卷隔离区。

## 项目结构

```text
run_archive.py                     # 完整增量归档入口
migrate_archive.py                 # 旧目录迁移入口
serve_archive_review.py            # 本地审核/汇总页面服务
logging_config.py                  # 统一控制台与滚动文件日志
config.yaml                        # app 与 flows 分层配置
common.env.example                 # 脱敏的本机配置模板
docs/                              # 架构、部署和开发文档
src/warranty_application_archive/
├─ config_loader.py                # 配置、环境变量和路径集中解析
├─ context.py                      # 入口与编排共享上下文
├─ modules/                        # 单一职责的基础能力
│  ├─ naming.py                    # 案卷及材料命名
│  ├─ repository.py                # JSON 唯一数据源
│  ├─ recognition.py               # 共享 AI/OCR 与缓存
│  ├─ summary_html.py              # JSON → 正式汇总 HTML
│  └─ validation.py                # 数据和成果校验
└─ flows/                          # 场景编排
   ├─ application_flow.py          # 完整归档、迁移和页面服务入口编排
   ├─ archive_flow.py              # 申报、名单及审批 PDF 入库
   ├─ migration_flow.py            # 旧目录迁移计划与执行
   ├─ approval_review_flow.py      # 审批候选、决定与回写
   └─ approval_review_web_flow.py  # 本地审核页面服务
```

## 开发验证

```powershell
python -m compileall -q src tests logging_config.py run_archive.py migrate_archive.py serve_archive_review.py
python -m pytest -q
python -m flake8 logging_config.py run_archive.py migrate_archive.py serve_archive_review.py src tests
```
