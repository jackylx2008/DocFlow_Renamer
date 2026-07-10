# DocFlow Renamer

质保作业申请单批处理工具。脚本会读取输入目录中的 Word 申请单，解析表单字段，按日期和施工内容重命名文件，并导出 Excel 汇总表。汇总表还会按施工区域、施工内容和施工开始日期匹配同目录下的工程类主体质保施工 PDF。

## 功能

- 解析 Word 申请单中的项目名称、质保单位、施工区域、施工日期、施工内容、负责人、危险作业等字段。
- 将申请单重命名为 `YYYY-MM-DD_施工内容_质保作业申请单.docx`。
- 自动关联同名或同目录下的申请单图片、附件目录。
- 对随机文件名的申请单图片调用本地 LLAMACPP 多模态模型识别文字，并按匹配到的申请单复制/重命名为同名图片。
- 将文件名包含 `人员名单` 或 `工人名单` 的图片，按修改日期匹配同一天的申请单，复制生成为对应的申请单工人名单图片。
- 导出 `质保作业申请汇总.xlsx`，如果目标文件已存在，会生成带时间戳的新汇总文件。
- 匹配工程类主体质保施工 PDF，并在 Excel 的 `匹配PDF文件名` 列写入结果和本地链接。
- 普通 PDF 文本不可用时自动使用本地 LLAMACPP 多模态模型识别。
- `docflow_renamer.py` 会自动识别包含 `工程类-主体质保施工` 和 `申请编号` 的 PDF，并重命名为 `工程类-主体质保施工_编号：12位编号.pdf`；已符合该命名规则的 PDF 会跳过识别。
- 可通过 `manual_matches.env` 配置少量手工特例，直接指定 Excel 汇总中的申请单附图和匹配 PDF。
- 可单独识别输入目录第一层的 `1.pdf`、`2.pdf` 等纯数字 PDF，提取 `申请编号` 后的 12 位编码，并重命名为 `工程类-主体质保施工_编号：申请编号.pdf`。
- 脚本运行日志会写入 `log` 目录，日志文件名前缀与入口 Python 文件名一致。

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 配置

默认读取仓库根目录下的 `common.env`：

```env
INPUT_PATH=D:\path\to\质保作业申请单
SKIP_PDF_FILES=["01 安全生产及消防安全协议（建工）.pdf"]
```

`INPUT_PATH` 指向申请单、图片、附件和 PDF 所在目录。`SKIP_PDF_FILES` 使用一行列表配置需要排除匹配的 PDF 文件名；旧的逗号、分号或中文分隔符写法仍兼容。

手工特例读取仓库根目录下的 `manual_matches.env`：

```env
MANUAL_PDF_MATCHES=[{"application_image":"2026-07-08刷漆质保作业申请单.ipg","pdf":"工程类-主体质保施工_编号：202607070599.pdf"}]
```

`application_image` 用于定位申请单记录，`pdf` 用于直接写入 Excel 的 `匹配PDF文件名` 并设置本地链接。配置中的图片扩展名如果写成 `.ipg`，脚本会自动尝试同名 `.jpg`、`.jpeg`、`.png` 文件。

也可以通过命令行指定输入目录：

```powershell
python docflow_renamer.py --input-dir "D:\path\to\质保作业申请单"
```

## 运行

```powershell
python docflow_renamer.py
```

运行完成后会输出 JSON 摘要，包含输入目录、处理数量、PDF 重命名数量、图片重命名数量、手工 PDF 匹配数量和生成的 Excel 路径。

### 复制生成工人名单图片

如果输入目录下存在文件名包含 `人员名单` 或 `工人名单` 的图片：

```text
人员名单.jpg
工人名单.jpg
现场人员名单.png
```

可以运行：

```powershell
python copy_worker_list_images.py
```

脚本会读取 `INPUT_PATH`，按文件修改日期匹配同一天的 `YYYY-MM-DD_施工内容_质保作业申请单.docx`，并复制生成为：

```text
YYYY-MM-DD_施工内容_质保作业申请单_工人名单.源扩展名
```

同一天一张名单图片会复制生成给同一天所有申请单；如果同一天存在多张名单图片，脚本只使用按修改时间和文件名排序后的第一张。如果目标图片已存在，脚本会记录日志并跳过，不覆盖已有文件。确认该原始图片对应的所有目标图片都存在后，脚本会删除原始名单图片。

### 重命名待规范 PDF

如果输入目录第一层存在待识别的纯数字 PDF，或带 `申请编号` 前缀但尚未规范命名的 PDF：

```text
1.pdf
2.pdf
12.pdf
申请编号： 202606300121.pdf
```

可以运行：

```powershell
python rename_pdfs.py
```

脚本会调用本地 LLAMACPP 多模态模型识别纯数字 PDF 页面，提取 `申请编号:` 或 `申请编号：` 后面的 12 位数字编码，并重命名为：

```text
工程类-主体质保施工_编号：123456789012.pdf
```

`申请编号： 202606300121.pdf` 这类文件会直接从文件名提取编号，不再调用 AI 识别。只扫描输入目录第一层的纯数字 PDF 和 `申请编号` 前缀 PDF；`d1.pdf`、`a1.pdf`、子目录内 PDF 不会处理。如果目标文件已存在，脚本会跳过，不覆盖已有文件。

默认识别每个 PDF 前 2 页，可以通过参数调整：

```powershell
python rename_pdfs.py --page-limit 3
python rename_pdfs.py --input-dir "D:\path\to\质保作业申请单"
```

## 图片匹配规则

`docflow_renamer.py` 会扫描输入目录第一层中尚未按日期命名的 `.jpg`、`.jpeg`、`.png` 图片，调用本地 LLAMACPP 多模态模型识别图片文字，再用识别文本中的施工内容匹配近期申请单文件名。

申请单候选默认限定为运行日前 2 天至后 14 天，适合提前制作未来日期申请单的场景，同时避免历史文件过多导致误匹配。

## PDF 匹配规则

`docflow_renamer.py` 会先处理待规范 PDF 文件名。对尚未命名为 `工程类-主体质保施工_编号：12位编号.pdf` 的 PDF，脚本会调用本地 LLAMACPP 多模态模型识别 PDF 前两页；识别文本同时包含 `工程类-主体质保施工` 和 `申请编号：12位数字` 或 `申请编号:12位数字` 时，会将 PDF 重命名为：

```text
工程类-主体质保施工_编号：123456789012.pdf
```

已经符合 `工程类-主体质保施工_编号：12位编号.pdf` 命名规则的 PDF 不会进入 AI 识别和重命名流程。

PDF 匹配会先建立输入目录下所有 PDF 的文本索引。对中文文本提取失败的 PDF，脚本会将 PDF 前两页渲染为图片，并通过本地 LLAMACPP 多模态模型识别文字。

匹配时会对文本做归一化处理：

- 移除空白并统一大小写。
- 统一横线符号。
- 统一全角/半角括号。
- 统一中文逗号、顿号和英文逗号。

候选 PDF 需要同时满足：

- PDF 文本包含申请单的施工区域。
- PDF 文本包含申请单的施工内容。

如果申请单施工内容以施工区域开头，例如 `给水泵房、中水泵房除锈刷漆`，脚本会额外尝试去掉区域前缀后的内容 `除锈刷漆`，用于匹配 PDF 中单独填写的施工内容。

如果多个 PDF 都满足施工区域和施工内容，脚本会继续用施工开始日期筛选，避免同区域、同内容、不同日期的 PDF 互相误匹配。

已有汇总表中的非空 PDF 匹配结果会作为缓存读取，缓存键包含施工开始日期、施工区域和施工内容。

### 手工 PDF 匹配特例

少量无法通过严格匹配规则对应上的文件，可以写入 `manual_matches.env` 的 `MANUAL_PDF_MATCHES`。这些特例在 Excel 导出前应用，只覆盖汇总记录中的 `申请单附图`、`匹配PDF文件名` 和 `匹配PDF文件链接`，不会改变 AI 识别或严格匹配规则。

当前内置特例：

```text
2026-07-08刷漆质保作业申请单.ipg -> 工程类-主体质保施工_编号：202607070599.pdf
```

## 日志

默认日志目录为仓库根目录下的 `log`，也可以通过 `common.env` 的 `LOG_DIR` 或 `config.yaml` 的 `log_dir` 调整。

日志文件名使用入口脚本名加时间戳：

```text
docflow_renamer_YYYYMMDD_HHMMSS.log
copy_worker_list_images_YYYYMMDD_HHMMSS.log
rename_pdfs_YYYYMMDD_HHMMSS.log
```

本地 llama.cpp 服务的 stdout/stderr 仍分别写入：

```text
llama_server.out.log
llama_server.err.log
```

## 验证

```powershell
python -m py_compile docflow_renamer.py copy_worker_list_images.py rename_pdfs.py
python -m flake8 docflow_renamer.py
python docflow_renamer.py
```
