# DocFlow Renamer

质保作业申请单批处理工具。脚本会读取输入目录中的 Word 申请单，解析表单字段，按日期和施工内容重命名文件，并导出 Excel 汇总表。汇总表还会按施工区域、施工内容和施工开始日期匹配同目录下的工程类主体质保施工 PDF。

## 功能

- 解析 Word 申请单中的项目名称、质保单位、施工区域、施工日期、施工内容、负责人、危险作业等字段。
- 将申请单重命名为 `YYYY-MM-DD_施工内容_质保作业申请单.docx`。
- 自动关联同名或同目录下的申请单图片、附件目录。
- 对随机文件名的申请单图片调用本地 LLAMACPP 多模态模型识别文字，并按匹配到的申请单复制/重命名为同名图片。
- 将同日期的 `YYYY-MM-DD_人员名单.jpg` 或 `YYYY-MM-DD_工人名单.jpg` 复制为各申请单对应的 `YYYY-MM-DD_施工内容_质保作业申请单_工人名单.jpg`。
- 导出 `质保作业申请汇总.xlsx`，如果目标文件已存在，会生成带时间戳的新汇总文件。
- 匹配工程类主体质保施工 PDF，并在 Excel 的 `匹配PDF文件名` 列写入结果和本地链接。
- 普通 PDF 文本不可用时自动使用本地 LLAMACPP 多模态模型识别。

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

也可以通过命令行指定输入目录：

```powershell
python docflow_renamer.py --input-dir "D:\path\to\质保作业申请单"
```

## 运行

```powershell
python docflow_renamer.py
```

运行完成后会输出 JSON 摘要，包含输入目录、处理数量、图片重命名数量和生成的 Excel 路径。

### 复制工人名单图片

如果输入目录下存在按日期命名的人员/工人名单图片：

```text
YYYY-MM-DD_人员名单.jpg
YYYY-MM-DD_工人名单.jpg
```

可以运行：

```powershell
python copy_worker_list_images.py
```

脚本会读取 `INPUT_PATH`，按同日期匹配所有 `YYYY-MM-DD_施工内容_质保作业申请单.docx`，并复制生成：

```text
YYYY-MM-DD_施工内容_质保作业申请单_工人名单.jpg
```

如果目标图片已存在，脚本会记录日志并跳过，不覆盖已有文件。

## 图片匹配规则

`docflow_renamer.py` 会扫描输入目录第一层中尚未按日期命名的 `.jpg`、`.jpeg`、`.png` 图片，调用本地 LLAMACPP 多模态模型识别图片文字，再用识别文本中的施工内容匹配近期申请单文件名。

申请单候选默认限定为运行日前 2 天至后 14 天，适合提前制作未来日期申请单的场景，同时避免历史文件过多导致误匹配。

## PDF 匹配规则

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

## 验证

```powershell
python -m py_compile docflow_renamer.py
python -m flake8 docflow_renamer.py
python docflow_renamer.py
```
