"""质保作业申请案卷归档工具的统一命令行入口。

项目功能
--------
以一份质保作业申请为一个案卷，处理人工投放到 ``_input`` 的 Word、手签
申请单、施工人员名单、专项作业材料和审批 PDF。程序负责文件分类、本地
AI/OCR 识别、统一命名、重复申请合并、材料完整性判断、审批 PDF 匹配、
人工审核回写，以及从正式 JSON 数据生成适合 2K 屏幕查阅的 HTML 页面。

基本用法
--------
在项目根目录执行::

    python warranty_application_archive.py run

默认资料目录来自 ``common.env``，也可以显式指定::

    python warranty_application_archive.py --input-dir "资料目录" run

常用子命令包括 ``applications``（申请材料入库）、``worker-lists``（人员
名单入库）、``approval-pdfs``（审批 PDF 归档）、``approval-review``
（生成人工审核数据和页面）、``apply-approval-review``（执行审核结果）、
``export``（重新生成汇总 HTML）、``status``（查看状态）和 ``validate``
（校验 JSON、案卷文件及 HTML）。``reclassify-materials`` 可根据 OCR 中
专项作业选项后的方框勾选状态，重新判断既往材料。Windows 和 macOS 用户
也可以运行资料目录中的独立 ``.cmd`` 或 ``.command`` 启动器打开汇总及
人工审核页面。

主要成果文件
------------
``质保作业申请数据.json``
    唯一正式数据源，记录案卷业务字段、材料、审批信息、哈希和变更历史。
``质保作业申请汇总.html``
    由正式 JSON 单向生成的人工查阅和线上流程辅助填报页面。
``待人工审核匹配PDF.json`` / ``待人工审核匹配PDF.html``
    未自动匹配审批 PDF 的审核数据和本地交互页面。
``_cases/``
    每份申请的独立案卷目录，保存规范命名后的申请材料和审批 PDF。
``.docflow/quarantine/``
    可恢复的重复文件及较少资料的重复案卷隔离区。
"""

from src.warranty_application_archive.cli import main


if __name__ == "__main__":
    raise SystemExit(main())
