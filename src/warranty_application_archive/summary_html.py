from __future__ import annotations

import json
import shutil
import sys
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any, Iterable
from urllib.parse import quote

from .constants import (
    INTERNAL_DIR_NAME,
    LEGACY_DIR_NAME,
    LEGACY_SUMMARY_EXCEL_FILE_NAME,
    SUMMARY_HTML_FILE_NAME,
    SUMMARY_LAUNCHER_FILE_NAME,
    SUMMARY_MACOS_LAUNCHER_FILE_NAME,
)
from .file_utils import atomic_replace_text
from .launchers import write_page_launchers


STATUS_LABELS = {
    "materials_incomplete": "材料待补充",
    "materials_ready": "材料齐全，待审批PDF",
    "approval_pdf_unmatched": "审批PDF待确认",
    "approved": "审批完成",
    "needs_review": "待人工确认",
    "terminated": "终止",
}
ROLE_LABELS = {
    "word": "Word申请单",
    "signed_application": "手签申请单",
    "worker_list": "施工人员名单",
    "safety_agreement": "安全生产及消防安全协议",
    "confined_space": "有限空间申请",
    "high_altitude": "高处作业申请",
    "special_work": "专项作业材料",
    "approval_pdf": "审批PDF",
}
SUMMARY_HEADERS = [
    "案卷状态",
    "材料完整性",
    "项目名称",
    "施工区域",
    "施工开始时间",
    "施工结束时间",
    "施工内容",
    "危险作业",
    "缺少材料",
    "Word申请单",
    "手签申请单",
    "施工人员名单",
    "专项作业材料",
    "安全协议",
    "审批编号",
    "审批PDF",
    "案卷目录",
    "案卷ID",
]
SUMMARY_SHEET_NAMES = [
    "申请汇总",
    "待补材料",
    "待审批PDF",
    "已完成",
]


def _files(application: dict[str, Any], role: str) -> list[dict[str, Any]]:
    return list((application.get("materials") or {}).get(role) or [])


def _file_links(files: Iterable[dict[str, Any]]) -> list[dict[str, str]]:
    links: list[dict[str, str]] = []
    for item in files:
        relative = str(item.get("path") or "")
        if not relative:
            continue
        links.append(
            {
                "text": str(
                    item.get("current_name")
                    or Path(relative).name
                ),
                "href": _relative_href(relative),
                "path": relative,
            }
        )
    return links


def _relative_href(relative: str, directory: bool = False) -> str:
    normalized = str(relative or "").replace("\\", "/").lstrip("/")
    href = "/".join(
        quote(part, safe="")
        for part in PurePosixPath(normalized).parts
        if part not in {"", "."}
    )
    if directory and href and not href.endswith("/"):
        href += "/"
    return href


def _cell(
    value: Any = "",
    *,
    links: list[dict[str, str]] | None = None,
    tone: str = "",
) -> dict[str, Any]:
    return {
        "text": str(value or ""),
        "links": links or [],
        "tone": tone,
    }


def _summary_row(application: dict[str, Any]) -> list[dict[str, Any]]:
    business = application.get("application") or {}
    approval = application.get("approval") or {}
    word_files = _files(application, "word")
    signed_files = _files(application, "signed_application")
    worker_files = _files(application, "worker_list")
    safety_files = _files(application, "safety_agreement")
    special_files = [
        *_files(application, "confined_space"),
        *_files(application, "high_altitude"),
        *_files(application, "special_work"),
    ]
    approval_files = list(approval.get("pdfs") or [])
    missing = [
        ROLE_LABELS.get(role, role)
        for role in application.get("missing_material_types") or []
    ]
    status = str(application.get("status") or "")
    status_tone = {
        "approved": "success",
        "materials_ready": "info",
        "materials_incomplete": "warning",
        "needs_review": "danger",
    }.get(status, "")
    case_directory = str(application.get("case_directory") or "")
    cells = [
        _cell(STATUS_LABELS.get(status, status), tone=status_tone),
        _cell(
            "完整" if not missing else "缺少材料",
            tone="success" if not missing else "warning",
        ),
        _cell(business.get("项目名称", "")),
        _cell(business.get("施工区域", "")),
        _cell(business.get("施工开始时间", "")),
        _cell(business.get("施工结束时间", "")),
        _cell(business.get("施工内容", "")),
        _cell(business.get("危险作业", "")),
        _cell("；".join(missing), tone="warning" if missing else ""),
        _cell(links=_file_links(word_files)),
        _cell(links=_file_links(signed_files)),
        _cell(links=_file_links(worker_files)),
        _cell(links=_file_links(special_files)),
        _cell(links=_file_links(safety_files)),
        _cell(approval.get("application_no", "")),
        _cell(links=_file_links(approval_files)),
        _cell(
            links=[
                {
                    "text": case_directory,
                    "href": _relative_href(case_directory, directory=True),
                    "path": case_directory,
                }
            ]
            if case_directory
            else []
        ),
        _cell(application.get("case_id", "")),
    ]
    if status == "terminated":
        for cell in cells:
            cell["tone"] = "terminated"
    return cells


def _summary_sheet(
    title: str,
    applications: Iterable[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "title": title,
        "headers": SUMMARY_HEADERS,
        "rows": [_summary_row(item) for item in applications],
    }


def build_summary_view(data: dict[str, Any]) -> dict[str, Any]:
    applications = list(data.get("applications") or [])
    sheets = [
        _summary_sheet("申请汇总", applications),
        _summary_sheet(
            "待补材料",
            [
                item
                for item in applications
                if (
                    item.get("status") != "terminated"
                    and item.get("missing_material_types")
                )
            ],
        ),
        _summary_sheet(
            "待审批PDF",
            [
                item
                for item in applications
                if item.get("status") == "materials_ready"
            ],
        ),
        _summary_sheet(
            "已完成",
            [
                item
                for item in applications
                if item.get("status") == "approved"
            ],
        ),
    ]
    return {
        "schema_version": 1,
        "dataset_revision": int(data.get("dataset_revision") or 0),
        "generated_at": datetime.now().astimezone().isoformat(
            timespec="seconds"
        ),
        "application_count": len(applications),
        "metrics": {
            "total": len(applications),
            "approved": sum(
                item.get("status") == "approved"
                for item in applications
            ),
            "pending": sum(
                item.get("status") == "materials_ready"
                for item in applications
            ),
            "incomplete": sum(
                (
                    item.get("status") != "terminated"
                    and bool(item.get("missing_material_types"))
                )
                for item in applications
            ),
        },
        "sheets": sheets,
    }


def _json_for_script(value: object) -> str:
    return (
        json.dumps(value, ensure_ascii=False)
        .replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def _archive_retired_excel(root: Path) -> None:
    source = root / LEGACY_SUMMARY_EXCEL_FILE_NAME
    if not source.is_file():
        return
    archive_dir = root / INTERNAL_DIR_NAME / LEGACY_DIR_NAME
    archive_dir.mkdir(parents=True, exist_ok=True)
    target = archive_dir / source.name
    if target.exists():
        target = archive_dir / (
            f"{source.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            f"{source.suffix}"
        )
    try:
        shutil.move(str(source), str(target))
    except PermissionError as exc:
        raise PermissionError(
            f"旧汇总 Excel 正在使用，请关闭后重新生成 HTML: {source}"
        ) from exc


def export_summary_html(data: dict[str, Any], root: Path) -> Path:
    root = root.resolve()
    _archive_retired_excel(root)
    view = build_summary_view(data)
    content = _HTML_TEMPLATE.replace(
        "__SUMMARY_DATA__",
        _json_for_script(view),
    )
    output = root / SUMMARY_HTML_FILE_NAME
    atomic_replace_text(output, content)
    entry_point = (
        Path(__file__).resolve().parents[2]
        / "warranty_application_archive.py"
    )
    write_page_launchers(
        root,
        entry_point.parent,
        [
            sys.executable,
            str(entry_point),
            "--input-dir",
            str(root),
            "approval-review-server",
            "--page",
            "summary",
            "--port",
            "8766",
        ],
        windows_name=SUMMARY_LAUNCHER_FILE_NAME,
        macos_name=SUMMARY_MACOS_LAUNCHER_FILE_NAME,
    )
    return output


_HTML_TEMPLATE = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>质保作业申请汇总</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #162234;
      --muted: #68798d;
      --navy: #44546a;
      --navy-dark: #354357;
      --blue-row: #d9e7f5;
      --line: #afbdcc;
      --paper: #fff;
      --canvas: #edf2f7;
      --success: #267147;
      --success-bg: #e4f3e9;
      --warning: #8a5d13;
      --warning-bg: #fff2ce;
      --danger: #a03936;
      --danger-bg: #fde7e5;
      --info: #285f8e;
      --info-bg: #e3f0fb;
      --shadow: 0 12px 30px rgba(32, 51, 73, .11);
    }
    * { box-sizing: border-box; }
    html, body { max-width: 100%; overflow-x: hidden; }
    body {
      margin: 0;
      color: var(--ink);
      background: var(--canvas);
      font-family: "Microsoft YaHei UI", "Microsoft YaHei", system-ui, sans-serif;
    }
    header {
      padding: 25px max(18px, calc((100vw - 2520px) / 2));
      color: white;
      background: linear-gradient(135deg, var(--navy-dark), #536b86);
    }
    h1 { margin: 0 0 7px; font-size: 25px; }
    header p { margin: 0; color: #dce8f4; font-size: 14px; }
    main {
      width: min(2520px, calc(100% - 28px));
      margin: 18px auto 50px;
    }
    .metrics {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      margin-bottom: 16px;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 11px;
      background: var(--paper);
      box-shadow: var(--shadow);
    }
    .metric { padding: 16px 20px; border-right: 1px solid var(--line); }
    .metric:last-child { border-right: 0; }
    .metric strong { display: block; color: var(--navy); font-size: 27px; }
    .metric span { color: var(--muted); font-size: 13px; }
    .workspace {
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 11px;
      background: var(--paper);
      box-shadow: var(--shadow);
    }
    .toolbar {
      display: flex;
      gap: 12px;
      align-items: center;
      justify-content: space-between;
      padding: 12px 14px;
      border-bottom: 1px solid var(--line);
      background: #f7f9fc;
    }
    .tabs { display: flex; gap: 7px; flex-wrap: wrap; }
    button, input { font: inherit; }
    .tab {
      padding: 9px 14px;
      border: 1px solid #c1ccda;
      border-radius: 7px;
      color: var(--ink);
      background: #e6edf5;
      cursor: pointer;
      font-weight: 700;
    }
    .tab.active { color: white; border-color: var(--navy); background: var(--navy); }
    .tab .count {
      display: inline-block;
      min-width: 22px;
      margin-left: 6px;
      padding: 1px 6px;
      border-radius: 999px;
      background: rgba(255,255,255,.28);
      font-size: 12px;
    }
    .search {
      width: min(330px, 100%);
      padding: 9px 12px;
      border: 1px solid #aebccd;
      border-radius: 7px;
      outline: none;
      background: white;
    }
    .search:focus { border-color: #5b87b4; box-shadow: 0 0 0 3px #dfeaf5; }
    .sheet-status {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      padding: 9px 14px;
      color: var(--muted);
      border-bottom: 1px solid var(--line);
      font-size: 13px;
    }
    .table-wrap {
      max-height: calc(100vh - 285px);
      overflow-x: hidden;
      overflow-y: auto;
    }
    table {
      width: 100%;
      table-layout: fixed;
      border-collapse: separate;
      border-spacing: 0;
      font-size: clamp(11px, .58vw, 13px);
    }
    th, td {
      min-width: 0;
      max-width: none;
      padding: 9px 7px;
      border-right: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      text-align: left;
      vertical-align: middle;
      overflow-wrap: anywhere;
      word-break: break-word;
    }
    th:nth-child(1), td:nth-child(1),
    th:nth-child(2), td:nth-child(2),
    th:nth-child(4), td:nth-child(4),
    th:nth-child(5), td:nth-child(5),
    th:nth-child(6), td:nth-child(6),
    th:nth-child(8), td:nth-child(8),
    th:nth-child(9), td:nth-child(9),
    th:nth-child(15), td:nth-child(15),
    th:nth-child(18), td:nth-child(18) { width: 5%; }
    th:nth-child(3), td:nth-child(3),
    th:nth-child(7), td:nth-child(7),
    th:nth-child(10), td:nth-child(10),
    th:nth-child(11), td:nth-child(11),
    th:nth-child(12), td:nth-child(12),
    th:nth-child(13), td:nth-child(13),
    th:nth-child(14), td:nth-child(14),
    th:nth-child(16), td:nth-child(16) { width: 6%; }
    th:nth-child(17), td:nth-child(17) { width: 7%; }
    th {
      position: sticky;
      top: 0;
      z-index: 2;
      color: white;
      background: var(--navy);
      text-align: center;
      font-weight: 700;
    }
    tbody tr:nth-child(odd) td { background: var(--blue-row); }
    tbody tr:nth-child(even) td { background: white; }
    tbody tr:hover td { background: #fff4d7; }
    td.tone-success { color: var(--success); background: var(--success-bg) !important; font-weight: 700; }
    td.tone-warning { color: var(--warning); background: var(--warning-bg) !important; font-weight: 700; }
    td.tone-danger { color: var(--danger); background: var(--danger-bg) !important; font-weight: 700; }
    td.tone-info { color: var(--info); background: var(--info-bg) !important; font-weight: 700; }
    td.tone-terminated { color: #4f5965; background: #d7dbe0 !important; }
    td.tone-terminated a { color: #4f5965; }
    a { color: #075fa8; text-decoration: underline; text-underline-offset: 2px; }
    .links { display: grid; min-width: 0; gap: 5px; }
    .links a { min-width: 0; overflow-wrap: anywhere; }
    .file-menu {
      position: fixed;
      z-index: 20;
      min-width: 210px;
      padding: 6px;
      border: 1px solid #aebccd;
      border-radius: 8px;
      background: white;
      box-shadow: 0 10px 28px rgba(21, 35, 52, .22);
    }
    .file-menu button {
      width: 100%;
      padding: 9px 12px;
      border: 0;
      border-radius: 5px;
      color: var(--ink);
      background: transparent;
      text-align: left;
      cursor: pointer;
    }
    .file-menu button:hover { background: #e6edf5; }
    .copy-toast {
      position: fixed;
      right: 20px;
      bottom: 20px;
      z-index: 21;
      max-width: min(420px, calc(100% - 40px));
      padding: 11px 16px;
      border-radius: 8px;
      color: white;
      background: #267147;
      box-shadow: var(--shadow);
    }
    .copy-toast.error { background: #a03936; }
    .empty { padding: 34px; color: var(--muted); text-align: center; }
    [hidden] { display: none !important; }
    @media (max-width: 850px) {
      .metrics { grid-template-columns: 1fr 1fr; }
      .metric:nth-child(2) { border-right: 0; }
      .metric:nth-child(-n+2) { border-bottom: 1px solid var(--line); }
      .toolbar { align-items: stretch; flex-direction: column; }
      .search { width: 100%; }
      .table-wrap { max-height: none; }
      table { font-size: 10px; }
      th, td { padding: 7px 4px; }
    }
    @media print {
      header, .metrics, .toolbar, .sheet-status { display: none; }
      body, main, .workspace { margin: 0; width: 100%; background: white; box-shadow: none; border: 0; }
      .table-wrap { max-height: none; overflow: visible; }
      th { position: static; }
    }
  </style>
</head>
<body>
  <header>
    <h1>质保作业申请汇总</h1>
    <p id="subtitle"></p>
  </header>
  <main>
    <section class="metrics" aria-label="汇总指标">
      <div class="metric"><strong id="totalMetric">0</strong><span>案卷总数</span></div>
      <div class="metric"><strong id="approvedMetric">0</strong><span>审批完成</span></div>
      <div class="metric"><strong id="pendingMetric">0</strong><span>待审批 PDF</span></div>
      <div class="metric"><strong id="incompleteMetric">0</strong><span>材料待补充</span></div>
    </section>
    <section class="workspace">
      <div class="toolbar">
        <nav id="tabs" class="tabs" aria-label="汇总页签"></nav>
        <input id="search" class="search" type="search" placeholder="搜索当前页签…" aria-label="搜索当前页签">
      </div>
      <div class="sheet-status">
        <strong id="sheetTitle"></strong>
        <span id="rowCount"></span>
      </div>
      <div class="table-wrap">
        <table>
          <thead id="tableHead"></thead>
          <tbody id="tableBody"></tbody>
        </table>
        <div id="emptyState" class="empty" hidden>当前页签没有记录</div>
      </div>
    </section>
  </main>
  <div id="fileMenu" class="file-menu" role="menu" hidden>
    <button id="copyFileButton" type="button" role="menuitem">复制文件（可直接粘贴）</button>
  </div>
  <div id="copyToast" class="copy-toast" role="status" hidden></div>
  <script id="summaryData" type="application/json">__SUMMARY_DATA__</script>
  <script>
    const data = JSON.parse(document.getElementById("summaryData").textContent);
    let activeIndex = 0;
    const byId = (id) => document.getElementById(id);
    const fileUrl = (path, directory = false) => {
      const normalized = String(path || "").replaceAll("\\", "/").replace(/^\/+/, "");
      const encoded = normalized.split("/").map(encodeURIComponent).join("/");
      if (location.protocol === "file:") return `${encoded}${directory ? "/" : ""}`;
      if (directory) return "#";
      return `/files/${encoded}`;
    };
    const text = (tag, value, className = "") => {
      const node = document.createElement(tag);
      node.textContent = value ?? "";
      if (className) node.className = className;
      return node;
    };
    function renderTabs() {
      const tabs = byId("tabs");
      tabs.replaceChildren();
      data.sheets.forEach((sheet, index) => {
        const button = text("button", sheet.title, `tab${index === activeIndex ? " active" : ""}`);
        button.type = "button";
        button.append(text("span", sheet.rows.length, "count"));
        button.addEventListener("click", () => {
          activeIndex = index;
          byId("search").value = "";
          renderTabs();
          renderSheet();
        });
        tabs.append(button);
      });
    }
    function renderCell(cell) {
      const td = document.createElement("td");
      if (cell.tone) td.className = `tone-${cell.tone}`;
      if (cell.links && cell.links.length) {
        const links = text("div", "", "links");
        cell.links.forEach((item) => {
          const anchor = text("a", item.text || item.path || "打开");
          const directory = String(item.href || "").endsWith("/");
          anchor.href = fileUrl(item.path || item.href, directory);
          if (!directory || location.protocol === "file:") {
            anchor.target = "_blank";
          }
          anchor.rel = "noopener";
          anchor.dataset.filePath = item.path || "";
          anchor.dataset.fileKind = directory ? "directory" : "file";
          if (directory && location.protocol !== "file:") {
            anchor.addEventListener("click", async (event) => {
              event.preventDefault();
              try {
                const response = await fetch("/api/open-path", {
                  method: "POST",
                  headers: {"Content-Type": "application/json"},
                  body: JSON.stringify({path: anchor.dataset.filePath}),
                });
                const result = await response.json();
                if (!response.ok || !result.ok) {
                  throw new Error(result.error || "打开文件夹失败");
                }
              } catch (error) {
                showCopyMessage(error.message || "打开文件夹失败", true);
              }
            });
          }
          links.append(anchor);
        });
        td.append(links);
      } else {
        td.textContent = cell.text || "";
      }
      return td;
    }
    function renderSheet() {
      const sheet = data.sheets[activeIndex];
      const query = byId("search").value.trim().toLocaleLowerCase("zh-CN");
      const rows = sheet.rows.filter((row) => {
        if (!query) return true;
        return row.some((cell) => {
          const linkText = (cell.links || []).map((item) => item.text).join(" ");
          return `${cell.text || ""} ${linkText}`.toLocaleLowerCase("zh-CN").includes(query);
        });
      });
      const headerRow = document.createElement("tr");
      sheet.headers.forEach((header) => headerRow.append(text("th", header)));
      byId("tableHead").replaceChildren(headerRow);
      const body = byId("tableBody");
      body.replaceChildren();
      rows.forEach((row) => {
        const tr = document.createElement("tr");
        row.forEach((cell) => tr.append(renderCell(cell)));
        body.append(tr);
      });
      byId("sheetTitle").textContent = sheet.title;
      byId("rowCount").textContent = `显示 ${rows.length} / ${sheet.rows.length} 条`;
      byId("emptyState").hidden = rows.length !== 0;
      byId("tableHead").hidden = rows.length === 0;
    }
    byId("search").addEventListener("input", renderSheet);
    let selectedFilePath = "";
    let toastTimer = 0;
    const hideFileMenu = () => { byId("fileMenu").hidden = true; };
    const showCopyMessage = (message, error = false) => {
      const toast = byId("copyToast");
      toast.textContent = message;
      toast.className = `copy-toast${error ? " error" : ""}`;
      toast.hidden = false;
      window.clearTimeout(toastTimer);
      toastTimer = window.setTimeout(() => { toast.hidden = true; }, 3200);
    };
    document.addEventListener("contextmenu", (event) => {
      const anchor = event.target.closest("a[data-file-path]");
      if (!anchor || !anchor.dataset.filePath) return;
      event.preventDefault();
      selectedFilePath = anchor.dataset.filePath;
      byId("copyFileButton").textContent = (
        anchor.dataset.fileKind === "directory"
          ? "复制文件夹（可直接粘贴）"
          : "复制文件（可直接粘贴）"
      );
      const menu = byId("fileMenu");
      menu.hidden = false;
      const bounds = menu.getBoundingClientRect();
      menu.style.left = `${Math.min(event.clientX, window.innerWidth - bounds.width - 8)}px`;
      menu.style.top = `${Math.min(event.clientY, window.innerHeight - bounds.height - 8)}px`;
    });
    document.addEventListener("click", (event) => {
      if (!event.target.closest("#fileMenu")) hideFileMenu();
    });
    window.addEventListener("blur", hideFileMenu);
    byId("copyFileButton").addEventListener("click", async () => {
      hideFileMenu();
      if (location.protocol === "file:") {
        showCopyMessage(
          "请通过启动器打开：Windows 使用 .cmd，macOS 使用 .command",
          true
        );
        return;
      }
      try {
        const response = await fetch("/api/copy-file", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify({path: selectedFilePath}),
        });
        const result = await response.json();
        if (!response.ok || !result.ok) throw new Error(result.error || "复制失败");
        showCopyMessage(`已复制文件：${result.file_name}`);
      } catch (error) {
        showCopyMessage(error.message || "复制失败，请查看程序日志", true);
      }
    });
    byId("totalMetric").textContent = data.metrics.total;
    byId("approvedMetric").textContent = data.metrics.approved;
    byId("pendingMetric").textContent = data.metrics.pending;
    byId("incompleteMetric").textContent = data.metrics.incomplete;
    byId("subtitle").textContent = `数据版本 ${data.dataset_revision} · 生成时间 ${data.generated_at}`;
    renderTabs();
    renderSheet();
  </script>
</body>
</html>
"""
