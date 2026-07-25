from __future__ import annotations

import json
import logging
import mimetypes
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, unquote, urlparse

from .approval_review import (
    ApprovalReviewRepository,
    DECISION_OPTIONS,
    _archive_legacy_artifact,
    apply_review_decisions,
    build_approval_review,
    import_json_decisions,
)
from .constants import (
    APPROVAL_REVIEW_HTML_FILE_NAME,
    APPROVAL_REVIEW_LAUNCHER_FILE_NAME,
    LEGACY_APPROVAL_REVIEW_EXCEL_FILE_NAME,
    RETIRED_APPROVAL_REVIEW_EXCEL_FILE_NAME,
    TRASH_DIR_NAME,
)
from .file_utils import atomic_replace_text, ensure_within
from .repository import JsonRepository
from .summary_html import export_summary_html
from .workflows import append_run


LOGGER = logging.getLogger(__name__)
MAX_REQUEST_BYTES = 1024 * 1024


def _json_for_script(value: object) -> str:
    return (
        json.dumps(value, ensure_ascii=False)
        .replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def save_review_payload(
    review: dict[str, Any],
    payload: dict[str, Any],
    current_dataset_revision: int,
) -> int:
    source_revision = int(review.get("source_dataset_revision") or 0)
    submitted_revision = int(payload.get("source_dataset_revision") or 0)
    if source_revision != current_dataset_revision:
        raise ValueError("正式数据已更新，请重新生成审核页面后再保存")
    if submitted_revision != source_revision:
        raise ValueError("审核页面已过期，请刷新页面后再保存")

    expected = {
        str(item.get("review_id") or ""): item
        for item in [
            *(review.get("pending_reviews") or []),
            *(review.get("unresolved_pdfs") or []),
        ]
    }
    unresolved_ids = {
        str(item.get("review_id") or "")
        for item in review.get("unresolved_pdfs") or []
    }
    submitted = payload.get("decisions")
    if not isinstance(submitted, list):
        raise ValueError("审核结果格式错误")

    seen: set[str] = set()
    for decision_item in submitted:
        if not isinstance(decision_item, dict):
            raise ValueError("审核结果条目格式错误")
        review_id = str(decision_item.get("review_id") or "").strip()
        if not review_id or review_id in seen:
            raise ValueError(f"审核ID缺失或重复: {review_id}")
        seen.add(review_id)
        candidate = expected.get(review_id)
        if candidate is None:
            raise ValueError(f"审核页面含有未知或过期审核ID: {review_id}")

        decision = str(decision_item.get("decision") or "").strip()
        if decision not in DECISION_OPTIONS:
            raise ValueError(f"审核结果无效: {decision}")
        if (
            review_id in unresolved_ids
            and decision not in {"待审核", "移至_trash"}
        ):
            raise ValueError("无严格候选 PDF 只允许待审核或移至_trash")
        pdf_hash = str(decision_item.get("pdf_sha256") or "").strip()
        case_id = str(decision_item.get("case_id") or "").strip()
        if pdf_hash != str((candidate.get("pdf") or {}).get("sha256") or ""):
            raise ValueError(f"PDF SHA-256 校验失败: {review_id}")
        if case_id != str((candidate.get("case") or {}).get("case_id") or ""):
            raise ValueError(f"案卷ID校验失败: {review_id}")

        candidate["decision"] = decision
        candidate["review_note"] = str(
            decision_item.get("review_note") or ""
        ).strip()[:2000]

    review["last_review_saved_at"] = datetime.now().astimezone().isoformat(
        timespec="seconds"
    )
    return sum(
        1
        for item in [
            *(review.get("pending_reviews") or []),
            *(review.get("unresolved_pdfs") or []),
        ]
        if item.get("decision") != "待审核"
    )


def export_approval_review_html(
    review: dict[str, Any],
    root: Path,
) -> Path:
    root = root.resolve()
    (root / TRASH_DIR_NAME).mkdir(parents=True, exist_ok=True)
    html = _HTML_TEMPLATE.replace(
        "__REVIEW_DATA__",
        _json_for_script(review),
    )
    output = root / APPROVAL_REVIEW_HTML_FILE_NAME
    atomic_replace_text(output, html)
    entry_point = (
        Path(__file__).resolve().parents[2]
        / "warranty_application_archive.py"
    )
    launch_command = subprocess.list2cmdline(
        [
            sys.executable,
            str(entry_point),
            "--input-dir",
            str(root),
            "approval-review-server",
        ]
    )
    launcher = root / APPROVAL_REVIEW_LAUNCHER_FILE_NAME
    atomic_replace_text(
        launcher,
        "\n".join(
            [
                "@echo off",
                "chcp 65001 >nul",
                f'cd /d "{entry_point.parent}"',
                launch_command,
                "pause",
                "",
            ]
        ),
    )
    _archive_legacy_artifact(
        root,
        root / RETIRED_APPROVAL_REVIEW_EXCEL_FILE_NAME,
    )
    _archive_legacy_artifact(
        root,
        root / LEGACY_APPROVAL_REVIEW_EXCEL_FILE_NAME,
    )
    return output


def save_and_apply_review_payload(
    root: Path,
    repo_root: Path,
    payload: dict[str, Any],
) -> dict[str, Any]:
    """Persist submitted decisions and immediately execute non-pending ones."""
    root = root.resolve()
    data_repository = JsonRepository(root)
    review_repository = ApprovalReviewRepository(root)
    data = data_repository.load()
    review = review_repository.load()
    current_revision = int(data.get("dataset_revision") or 0)
    saved = save_review_payload(
        review,
        payload,
        current_revision,
    )
    decisions = import_json_decisions(review)
    apply_result = {"confirmed": 0, "trashed": 0}
    formal_html: Path | None = None

    if decisions:
        apply_result = apply_review_decisions(
            data,
            review,
            decisions,
            root,
        )
        if apply_result["confirmed"] or apply_result["trashed"]:
            data["dataset_revision"] = current_revision + 1
        append_run(
            data,
            "apply-approval-review-web",
            {
                "review_decisions_imported": len(decisions),
                "approval_pdfs_human_confirmed": apply_result["confirmed"],
                "approval_pdfs_moved_to_trash": apply_result["trashed"],
            },
        )
        data_repository.save(data)
        formal_html = export_summary_html(data, root)
        review = build_approval_review(
            data,
            root,
            repo_root,
            existing=review,
        )
        # Recognition can populate cache entries while the next candidate is
        # rebuilt, so persist the dataset once more after rebuilding.
        data_repository.save(data)

    review_repository.save(review)
    export_approval_review_html(review, root)
    return {
        "saved_decisions": saved,
        "review_decisions_applied": len(decisions),
        "approval_pdfs_human_confirmed": apply_result["confirmed"],
        "approval_pdfs_moved_to_trash": apply_result["trashed"],
        "formal_html": str(formal_html) if formal_html else "",
        "review": review,
    }


def serve_approval_review(
    root: Path,
    host: str = "127.0.0.1",
    port: int = 8765,
    open_browser: bool = True,
) -> None:
    root = root.resolve()
    review_repository = ApprovalReviewRepository(root)
    repo_root = Path(__file__).resolve().parents[2]
    lock = threading.Lock()

    class ReviewHandler(BaseHTTPRequestHandler):
        server_version = "WarrantyApprovalReview/1.0"

        def log_message(self, format: str, *args: object) -> None:
            LOGGER.debug("审核页面: " + format, *args)

        def _send_json(
            self,
            value: object,
            status: HTTPStatus = HTTPStatus.OK,
        ) -> None:
            body = json.dumps(value, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header(
                "Content-Type",
                "application/json; charset=utf-8",
            )
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(body)

        def _send_file(self, path: Path, content_type: str) -> None:
            body = path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header(
                "Content-Disposition",
                f"inline; filename*=UTF-8''{quote(path.name)}",
            )
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self) -> None:
            parsed = urlparse(self.path)
            if parsed.path in {"/", f"/{APPROVAL_REVIEW_HTML_FILE_NAME}"}:
                html_path = root / APPROVAL_REVIEW_HTML_FILE_NAME
                if not html_path.is_file():
                    self.send_error(HTTPStatus.NOT_FOUND, "审核页面尚未生成")
                    return
                self._send_file(html_path, "text/html; charset=utf-8")
                return
            if parsed.path == "/api/review":
                self._send_json({"ok": True, "review": review_repository.load()})
                return
            if parsed.path.startswith("/files/"):
                relative = unquote(parsed.path[len("/files/") :])
                try:
                    file_path = ensure_within(root / Path(relative), root)
                except ValueError:
                    self.send_error(HTTPStatus.FORBIDDEN)
                    return
                if not file_path.is_file():
                    self.send_error(HTTPStatus.NOT_FOUND)
                    return
                content_type = (
                    mimetypes.guess_type(file_path.name)[0]
                    or "application/octet-stream"
                )
                self._send_file(file_path, content_type)
                return
            self.send_error(HTTPStatus.NOT_FOUND)

        def do_POST(self) -> None:
            if urlparse(self.path).path != "/api/decisions":
                self.send_error(HTTPStatus.NOT_FOUND)
                return
            try:
                length = int(self.headers.get("Content-Length") or 0)
                if length <= 0 or length > MAX_REQUEST_BYTES:
                    raise ValueError("请求大小无效")
                payload = json.loads(
                    self.rfile.read(length).decode("utf-8")
                )
                if not isinstance(payload, dict):
                    raise ValueError("审核结果格式错误")
                with lock:
                    result = save_and_apply_review_payload(
                        root,
                        repo_root,
                        payload,
                    )
                self._send_json({"ok": True, **result})
            except (ValueError, json.JSONDecodeError) as exc:
                self._send_json(
                    {"ok": False, "error": str(exc)},
                    HTTPStatus.BAD_REQUEST,
                )
            except Exception:
                LOGGER.exception("保存人工审核结果失败")
                self._send_json(
                    {"ok": False, "error": "保存失败，请查看程序日志"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )

    server = ThreadingHTTPServer((host, port), ReviewHandler)
    display_host = "127.0.0.1" if host in {"0.0.0.0", "::"} else host
    url = f"http://{display_host}:{server.server_port}/"
    LOGGER.info("人工审核服务已启动: %s", url)
    LOGGER.info("审核结果将保存到: %s", review_repository.path)
    if open_browser:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        LOGGER.info("人工审核服务已停止")
    finally:
        server.server_close()


_HTML_TEMPLATE = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>待人工审核匹配 PDF</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #182536;
      --muted: #627287;
      --navy: #44546a;
      --blue: #5b9bd5;
      --line: #c9d5e3;
      --soft: #eaf2f9;
      --paper: #ffffff;
      --ok: #2f7d4b;
      --ok-bg: #e7f4eb;
      --warn: #9a6714;
      --warn-bg: #fff4d6;
      --bad: #a33a38;
      --bad-bg: #fde9e7;
      --shadow: 0 12px 32px rgba(35, 54, 78, .10);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: #f3f6fa;
      color: var(--ink);
      font-family: "Microsoft YaHei UI", "Microsoft YaHei", system-ui, sans-serif;
    }
    header {
      padding: 28px max(24px, calc((100vw - 1440px) / 2));
      color: white;
      background: linear-gradient(135deg, #35465d, #536b87);
    }
    header h1 { margin: 0 0 8px; font-size: 26px; }
    header p { margin: 0; color: #dce8f5; }
    main {
      width: min(1440px, calc(100% - 32px));
      margin: 22px auto 56px;
    }
    .toolbar, .summary, .card, .empty {
      background: var(--paper);
      border: 1px solid var(--line);
      border-radius: 12px;
      box-shadow: var(--shadow);
    }
    .toolbar {
      position: sticky;
      top: 10px;
      z-index: 5;
      display: flex;
      gap: 12px;
      align-items: center;
      justify-content: space-between;
      padding: 12px 16px;
      margin-bottom: 16px;
    }
    .status { display: flex; align-items: center; gap: 10px; }
    .dot { width: 10px; height: 10px; border-radius: 50%; background: #9aa8b7; }
    .dot.online { background: #4caf70; box-shadow: 0 0 0 4px #e0f2e6; }
    button, select, textarea { font: inherit; }
    button {
      border: 0;
      border-radius: 8px;
      padding: 10px 18px;
      color: white;
      background: var(--navy);
      cursor: pointer;
      font-weight: 700;
    }
    button:hover { background: #334258; }
    button:disabled { opacity: .48; cursor: not-allowed; }
    .summary {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      overflow: hidden;
      margin-bottom: 24px;
    }
    .metric { padding: 18px 22px; border-right: 1px solid var(--line); }
    .metric:last-child { border-right: 0; }
    .metric strong { display: block; font-size: 28px; color: var(--navy); }
    .metric span { color: var(--muted); font-size: 14px; }
    nav { display: flex; gap: 8px; margin: 0 0 14px; }
    nav button { background: #dfe7f0; color: var(--ink); }
    nav button.active { background: var(--navy); color: white; }
    .section-title { margin: 24px 0 12px; font-size: 20px; }
    .cards { display: grid; gap: 16px; }
    .card { overflow: hidden; }
    .card-head {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 18px;
      padding: 16px 18px;
      border-bottom: 1px solid var(--line);
      background: var(--soft);
    }
    .card-head h3 { margin: 0 0 5px; font-size: 17px; }
    .card-head small { color: var(--muted); }
    .confidence {
      min-width: 92px;
      text-align: center;
      align-self: center;
      font-size: 22px;
      font-weight: 800;
      color: var(--navy);
    }
    .confidence small { display: block; font-size: 12px; font-weight: 500; }
    .compare {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 0;
    }
    .panel { padding: 16px 18px; min-width: 0; }
    .panel:first-child { border-right: 1px solid var(--line); }
    .panel h4 { margin: 0 0 12px; color: var(--navy); }
    dl {
      display: grid;
      grid-template-columns: 92px minmax(0, 1fr);
      gap: 8px 12px;
      margin: 0;
      line-height: 1.55;
    }
    dt { color: var(--muted); }
    dd { margin: 0; word-break: break-word; }
    a { color: #1769aa; }
    .evidence {
      margin: 0;
      padding: 12px 18px;
      color: #45566b;
      background: #f8fafc;
      border-top: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      line-height: 1.6;
    }
    .decision {
      display: grid;
      grid-template-columns: 280px minmax(260px, 1fr);
      gap: 14px;
      align-items: center;
      padding: 14px 18px 18px;
    }
    .decision-actions {
      display: grid;
      grid-template-columns: 1fr auto;
      gap: 8px;
    }
    select, textarea {
      width: 100%;
      border: 1px solid #aebccd;
      border-radius: 7px;
      padding: 9px 11px;
      background: white;
      color: var(--ink);
    }
    textarea { min-height: 42px; resize: vertical; }
    .trash-button { background: var(--bad); white-space: nowrap; }
    .trash-button:hover { background: #842d2b; }
    .decision-confirm { border-left: 5px solid var(--ok); }
    .decision-exclude { border-left: 5px solid var(--bad); }
    .decision-trash { border-left: 8px solid #6f1d1b; background: #fffafa; }
    .empty { padding: 28px; text-align: center; color: var(--muted); }
    .unresolved { border-left: 5px solid var(--bad); }
    .history { border-left: 5px solid var(--ok); }
    .message { min-height: 22px; color: var(--muted); }
    .message.ok { color: var(--ok); }
    .message.error { color: var(--bad); }
    [hidden] { display: none !important; }
    @media (max-width: 820px) {
      .summary { grid-template-columns: 1fr 1fr; }
      .metric:nth-child(2) { border-right: 0; }
      .metric:nth-child(-n+2) { border-bottom: 1px solid var(--line); }
      .toolbar { position: static; align-items: flex-start; flex-direction: column; }
      .compare, .decision { grid-template-columns: 1fr; }
      .panel:first-child { border-right: 0; border-bottom: 1px solid var(--line); }
    }
  </style>
</head>
<body>
  <header>
    <h1>待人工审核匹配 PDF</h1>
    <p>只展示 _inbox 第一层未自动匹配 PDF 的最高置信度严格候选</p>
  </header>
  <main>
    <div class="toolbar">
      <div>
        <div class="status"><span id="dot" class="dot"></span><strong id="serverStatus">正在连接本地审核服务…</strong></div>
        <div id="message" class="message"></div>
      </div>
      <button id="saveButton" type="button" disabled>保存并执行审核结果</button>
    </div>
    <section class="summary" aria-label="审核概况">
      <div class="metric"><strong id="pendingCount">0</strong><span>待审核候选</span></div>
      <div class="metric"><strong id="selectedCount">0</strong><span>已填写决定</span></div>
      <div class="metric"><strong id="unresolvedCount">0</strong><span>无严格候选</span></div>
      <div class="metric"><strong id="revision">0</strong><span>正式数据版本</span></div>
    </section>
    <nav aria-label="审核内容">
      <button class="tab active" data-section="pending" type="button">最高候选</button>
      <button class="tab" data-section="unresolved" type="button">无严格候选</button>
      <button class="tab" data-section="history" type="button">已处理记录</button>
    </nav>
    <section id="pendingSection"><div id="pendingCards" class="cards"></div></section>
    <section id="unresolvedSection" hidden><div id="unresolvedCards" class="cards"></div></section>
    <section id="historySection" hidden><div id="historyCards" class="cards"></div></section>
  </main>
  <script id="reviewData" type="application/json">__REVIEW_DATA__</script>
  <script>
    let state = JSON.parse(document.getElementById("reviewData").textContent);
    let saveMode = "readonly";
    const byId = (id) => document.getElementById(id);
    const text = (tag, value, className) => {
      const node = document.createElement(tag);
      node.textContent = value ?? "";
      if (className) node.className = className;
      return node;
    };
    const addField = (list, label, value) => {
      list.append(text("dt", label), text("dd", value || "—"));
    };
    const fileUrl = (path) => {
      const normalized = String(path || "").replaceAll("\\", "/");
      if (location.protocol === "http:" || location.protocol === "https:") {
        return "/files/" + normalized.split("/").map(encodeURIComponent).join("/");
      }
      return normalized.split("/").map(encodeURIComponent).join("/");
    };
    function caseStatus(value) {
      return {
        materials_incomplete: "材料待补充",
        materials_ready: "材料齐全，待审批 PDF",
        approval_pdf_unmatched: "审批 PDF 待确认",
        approved: "审批完成",
        needs_review: "待人工确认",
        terminated: "终止"
      }[value] || value || "—";
    }
    function renderPendingCard(item) {
      const pdf = item.pdf || {};
      const candidate = item.case || {};
      const card = text("article", "", "card review-decision");
      card.dataset.reviewId = item.review_id;
      const head = text("div", "", "card-head");
      const title = text("div");
      title.append(text("h3", pdf.file_name || "未命名 PDF"));
      title.append(text("small", `审批编号：${pdf.application_no || "未识别"} · 识别方式：${pdf.recognition_method || "未知"}`));
      const confidence = text("div", `${item.confidence || 0}%`, "confidence");
      confidence.append(text("small", `${item.confidence_level || ""}置信度 · 严格候选 ${item.strict_candidate_count || 0} 个`));
      head.append(title, confidence);
      const compare = text("div", "", "compare");
      const pdfPanel = text("div", "", "panel");
      pdfPanel.append(text("h4", "审批 PDF 识别内容"));
      const pdfList = text("dl");
      addField(pdfList, "施工区域", pdf.area);
      addField(pdfList, "开始时间", pdf.start);
      addField(pdfList, "结束时间", pdf.end);
      addField(pdfList, "施工内容", pdf.content);
      const pdfLink = text("a", "在浏览器中打开审批 PDF");
      pdfLink.href = fileUrl(pdf.path);
      pdfLink.target = "_blank";
      pdfLink.rel = "noopener";
      const pdfLinkRow = text("dd");
      pdfLinkRow.append(pdfLink);
      pdfList.append(text("dt", "原文件"), pdfLinkRow);
      pdfPanel.append(pdfList);
      const casePanel = text("div", "", "panel");
      casePanel.append(text("h4", "推荐匹配案卷"));
      const caseList = text("dl");
      addField(caseList, "案卷名称", candidate.case_name);
      addField(caseList, "案卷状态", caseStatus(candidate.status));
      addField(caseList, "施工区域", candidate.area);
      addField(caseList, "开始时间", candidate.start);
      addField(caseList, "结束时间", candidate.end);
      addField(caseList, "施工内容", candidate.content);
      addField(caseList, "案卷目录", candidate.case_directory);
      casePanel.append(caseList);
      compare.append(pdfPanel, casePanel);
      const evidence = text("p", item.matching_evidence || "—", "evidence");
      const controls = text("div", "", "decision");
      const select = document.createElement("select");
      select.setAttribute("aria-label", "审核结果");
      ["待审核", "确认匹配", "排除", "移至_trash"].forEach((optionValue) => {
        const option = text("option", optionValue);
        option.value = optionValue;
        select.append(option);
      });
      select.value = item.decision || "待审核";
      const note = document.createElement("textarea");
      note.placeholder = "人工备注（可选）";
      note.value = item.review_note || "";
      note.setAttribute("aria-label", "人工备注");
      const updateStyle = () => {
        card.classList.toggle("decision-confirm", select.value === "确认匹配");
        card.classList.toggle("decision-exclude", select.value === "排除");
        card.classList.toggle("decision-trash", select.value === "移至_trash");
        updateMetrics();
      };
      select.addEventListener("change", updateStyle);
      const decisionActions = text("div", "", "decision-actions");
      const trashButton = text("button", "移至 _trash", "trash-button");
      trashButton.type = "button";
      trashButton.addEventListener("click", () => {
        if (window.confirm(`确认将“${pdf.file_name || "该 PDF"}”移至 _trash？\n点击“保存并执行审核结果”后立即移动，可从 _trash 恢复。`)) {
          select.value = "移至_trash";
          if (!note.value.trim()) note.value = "人工审核决定移至 _trash";
          updateStyle();
        }
      });
      decisionActions.append(select, trashButton);
      controls.append(decisionActions, note);
      card.append(head, compare, evidence, controls);
      updateStyle();
      return card;
    }
    function renderSimpleCard(item, mode) {
      const pdf = item.pdf || {};
      const card = text("article", "", `card ${mode === "history" ? "history" : "unresolved review-decision"}`);
      if (mode !== "history") card.dataset.reviewId = item.review_id;
      const head = text("div", "", "card-head");
      const title = text("div");
      if (mode === "history") {
        title.append(text("h3", item.pdf_file_name || "已处理审批 PDF"));
        title.append(text("small", `${item.decision || ""} · ${item.decided_at || ""}`));
      } else {
        title.append(text("h3", pdf.file_name || "未命名 PDF"));
        title.append(text("small", item.reason || "未形成严格候选"));
      }
      head.append(title);
      const panel = text("div", "", "panel");
      const list = text("dl");
      if (mode === "history") {
        addField(list, "审核结果", item.decision);
        addField(list, "候选案卷", item.case_name);
        addField(list, "人工备注", item.review_note);
        addField(list, "处理后路径", item.result_path);
      } else {
        addField(list, "施工区域", pdf.area);
        addField(list, "开始时间", pdf.start);
        addField(list, "结束时间", pdf.end);
        addField(list, "施工内容", pdf.content);
        const link = text("a", "在浏览器中打开审批 PDF");
        link.href = fileUrl(pdf.path);
        link.target = "_blank";
        link.rel = "noopener";
        const linkRow = text("dd");
        linkRow.append(link);
        list.append(text("dt", "原文件"), linkRow);
      }
      panel.append(list);
      card.append(head, panel);
      if (mode !== "history") {
        const controls = text("div", "", "decision");
        const decisionActions = text("div", "", "decision-actions");
        const select = document.createElement("select");
        select.setAttribute("aria-label", "无严格候选 PDF 处理结果");
        ["待审核", "移至_trash"].forEach((optionValue) => {
          const option = text("option", optionValue);
          option.value = optionValue;
          select.append(option);
        });
        select.value = item.decision || "待审核";
        const note = document.createElement("textarea");
        note.placeholder = "移至 _trash 的原因（可选）";
        note.value = item.review_note || "";
        note.setAttribute("aria-label", "人工备注");
        const updateStyle = () => {
          card.classList.toggle("decision-trash", select.value === "移至_trash");
          updateMetrics();
        };
        select.addEventListener("change", updateStyle);
        const trashButton = text("button", "移至 _trash", "trash-button");
        trashButton.type = "button";
        trashButton.addEventListener("click", () => {
          if (window.confirm(`确认将“${pdf.file_name || "该 PDF"}”移至 _trash？\n点击“保存并执行审核结果”后立即移动，可从 _trash 恢复。`)) {
            select.value = "移至_trash";
            if (!note.value.trim()) note.value = "人工审核决定移至 _trash";
            updateStyle();
          }
        });
        decisionActions.append(select, trashButton);
        controls.append(decisionActions, note);
        card.append(controls);
        updateStyle();
      }
      return card;
    }
    function renderList(target, items, renderer, emptyText) {
      target.replaceChildren();
      if (!items.length) {
        target.append(text("div", emptyText, "empty"));
        return;
      }
      items.forEach((item) => target.append(renderer(item)));
    }
    function updateMetrics() {
      const selects = [...document.querySelectorAll(".review-decision select")];
      byId("pendingCount").textContent = (state.pending_reviews || []).length;
      byId("selectedCount").textContent = selects.filter((select) => select.value !== "待审核").length;
      byId("unresolvedCount").textContent = (state.unresolved_pdfs || []).length;
      byId("revision").textContent = state.source_dataset_revision || 0;
    }
    function render() {
      renderList(byId("pendingCards"), state.pending_reviews || [], renderPendingCard, "当前没有需要人工审核的最高候选");
      renderList(byId("unresolvedCards"), state.unresolved_pdfs || [], (item) => renderSimpleCard(item, "unresolved"), "当前没有无严格候选的 PDF");
      renderList(byId("historyCards"), state.decisions || [], (item) => renderSimpleCard(item, "history"), "当前还没有已处理审核记录");
      updateMetrics();
    }
    function setMessage(value, kind = "") {
      byId("message").textContent = value;
      byId("message").className = `message ${kind}`;
    }
    async function connect() {
      if (location.protocol === "file:") {
        byId("serverStatus").textContent = "请通过审核启动器打开";
        setMessage("双击同目录的“打开待人工审核匹配PDF.cmd”；保存后会立即归档或移至 _trash。", "error");
        return;
      }
      try {
        const response = await fetch("/api/review", {cache: "no-store"});
        const result = await response.json();
        if (!response.ok || !result.ok) throw new Error(result.error || "连接失败");
        state = result.review;
        saveMode = "server";
        byId("dot").classList.add("online");
        byId("serverStatus").textContent = "本地审核服务已连接";
        byId("saveButton").disabled = false;
        setMessage("保存后立即执行决定，并同步正式 JSON、汇总 HTML 和本审核页面。");
        render();
      } catch (error) {
        setMessage(error.message || "无法连接本地审核服务", "error");
      }
    }
    function collectDecisions() {
      const cards = [...document.querySelectorAll(".review-decision")];
      const expected = new Map([
        ...(state.pending_reviews || []),
        ...(state.unresolved_pdfs || [])
      ].map((item) => [item.review_id, item]));
      return cards.map((card) => {
        const item = expected.get(card.dataset.reviewId);
        return {
          review_id: item.review_id,
          decision: card.querySelector("select").value,
          review_note: card.querySelector("textarea").value,
          pdf_sha256: (item.pdf || {}).sha256 || "",
          case_id: (item.case || {}).case_id || ""
        };
      });
    }
    async function save() {
      if (saveMode === "readonly") return;
      const decisions = collectDecisions();
      byId("saveButton").disabled = true;
      setMessage("正在校验、归档并刷新数据，请稍候…");
      try {
        const response = await fetch("/api/decisions", {
          method: "POST",
          headers: {"Content-Type": "application/json"},
          body: JSON.stringify({
            source_dataset_revision: state.source_dataset_revision,
            decisions
          })
        });
        const result = await response.json();
        if (!response.ok || !result.ok) throw new Error(result.error || "保存失败");
        state = result.review;
        render();
        setMessage(
          `已处理 ${result.review_decisions_applied} 条：确认归档 ${result.approval_pdfs_human_confirmed} 条，移至 _trash ${result.approval_pdfs_moved_to_trash} 条；待审核列表已刷新。`,
          "ok"
        );
      } catch (error) {
        setMessage(error.message || "保存失败", "error");
      } finally {
        byId("saveButton").disabled = false;
      }
    }
    document.querySelectorAll(".tab").forEach((button) => {
      button.addEventListener("click", () => {
        document.querySelectorAll(".tab").forEach((item) => item.classList.toggle("active", item === button));
        ["pending", "unresolved", "history"].forEach((name) => {
          byId(`${name}Section`).hidden = name !== button.dataset.section;
        });
      });
    });
    byId("saveButton").addEventListener("click", save);
    render();
    connect();
  </script>
</body>
</html>
"""
