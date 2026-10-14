"""质保作业申请增量归档入口。

Python 源码默认运行命令行归档；--gui 或打包后的 EXE 显示图形控制窗口。
"""

from __future__ import annotations

import argparse
import logging
import os
import queue
import sys
import threading
import webbrowser
from http.server import ThreadingHTTPServer
from pathlib import Path
from tkinter import Button, Frame, Label, StringVar, Tk, messagebox
from urllib.parse import quote
from urllib.request import urlopen

from logging_config import setup_logger


if not getattr(sys, "frozen", False):
    sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from warranty_application_archive.config_loader import (  # noqa: E402
    AppConfig,
    runtime_project_root,
)

from warranty_application_archive.flows.application_flow import (  # noqa: E402
    run_archive,
)
from warranty_application_archive.flows.approval_review_web_flow import (  # noqa: E402
    serve_approval_review,
)
from warranty_application_archive.modules.constants import (  # noqa: E402
    APPROVAL_REVIEW_HTML_FILE_NAME,
)


PROJECT_ROOT = (
    runtime_project_root()
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
LOGGER = logging.getLogger(__name__)


def _icon_path() -> Path:
    bundle_root = Path(getattr(sys, "_MEIPASS", PROJECT_ROOT))
    return bundle_root / "assets" / "docflow_archive.ico"


def run_gui(
    data_root: Path,
    *,
    open_browser: bool,
    skip_archive: bool,
    self_test: bool = False,
) -> int:
    root = Tk()
    root.title("质保作业申请归档")
    root.geometry("480x220")
    root.resizable(False, False)
    icon = _icon_path()
    if icon.is_file():
        try:
            root.iconbitmap(default=str(icon))
        except Exception:
            LOGGER.warning("窗口图标加载失败: %s", icon, exc_info=True)

    status = StringVar(value="准备启动归档服务…")
    address = StringVar(value=str(data_root))
    Label(root, textvariable=status, font=("Microsoft YaHei UI", 13, "bold")).pack(
        pady=(23, 8)
    )
    Label(root, textvariable=address, font=("Segoe UI", 9), fg="#285a80").pack(
        pady=(0, 18)
    )
    actions = Frame(root)
    actions.pack()
    events: queue.Queue[tuple[str, object]] = queue.Queue()
    server: ThreadingHTTPServer | None = None
    url = ""
    busy = False
    closing = False
    self_test_ok = False

    def open_summary() -> None:
        if url:
            webbrowser.open(url)

    def open_review() -> None:
        if url:
            webbrowser.open(url + quote(APPROVAL_REVIEW_HTML_FILE_NAME))

    def server_ready(value: ThreadingHTTPServer, address_url: str) -> None:
        events.put(("server_ready", (value, address_url)))

    def serve() -> None:
        try:
            serve_approval_review(
                data_root,
                host="127.0.0.1",
                port=0,
                open_browser=False,
                initial_page="summary",
                on_server_ready=server_ready,
            )
        except Exception as exc:
            LOGGER.exception("本地页面服务启动失败")
            events.put(("error", str(exc)))

    def archive() -> None:
        try:
            result = run_archive(data_root)
            if result != 0:
                raise RuntimeError(f"归档流程退出代码：{result}")
            events.put(("archive_done", None))
        except Exception as exc:
            LOGGER.exception("归档执行失败")
            events.put(("error", str(exc)))

    def start_archive() -> None:
        nonlocal busy
        if busy or closing:
            return
        busy = True
        status.set("正在识别并归档，请稍候…")
        archive_button.configure(state="disabled")
        threading.Thread(target=archive, daemon=True).start()

    def close() -> None:
        nonlocal closing
        if busy:
            messagebox.showinfo("归档进行中", "请等待本次归档完成后再关闭窗口。")
            return
        if closing:
            return
        closing = True
        if server is None:
            root.destroy()
            return
        status.set("正在停止本地服务…")
        archive_button.configure(state="disabled")
        summary_button.configure(state="disabled")
        review_button.configure(state="disabled")
        stop_button.configure(state="disabled")

        def shutdown() -> None:
            server.shutdown()
            events.put(("stopped", None))

        threading.Thread(target=shutdown, daemon=True).start()

    archive_button = Button(actions, text="更新归档", width=12, command=start_archive)
    archive_button.pack(side="left", padx=3)
    summary_button = Button(
        actions, text="打开汇总", width=12, command=open_summary, state="disabled"
    )
    summary_button.pack(side="left", padx=3)
    review_button = Button(
        actions, text="人工审核", width=12, command=open_review, state="disabled"
    )
    review_button.pack(side="left", padx=3)
    stop_button = Button(actions, text="退出", width=8, command=close)
    stop_button.pack(side="left", padx=3)
    root.protocol("WM_DELETE_WINDOW", close)

    def poll() -> None:
        nonlocal busy, server, url, self_test_ok
        while True:
            try:
                event, payload = events.get_nowait()
            except queue.Empty:
                break
            if event == "archive_done":
                busy = False
                archive_button.configure(state="normal")
                status.set("归档完成，本地页面已就绪" if url else "归档完成，正在启动页面…")
                if server is None:
                    threading.Thread(target=serve, daemon=True).start()
            elif event == "server_ready":
                server, url = payload  # type: ignore[misc]
                address.set(url)
                status.set("本地归档页面正在运行")
                summary_button.configure(state="normal")
                review_button.configure(state="normal")
                if self_test:
                    try:
                        with urlopen(url, timeout=5) as response:
                            self_test_ok = response.status == 200
                    except Exception:
                        LOGGER.exception("GUI 自检无法打开本地页面")
                    root.after(100, close)
                if open_browser:
                    root.after(250, open_summary)
            elif event == "error":
                busy = False
                archive_button.configure(state="normal")
                status.set("启动或归档失败，请查看日志")
                if self_test:
                    LOGGER.error("GUI 自检失败: %s", payload)
                    root.after(100, close)
                else:
                    messagebox.showerror("质保作业申请归档", str(payload))
            elif event == "stopped":
                root.destroy()
                return
        root.after(100, poll)

    root.after(100, poll)
    if skip_archive:
        root.after(200, lambda: threading.Thread(target=serve, daemon=True).start())
    else:
        root.after(200, start_archive)
    root.mainloop()
    return 0 if not self_test or self_test_ok else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, help="临时指定资料根目录")
    parser.add_argument("--gui", action="store_true", help="启动图形归档窗口")
    parser.add_argument("--no-open", action="store_true", help="不自动打开浏览器")
    parser.add_argument("--skip-archive", action="store_true", help="只打开现有页面")
    parser.add_argument("--check-config", action="store_true", help="仅检查配置后退出")
    parser.add_argument("--self-test", action="store_true", help="启动页面自检后退出")
    parser.add_argument(
        "--self-test-archive",
        action="store_true",
        help="执行归档和页面自检后退出",
    )
    args = parser.parse_args()
    if not getattr(sys, "frozen", False) and not any(
        (
            args.gui,
            args.no_open,
            args.skip_archive,
            args.check_config,
            args.self_test,
            args.self_test_archive,
        )
    ):
        return run_archive(args.input_dir)
    os.chdir(PROJECT_ROOT)
    try:
        config = AppConfig.resolve(PROJECT_ROOT, args.input_dir)
        if not config.data_root.is_dir():
            raise NotADirectoryError(f"资料根目录不存在：{config.data_root}")
        if args.check_config:
            return 0
        setup_logger(
            log_level=config.log_level,
            log_file=config.log_dir / "run_archive_gui.log",
        )
        return run_gui(
            config.data_root,
            open_browser=(
                not args.no_open
                and not args.self_test
                and not args.self_test_archive
            ),
            skip_archive=args.skip_archive or args.self_test,
            self_test=args.self_test or args.self_test_archive,
        )
    except Exception as exc:
        LOGGER.exception("图形启动器失败")
        if sys.platform == "win32" and getattr(sys, "frozen", False):
            import ctypes

            ctypes.windll.user32.MessageBoxW(0, str(exc), "质保归档启动失败", 0x10)
        else:
            print(f"质保归档启动失败：{exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
