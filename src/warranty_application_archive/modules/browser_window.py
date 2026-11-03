from __future__ import annotations

import shutil
import os
import subprocess
import sys
import threading
import time
import webbrowser
from logging import Logger
from pathlib import Path


def _chrome_path() -> Path | None:
    if sys.platform != "win32":
        return None
    candidates = [shutil.which("chrome.exe"), shutil.which("chrome")]
    for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(variable)
        if base:
            candidates.append(
                str(Path(base) / "Google/Chrome/Application/chrome.exe")
            )
    return next(
        (Path(candidate) for candidate in candidates if candidate and Path(candidate).is_file()),
        None,
    )


def _matching_chrome_windows(title_markers: tuple[str, ...]) -> set[int]:
    if sys.platform != "win32":
        return set()

    import ctypes
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.EnumWindows.argtypes = [callback_type, wintypes.LPARAM]
    user32.EnumWindows.restype = wintypes.BOOL
    found: set[int] = set()

    def collect(hwnd: int, _lparam: int) -> bool:
        if not user32.IsWindowVisible(hwnd):
            return True
        length = user32.GetWindowTextLengthW(hwnd)
        if not length:
            return True
        title = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title, length + 1)
        if any(marker in title.value for marker in title_markers):
            found.add(int(hwnd))
        return True

    user32.EnumWindows(callback_type(collect), 0)
    return found


def _move_to_primary_monitor(hwnd: int) -> None:
    import ctypes
    from ctypes import wintypes

    class Rect(ctypes.Structure):
        _fields_ = [(name, wintypes.LONG) for name in ("left", "top", "right", "bottom")]

    class MonitorInfo(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", Rect),
            ("rcWork", Rect),
            ("dwFlags", wintypes.DWORD),
        ]

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(
        wintypes.BOOL,
        wintypes.HANDLE,
        wintypes.HDC,
        ctypes.POINTER(Rect),
        wintypes.LPARAM,
    )
    user32.GetMonitorInfoW.argtypes = [wintypes.HANDLE, ctypes.POINTER(MonitorInfo)]
    user32.GetMonitorInfoW.restype = wintypes.BOOL
    user32.EnumDisplayMonitors.argtypes = [
        wintypes.HDC,
        ctypes.POINTER(Rect),
        callback_type,
        wintypes.LPARAM,
    ]
    user32.EnumDisplayMonitors.restype = wintypes.BOOL
    primary: list[Rect] = []

    def collect(monitor: int, _hdc: int, _rect: object, _lparam: int) -> bool:
        info = MonitorInfo()
        info.cbSize = ctypes.sizeof(MonitorInfo)
        if user32.GetMonitorInfoW(monitor, ctypes.byref(info)) and info.dwFlags & 1:
            primary.append(info.rcWork)
            return False
        return True

    user32.EnumDisplayMonitors(None, None, callback_type(collect), 0)
    if not primary:
        return
    work = primary[0]
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.SetWindowPos.argtypes = [
        wintypes.HWND,
        wintypes.HWND,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        ctypes.c_int,
        wintypes.UINT,
    ]
    user32.SetWindowPos.restype = wintypes.BOOL
    user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    user32.SetWindowPos(
        hwnd,
        None,
        work.left,
        work.top,
        work.right - work.left,
        work.bottom - work.top,
        0x0014,
    )
    user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE


def open_page_on_primary_monitor(
    url: str,
    *,
    title_markers: tuple[str, ...] = ("质保作业申请汇总", "待人工审核匹配 PDF"),
    logger: Logger | None = None,
) -> None:
    """Open a dedicated Chrome window and place it on the Windows primary display."""
    chrome = _chrome_path()
    if chrome is None:
        webbrowser.open_new(url)
        return

    before = _matching_chrome_windows(title_markers)
    try:
        subprocess.Popen(
            [str(chrome), "--new-window", url],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except OSError:
        if logger:
            logger.warning("启动 Chrome 失败，将改用默认浏览器", exc_info=True)
        webbrowser.open_new(url)
        return

    def position_window() -> None:
        for _ in range(40):
            opened = _matching_chrome_windows(title_markers) - before
            if opened:
                for hwnd in opened:
                    _move_to_primary_monitor(hwnd)
                return
            time.sleep(0.25)
        if logger:
            logger.warning("未能定位新开的 Chrome 页面窗口；页面仍保持打开")

    threading.Thread(target=position_window, name="position-html-window", daemon=True).start()
