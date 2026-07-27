from __future__ import annotations

import os
import shlex
import stat
import subprocess
from pathlib import Path
from typing import Sequence

from .file_utils import atomic_replace_text


def write_page_launchers(
    root: Path,
    working_directory: Path,
    windows_command: Sequence[str],
    macos_server_arguments: Sequence[str],
    *,
    windows_name: str,
    macos_name: str,
) -> tuple[Path, Path]:
    """Generate equivalent Windows and macOS launchers for a local page."""
    windows_path = root / windows_name
    atomic_replace_text(
        windows_path,
        "\n".join(
            [
                "@echo off",
                "chcp 65001 >nul",
                f'cd /d "{working_directory}"',
                subprocess.list2cmdline(list(windows_command)),
                "pause",
                "",
            ]
        ),
    )

    macos_arguments = shlex.join(list(macos_server_arguments))
    macos_path = root / macos_name
    atomic_replace_text(
        macos_path,
        "\n".join(
            [
                "#!/bin/zsh",
                "set -u",
                'SCRIPT_DIR="${0:A:h}"',
                'PROJECT_ROOT="${DOCFLOW_PROJECT_ROOT:-}"',
                'if [[ -z "$PROJECT_ROOT" ]]; then',
                '  SEARCH_ROOT="$SCRIPT_DIR"',
                '  while [[ "$SEARCH_ROOT" != "/" ]]; do',
                '    CANDIDATE="$SEARCH_ROOT/Python/Project/DocFlow_Renamer"',
                '    if [[ -f "$CANDIDATE/warranty_application_archive.py" ]]; then',
                '      PROJECT_ROOT="$CANDIDATE"',
                "      break",
                "    fi",
                '    SEARCH_ROOT="${SEARCH_ROOT:h}"',
                "  done",
                "fi",
                'if [[ -z "$PROJECT_ROOT" ]]; then',
                "  for CANDIDATE in \\",
                '    "$HOME/Clooustation/Python/Project/DocFlow_Renamer" \\',
                '    "$HOME/CloudStation/Python/Project/DocFlow_Renamer" \\',
                '    "$HOME/Cloudstation/Python/Project/DocFlow_Renamer"; do',
                '    if [[ -f "$CANDIDATE/warranty_application_archive.py" ]]; then',
                '      PROJECT_ROOT="$CANDIDATE"',
                "      break",
                "    fi",
                "  done",
                "fi",
                'if [[ ! -f "$PROJECT_ROOT/warranty_application_archive.py" ]]; then',
                '  echo "未找到 DocFlow_Renamer 项目目录。"',
                '  echo "可设置 DOCFLOW_PROJECT_ROOT 后重新运行。"',
                "  exit 1",
                "fi",
                'PYTHON_BIN="${DOCFLOW_PYTHON:-}"',
                'if [[ -z "$PYTHON_BIN" ]]; then',
                "  for CANDIDATE in \\",
                '    "$HOME/anaconda3/bin/python" \\',
                '    "$HOME/opt/anaconda3/bin/python" \\',
                '    "/opt/anaconda3/bin/python" \\',
                '    "${commands[python3]:-}"; do',
                '    if [[ -n "$CANDIDATE" && -x "$CANDIDATE" ]]; then',
                '      PYTHON_BIN="$CANDIDATE"',
                "      break",
                "    fi",
                "  done",
                "fi",
                'if [[ -z "$PYTHON_BIN" ]]; then',
                '  echo "未找到可用的 Python。"',
                '  echo "可设置 DOCFLOW_PYTHON 后重新运行。"',
                "  exit 1",
                "fi",
                'cd "$PROJECT_ROOT"',
                (
                    'exec "$PYTHON_BIN" '
                    '"$PROJECT_ROOT/warranty_application_archive.py" '
                    '--input-dir "$SCRIPT_DIR" '
                    f"{macos_arguments}"
                ),
                "",
            ]
        ),
    )
    current_mode = macos_path.stat().st_mode
    os.chmod(
        macos_path,
        current_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH,
    )
    return windows_path, macos_path
