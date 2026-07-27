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
    command: Sequence[str],
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
                subprocess.list2cmdline(list(command)),
                "pause",
                "",
            ]
        ),
    )

    macos_path = root / macos_name
    atomic_replace_text(
        macos_path,
        "\n".join(
            [
                "#!/bin/zsh",
                f"cd {shlex.quote(str(working_directory))}",
                f"exec {shlex.join(list(command))}",
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
