"""根目录独立入口共享的源码路径启动辅助。"""

from __future__ import annotations

import sys
from pathlib import Path


SRC_DIR = Path(__file__).resolve().parent / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from warranty_application_archive.flows.command_flow import (  # noqa: E402
    main as run_command,
)
from warranty_application_archive.modules.entry_support import (  # noqa: E402
    run_simple_entry,
)


__all__ = ["run_command", "run_simple_entry"]
