from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from ..flows.command_flow import main as run_command


def run_simple_entry(
    command: str,
    description: str,
    argv: Sequence[str] | None = None,
) -> int:
    parser = argparse.ArgumentParser(
        description=description,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        help="资料根目录；默认从 config.yaml/common.env 读取",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)
    forwarded: list[str] = []
    if args.input_dir:
        forwarded.extend(["--input-dir", str(args.input_dir)])
    forwarded.append(command)
    return run_command(forwarded)
