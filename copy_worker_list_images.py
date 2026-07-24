"""Compatibility entry point for the worker-list intake subworkflow."""

import sys

from src.docflow_renamer.cli import main


if __name__ == "__main__":
    raise SystemExit(main([*sys.argv[1:], "worker-lists"]))
