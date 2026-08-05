import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from warranty_application_archive.config_loader import AppConfig


CONFIG_TEXT = """app:
  log_level: ${LOG_LEVEL:-INFO}
  input_path: ${INPUT_PATH:-input}
  output_dir: ${OUTPUT_DIR:-output}
  log_dir: ${LOG_DIR:-logs}
flows:
  archive: {}
"""


class ConfigLoaderTest(unittest.TestCase):
    def test_nested_cloudstation_path_and_platform_value_are_expanded(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            repo_root = Path(temporary_dir)
            cloud_root = repo_root / "cloud"
            (repo_root / "config.yaml").write_text(
                CONFIG_TEXT,
                encoding="utf-8",
            )
            (repo_root / "common.env").write_text(
                "\n".join(
                    [
                        f"CLOUDSTATION_ROOT_WINDOWS={cloud_root}",
                        "INPUT_PATH=${CLOUDSTATION_ROOT}/records",
                        "LOG_DIR=logs",
                    ]
                ),
                encoding="utf-8",
            )

            with (
                patch.dict(os.environ, {}, clear=True),
                patch(
                    "warranty_application_archive.config_loader."
                    "sys.platform",
                    "win32",
                ),
            ):
                config = AppConfig.resolve(repo_root)

            self.assertEqual(config.data_root, (cloud_root / "records").resolve())
            self.assertEqual(config.log_dir, (repo_root / "logs").resolve())

    def test_process_environment_has_priority_over_common_env(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_dir:
            repo_root = Path(temporary_dir)
            (repo_root / "config.yaml").write_text(
                CONFIG_TEXT,
                encoding="utf-8",
            )
            (repo_root / "common.env").write_text(
                "LOG_LEVEL=WARNING\nINPUT_PATH=from-file",
                encoding="utf-8",
            )

            with patch.dict(
                os.environ,
                {"LOG_LEVEL": "DEBUG", "INPUT_PATH": "from-process"},
                clear=True,
            ):
                config = AppConfig.resolve(repo_root)

            self.assertEqual(config.log_level, "DEBUG")
            self.assertEqual(
                config.data_root,
                (repo_root / "from-process").resolve(),
            )


if __name__ == "__main__":
    unittest.main()
