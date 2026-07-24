from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

from . import legacy
from .file_utils import sha256_file


class RecognitionService:
    """One shared OCR service and one persistent cache for every workflow."""

    def __init__(
        self,
        dataset: dict[str, Any],
        repo_root: Path,
    ) -> None:
        self.dataset = dataset
        self.repo_root = repo_root
        self.cache: dict[str, dict[str, Any]] = dataset.setdefault(
            "recognition_cache", {}
        )
        self.client: legacy.LlamaCppClient | None = None

    def __enter__(self) -> "RecognitionService":
        return self

    def __exit__(self, *_args: object) -> None:
        self.close()

    def close(self) -> None:
        if self.client is not None:
            self.client.shutdown_server()
            self.client = None

    def _ensure_client(self) -> legacy.LlamaCppClient:
        if self.client is None:
            self.client = legacy.LlamaCppClient(
                legacy.LlamaCppConfig.from_repo(self.repo_root),
                self.repo_root,
            )
            self.client.ensure_server()
            self.client.assert_model_available()
        return self.client

    def _cached(self, fingerprint: str) -> str | None:
        item = self.cache.get(fingerprint)
        if not item or not isinstance(item.get("text"), str):
            return None
        return str(item["text"])

    def _store(
        self,
        path: Path,
        fingerprint: str,
        text: str,
        method: str,
    ) -> str:
        normalized = legacy.normalize_match_text(text)
        self.cache[fingerprint] = {
            "text": normalized,
            "method": method,
            "file_name": path.name,
            "recognized_at": datetime.now().astimezone().isoformat(
                timespec="seconds"
            ),
            "recognition_version": 1,
        }
        return normalized

    def pdf_text(self, path: Path) -> str:
        fingerprint = sha256_file(path)
        cached = self._cached(fingerprint)
        if cached is not None:
            return cached
        plain_text = legacy.read_pdf_plain_text(path)
        if legacy.count_cjk_chars(plain_text) >= legacy.MIN_PLAIN_PDF_CJK_CHARS:
            return self._store(path, fingerprint, plain_text, "plain")
        ocr_text = legacy.read_pdf_ai_ocr_text(path, self._ensure_client())
        return self._store(path, fingerprint, ocr_text, "ocr")

    def image_text(self, path: Path) -> str:
        fingerprint = sha256_file(path)
        cached = self._cached(fingerprint)
        if cached is not None:
            return cached
        text = self._ensure_client().extract_image_text(path)
        return self._store(path, fingerprint, text, "ocr")
