from __future__ import annotations

import logging
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from . import legacy
from .file_utils import sha256_file


LOGGER = logging.getLogger(__name__)


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
            LOGGER.info("正在连接或启动本地 AI 服务")
            self.client = legacy.LlamaCppClient(
                legacy.LlamaCppConfig.from_repo(self.repo_root),
                self.repo_root,
            )
            self.client.ensure_server()
            self.client.assert_model_available()
            LOGGER.info(
                "本地 AI 服务已就绪，模型: %s",
                self.client.config.model,
            )
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
        started_at = time.perf_counter()
        fingerprint = sha256_file(path)
        cached = self._cached(fingerprint)
        if cached is not None:
            method = str(self.cache[fingerprint].get("method") or "unknown")
            LOGGER.info(
                "PDF 文本识别命中缓存: %s（识别方式: %s）",
                path.name,
                method,
            )
            return cached
        LOGGER.info("正在提取 PDF 内嵌文本: %s", path.name)
        plain_text = legacy.read_pdf_plain_text(path)
        cjk_chars = legacy.count_cjk_chars(plain_text)
        if cjk_chars >= legacy.MIN_PLAIN_PDF_CJK_CHARS:
            LOGGER.info(
                "PDF 内嵌文本可用，无需调用本地 AI: %s（中文字符 %s 个）",
                path.name,
                cjk_chars,
            )
            return self._store(path, fingerprint, plain_text, "plain")
        LOGGER.info(
            "PDF 内嵌文本不足，准备调用本地 AI OCR: %s（中文字符 %s 个）",
            path.name,
            cjk_chars,
        )
        client = self._ensure_client()
        LOGGER.info("本地 AI OCR 开始识别 PDF: %s", path.name)
        ocr_started_at = time.perf_counter()
        ocr_text = legacy.read_pdf_ai_ocr_text(path, client)
        LOGGER.info(
            "本地 AI OCR 识别完成: %s（OCR 耗时 %.1f 秒，总耗时 %.1f 秒）",
            path.name,
            time.perf_counter() - ocr_started_at,
            time.perf_counter() - started_at,
        )
        return self._store(path, fingerprint, ocr_text, "ocr")

    def image_text(self, path: Path) -> str:
        fingerprint = sha256_file(path)
        cached = self._cached(fingerprint)
        if cached is not None:
            LOGGER.info("图片文字识别命中缓存: %s", path.name)
            return cached
        client = self._ensure_client()
        LOGGER.info("本地 AI OCR 开始识别图片: %s", path.name)
        started_at = time.perf_counter()
        text = client.extract_image_text(path)
        LOGGER.info(
            "本地 AI OCR 识别图片完成: %s（耗时 %.1f 秒）",
            path.name,
            time.perf_counter() - started_at,
        )
        return self._store(path, fingerprint, text, "ocr")
