"""
Exception hierarchy and retry utilities for MyPDF-OCR / THSummarizer.
"""

from __future__ import annotations

import time
import urllib.error
from typing import Any, Callable, TypeVar
from tenacity import retry, retry_if_exception, stop_after_attempt, wait_exponential

T = TypeVar("T")


class OCRError(Exception):
    """Base exception for all OCR and document processing errors."""
    pass


class TransientOCRError(OCRError):
    """Temporary errors (Rate limit 429, 503 Service Unavailable, network timeout)."""
    pass


class PermanentOCRError(OCRError):
    """Permanent errors (401 Unauthorized, 404 Model Not Found, 400 Invalid Arguments)."""
    pass


def is_transient_error(exc: BaseException) -> bool:
    """Determine whether an error is transient and eligible for automatic retry."""
    if isinstance(exc, TransientOCRError):
        return True
    if isinstance(exc, urllib.error.HTTPError):
        # 429 Rate Limit, 500 Internal, 502 Bad Gateway, 503 Unavailable, 504 Gateway Timeout
        return exc.code in (429, 500, 502, 503, 504)
    if isinstance(exc, (urllib.error.URLError, TimeoutError, ConnectionError, OSError)):
        return True
    return False


def with_api_retry(
    max_attempts: int = 3,
    min_wait: float = 1.5,
    max_wait: float = 10.0,
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """Decorator to retry transient API calls with exponential backoff."""
    return retry(
        reraise=True,
        stop=stop_after_attempt(max_attempts),
        wait=wait_exponential(multiplier=1.5, min=min_wait, max=max_wait),
        retry=retry_if_exception(is_transient_error),
    )
