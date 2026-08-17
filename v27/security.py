"""Filesystem and log-sanitization boundaries for V2.7."""

import re
from pathlib import Path
from typing import Any

from fastapi import HTTPException


_SENSITIVE_KEY = re.compile(
    r"(token|secret|authorization|credential|client_secret|refresh_token)",
    re.IGNORECASE,
)
_BEARER = re.compile(r"Bearer\s+[A-Za-z0-9._~+\-/]+=*", re.IGNORECASE)


def sanitize_for_log(value: Any) -> Any:
    """Recursively redact secrets before structured events leave the app."""
    if isinstance(value, dict):
        return {
            key: "[REDACTED]" if _SENSITIVE_KEY.search(str(key)) else sanitize_for_log(item)
            for key, item in value.items()
        }
    if isinstance(value, list):
        return [sanitize_for_log(item) for item in value]
    if isinstance(value, tuple):
        return tuple(sanitize_for_log(item) for item in value)
    if isinstance(value, str):
        return _BEARER.sub("Bearer [REDACTED]", value)
    return value


def resolve_safe_output_path(outputs_dir: Path, filename: str) -> Path:
    if not filename or Path(filename).name != filename:
        raise HTTPException(status_code=400, detail="Nome de arquivo inválido.")

    base = outputs_dir.resolve()
    candidate = (base / filename).resolve(strict=False)

    if candidate.parent != base:
        raise HTTPException(status_code=400, detail="Nome de arquivo inválido.")

    return candidate
