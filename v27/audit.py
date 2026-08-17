"""Structured and persistent audit-log contracts for V2.7."""

from __future__ import annotations

import json
import logging
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Protocol

from .security import sanitize_for_log


class AuditLogger(Protocol):
    def record(self, event: dict[str, Any]) -> None: ...


def _envelope(event: dict[str, Any]) -> dict[str, Any]:
    return sanitize_for_log({
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "version": "v2.7",
        **event,
    })


class StructuredAuditLogger:
    def __init__(self, logger: logging.Logger | None = None):
        self._logger = logger or logging.getLogger("estate_gover.audit.v27")

    def record(self, event: dict[str, Any]) -> None:
        self._logger.info(
            json.dumps(_envelope(event), ensure_ascii=False, sort_keys=True)
        )


class JsonlAuditLogger:
    """Append-only JSONL audit store for homologation.

    Existing log history is never deleted or rewritten by this class.
    """

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()

    def record(self, event: dict[str, Any]) -> None:
        payload = _envelope(event)
        line = json.dumps(payload, ensure_ascii=False, sort_keys=True)
        with self._lock:
            with self.path.open("a", encoding="utf-8") as handle:
                handle.write(line + "\n")


class InMemoryAuditLogger:
    """Test-only implementation of the same audit contract."""

    def __init__(self):
        self.events: list[dict[str, Any]] = []

    def record(self, event: dict[str, Any]) -> None:
        self.events.append(sanitize_for_log(dict(event)))
