"""Pilot serial-execution gate with a persistent-store extension point."""

from __future__ import annotations

import threading
from contextlib import contextmanager
from typing import Iterator, Protocol


class OperationBusy(RuntimeError):
    pass


class IdempotencyStore(Protocol):
    def reserve(self, key: str, operation_id: str) -> bool: ...
    def complete(self, key: str, operation_id: str, result: str) -> None: ...


class PilotSerialGuard:
    """Single-process guard for the approved one-shot homologation pilot.

    It deliberately does not claim multi-instance persistence. IdempotencyStore
    is the boundary to replace before scaled operation.
    """

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._active: set[str] = set()

    @contextmanager
    def claim(self, key: str) -> Iterator[None]:
        with self._lock:
            if self._active:
                raise OperationBusy("piloto V2.7 já possui operação em andamento")
            self._active.add(key)
        try:
            yield
        finally:
            with self._lock:
                self._active.discard(key)
