"""Safe V2.7 dependency bootstrap isolated from V1/V2.6."""

from __future__ import annotations

import os
from dataclasses import dataclass

from .audit import AuditLogger, StructuredAuditLogger
from .config import V27Config
from .google_auth import build_google_drive_service
from .idempotency import PilotSerialGuard
from .providers import DriveProvider, GoogleDriveProvider
from .renderer import CapaRendererV27


@dataclass(frozen=True)
class V27Runtime:
    config: V27Config
    audit_logger: AuditLogger
    drive_provider: DriveProvider | None
    renderer: CapaRendererV27 | None
    serial_guard: PilotSerialGuard


def build_v27_runtime(config: V27Config | None = None) -> V27Runtime:
    config = config or V27Config.from_env()
    audit = StructuredAuditLogger()
    provider = None
    renderer = None
    if config.drive_enabled:
        service = build_google_drive_service(config)
        provider = GoogleDriveProvider(
            service,
            allowed_asset_folder_ids=set(config.allowed_asset_folder_ids),
        )
        renderer = CapaRendererV27()
    return V27Runtime(config, audit, provider, renderer, PilotSerialGuard())


def build_v27_runtime_safely() -> V27Runtime:
    """Keep V1/V2.6 available if V2.7 configuration/auth bootstrap fails."""
    try:
        return build_v27_runtime()
    except Exception as exc:
        audit = StructuredAuditLogger()
        audit.record({
            "event": "v27_bootstrap_failed",
            "result": "erro",
            "error_type": type(exc).__name__,
            "drive_enabled": False,
        })
        fallback_env = dict(os.environ)
        fallback_env["GOOGLE_DRIVE_ENABLED"] = "false"
        config = V27Config.from_env(fallback_env)
        return V27Runtime(config, audit, None, None, PilotSerialGuard())
