"""Fail-closed environment configuration for V2.7 homologation."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Mapping


class ConfigurationError(RuntimeError):
    pass


def _enabled(value: str | None) -> bool:
    return (value or "false").strip().lower() in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class V27Config:
    environment: str
    drive_enabled: bool
    drive_auth_mode: str
    allowed_asset_folder_ids: frozenset[str]
    public_base_url: str
    oauth_client_id: str | None
    oauth_client_secret: str | None
    oauth_refresh_token: str | None

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> "V27Config":
        source = os.environ if env is None else env
        allowed = frozenset(
            item.strip()
            for item in source.get("GOOGLE_DRIVE_ALLOWED_ASSET_FOLDER_IDS", "").split(",")
            if item.strip()
        )
        config = cls(
            environment=source.get("ESTATE_GOVER_ENV", "local").strip() or "local",
            drive_enabled=_enabled(source.get("GOOGLE_DRIVE_ENABLED")),
            drive_auth_mode=source.get("GOOGLE_DRIVE_AUTH_MODE", "oauth").strip().lower(),
            allowed_asset_folder_ids=allowed,
            public_base_url=source.get(
                "PUBLIC_BASE_URL", "https://estate-gover-pptx-engine.onrender.com"
            ).rstrip("/"),
            oauth_client_id=source.get("GOOGLE_OAUTH_CLIENT_ID"),
            oauth_client_secret=source.get("GOOGLE_OAUTH_CLIENT_SECRET"),
            oauth_refresh_token=source.get("GOOGLE_OAUTH_REFRESH_TOKEN"),
        )
        config.validate()
        return config

    def validate(self) -> None:
        if not self.drive_enabled:
            return
        if self.drive_auth_mode != "oauth":
            raise ConfigurationError("V2.7 HOMOLOG suporta somente GOOGLE_DRIVE_AUTH_MODE=oauth")
        if not self.allowed_asset_folder_ids:
            raise ConfigurationError("allowlist do Google Drive é obrigatória quando habilitado")
        if not all((self.oauth_client_id, self.oauth_client_secret, self.oauth_refresh_token)):
            raise ConfigurationError("credenciais OAuth de homologação estão incompletas")
