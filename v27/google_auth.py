"""OAuth factory for the current institutional homologation account."""

from __future__ import annotations

from typing import Any

from .config import V27Config


# Access by direct pre-existing folder_id plus file creation requires Drive scope.
# Credentials remain external so a future identity migration changes configuration,
# not provider behavior.
GOOGLE_DRIVE_OAUTH_SCOPES = ("https://www.googleapis.com/auth/drive",)


def build_google_drive_service(config: V27Config) -> Any:
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError("dependências Google OAuth não instaladas") from exc

    credentials = Credentials(
        token=None,
        refresh_token=config.oauth_refresh_token,
        token_uri="https://oauth2.googleapis.com/token",
        client_id=config.oauth_client_id,
        client_secret=config.oauth_client_secret,
        scopes=list(GOOGLE_DRIVE_OAUTH_SCOPES),
    )
    return build("drive", "v3", credentials=credentials, cache_discovery=False)
