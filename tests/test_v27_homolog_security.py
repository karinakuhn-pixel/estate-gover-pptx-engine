import threading

import pytest

from v27.audit import InMemoryAuditLogger
from v27.config import ConfigurationError, V27Config
from v27.google_auth import GOOGLE_DRIVE_OAUTH_SCOPES
from v27.idempotency import OperationBusy, PilotSerialGuard
from v27.providers import AssetFolderNotFound, GoogleDriveProvider


OAUTH_ENV = {
    "ESTATE_GOVER_ENV": "homologacao",
    "GOOGLE_DRIVE_ENABLED": "true",
    "GOOGLE_DRIVE_AUTH_MODE": "oauth",
    "GOOGLE_DRIVE_ALLOWED_ASSET_FOLDER_IDS": "asset-1",
    "PUBLIC_BASE_URL": "https://homolog.example",
    "GOOGLE_OAUTH_CLIENT_ID": "client-id",
    "GOOGLE_OAUTH_CLIENT_SECRET": "client-secret",
    "GOOGLE_OAUTH_REFRESH_TOKEN": "refresh-token",
}


def test_drive_is_disabled_by_default():
    config = V27Config.from_env({})
    assert config.drive_enabled is False
    assert config.allowed_asset_folder_ids == frozenset()


def test_enabled_drive_requires_nonempty_allowlist():
    with pytest.raises(ConfigurationError, match="allowlist"):
        V27Config.from_env({**OAUTH_ENV, "GOOGLE_DRIVE_ALLOWED_ASSET_FOLDER_IDS": ""})


def test_enabled_drive_requires_complete_oauth_secrets():
    incomplete = dict(OAUTH_ENV)
    incomplete.pop("GOOGLE_OAUTH_REFRESH_TOKEN")
    with pytest.raises(ConfigurationError, match="OAuth"):
        V27Config.from_env(incomplete)


def test_service_account_mode_is_not_available_in_homologation():
    with pytest.raises(ConfigurationError, match="oauth"):
        V27Config.from_env({**OAUTH_ENV, "GOOGLE_DRIVE_AUTH_MODE": "service_account"})


def test_oauth_scope_supports_direct_preexisting_folder_access():
    assert GOOGLE_DRIVE_OAUTH_SCOPES == (
        "https://www.googleapis.com/auth/drive",
    )


def test_google_provider_allowlist_is_fail_closed_when_empty():
    provider = GoogleDriveProvider(service=object(), allowed_asset_folder_ids=set())
    with pytest.raises(AssetFolderNotFound, match="allowlist"):
        provider.validate_asset_folder("asset-1")


def test_audit_recursively_redacts_oauth_secrets_and_bearer_tokens():
    audit = InMemoryAuditLogger()
    audit.record({
        "event": "auth_failed",
        "client_secret": "must-not-leak",
        "nested": {"refresh_token": "must-not-leak"},
        "message": "Authorization: Bearer abc.def.ghi",
    })
    event = audit.events[0]
    assert event["client_secret"] == "[REDACTED]"
    assert event["nested"]["refresh_token"] == "[REDACTED]"
    assert "abc.def.ghi" not in event["message"]


def test_pilot_serial_guard_blocks_a_concurrent_operation():
    guard = PilotSerialGuard()
    entered = threading.Event()
    release = threading.Event()

    def hold_claim():
        with guard.claim("first"):
            entered.set()
            release.wait(timeout=2)

    worker = threading.Thread(target=hold_claim)
    worker.start()
    assert entered.wait(timeout=2)
    try:
        with pytest.raises(OperationBusy):
            with guard.claim("second"):
                pass
    finally:
        release.set()
        worker.join(timeout=2)


def test_pilot_serial_guard_releases_after_failure():
    guard = PilotSerialGuard()
    with pytest.raises(RuntimeError):
        with guard.claim("first"):
            raise RuntimeError("controlled")
    with guard.claim("second"):
        pass
