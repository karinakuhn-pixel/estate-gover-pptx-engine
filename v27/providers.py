"""Provider boundaries and local homologation adapter for V2.7.

The local adapter is intentionally filesystem-only. It exercises folder_id
resolution, exact governed destinations, conflict blocking and draft upload
without claiming that Google Drive API credentials are configured.
"""

from __future__ import annotations

import hashlib
import shutil
from pathlib import Path
from typing import Any, Mapping, Protocol


class DriveProvider(Protocol):
    def validate_asset_folder(self, folder_id: str) -> None: ...
    def resolve_output_folder(self, folder_id: str, output_type: str) -> str: ...
    def file_exists(self, folder_id: str, filename: str) -> bool: ...
    def upload_draft(self, folder_id: str, path: Path) -> dict[str, Any]: ...


class ProviderError(RuntimeError):
    pass


class AssetFolderNotFound(ProviderError):
    pass


class OutputFolderNotFound(ProviderError):
    pass


class OutputConflict(ProviderError):
    pass


class LocalHomologationDriveProvider:
    """Local filesystem adapter mirroring governed Drive behavior.

    asset_roots maps official folder_id -> local asset root.
    No folders are created by this provider.
    """

    OUTPUT_DIRS = {
        "CAPA": "04_CAPA",
        "ESTRATEGICA": "05_ESTRATEGICA",
    }

    def __init__(self, asset_roots: Mapping[str, str | Path]):
        self._asset_roots = {
            str(folder_id): Path(root).resolve()
            for folder_id, root in asset_roots.items()
        }

    def _asset_root(self, folder_id: str) -> Path:
        root = self._asset_roots.get(folder_id)
        if root is None or not root.is_dir():
            raise AssetFolderNotFound(
                f"folder_id não autorizado ou inexistente: {folder_id}"
            )
        return root

    def validate_asset_folder(self, folder_id: str) -> None:
        self._asset_root(folder_id)

    def _output_path(self, folder_id: str, output_type: str) -> Path:
        root = self._asset_root(folder_id)
        dirname = self.OUTPUT_DIRS.get(output_type)
        if dirname is None:
            raise OutputFolderNotFound(
                f"tipo de saída sem destino governado: {output_type}"
            )
        destination = (root / dirname).resolve()
        if destination.parent != root or not destination.is_dir():
            raise OutputFolderNotFound(
                f"pasta governada ausente: {dirname}"
            )
        return destination

    def resolve_output_folder(self, folder_id: str, output_type: str) -> str:
        return str(self._output_path(folder_id, output_type))

    def file_exists(self, folder_id: str, filename: str) -> bool:
        destination = self._output_path(folder_id, "CAPA")
        if Path(filename).name != filename:
            raise ProviderError("nome de arquivo inválido")
        return (destination / filename).exists()

    def upload_draft(self, folder_id: str, path: Path) -> dict[str, Any]:
        source = Path(path).resolve()
        if not source.is_file():
            raise ProviderError("arquivo de origem inexistente")
        if not source.name.endswith("_RASCUNHO.pptx"):
            raise ProviderError("upload V2.7 deve permanecer RASCUNHO")

        destination_dir = self._output_path(folder_id, "CAPA")
        destination = destination_dir / source.name

        # Exclusividade: nunca sobrescrever.
        try:
            with destination.open("xb") as target, source.open("rb") as origin:
                shutil.copyfileobj(origin, target)
        except FileExistsError as exc:
            raise OutputConflict(
                f"arquivo já existe: {source.name}"
            ) from exc

        digest = hashlib.sha256(destination.read_bytes()).hexdigest()[:24]
        return {
            "file_id": f"local-homologation:{digest}",
            "name": destination.name,
            "webViewLink": destination.as_uri(),
            "parent_folder_id": str(destination_dir),
        }


class EmailProvider(Protocol):
    def send(self, message: dict[str, Any]) -> str: ...


class WhatsAppProvider(Protocol):
    def send_approved_template(self, message: dict[str, Any]) -> str: ...

class GoogleDriveProvider:
    """Google Drive v3 adapter with injected authenticated service.

    Credentials/tokens are external to this class. The adapter only operates
    on allowlisted asset folder_ids and exact governed child folders.
    """

    FOLDER_MIME = "application/vnd.google-apps.folder"
    OUTPUT_DIRS = {"CAPA": "04_CAPA", "ESTRATEGICA": "05_ESTRATEGICA"}

    def __init__(
        self,
        service: Any,
        *,
        allowed_asset_folder_ids: set[str] | None = None,
        media_upload_factory: Any | None = None,
    ):
        self._service = service
        self._allowed = set(allowed_asset_folder_ids or ())
        self._media_upload_factory = media_upload_factory
        self._resolved_outputs: dict[tuple[str, str], str] = {}

    def _assert_allowed(self, folder_id: str) -> None:
        # Fail closed: an absent or empty allowlist authorizes nothing.
        if not self._allowed or folder_id not in self._allowed:
            raise AssetFolderNotFound(f"folder_id fora da allowlist: {folder_id}")

    def validate_asset_folder(self, folder_id: str) -> None:
        self._assert_allowed(folder_id)
        try:
            meta = (
                self._service.files()
                .get(
                    fileId=folder_id,
                    fields="id,name,mimeType,trashed",
                    supportsAllDrives=True,
                )
                .execute()
            )
        except Exception as exc:
            raise AssetFolderNotFound(
                f"folder_id inexistente ou inacessível: {folder_id}"
            ) from exc
        if meta.get("trashed") or meta.get("mimeType") != self.FOLDER_MIME:
            raise AssetFolderNotFound(f"folder_id não é pasta ativa: {folder_id}")

    @staticmethod
    def _escape(value: str) -> str:
        return value.replace("\\", "\\\\").replace("'", "\\'")

    def _find_direct_child_folder(self, parent_id: str, child_name: str) -> str:
        q = (
            f"'{parent_id}' in parents and trashed = false and "
            f"mimeType = '{self.FOLDER_MIME}' and name = '{self._escape(child_name)}'"
        )
        result = (
            self._service.files()
            .list(
                q=q,
                spaces="drive",
                fields="files(id,name,mimeType,parents)",
                pageSize=10,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        files = result.get("files", [])
        if len(files) != 1:
            raise OutputFolderNotFound(
                f"destino governado '{child_name}' não encontrado de forma unívoca"
            )
        return files[0]["id"]

    def resolve_output_folder(self, folder_id: str, output_type: str) -> str:
        self.validate_asset_folder(folder_id)
        dirname = self.OUTPUT_DIRS.get(output_type)
        if dirname is None:
            raise OutputFolderNotFound(
                f"tipo de saída sem destino governado: {output_type}"
            )
        key = (folder_id, output_type)
        if key not in self._resolved_outputs:
            self._resolved_outputs[key] = self._find_direct_child_folder(
                folder_id, dirname
            )
        return self._resolved_outputs[key]

    def file_exists(self, folder_id: str, filename: str) -> bool:
        if Path(filename).name != filename:
            raise ProviderError("nome de arquivo inválido")
        output_folder_id = self.resolve_output_folder(folder_id, "CAPA")
        q = (
            f"'{output_folder_id}' in parents and trashed = false and "
            f"name = '{self._escape(filename)}'"
        )
        result = (
            self._service.files()
            .list(
                q=q,
                spaces="drive",
                fields="files(id,name)",
                pageSize=10,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )
        return bool(result.get("files", []))

    def _media_upload(self, path: Path):
        if self._media_upload_factory is not None:
            return self._media_upload_factory(
                str(path),
                mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                resumable=False,
            )
        try:
            from googleapiclient.http import MediaFileUpload
        except ImportError as exc:
            raise ProviderError("google-api-python-client não instalado no runtime") from exc
        return MediaFileUpload(
            str(path),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            resumable=False,
        )

    def upload_draft(self, folder_id: str, path: Path) -> dict[str, Any]:
        source = Path(path).resolve()
        if not source.is_file():
            raise ProviderError("arquivo de origem inexistente")
        if not source.name.endswith("_RASCUNHO.pptx"):
            raise ProviderError("upload V2.7 deve permanecer RASCUNHO")

        output_folder_id = self.resolve_output_folder(folder_id, "CAPA")
        if self.file_exists(folder_id, source.name):
            raise OutputConflict(f"arquivo já existe: {source.name}")

        body = {
            "name": source.name,
            "parents": [output_folder_id],
            "mimeType": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }
        try:
            created = (
                self._service.files()
                .create(
                    body=body,
                    media_body=self._media_upload(source),
                    fields="id,name,webViewLink,parents,mimeType",
                    supportsAllDrives=True,
                )
                .execute()
            )
        except Exception as exc:
            raise ProviderError(
                f"falha no upload governado: {type(exc).__name__}"
            ) from exc

        return {
            "file_id": created["id"],
            "name": created.get("name", source.name),
            "webViewLink": created.get("webViewLink"),
            "parent_folder_id": (created.get("parents") or [output_folder_id])[0],
        }
