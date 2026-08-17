import pytest
from v27.providers import (
    AssetFolderNotFound, GoogleDriveProvider, OutputConflict, OutputFolderNotFound
)

ASSET_ID = "asset-folder-id"
CAPA_ID = "capa-folder-id"

class FakeRequest:
    def __init__(self, result=None, exc=None):
        self.result, self.exc = result, exc
    def execute(self):
        if self.exc:
            raise self.exc
        return self.result

class FakeFiles:
    def __init__(self):
        self.created = []
        self.existing_names = set()
        self.direct_child_mode = "ok"

    def get(self, **kwargs):
        if kwargs["fileId"] != ASSET_ID:
            return FakeRequest(exc=RuntimeError("404"))
        return FakeRequest({
            "id": ASSET_ID,
            "name": "EG0029",
            "mimeType": "application/vnd.google-apps.folder",
            "trashed": False,
        })

    def list(self, **kwargs):
        q = kwargs.get("q", "")
        if "mimeType = 'application/vnd.google-apps.folder'" in q:
            if self.direct_child_mode == "missing":
                return FakeRequest({"files": []})
            if self.direct_child_mode == "duplicate":
                return FakeRequest({"files": [
                    {"id": CAPA_ID, "name": "04_CAPA"},
                    {"id": "dup", "name": "04_CAPA"},
                ]})
            return FakeRequest({"files": [{
                "id": CAPA_ID, "name": "04_CAPA",
                "mimeType": "application/vnd.google-apps.folder",
                "parents": [ASSET_ID],
            }]})
        found = []
        for name in self.existing_names:
            if f"name = '{name}'" in q:
                found.append({"id": "existing", "name": name})
        return FakeRequest({"files": found})

    def create(self, **kwargs):
        body = kwargs["body"]
        assert body["parents"] == [CAPA_ID]
        assert "permissions" not in body
        self.created.append(body)
        return FakeRequest({
            "id": "new-drive-file-id",
            "name": body["name"],
            "webViewLink": "https://drive.google.com/file/d/new-drive-file-id/view",
            "parents": [CAPA_ID],
            "mimeType": body["mimeType"],
        })

class FakeDriveService:
    def __init__(self):
        self._files = FakeFiles()
    def files(self):
        return self._files

def fake_media_factory(path, mimetype=None, resumable=False):
    return {"path": path, "mimetype": mimetype, "resumable": resumable}

def make_provider():
    service = FakeDriveService()
    provider = GoogleDriveProvider(
        service,
        allowed_asset_folder_ids={ASSET_ID},
        media_upload_factory=fake_media_factory,
    )
    return provider, service

def test_validates_only_allowlisted_asset_folder():
    provider, _ = make_provider()
    provider.validate_asset_folder(ASSET_ID)
    with pytest.raises(AssetFolderNotFound):
        provider.validate_asset_folder("outro-folder")

def test_resolves_04_capa_as_direct_child_only():
    provider, _ = make_provider()
    assert provider.resolve_output_folder(ASSET_ID, "CAPA") == CAPA_ID

def test_missing_or_duplicate_output_folder_blocks():
    provider, service = make_provider()
    service._files.direct_child_mode = "missing"
    with pytest.raises(OutputFolderNotFound):
        provider.resolve_output_folder(ASSET_ID, "CAPA")
    provider, service = make_provider()
    service._files.direct_child_mode = "duplicate"
    with pytest.raises(OutputFolderNotFound):
        provider.resolve_output_folder(ASSET_ID, "CAPA")

def test_exact_filename_conflict_is_detected():
    provider, service = make_provider()
    service._files.existing_names.add("EG0029_TESTE_CAPA_V01_RASCUNHO.pptx")
    assert provider.file_exists(ASSET_ID, "EG0029_TESTE_CAPA_V01_RASCUNHO.pptx")

def test_upload_is_new_file_in_04_capa_without_permissions(tmp_path):
    provider, service = make_provider()
    path = tmp_path / "EG0029_TESTE_CAPA_V01_RASCUNHO.pptx"
    path.write_bytes(b"fake-pptx")
    result = provider.upload_draft(ASSET_ID, path)
    assert result["file_id"] == "new-drive-file-id"
    assert result["parent_folder_id"] == CAPA_ID
    assert service._files.created[0]["name"] == path.name
    assert "permissions" not in service._files.created[0]

def test_upload_conflict_blocks_before_create(tmp_path):
    provider, service = make_provider()
    name = "EG0029_TESTE_CAPA_V01_RASCUNHO.pptx"
    service._files.existing_names.add(name)
    path = tmp_path / name
    path.write_bytes(b"fake-pptx")
    with pytest.raises(OutputConflict):
        provider.upload_draft(ASSET_ID, path)
    assert service._files.created == []
