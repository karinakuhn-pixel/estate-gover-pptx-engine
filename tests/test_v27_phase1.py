from pathlib import Path

import pytest
from fastapi import FastAPI, HTTPException
from fastapi.testclient import TestClient
from pydantic import ValidationError

from v27.audit import InMemoryAuditLogger
from v27.capabilities import (
    CapabilityStatus,
    WHATSAPP_COMMUNITIES_GROUPS,
)
from v27.middleware import V27OperationIdMiddleware
from v27.models import PresentationRequestV27
from v27.routes import build_v27_router
from v27.security import resolve_safe_output_path


VALID_PAYLOAD = {
    "codigo_ativo": "EG0029",
    "folder_id": "1G9Ft-kriDZS4BYZBVmEsyWVq9XY5grxE",
    "nome_ativo": "AREA URBANA CARNIEL 5,2HA",
    "municipio": "Gramado",
    "uf": "RS",
    "tipo_saida": "CAPA",
    "versao_saida": "V01",
    "formato_saida": "pptx",
}


@pytest.fixture
def phase1_app():
    audit = InMemoryAuditLogger()
    app = FastAPI()
    app.add_middleware(V27OperationIdMiddleware, audit_logger=audit)
    app.include_router(build_v27_router(audit))
    return TestClient(app), audit


def test_folder_id_is_required():
    payload = {key: value for key, value in VALID_PAYLOAD.items() if key != "folder_id"}
    with pytest.raises(ValidationError):
        PresentationRequestV27.model_validate(payload)


def test_blank_folder_id_is_rejected():
    with pytest.raises(ValidationError):
        PresentationRequestV27.model_validate({**VALID_PAYLOAD, "folder_id": "   "})


@pytest.mark.parametrize(
    ("supplied", "expected"),
    [(1, "V01"), ("01", "V01"), ("V1", "V01"), ("V01", "V01"), (12, "V12")],
)
def test_version_is_normalized_in_pydantic(supplied, expected):
    model = PresentationRequestV27.model_validate(
        {**VALID_PAYLOAD, "versao_saida": supplied}
    )
    assert model.versao_saida == expected


@pytest.mark.parametrize("supplied", [0, "V00", -1, "V-1", "FINAL", "1A", True])
def test_invalid_version_is_rejected_in_pydantic(supplied):
    with pytest.raises(ValidationError):
        PresentationRequestV27.model_validate(
            {**VALID_PAYLOAD, "versao_saida": supplied}
        )


@pytest.mark.parametrize(
    "forbidden",
    [{"status_arquivo": "APROVADO"}, {"publicacao_automatica": True}],
)
def test_backend_controlled_fields_are_forbidden(forbidden):
    with pytest.raises(ValidationError) as exc_info:
        PresentationRequestV27.model_validate({**VALID_PAYLOAD, **forbidden})
    assert "extra_forbidden" in str(exc_info.value)


def test_phase1_route_is_isolated_and_does_not_generate(phase1_app):
    client, audit = phase1_app
    response = client.post(
        "/v2.7/gerar-apresentacao-estate-gover",
        json=VALID_PAYLOAD,
    )

    assert response.status_code == 501
    body = response.json()
    assert body["status"] == "nao_implementado"
    assert body["status_arquivo"] == "RASCUNHO"
    assert body["publicacao_automatica"] is False
    assert body["arquivos"] == []
    assert body["operation_id"] == response.headers["X-Operation-ID"]
    assert any(event["event"] == "governed_request_validated" for event in audit.events)


def test_invalid_request_also_receives_operation_id(phase1_app):
    client, audit = phase1_app
    response = client.post(
        "/v2.7/gerar-apresentacao-estate-gover",
        json={**VALID_PAYLOAD, "status_arquivo": "APROVADO"},
    )

    assert response.status_code == 422
    assert response.headers["X-Operation-ID"]
    assert audit.events[0]["event"] == "request_started"
    assert audit.events[-1]["event"] == "request_completed"
    assert audit.events[-1]["result"] == "erro"


def test_audit_events_never_change_governed_status(phase1_app):
    client, audit = phase1_app
    client.post("/v2.7/gerar-apresentacao-estate-gover", json=VALID_PAYLOAD)
    governed_events = [event for event in audit.events if "status_arquivo" in event]
    assert governed_events
    assert all(event["status_arquivo"] == "RASCUNHO" for event in governed_events)
    assert all(event["publicacao_automatica"] is False for event in governed_events)


def test_safe_output_path_accepts_direct_child(tmp_path):
    assert resolve_safe_output_path(tmp_path, "arquivo.pptx") == (
        tmp_path / "arquivo.pptx"
    ).resolve()


@pytest.mark.parametrize(
    "filename",
    ["../arquivo", "../../etc/passwd", "subdir/arquivo.pptx", "", "."],
)
def test_safe_output_path_blocks_traversal(tmp_path, filename):
    with pytest.raises(HTTPException) as exc_info:
        resolve_safe_output_path(tmp_path, filename)
    assert exc_info.value.status_code == 400
    assert str(tmp_path) not in str(exc_info.value.detail)


def test_symlink_escaping_outputs_is_blocked(tmp_path):
    outside = tmp_path.parent / "outside-v27.txt"
    outside.write_text("secret", encoding="utf-8")
    link = tmp_path / "link.pptx"
    link.symlink_to(outside)

    with pytest.raises(HTTPException):
        resolve_safe_output_path(tmp_path, link.name)


def test_whatsapp_communities_capability_is_not_claimed_as_available():
    assert WHATSAPP_COMMUNITIES_GROUPS == (
        CapabilityStatus.ESPECIFICADO,
        CapabilityStatus.PENDENTE_DE_VALIDACAO_OFICIAL,
    )

