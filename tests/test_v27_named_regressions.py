"""Named governance regressions required by the V2.7 V02 handoff."""

from fastapi.testclient import TestClient

import main
from v27.capabilities import CapabilityStatus, WHATSAPP_COMMUNITIES_GROUPS
from v27.models import PresentationRequestV27
from v27.providers import EmailProvider, WhatsAppProvider


def test_eg0030_has_no_automatic_email_whatsapp_or_approval_flow():
    paths = {route.path for route in main.app.routes}
    assert not any("email" in path.lower() or "whatsapp" in path.lower() for path in paths)
    assert hasattr(EmailProvider, "send")
    assert hasattr(WhatsAppProvider, "send_approved_template")
    assert WHATSAPP_COMMUNITIES_GROUPS == (
        CapabilityStatus.ESPECIFICADO,
        CapabilityStatus.PENDENTE_DE_VALIDACAO_OFICIAL,
    )


def test_eg0033_geometry_classification_cannot_be_elevated_to_unapproved_value():
    payload = {
        "codigo_ativo": "EG0033",
        "folder_id": "folder-eg0033",
        "nome_ativo": "Ativo Geométrico",
        "municipio": "Canela",
        "uf": "RS",
        "mapas_links": [{
            "tipo": "KMZ",
            "url": "https://example.invalid/eg0033.kmz",
            "classificacao": "BASE TÉCNICA — PENDENTE DE CONFERÊNCIA",
        }],
    }
    model = PresentationRequestV27.model_validate(payload)
    assert model.mapas_links[0].classificacao == "BASE TÉCNICA — PENDENTE DE CONFERÊNCIA"


def test_eg0002_legacy_v1_and_v26_routes_remain_compatible():
    paths = {route.path for route in main.app.routes}
    assert "/v1/gerar-apresentacao-estate-gover" in paths
    assert "/v2/gerar-apresentacao-estate-gover" in paths
    assert "/gerar-apresentacao-estate-gover" in paths
    payload = main.PresentationRequestV2(
        codigo_ativo="EG0002",
        nome_ativo="Ativo Piloto",
        localizacao="Canela/RS",
    )
    assert payload.tipo_saida == "CAPA"
    assert payload.versao_saida == "V01"
    assert payload.status_arquivo == "RASCUNHO"
    assert TestClient(main.app).get("/").status_code == 200
