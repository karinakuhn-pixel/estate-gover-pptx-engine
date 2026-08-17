from fastapi.testclient import TestClient

import main


def test_v1_v26_and_default_routes_remain_registered():
    paths = {route.path for route in main.app.routes}
    assert "/v1/gerar-apresentacao-estate-gover" in paths
    assert "/v2/gerar-apresentacao-estate-gover" in paths
    assert "/gerar-apresentacao-estate-gover" in paths
    assert "/v2.7/gerar-apresentacao-estate-gover" in paths


def test_v26_model_defaults_remain_unchanged():
    payload = main.PresentationRequestV2(
        codigo_ativo="EG0002",
        nome_ativo="Ativo Piloto",
        localizacao="Canela/RS",
    )
    assert payload.tipo_saida == "CAPA"
    assert payload.versao_saida == "V01"
    assert payload.status_arquivo == "RASCUNHO"


def test_healthcheck_keeps_v26_default_and_reports_v27_phase1():
    response = TestClient(main.app).get("/")
    assert response.status_code == 200
    body = response.json()
    assert body["version"] == "2.6.0"
    assert body["default_route"].endswith("uses V2.6")
    assert body["v2_7"] == {
        "estado": "FASE_1_IMPLEMENTADA",
        "geracao_disponivel": False,
        "drive_disponivel": False,
        "publicacao_automatica": False,
    }

