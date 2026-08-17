from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import main


BASE_PAYLOAD = {
    "codigo_ativo": "EG0002",
    "nome_ativo": "Ativo Piloto",
    "localizacao": "Canela/RS",
}


@pytest.fixture
def client_and_calls(tmp_path, monkeypatch):
    outputs = tmp_path / "outputs"
    outputs.mkdir()
    monkeypatch.setattr(main, "OUTPUTS_DIR", outputs)
    monkeypatch.setattr(
        main,
        "PUBLIC_BASE_URL",
        "https://estate-gover-pptx-engine.onrender.com",
    )

    calls = []

    def fake_find_template(kind):
        calls.append(("template", kind))
        template = tmp_path / f"{kind}.pptx"
        template.write_bytes(b"template")
        return template

    def fake_process(
        template_path,
        output_path,
        payload,
        warnings,
        presentation_kind,
    ):
        calls.append(("process", presentation_kind))
        output_path.write_bytes(
            f"pptx-{presentation_kind}".encode()
        )
        return {
            "output": output_path.name,
            "kind": presentation_kind,
            "url": main.output_absolute_url(output_path),
        }

    monkeypatch.setattr(
        main,
        "find_template",
        fake_find_template,
    )
    monkeypatch.setattr(
        main,
        "process_pptx_v2",
        fake_process,
    )

    return TestClient(main.app), calls, outputs


def test_default_generates_only_capa_with_governance_metadata(
    client_and_calls,
):
    client, calls, outputs = client_and_calls

    response = client.post(
        "/gerar-apresentacao-estate-gover",
        json=BASE_PAYLOAD,
    )

    assert response.status_code == 200

    body = response.json()

    assert body["versao"] == "v2.6"
    assert body["versao_saida"] == "V01"
    assert body["status_arquivo"] == "RASCUNHO"
    assert body["publicacao_automatica"] is False
    assert body["sobrescrita_silenciosa"] is False
    assert body["tipo_saida"] == "CAPA"
    assert body["produtos_gerados"] == ["CAPA"]

    assert set(body["arquivos"]) == {
        "capa_pptx",
        "texto_base",
    }

    expected_pptx = (
        "EG0002_ATIVO-PILOTO_CAPA_V01_RASCUNHO.pptx"
    )

    assert (
        body["arquivos"]["capa_pptx"]["arquivo"]
        == expected_pptx
    )

    assert (
        body["arquivos"]["capa_pptx"]["url_relativa"]
        == f"/outputs/{expected_pptx}"
    )

    assert (
        body["arquivos"]["capa_pptx"]["url"]
        == (
            "https://estate-gover-pptx-engine.onrender.com"
            f"/outputs/{expected_pptx}"
        )
    )

    assert calls == [
        ("template", "capa"),
        ("process", "capa"),
    ]

    assert (
        outputs / expected_pptx
    ).exists()


def test_estrategica_generates_only_estrategica(
    client_and_calls,
):
    client, calls, outputs = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "tipo_saida": "ESTRATEGICA",
        },
    )

    assert response.status_code == 200

    body = response.json()

    assert body["tipo_saida"] == "ESTRATEGICA"
    assert body["produtos_gerados"] == [
        "ESTRATEGICA"
    ]

    assert set(body["arquivos"]) == {
        "estrategica_pptx",
        "texto_base",
    }

    assert set(body["resultados"]) == {
        "estrategica"
    }

    expected_pptx = (
        "EG0002_ATIVO-PILOTO_"
        "ESTRATEGICA_V01_RASCUNHO.pptx"
    )

    assert (
        body["arquivos"]["estrategica_pptx"]["arquivo"]
        == expected_pptx
    )

    assert (
        outputs / expected_pptx
    ).exists()

    assert calls == [
        ("template", "estrategica"),
        ("process", "estrategica"),
    ]


def test_ambos_preserves_capa_before_estrategica(
    client_and_calls,
):
    client, calls, _ = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "tipo_saida": "AMBOS",
        },
    )

    assert response.status_code == 200

    body = response.json()

    assert body["produtos_gerados"] == [
        "CAPA",
        "ESTRATEGICA",
    ]

    assert calls == [
        ("template", "capa"),
        ("process", "capa"),
        ("template", "estrategica"),
        ("process", "estrategica"),
    ]


def test_custom_version_and_status_are_reflected_in_filename(
    client_and_calls,
):
    client, _, outputs = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "versao_saida": "V02",
            "status_arquivo": "APROVADO",
        },
    )

    assert response.status_code == 200

    body = response.json()

    assert body["versao_saida"] == "V02"
    assert body["status_arquivo"] == "APROVADO"

    expected_pptx = (
        "EG0002_ATIVO-PILOTO_CAPA_V02_APROVADO.pptx"
    )

    assert (
        body["arquivos"]["capa_pptx"]["arquivo"]
        == expected_pptx
    )

    assert (
        outputs / expected_pptx
    ).exists()


def test_existing_output_returns_409_without_processing(
    client_and_calls,
):
    client, calls, outputs = client_and_calls

    existing = (
        outputs
        / "EG0002_ATIVO-PILOTO_CAPA_V01_RASCUNHO.pptx"
    )

    existing.write_bytes(b"original")

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json=BASE_PAYLOAD,
    )

    assert response.status_code == 409
    assert existing.read_bytes() == b"original"
    assert calls == []


def test_invalid_tipo_saida_is_rejected(
    client_and_calls,
):
    client, _, _ = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "tipo_saida": "PDF",
        },
    )

    assert response.status_code == 422


def test_invalid_version_is_rejected(
    client_and_calls,
):
    client, _, _ = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "versao_saida": "FINAL",
        },
    )

    assert response.status_code == 400


def test_failure_removes_only_reserved_partial_outputs(
    client_and_calls,
    monkeypatch,
):
    client, _, outputs = client_and_calls

    def fail_process(*args, **kwargs):
        raise RuntimeError(
            "falha simulada"
        )

    monkeypatch.setattr(
        main,
        "process_pptx_v2",
        fail_process,
    )

    with pytest.raises(
        RuntimeError,
        match="falha simulada",
    ):
        client.post(
            "/v2/gerar-apresentacao-estate-gover",
            json=BASE_PAYLOAD,
        )

    assert list(outputs.iterdir()) == []


def test_text_base_respects_capa_only(
    client_and_calls,
):
    client, _, _ = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "descricao_capa": "Texto exclusivo da CAPA",
            "texto_base_estrategica": (
                "ISTO NÃO DEVE APARECER"
            ),
        },
    )

    assert response.status_code == 200

    body = response.json()

    text_name = (
        body["arquivos"]["texto_base"]["arquivo"]
    )

    text_path = (
        main.OUTPUTS_DIR / text_name
    )

    content = text_path.read_text(
        encoding="utf-8"
    )

    assert "=== CAPA ===" in content
    assert "Texto exclusivo da CAPA" in content
    assert "=== ESTRATÉGICA ===" not in content
    assert "ISTO NÃO DEVE APARECER" not in content


def test_text_base_respects_estrategica_only(
    client_and_calls,
):
    client, _, _ = client_and_calls

    response = client.post(
        "/v2/gerar-apresentacao-estate-gover",
        json={
            **BASE_PAYLOAD,
            "tipo_saida": "ESTRATEGICA",
            "resumo_executivo": (
                "Resumo exclusivo da estratégica"
            ),
            "texto_base_capa": (
                "ISTO NÃO DEVE APARECER"
            ),
        },
    )

    assert response.status_code == 200

    body = response.json()

    text_name = (
        body["arquivos"]["texto_base"]["arquivo"]
    )

    content = (
        main.OUTPUTS_DIR / text_name
    ).read_text(
        encoding="utf-8"
    )

    assert "=== ESTRATÉGICA ===" in content
    assert (
        "Resumo exclusivo da estratégica"
        in content
    )
    assert "=== CAPA ===" not in content
    assert "ISTO NÃO DEVE APARECER" not in content


def test_v1_route_remains_available_and_relative(
    tmp_path,
    monkeypatch,
):
    outputs = tmp_path / "outputs"
    outputs.mkdir()

    capa = tmp_path / "capa.pptx"
    estrategica = tmp_path / "estrategica.pptx"

    capa.write_bytes(b"capa")
    estrategica.write_bytes(b"estrategica")

    monkeypatch.setattr(
        main,
        "OUTPUTS_DIR",
        outputs,
    )

    monkeypatch.setattr(
        main,
        "find_template",
        lambda kind: (
            capa
            if kind == "capa"
            else estrategica
        ),
    )

    payload = {
        **BASE_PAYLOAD,
        "texto_base_capa": "Texto capa",
        "texto_base_estrategica": (
            "Texto estratégica"
        ),
    }

    response = TestClient(
        main.app
    ).post(
        "/v1/gerar-apresentacao-estate-gover",
        json=payload,
    )

    assert response.status_code == 200

    body = response.json()

    assert body["versao"] == "v1"
    assert body["capa_pptx_url"].startswith(
        "/outputs/"
    )
    assert body[
        "estrategica_pptx_url"
    ].startswith(
        "/outputs/"
    )


def test_healthcheck_reports_v26_governance():
    response = TestClient(
        main.app
    ).get("/")

    assert response.status_code == 200

    body = response.json()

    assert body["version"] == "2.6.0"
    assert (
        body["governanca"]["tipo_saida_default"]
        == "CAPA"
    )
    assert (
        body["governanca"]["versao_saida_default"]
        == "V01"
    )
    assert (
        body["governanca"]["status_arquivo_default"]
        == "RASCUNHO"
    )
    assert (
        body["governanca"]["publicacao_automatica"]
        is False
    )
    assert (
        body["governanca"]["sobrescrita_silenciosa"]
        is False
    )
