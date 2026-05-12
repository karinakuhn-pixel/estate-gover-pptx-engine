from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from typing import Optional
from pathlib import Path
import shutil
import uuid

app = FastAPI(
    title="Estate Gover Presentation Generator",
    version="1.0.0"
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

class PresentationRequest(BaseModel):
    codigo_ativo: str
    nome_ativo: str
    localizacao: str
    tipo_ativo: Optional[str] = None
    area_aproximada: Optional[str] = None
    valor_referencia: Optional[str] = None
    texto_base_capa: str
    texto_base_estrategica: str
    observacoes_urbanisticas: Optional[str] = None
    preservar_multimidia: bool = True
    formato_saida: str = "pptx"

@app.get("/")
def healthcheck():
    return {"status": "ok", "service": "Estate Gover Presentation Generator"}

@app.get("/outputs/{filename}")
def baixar_arquivo(filename: str):
    file_path = OUTPUTS_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")
    return FileResponse(file_path)

@app.post("/gerar-apresentacao-estate-gover")
def gerar_apresentacao_estate_gover(payload: PresentationRequest):
    if payload.formato_saida != "pptx":
        raise HTTPException(status_code=400, detail="Apenas formato pptx é permitido.")

    capa_template = TEMPLATES_DIR / "estate-gover-capa-modelo-final.pptx"
    estrategica_template = TEMPLATES_DIR / "estate-gover-estrategica-modelo-final.pptx"

    if not capa_template.exists() or not estrategica_template.exists():
        raise HTTPException(
            status_code=500,
            detail="Modelos oficiais PPTX não encontrados na pasta templates."
        )

    job_id = f"{payload.codigo_ativo}_{uuid.uuid4().hex[:8]}"
    capa_out = OUTPUTS_DIR / f"{job_id}_CAPA.pptx"
    estrategica_out = OUTPUTS_DIR / f"{job_id}_ESTRATEGICA.pptx"

    # Versão inicial segura:
    # duplica os modelos oficiais para preservar layout, vídeos, QR codes, WhatsApp e hyperlinks.
    shutil.copyfile(capa_template, capa_out)
    shutil.copyfile(estrategica_template, estrategica_out)

    txt_out = OUTPUTS_DIR / f"{job_id}_texto_base.txt"
    txt_out.write_text(
        f'''ESTATE GOVER - TEXTO BASE

Código: {payload.codigo_ativo}
Ativo: {payload.nome_ativo}
Localização: {payload.localizacao}
Tipo: {payload.tipo_ativo or ""}
Área aproximada: {payload.area_aproximada or ""}
Valor de referência: {payload.valor_referencia or ""}

=== CAPA ===
{payload.texto_base_capa}

=== ESTRATÉGICA ===
{payload.texto_base_estrategica}

=== OBSERVAÇÕES URBANÍSTICAS ===
{payload.observacoes_urbanisticas or ""}
''',
        encoding="utf-8"
    )

    return JSONResponse({
        "status": "ok",
        "mensagem": "Arquivos gerados a partir dos modelos oficiais.",
        "capa_pptx_url": f"/outputs/{capa_out.name}",
        "estrategica_pptx_url": f"/outputs/{estrategica_out.name}",
        "texto_base_url": f"/outputs/{txt_out.name}"
    })
