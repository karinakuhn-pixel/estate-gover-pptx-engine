from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any
from pathlib import Path
from pptx import Presentation
import shutil
import uuid
import unicodedata
import re
import zipfile
import xml.etree.ElementTree as ET


app = FastAPI(
    title="Estate Gover Presentation Generator",
    version="2.4.0"
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# MODELOS DE DADOS
# ============================================================

class PresentationRequest(BaseModel):
    """
    V1 — mantida apenas para rollback seguro em /v1.
    Não deve ser usada como rota principal.
    """

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


class MediaItem(BaseModel):
    """
    Item de mídia V2.

    Regra:
    - usar somente mídia validada do ativo atual;
    - nunca aproveitar mídia de outro ativo;
    - mídia só é aplicada em shapes marcados como EG_SLOT_* ou EG_ASSET_*.
    """

    slot: str
    tipo: Optional[str] = None
    path: Optional[str] = None
    codigo_ativo: Optional[str] = None
    legenda: Optional[str] = None


class PresentationRequestV2(BaseModel):
    """
    V2 — substitui placeholders e aplica política de mídia asset_only.
    """

    codigo_ativo: str
    nome_ativo: str
    localizacao: str
    tipo_ativo: Optional[str] = None
    area_aproximada: Optional[str] = None
    valor_referencia: Optional[str] = None

    descricao_capa: Optional[str] = None
    resumo_executivo: Optional[str] = None
    conheca_ativo: Optional[str] = None
    localizacao_texto: Optional[str] = None
    dimensao_status: Optional[str] = None
    potencial_urbanistico: Optional[str] = None
    diferenciais: Optional[str] = None
    escala_ativo: Optional[str] = None
    condicoes_negocio: Optional[str] = None
    observacoes_urbanisticas: Optional[str] = None

    # Campos legados da V1, mantidos por compatibilidade com a Action.
    texto_base_capa: Optional[str] = None
    texto_base_estrategica: Optional[str] = None

    # Campos V2.
    midias: List[MediaItem] = Field(default_factory=list)
    media_policy: str = "asset_only"
    preservar_multimidia: bool = True
    formato_saida: str = "pptx"


PLACEHOLDER_FIELDS = [
    "codigo_ativo",
    "nome_ativo",
    "localizacao",
    "tipo_ativo",
    "area_aproximada",
    "valor_referencia",
    "descricao_capa",
    "resumo_executivo",
    "conheca_ativo",
    "localizacao_texto",
    "dimensao_status",
    "potencial_urbanistico",
    "diferenciais",
    "escala_ativo",
    "condicoes_negocio",
    "observacoes_urbanisticas",
    "texto_base_capa",
    "texto_base_estrategica",
]

FIXED_PREFIXES = (
    "EG_FIXED_",
    "EG_FIXED_LOGO",
    "EG_FIXED_QR",
    "EG_FIXED_WHATSAPP",
    "EG_FIXED_CONTATO",
)

ASSET_SLOT_PREFIXES = (
    "EG_SLOT_",
    "EG_ASSET_",
)

WHATSAPP_URL_RE = re.compile(
    r"https?://(?:api\.whatsapp\.com/send\?phone=|wa\.me/)[^\s<>\"]+",
    flags=re.IGNORECASE,
)

WHATSAPP_LABEL = "Abrir WhatsApp"

IMAGE_CTA_PATTERNS = (
    "aperte na imagem",
    "clique na imagem",
    "descubra o que te espera",
)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFD", value or "")
    without_accents = normalized.encode("ascii", "ignore").decode("utf-8")
    return without_accents.lower().strip()


def normalize_filename(value: str) -> str:
    return normalize_text(value).replace(" ", "-")


def find_template(kind: str) -> Path:
    """
    Localiza os PPTX oficiais dentro da pasta templates.

    Regra V2:
    - usar somente nomes Estate Gover;
    - não aceitar mais "estate-governador" como fallback silencioso.
    """

    if kind == "capa":
        candidates = [
            "estate-gover-capa-modelo-final.pptx",
        ]
        keywords = ["estate-gover", "capa"]
    elif kind == "estrategica":
        candidates = [
            "estate-gover-estrategica-modelo-final.pptx",
            "estate-gover-estratégica-modelo-final.pptx",
        ]
        keywords = ["estate-gover", "estrateg"]
    else:
        raise ValueError("Tipo de modelo inválido.")

    for filename in candidates:
        path = TEMPLATES_DIR / filename
        if path.exists():
            return path

    for path in TEMPLATES_DIR.glob("*.pptx"):
        normalized = normalize_filename(path.name)
        if "governador" in normalized:
            continue
        if all(keyword in normalized for keyword in keywords):
            return path

    raise HTTPException(
        status_code=500,
        detail=f"Modelo oficial {kind} não encontrado na pasta templates com nomenclatura Estate Gover."
    )


def safe_value(value: Any, fallback: str = "[pendente de confirmação]") -> str:
    if value is None:
        return fallback

    text = str(value).strip()
    if not text:
        return fallback

    return text


def payload_to_dict(payload: BaseModel) -> Dict[str, Any]:
    """
    Compatível com Pydantic v1 e v2.
    """

    if hasattr(payload, "model_dump"):
        return payload.model_dump()
    return payload.dict()


def build_placeholder_mapping(payload: PresentationRequestV2) -> Dict[str, str]:
    data = payload_to_dict(payload)
    mapping = {}

    for field in PLACEHOLDER_FIELDS:
        mapping[f"{{{{{field}}}}}"] = safe_value(data.get(field))

    return mapping


def replace_placeholders_in_presentation(prs: Presentation, mapping: Dict[str, str]) -> int:
    """
    Substitui placeholders dentro de caixas de texto existentes.

    Não cria slides, não altera master, não muda layout, não remove QR,
    WhatsApp, hyperlinks, logos, rodapés ou objetos interativos.
    """

    replacements = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue

            if shape.text_frame is None:
                continue

            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    original = run.text

                    for placeholder, value in mapping.items():
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)

                    if run.text != original:
                        replacements += 1

    return replacements


def is_fixed_shape(shape) -> bool:
    name = getattr(shape, "name", "") or ""
    upper_name = name.upper()
    return any(upper_name.startswith(prefix.upper()) for prefix in FIXED_PREFIXES)


def get_asset_slot_name(shape) -> Optional[str]:
    name = getattr(shape, "name", "") or ""

    for prefix in ASSET_SLOT_PREFIXES:
        if name.upper().startswith(prefix.upper()):
            return name[len(prefix):]

    return None


def remove_shape(shape) -> None:
    element = shape._element
    element.getparent().remove(element)


def clear_text_shape(shape, message: str = "") -> bool:
    if not hasattr(shape, "text_frame") or shape.text_frame is None:
        return False

    shape.text_frame.clear()
    if message:
        shape.text = message

    return True


def apply_media_policy(
    prs: Presentation,
    payload: PresentationRequestV2,
    warnings: List[str]
) -> Dict[str, int]:
    """
    Política de mídia V2.4 — modo seguro para integridade do PowerPoint.

    Regras:
    - EG_FIXED_* deve ser preservado.
    - EG_SLOT_* ou EG_ASSET_* representa mídia do ativo.
    - Não remover shapes do PPTX. Remoção direta pode deixar relações internas órfãs
      e fazer o PowerPoint pedir reparo.
    - Slide 01 é protegido como capa padrão: sem mídia nova válida, o slot visual original
      do modelo permanece intacto, sem ser movido e sem receber aviso de mídia ausente.
    - Se houver mídia válida, ela é aplicada por cima do espaço existente.
    - Se não houver mídia válida, nenhum slide recebe o texto "[mídia do ativo não fornecida]".
      O slot original do modelo é preservado para manter o padrão visual da matriz aprovada.
    """

    stats = {
        "fixed_preserved": 0,
        "asset_slots_found": 0,
        "asset_slots_replaced": 0,
        "asset_slots_neutralized": 0,
        "untagged_media_untouched": 0,
        "cover_slots_preserved_as_template": 0,
        "asset_slots_preserved_as_template": 0,
        "non_destructive_mode": 1,
    }

    midias_por_slot = {}

    for item in payload.midias:
        if item.codigo_ativo and item.codigo_ativo != payload.codigo_ativo:
            warnings.append(
                f"Mídia ignorada no slot {item.slot}: código do ativo diferente."
            )
            continue

        midias_por_slot[item.slot] = item

    def neutralize_visual_slot(slide, shape, message: str = "[mídia do ativo não fornecida]") -> bool:
        """
        Neutraliza slot sem remover shape:
        - se for caixa de texto, troca pelo aviso;
        - se for imagem/vídeo/forma sem texto, move para fora do slide e cria aviso no espaço original.
        """
        if clear_text_shape(shape, message):
            return True

        try:
            left = shape.left
            top = shape.top
            width = shape.width
            height = shape.height

            # Move o conteúdo antigo para fora da área visível sem quebrar suas relações internas.
            shape.left = prs.slide_width + 914400
            shape.top = prs.slide_height + 914400

            textbox = slide.shapes.add_textbox(left, top, width, height)
            textbox.text = message
            return True
        except Exception:
            # Último fallback: não remove nada para evitar corromper o pacote.
            return False

    for slide_index, slide in enumerate(prs.slides, start=1):
        for shape in list(slide.shapes):
            if is_fixed_shape(shape):
                stats["fixed_preserved"] += 1
                continue

            slot = get_asset_slot_name(shape)

            if not slot:
                continue

            stats["asset_slots_found"] += 1
            media = midias_por_slot.get(slot)

            if media and media.path:
                media_path = Path(media.path)

                if not media_path.exists():
                    warnings.append(
                        f"Mídia do slot {slot} não encontrada no caminho informado: {media.path}"
                    )

                    # V2.4: não inserir mensagem visual nem mover shapes.
                    # Preserva o slot original da matriz para manter o padrão visual.
                    stats["asset_slots_preserved_as_template"] += 1
                    if slide_index == 1:
                        stats["cover_slots_preserved_as_template"] += 1

                    continue

                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height

                # Não remove o shape original; aplica a mídia validada por cima.
                slide.shapes.add_picture(str(media_path), left, top, width=width, height=height)
                stats["asset_slots_replaced"] += 1
            else:
                # V2.4: sem mídia do ativo, preservar o slot original da matriz aprovada.
                # Não move shapes, não remove objetos e não escreve "[mídia do ativo não fornecida]".
                stats["asset_slots_preserved_as_template"] += 1
                if slide_index == 1:
                    stats["cover_slots_preserved_as_template"] += 1
                continue

    if stats["asset_slots_found"] == 0:
        warnings.append(
            "Nenhum slot de mídia EG_SLOT_* ou EG_ASSET_* foi encontrado. "
            "A V2 não removeu mídia visual sem marcação para evitar quebrar QR, WhatsApp ou hyperlinks."
        )

    return stats


def sanitize_visible_whatsapp_links(prs: Presentation) -> int:
    """
    Substitui URLs visíveis de WhatsApp por texto comercial limpo,
    preservando o hyperlink sempre que possível.

    V2.2:
    - trata URLs quebradas em múltiplos runs;
    - evita deixar a URL crua visível nos slides de contato.
    """

    replacements = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape.text_frame is None:
                continue

            for paragraph in shape.text_frame.paragraphs:
                runs = list(paragraph.runs)

                if runs:
                    paragraph_text = "".join((run.text or "") for run in runs)
                    match = WHATSAPP_URL_RE.search(paragraph_text)

                    if not match:
                        continue

                    url = match.group(0)

                    # Mantém somente um texto visual limpo no primeiro run.
                    runs[0].text = WHATSAPP_LABEL
                    for run in runs[1:]:
                        run.text = ""

                    try:
                        runs[0].hyperlink.address = url
                    except Exception:
                        pass

                    replacements += 1
                    continue

                current_text = getattr(shape, "text", "") or ""
                match = WHATSAPP_URL_RE.search(current_text)

                if match:
                    url = match.group(0)
                    shape.text = WHATSAPP_LABEL

                    try:
                        first_paragraph = shape.text_frame.paragraphs[0]
                        first_run = first_paragraph.runs[0]
                        first_run.hyperlink.address = url
                    except Exception:
                        pass

                    replacements += 1

    return replacements


def neutralize_image_ctas_without_media(prs: Presentation, media_stats: Dict[str, int]) -> int:
    """
    Remove chamadas como "Aperte na imagem" quando nenhum slot de mídia do ativo foi preenchido.
    Isso evita prometer interação visual quando a imagem foi neutralizada por falta de mídia válida.
    """

    if media_stats.get("asset_slots_replaced", 0) > 0:
        return 0

    replacements = 0

    for slide in prs.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame") or shape.text_frame is None:
                continue

            original_text = getattr(shape, "text", "") or ""
            normalized = normalize_text(original_text)

            if not normalized:
                continue

            if any(pattern in normalized for pattern in IMAGE_CTA_PATTERNS):
                clear_text_shape(shape, "")
                replacements += 1

    return replacements


def update_core_metadata(
    prs: Presentation,
    payload: PresentationRequestV2,
    presentation_kind: str
) -> bool:
    """
    Atualiza metadados internos do PPTX para evitar herança de dados do ativo-base
    do modelo, como "EG0013 Canela".
    """

    try:
        props = prs.core_properties

        kind_label = "CAPA" if presentation_kind == "capa" else "ESTRATÉGICA"

        props.title = f"Estate Gover {kind_label} - {payload.codigo_ativo}"
        props.subject = f"{payload.nome_ativo} | {payload.localizacao}"
        props.author = "Estate Gover"
        props.last_modified_by = "Estate Gover PPTX Engine"
        props.category = "Estate Gover Presentation"
        props.keywords = f"Estate Gover, {payload.codigo_ativo}, {kind_label}"
        props.comments = (
            "Arquivo gerado pelo Estate Gover PPTX Engine a partir dos modelos oficiais. "
            "CAPA e ESTRATÉGICA devem permanecer separadas."
        )

        return True
    except Exception:
        return False


def prune_unused_slide_relationships(path: Path, warnings: List[str]) -> int:
    """
    Remove relações órfãs de slides depois das alterações.
    Isso reduz risco de o PowerPoint pedir reparo quando algum hyperlink/mídia antigo
    deixa de ser referenciado pelo XML do slide.
    """

    pruned = 0

    try:
        tmp_path = path.with_suffix(".tmp.pptx")

        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        ET.register_namespace("", rel_ns)

        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            names = zin.namelist()

            for name in names:
                data = zin.read(name)

                if name.startswith("ppt/slides/_rels/slide") and name.endswith(".xml.rels"):
                    slide_name = name.replace("ppt/slides/_rels/", "ppt/slides/").replace(".rels", "")

                    if slide_name in names:
                        slide_root = ET.fromstring(zin.read(slide_name))
                        used_ids = set()

                        for elem in slide_root.iter():
                            for attr_name, attr_value in elem.attrib.items():
                                if attr_name.startswith("{" + r_ns + "}") and attr_value:
                                    used_ids.add(attr_value)

                        rel_root = ET.fromstring(data)
                        removed_any = False

                        for rel in list(rel_root):
                            rid = rel.attrib.get("Id")
                            rel_type = rel.attrib.get("Type", "")

                            is_prunable = (
                                rel_type.endswith("/image")
                                or rel_type.endswith("/video")
                                or rel_type.endswith("/media")
                                or rel_type.endswith("/hyperlink")
                            )

                            if is_prunable and rid and rid not in used_ids:
                                rel_root.remove(rel)
                                pruned += 1
                                removed_any = True

                        if removed_any:
                            data = ET.tostring(rel_root, encoding="utf-8", xml_declaration=True)

                zout.writestr(name, data)

        tmp_path.replace(path)
        return pruned
    except Exception as exc:
        warnings.append(f"Falha ao limpar relações órfãs do PPTX {path.name}: {exc}")
        return pruned


def validate_pptx_package(path: Path, warnings: List[str]) -> bool:
    """
    Validação técnica básica do pacote PPTX.
    Não substitui abertura no PowerPoint, mas ajuda a detectar ZIP corrompido.
    """

    try:
        with zipfile.ZipFile(path) as zf:
            bad_file = zf.testzip()

        if bad_file:
            warnings.append(f"PPTX com arquivo interno possivelmente corrompido: {bad_file}")
            return False

        return True
    except Exception as exc:
        warnings.append(f"Falha ao validar pacote PPTX {path.name}: {exc}")
        return False


def gerar_texto_base_v1(payload: PresentationRequest) -> str:
    return f"""ESTATE GOVER - TEXTO BASE

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
"""


def gerar_texto_base_v2(payload: PresentationRequestV2) -> str:
    return f"""ESTATE GOVER — TEXTO BASE V2

Código: {payload.codigo_ativo}
Ativo: {payload.nome_ativo}
Localização: {payload.localizacao}
Tipo: {payload.tipo_ativo or "[pendente de confirmação]"}
Área aproximada: {payload.area_aproximada or "[pendente de confirmação]"}
Valor de referência: {payload.valor_referencia or "[pendente de confirmação]"}

=== CAPA ===
{payload.descricao_capa or payload.texto_base_capa or "[pendente de confirmação]"}

=== ESTRATÉGICA ===

Resumo executivo:
{payload.resumo_executivo or payload.texto_base_estrategica or "[pendente de confirmação]"}

Conheça o ativo:
{payload.conheca_ativo or "[pendente de confirmação]"}

Localização:
{payload.localizacao_texto or "[pendente de confirmação]"}

Dimensão e status:
{payload.dimensao_status or "[pendente de confirmação]"}

Potencial construtivo e urbanístico:
{payload.potencial_urbanistico or "[necessário DM / consulta urbanística]"}

Diferenciais:
{payload.diferenciais or "[pendente de confirmação]"}

Escala do ativo:
{payload.escala_ativo or "[pendente de confirmação]"}

Condições de negócio:
{payload.condicoes_negocio or "[pendente de confirmação]"}

Observações urbanísticas:
{payload.observacoes_urbanisticas or "[necessário DM / consulta urbanística]"}
"""


def process_pptx_v2(
    template_path: Path,
    output_path: Path,
    payload: PresentationRequestV2,
    warnings: List[str],
    presentation_kind: str,
) -> Dict[str, Any]:
    """
    Duplicar matriz oficial, substituir placeholders, aplicar política de mídia V2
    e corrigir metadados/textos auxiliares sem alterar a identidade visual.
    """

    shutil.copyfile(template_path, output_path)

    prs = Presentation(str(output_path))

    metadata_updated = update_core_metadata(prs, payload, presentation_kind)

    mapping = build_placeholder_mapping(payload)
    replacements = replace_placeholders_in_presentation(prs, mapping)

    media_stats = apply_media_policy(prs, payload, warnings)

    whatsapp_links_sanitized = sanitize_visible_whatsapp_links(prs)
    image_ctas_neutralized = neutralize_image_ctas_without_media(prs, media_stats)

    prs.save(str(output_path))

    unused_relationships_pruned = prune_unused_slide_relationships(output_path, warnings)
    pptx_package_validated = validate_pptx_package(output_path, warnings)

    if replacements == 0:
        warnings.append(
            f"Nenhum placeholder textual foi substituído em {output_path.name}. "
            "Confirme se o PPTX possui placeholders como {{codigo_ativo}}, {{localizacao}} e {{valor_referencia}}."
        )

    if not metadata_updated:
        warnings.append(
            f"Metadados internos não puderam ser atualizados em {output_path.name}."
        )

    return {
        "template": template_path.name,
        "output": output_path.name,
        "text_replacements": replacements,
        "media_stats": media_stats,
        "metadata_updated": metadata_updated,
        "whatsapp_links_sanitized": whatsapp_links_sanitized,
        "image_ctas_neutralized": image_ctas_neutralized,
        "unused_relationships_pruned": unused_relationships_pruned,
        "pptx_package_validated": pptx_package_validated,
    }


# ============================================================
# ROTAS
# ============================================================

@app.get("/")
def healthcheck():
    return {
        "status": "ok",
        "service": "Estate Gover Presentation Generator",
        "version": "2.2.0",
        "routes": [
            "/gerar-apresentacao-estate-gover",
            "/v1/gerar-apresentacao-estate-gover",
            "/v2/gerar-apresentacao-estate-gover",
        ],
        "default_route": "/gerar-apresentacao-estate-gover now uses V2",
    }


@app.get("/outputs/{filename}")
def baixar_arquivo(filename: str):
    file_path = OUTPUTS_DIR / filename

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado.")

    return FileResponse(file_path)


@app.post("/v1/gerar-apresentacao-estate-gover")
def gerar_apresentacao_estate_gover_v1(payload: PresentationRequest):
    """
    V1 preservada para rollback:
    - duplica os modelos oficiais;
    - gera texto-base;
    - não substitui placeholders dentro do PPTX.
    """

    if payload.formato_saida != "pptx":
        raise HTTPException(status_code=400, detail="Apenas formato pptx é permitido.")

    capa_template = find_template("capa")
    estrategica_template = find_template("estrategica")

    job_id = f"{payload.codigo_ativo}_{uuid.uuid4().hex[:8]}"

    capa_out = OUTPUTS_DIR / f"{job_id}_CAPA.pptx"
    estrategica_out = OUTPUTS_DIR / f"{job_id}_ESTRATEGICA.pptx"

    # Versão inicial segura: duplica os modelos oficiais.
    shutil.copyfile(capa_template, capa_out)
    shutil.copyfile(estrategica_template, estrategica_out)

    txt_out = OUTPUTS_DIR / f"{job_id}_texto_base.txt"
    txt_out.write_text(gerar_texto_base_v1(payload), encoding="utf-8")

    return JSONResponse({
        "status": "ok",
        "versao": "v1",
        "mensagem": "Arquivos gerados a partir dos modelos oficiais.",
        "capa_pptx_url": f"/outputs/{capa_out.name}",
        "estrategica_pptx_url": f"/outputs/{estrategica_out.name}",
        "texto_base_url": f"/outputs/{txt_out.name}"
    })


@app.post("/gerar-apresentacao-estate-gover")
@app.post("/v2/gerar-apresentacao-estate-gover")
def gerar_apresentacao_estate_gover_v2(payload: PresentationRequestV2):
    """
    V2 — rota principal:
    - CAPA antes da ESTRATÉGICA;
    - duplica os modelos oficiais;
    - substitui placeholders existentes;
    - preserva elementos fixos;
    - limpa metadados herdados do modelo;
    - troca links visíveis de WhatsApp por texto comercial;
    - neutraliza chamadas de imagem quando não há mídia do ativo;
    - não remove shapes do PPTX; neutraliza mídia com política não destrutiva.
    """

    if payload.formato_saida != "pptx":
        raise HTTPException(
            status_code=400,
            detail="Apenas formato pptx é permitido."
        )

    warnings = []

    capa_template = find_template("capa")
    estrategica_template = find_template("estrategica")

    job_id = f"{payload.codigo_ativo}_{uuid.uuid4().hex[:8]}"

    capa_out = OUTPUTS_DIR / f"{job_id}_CAPA_V2.pptx"
    estrategica_out = OUTPUTS_DIR / f"{job_id}_ESTRATEGICA_V2.pptx"
    txt_out = OUTPUTS_DIR / f"{job_id}_texto_base_v2.txt"

    # Ordem obrigatória: CAPA antes da ESTRATÉGICA.
    capa_result = process_pptx_v2(
        capa_template,
        capa_out,
        payload,
        warnings,
        presentation_kind="capa",
    )

    estrategica_result = process_pptx_v2(
        estrategica_template,
        estrategica_out,
        payload,
        warnings,
        presentation_kind="estrategica",
    )

    txt_out.write_text(gerar_texto_base_v2(payload), encoding="utf-8")

    return JSONResponse({
        "status": "ok",
        "versao": "v2.4",
        "mensagem": "Arquivos V2 gerados a partir dos modelos oficiais.",
        "codigo_ativo": payload.codigo_ativo,
        "media_policy": payload.media_policy,
        "arquivos": {
            "capa_pptx_url": f"/outputs/{capa_out.name}",
            "estrategica_pptx_url": f"/outputs/{estrategica_out.name}",
            "texto_base_url": f"/outputs/{txt_out.name}"
        },
        "resultados": {
            "capa": capa_result,
            "estrategica": estrategica_result
        },
        "warnings": warnings
    })
