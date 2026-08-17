"""Governed request models for the isolated V2.7 contract."""

import re
from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator


class GovernedModel(BaseModel):
    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class OutputTypeV27(str, Enum):
    CAPA = "CAPA"
    ESTRATEGICA = "ESTRATEGICA"
    AMBOS = "AMBOS"


class SlideSpecV27(GovernedModel):
    tipo: str = Field(min_length=1)
    titulo: str | None = None
    mensagem_principal: str | None = None
    texto_apoio: str | None = None
    dados: dict[str, Any] = Field(default_factory=dict)
    imagem: str | None = None
    mapa: str | None = None
    ressalva: str | None = None


class NegotiationSpecV27(GovernedModel):
    preco_pedido: str | None = None
    avaliacao: str | None = None
    vgv: str | None = None
    valor_residual: str | None = None
    condicao_comercial: str | None = None


class SpatialLinkSpecV27(GovernedModel):
    tipo: str = Field(min_length=1)
    url: str = Field(min_length=1)
    classificacao: Literal[
        "LOCALIZAÇÃO APROXIMADA",
        "DELIMITAÇÃO COMERCIAL PRELIMINAR",
        "BASE RECEBIDA — NÃO VALIDADA",
        "BASE TÉCNICA — PENDENTE DE CONFERÊNCIA",
        "VALIDADO TECNICAMENTE",
    ]


class ContactSpecV27(GovernedModel):
    nome: str = Field(min_length=1)
    selecionado: bool = False
    papel: str | None = None
    telefone: str | None = None
    whatsapp: str | None = None
    email: str | None = None
    url: str | None = None


class ImageSpecV27(GovernedModel):
    referencia: str = Field(min_length=1)
    classificacao: Literal["REAL", "ANALÍTICA", "CONCEITUAL"]
    autorizada: bool
    ressalva: str | None = None


class PresentationRequestV27(GovernedModel):
    codigo_ativo: str = Field(min_length=1)
    folder_id: str = Field(min_length=1)
    nome_ativo: str = Field(min_length=1)
    municipio: str = Field(min_length=1)
    uf: str = Field(min_length=2, max_length=2)
    tipo_saida: OutputTypeV27 = OutputTypeV27.CAPA
    versao_saida: str = "V01"
    tese_criativa: str | None = None
    slides: list[SlideSpecV27] = Field(default_factory=list)
    negociacao: NegotiationSpecV27 = Field(default_factory=NegotiationSpecV27)
    mapas_links: list[SpatialLinkSpecV27] = Field(default_factory=list)
    contatos: list[ContactSpecV27] = Field(default_factory=list)
    imagens_autorizadas: list[ImageSpecV27] = Field(default_factory=list)
    formato_saida: Literal["pptx"] = "pptx"

    @field_validator("folder_id")
    @classmethod
    def validate_folder_id(cls, value: str) -> str:
        value = value.strip()
        if not value:
            raise ValueError("folder_id é obrigatório e não pode ser vazio")
        return value

    @field_validator("uf")
    @classmethod
    def normalize_uf(cls, value: str) -> str:
        return value.upper()

    @field_validator("versao_saida", mode="before")
    @classmethod
    def normalize_version(cls, value: Any) -> str:
        if isinstance(value, bool):
            raise ValueError("versao_saida inválida")

        text = str(value).strip().upper()
        match = re.fullmatch(r"V?(\d+)", text)
        if not match:
            raise ValueError("versao_saida deve seguir o padrão V01, V02, V03...")

        number = int(match.group(1))
        if number < 1:
            raise ValueError("versao_saida deve ser igual ou superior a V01")

        return f"V{number:02d}"

