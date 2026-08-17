"""Declared future capabilities; this module performs no external integration."""

from enum import Enum


class CapabilityStatus(str, Enum):
    ESPECIFICADO = "ESPECIFICADO"
    PENDENTE_DE_VALIDACAO_OFICIAL = "PENDENTE_DE_VALIDACAO_OFICIAL"


WHATSAPP_COMMUNITIES_GROUPS = (
    CapabilityStatus.ESPECIFICADO,
    CapabilityStatus.PENDENTE_DE_VALIDACAO_OFICIAL,
)

