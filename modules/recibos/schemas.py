"""Esquemas Pydantic estrictos de recibos."""

from __future__ import annotations

from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field

from modules.recibos.domain import (
    AuditoriaEntry,
    ReciboDocument,
    ReciboResumen,
)


class _EsquemaEstricto(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ReciboDocumentSchema(_EsquemaEstricto):
    concepto: str = Field(min_length=1, max_length=200)
    codigo: str = Field(min_length=1, max_length=50)
    entidad: str = Field(min_length=1, max_length=200)
    numero: str = Field(min_length=1)
    monto: Decimal

    @classmethod
    def from_domain(cls, documento: ReciboDocument) -> "ReciboDocumentSchema":
        return cls(
            concepto=documento.concepto.nombre,
            codigo=documento.concepto.codigo,
            entidad=documento.entidad.nombre,
            numero=documento.numero,
            monto=documento.monto,
        )


class ReciboResumenSchema(_EsquemaEstricto):
    archivo: str = Field(min_length=1)
    numero: str = Field(min_length=1)

    @classmethod
    def from_domain(cls, resumen: ReciboResumen) -> "ReciboResumenSchema":
        return cls(archivo=resumen.archivo, numero=resumen.numero)


class AuditoriaEntrySchema(_EsquemaEstricto):
    nivel: str = Field(min_length=1)
    mensaje: str = Field(min_length=1)

    @classmethod
    def from_domain(cls, entrada: AuditoriaEntry) -> "AuditoriaEntrySchema":
        return cls(nivel=entrada.nivel, mensaje=entrada.mensaje)


class RecibosExtraccionSchema(_EsquemaEstricto):
    documentos: list[ReciboDocumentSchema]
    auditoria: list[AuditoriaEntrySchema]
