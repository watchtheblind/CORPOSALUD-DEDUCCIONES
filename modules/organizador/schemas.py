"""Esquemas Pydantic estrictos del organizador."""

from __future__ import annotations

from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field

from modules.organizador.domain import (
    CasoUso,
    CategoriaPdf,
    EvaluacionPdf,
    OrganizacionResult,
)


class _EsquemaEstricto(BaseModel):
    model_config = ConfigDict(extra="forbid")


class OrganizacionRequestSchema(_EsquemaEstricto):
    base: Path
    caso: CasoUso
    destino_normal: Path
    destino_deuda: Path


class EvaluacionPdfSchema(_EsquemaEstricto):
    origen: Path
    valido: bool
    categoria: CategoriaPdf
    conceptos: list[str]
    detalle: str

    @classmethod
    def from_domain(cls, evaluacion: EvaluacionPdf) -> "EvaluacionPdfSchema":
        return cls(
            origen=evaluacion.origen,
            valido=evaluacion.valido,
            categoria=evaluacion.categoria,
            conceptos=list(evaluacion.conceptos),
            detalle=evaluacion.detalle,
        )


class OrganizacionResultSchema(_EsquemaEstricto):
    log: list[str]
    normales: list[str]
    deudas: list[str]
    unificados: list[str]
    total_normales: int = Field(ge=0)
    total_deudas: int = Field(ge=0)

    @classmethod
    def from_domain(cls, resultado: OrganizacionResult) -> "OrganizacionResultSchema":
        return cls(
            log=list(resultado.log),
            normales=[str(ruta) for ruta in resultado.normales],
            deudas=[str(ruta) for ruta in resultado.deudas],
            unificados=[str(ruta) for ruta in resultado.unificados],
            total_normales=len(resultado.normales),
            total_deudas=len(resultado.deudas),
        )
