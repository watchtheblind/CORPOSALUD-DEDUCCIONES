"""Esquemas Pydantic estrictos de casas comerciales."""

from __future__ import annotations

from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field

from modules.casas_comerciales.domain import (
    CasasComercialesRow,
    Centro,
    Concepto,
)


class _EsquemaEstricto(BaseModel):
    model_config = ConfigDict(extra="forbid")


class CasasComercialesRowSchema(_EsquemaEstricto):
    centro: str = Field(min_length=1, max_length=200)
    grupo: str
    referencia: str
    concepto: str = Field(min_length=1, max_length=200)
    cantidad: int = Field(ge=0)
    trabajador: Decimal
    empresa: Decimal

    @classmethod
    def from_domain(cls, fila: CasasComercialesRow) -> "CasasComercialesRowSchema":
        return cls(
            centro=fila.centro.nombre,
            grupo=fila.grupo,
            referencia=fila.referencia,
            concepto=fila.concepto.codigo,
            cantidad=fila.cantidad,
            trabajador=fila.trabajador,
            empresa=fila.empresa,
        )

    def to_domain(self) -> CasasComercialesRow:
        return CasasComercialesRow(
            centro=Centro(self.centro),
            grupo=self.grupo,
            referencia=self.referencia,
            concepto=Concepto(self.concepto),
            cantidad=self.cantidad,
            trabajador=self.trabajador,
            empresa=self.empresa,
        )


class CasasComercialesResultSchema(_EsquemaEstricto):
    documento: str = Field(min_length=1)
    registros: int = Field(ge=0)
    filas: list[CasasComercialesRowSchema]
