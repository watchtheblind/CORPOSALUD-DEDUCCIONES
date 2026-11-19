"""Esquemas Pydantic estrictos de aportes patronales.

Todos los esquemas prohiben campos extra y declaran explicitamente el centro,
para que el repo destino pueda aplicar scoping por departamento.
"""

from __future__ import annotations

from decimal import Decimal

from pydantic import BaseModel, ConfigDict, Field

from modules.aportes_patronales.domain import (
    AporteRow,
    AportesResult,
    Centro,
    Concepto,
)


class _EsquemaEstricto(BaseModel):
    model_config = ConfigDict(extra="forbid")


class CentroSchema(_EsquemaEstricto):
    nombre: str = Field(min_length=1, max_length=200)


class ConceptoSchema(_EsquemaEstricto):
    codigo: str = Field(min_length=1, max_length=200)


class AporteRowSchema(_EsquemaEstricto):
    centro: str = Field(min_length=1, max_length=200)
    grupo: str
    referencia: str
    concepto: str = Field(min_length=1, max_length=200)
    cantidad: int = Field(ge=0)
    trabajador: Decimal
    empresa: Decimal

    @classmethod
    def from_domain(cls, fila: AporteRow) -> AporteRowSchema:
        return cls(
            centro=fila.centro.nombre,
            grupo=fila.grupo,
            referencia=fila.referencia,
            concepto=fila.concepto.codigo,
            cantidad=fila.cantidad,
            trabajador=fila.trabajador,
            empresa=fila.empresa,
        )

    def to_domain(self) -> AporteRow:
        return AporteRow(
            centro=Centro(self.centro),
            grupo=self.grupo,
            referencia=self.referencia,
            concepto=Concepto(self.concepto),
            cantidad=self.cantidad,
            trabajador=self.trabajador,
            empresa=self.empresa,
        )


class AportesResultSchema(_EsquemaEstricto):
    filas: list[AporteRowSchema]
    documento: str = Field(min_length=1)
    registros: int = Field(ge=0)

    @classmethod
    def from_domain(cls, resultado: AportesResult) -> AportesResultSchema:
        return cls(
            filas=[AporteRowSchema.from_domain(fila) for fila in resultado.filas],
            documento=resultado.documento.nombre,
            registros=resultado.registros,
        )
