"""Modulo de recibos: extraccion, consolidado y reportes."""

from modules.recibos.domain import (
    ENCABEZADO_CONSOLIDADO,
    AuditoriaEntry,
    CatalogoCentros,
    CatalogoConceptos,
    Centro,
    ConceptoCode,
    ConsolidadoBuilder,
    NomenclaturaHojas,
    ReciboDocument,
    ReciboParser,
    ReciboParseResult,
    ReciboResumen,
    RecibosAuditoria,
    ReporteConsolidadoBuilder,
    ResumidorRecibos,
)
from modules.recibos.schemas import (
    AuditoriaEntrySchema,
    ReciboDocumentSchema,
    ReciboResumenSchema,
    RecibosExtraccionSchema,
)
from modules.recibos.service import RecibosService

__all__ = [
    "ENCABEZADO_CONSOLIDADO",
    "AuditoriaEntry",
    "AuditoriaEntrySchema",
    "CatalogoCentros",
    "CatalogoConceptos",
    "Centro",
    "ConceptoCode",
    "ConsolidadoBuilder",
    "NomenclaturaHojas",
    "ReciboDocument",
    "ReciboDocumentSchema",
    "ReciboParseResult",
    "ReciboParser",
    "ReciboResumen",
    "ReciboResumenSchema",
    "RecibosAuditoria",
    "RecibosExtraccionSchema",
    "RecibosService",
    "ReporteConsolidadoBuilder",
    "ResumidorRecibos",
]
