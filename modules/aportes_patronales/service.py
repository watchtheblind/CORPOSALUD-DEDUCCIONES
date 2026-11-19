"""Fachada de aplicacion de aportes patronales.

Solo orquesta: lee la configuracion inyectada, delega la extraccion en el
dominio, delega la construccion del Excel y traduce fallos a errores de core.
"""

from __future__ import annotations

from typing import Sequence

from core.config import ConfigProvider, Settings
from core.documents import PdfDocument
from core.errors import ConfigError, DomainError, PdfExtractionError
from core.logging import AuditLogger, NullAuditLogger
from core.money import CommaDotMoneyParser, MoneyParser
from modules.aportes_patronales.domain import (
    AportesExcelBuilder,
    AportesPdfExtractor,
    AportesResult,
    ReportBuilder,
)


class AportesService:
    """Caso de uso de aportes patronales, sin dependencias de FastAPI."""

    def __init__(
        self,
        extractor: AportesPdfExtractor,
        builder: ReportBuilder,
        logger: AuditLogger | None = None,
    ) -> None:
        self._extractor = extractor
        self._builder = builder
        self._logger = logger or NullAuditLogger()

    @classmethod
    def crear(
        cls, money: MoneyParser | None = None, logger: AuditLogger | None = None
    ) -> "AportesService":
        parser = money or CommaDotMoneyParser()
        return cls(
            extractor=AportesPdfExtractor(money=parser),
            builder=AportesExcelBuilder(money=parser),
            logger=logger,
        )

    def procesar(
        self, documentos: Sequence[PdfDocument], config: ConfigProvider
    ) -> AportesResult:
        settings = self._leer_settings(config)
        filas = []
        for pdf in documentos:
            try:
                extraidas = self._extractor.extract(pdf, settings.etiquetas_pdf)
            except DomainError:
                raise
            except Exception as exc:
                raise PdfExtractionError(
                    f"No se pudo extraer el PDF {pdf.nombre_archivo}: {exc}"
                ) from exc
            filas.extend(extraidas)
            self._logger.info(f"{pdf.nombre_archivo}: {len(extraidas)} registros")

        if not filas:
            raise PdfExtractionError(
                "No se extrajeron registros de los PDFs de aportes recibidos"
            )

        documento = self._builder.build(filas)
        self._logger.info(f"Consolidado de aportes generado con {len(filas)} registros")
        return AportesResult(filas=tuple(filas), documento=documento, registros=len(filas))

    def _leer_settings(self, config: ConfigProvider) -> Settings:
        try:
            settings = config.settings()
        except DomainError:
            raise
        except Exception as exc:
            raise ConfigError("No se pudo leer la configuracion de aportes") from exc
        if settings is None:
            raise ConfigError("La configuracion de aportes no puede ser nula")
        return settings
