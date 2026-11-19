"""Fachada de aplicacion de casas comerciales.

Orquesta configuracion, extraccion y generacion del Excel de auditoria.
"""

from __future__ import annotations

from typing import Sequence

from core.config import ConfigProvider, Settings
from core.documents import ExcelDocument, PdfDocument
from core.errors import ConfigError, DomainError, PdfExtractionError
from core.logging import AuditLogger, NullAuditLogger
from core.money import CommaDotMoneyParser, MoneyParser
from modules.casas_comerciales.domain import (
    AuditoriaExcelBuilder,
    CasasPdfExtractor,
)


class CasasComercialesService:
    """Caso de uso de casas comerciales, sin dependencias de FastAPI."""

    def __init__(
        self,
        extractor: CasasPdfExtractor,
        builder: AuditoriaExcelBuilder,
        logger: AuditLogger | None = None,
    ) -> None:
        self._extractor = extractor
        self._builder = builder
        self._logger = logger or NullAuditLogger()

    @classmethod
    def crear(
        cls, money: MoneyParser | None = None, logger: AuditLogger | None = None
    ) -> "CasasComercialesService":
        parser = money or CommaDotMoneyParser()
        return cls(
            extractor=CasasPdfExtractor(money=parser),
            builder=AuditoriaExcelBuilder(money=parser),
            logger=logger,
        )

    def procesar(
        self, documentos: Sequence[PdfDocument], config: ConfigProvider
    ) -> ExcelDocument:
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
                "No se extrajeron registros de los PDFs de casas comerciales"
            )

        documento = self._builder.build(filas)
        self._logger.info(
            f"Auditoria de casas comerciales generada con {len(filas)} registros"
        )
        return documento

    def _leer_settings(self, config: ConfigProvider) -> Settings:
        try:
            settings = config.settings()
        except DomainError:
            raise
        except Exception as exc:
            raise ConfigError("No se pudo leer la configuracion de casas comerciales") from exc
        if settings is None:
            raise ConfigError("La configuracion de casas comerciales no puede ser nula")
        return settings
