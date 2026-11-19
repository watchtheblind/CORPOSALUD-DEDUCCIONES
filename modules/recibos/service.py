"""Fachada de aplicacion de recibos.

Expone extraccion, generacion de reportes y resumen como operaciones sin
FastAPI. La auditoria se acumula en memoria (RecibosAuditoria) y tambien se
emite por el puerto AuditLogger: no se escriben archivos de log.
"""

from __future__ import annotations

from typing import Iterable, Sequence

from core.documents import ExcelDocument, PdfDocument
from core.errors import DomainError, ReportGenerationError
from core.logging import AuditLogger, NullAuditLogger
from core.money import LegacyReciboMoneyParser, MoneyParser
from modules.recibos.domain import (
    AuditoriaEntry,
    Centro,
    ConsolidadoBuilder,
    ReciboDocument,
    ReciboParser,
    ReciboResumen,
    RecibosAuditoria,
    ReporteConsolidadoBuilder,
    ResumidorRecibos,
)


class RecibosService:
    """Caso de uso de recibos de nomina y deuda, sin dependencias de FastAPI."""

    def __init__(
        self,
        parser: ReciboParser,
        consolidado: ConsolidadoBuilder,
        reportes: ReporteConsolidadoBuilder,
        resumidor: ResumidorRecibos,
        auditoria: RecibosAuditoria | None = None,
        logger: AuditLogger | None = None,
    ) -> None:
        self._parser = parser
        self._consolidado = consolidado
        self._reportes = reportes
        self._resumidor = resumidor
        self._auditoria = auditoria or RecibosAuditoria()
        self._logger = logger or NullAuditLogger()

    @classmethod
    def crear(
        cls,
        money: MoneyParser | None = None,
        logger: AuditLogger | None = None,
        auditoria: RecibosAuditoria | None = None,
    ) -> "RecibosService":
        parser_money = money or LegacyReciboMoneyParser()
        return cls(
            parser=ReciboParser(money=parser_money),
            consolidado=ConsolidadoBuilder(money=parser_money),
            reportes=ReporteConsolidadoBuilder(),
            resumidor=ResumidorRecibos(),
            auditoria=auditoria,
            logger=logger,
        )

    def extraer(
        self, documentos: Sequence[PdfDocument], es_deuda: bool
    ) -> list[ReciboDocument]:
        tipo = "DEUDAS" if es_deuda else "NOMINAS"
        self._registrar(
            f"--- INICIO ESCANEO CARPETA {tipo}: {len(documentos)} documentos ---"
        )
        extraidos: list[ReciboDocument] = []
        for pdf in documentos:
            try:
                resultado = self._parser.parse_documento(pdf, es_deuda)
            except DomainError as exc:
                self._registrar(
                    f"ERROR: No se pudo leer {pdf.nombre_archivo}. Detalle: {exc}",
                    "ERROR",
                )
                continue
            if resultado.documentos:
                extraidos.extend(resultado.documentos)
            else:
                self._registrar(
                    "AVISO: Recibo "
                    f"{resultado.numero} ({resultado.entidad.nombre}) "
                    "sin conceptos relevantes",
                    "AVISO",
                )
        return extraidos

    def generar_reportes(
        self, nominas: Sequence[ReciboDocument], deudas: Sequence[ReciboDocument]
    ) -> ExcelDocument:
        try:
            consolidado = self._consolidado.build(nominas, deudas)
            documento = self._reportes.build(consolidado, self._consolidado.nombre_hoja())
        except DomainError:
            raise
        except Exception as exc:
            raise ReportGenerationError(
                "No se pudieron generar los reportes de recibos"
            ) from exc
        self._registrar(
            f"PROCESO: reporte generado con {len(nominas)} nominas y {len(deudas)} deudas"
        )
        return documento

    def resumir(self, documentos: Sequence[PdfDocument]) -> list[ReciboResumen]:
        return self._resumidor.resumir(documentos)

    def auditar_cobertura(
        self, entidades: Iterable[Centro]
    ) -> tuple[AuditoriaEntry, ...]:
        vistas = {entidad.nombre for entidad in entidades}
        faltantes = [
            nombre for nombre in self._parser.nombres_entidades() if nombre not in vistas
        ]
        if faltantes:
            self._registrar("--- ENTIDADES NO ENCONTRADAS ---")
            for nombre in faltantes:
                self._registrar(f"AUSENTE: {nombre}", "AVISO")
        else:
            self._registrar("PROCESO: Todas las entidades fueron detectadas")
        return self._auditoria.entries()

    def auditoria(self) -> tuple[AuditoriaEntry, ...]:
        return self._auditoria.entries()

    def _registrar(self, mensaje: str, nivel: str = "INFO") -> None:
        self._auditoria.registrar(mensaje, nivel)
        if nivel == "ERROR":
            self._logger.error(mensaje)
        else:
            self._logger.info(mensaje)
