"""Fachada de aplicacion del organizador de PDFs.

Orquesta el recorrido de carpetas, la evaluacion de cada PDF, la copia con
nombre saneado y la unificacion por subcarpeta. El acceso a disco queda
limitado a las rutas de entrada y salida explicitas.
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Mapping, Sequence

from core.config import Settings
from core.errors import ConfigError
from core.logging import AuditLogger, NullAuditLogger
from modules.organizador.domain import (
    BuscadorPdfs,
    CategoriaPdf,
    CasoUso,
    EvaluacionPdf,
    LectorPdf,
    LectorPdfPyMuPDF,
    PdfEvaluador,
    ReglasCasos,
    RenombradorArchivos,
    OrganizacionResult,
    UnificadorPdfs,
)


class OrganizadorService:
    """Caso de uso de organizacion de PDFs, sin dependencias de FastAPI."""

    def __init__(
        self,
        reglas: ReglasCasos,
        evaluadores: Mapping[CasoUso, PdfEvaluador],
        renombrador: RenombradorArchivos,
        unificador: UnificadorPdfs,
        lector: LectorPdf,
        buscador: BuscadorPdfs | None = None,
        logger: AuditLogger | None = None,
    ) -> None:
        self._reglas = reglas
        self._evaluadores = dict(evaluadores)
        self._renombrador = renombrador
        self._unificador = unificador
        self._lector = lector
        self._buscador = buscador or BuscadorPdfs()
        self._logger = logger or NullAuditLogger()

    @classmethod
    def crear(
        cls, settings: Settings | None = None, logger: AuditLogger | None = None
    ) -> "OrganizadorService":
        reglas = ReglasCasos.desde_settings(settings or Settings())
        evaluadores = {
            caso: PdfEvaluador(reglas.obtener(caso)) for caso in reglas.casos()
        }
        return cls(
            reglas=reglas,
            evaluadores=evaluadores,
            renombrador=RenombradorArchivos(),
            unificador=UnificadorPdfs(),
            lector=LectorPdfPyMuPDF(),
            logger=logger,
        )

    def organizar(
        self,
        base: Path,
        caso: CasoUso,
        destino_normal: Path,
        destino_deuda: Path,
    ) -> OrganizacionResult:
        regla = self._reglas.obtener(caso)
        evaluador = self._evaluador(caso)
        if not base.is_dir():
            raise ConfigError(f"La carpeta base no existe o no es un directorio: {base}")

        log: list[str] = []
        normales: list[Path] = []
        deudas: list[Path] = []
        unificados: list[Path] = []

        for subcarpeta in self._buscador.subcarpetas(base):
            rutas = self._buscador.pdfs(subcarpeta)
            log.extend(self._cabecera(subcarpeta, len(rutas)))
            if not rutas:
                log.append("  [VACÍO] No se encontraron archivos .pdf.\n")
                continue

            copiados_normal: list[Path] = []
            copiados_deuda: list[Path] = []
            for ruta in rutas:
                evaluacion = self._evaluar(evaluador, ruta)
                if not evaluacion.valido:
                    self._logger.error(f"[RECHAZADO] {ruta.name}: {evaluacion.detalle}")
                    log.append(f"  [RECHAZADO] {ruta.name}")
                    log.append(f"              Ruta: {ruta}")
                    log.append(f"              Razón: {evaluacion.detalle}\n")
                    continue

                destino = self._copiar(
                    ruta,
                    (destino_deuda if evaluacion.categoria is CategoriaPdf.DEUDA else destino_normal)
                    / subcarpeta.name,
                    self._renombrador.nombre_base(regla, evaluacion, subcarpeta, ruta),
                )
                self._logger.info(
                    f"[ACEPTADO - {evaluacion.categoria.value}] {ruta.name}"
                )
                log.append(f"  [ACEPTADO - {evaluacion.categoria.value}] {ruta.name}")
                log.append(f"             Ruta: {ruta}")
                log.append(f"             Detalle: {evaluacion.detalle}\n")

                if evaluacion.categoria is CategoriaPdf.DEUDA:
                    copiados_deuda.append(destino)
                    deudas.append(destino)
                else:
                    copiados_normal.append(destino)
                    normales.append(destino)

            if regla.unificar_pdf:
                if copiados_normal:
                    unificado = self._unificar(
                        copiados_normal,
                        destino_normal / subcarpeta.name,
                        f"{subcarpeta.name}_UNIFICADO",
                    )
                    unificados.append(unificado)
                    log.append(f"  [MERGE NORMAL] Generado -> {unificado.name}")
                    self._logger.info(f"[MERGE NORMAL] {unificado}")
                if copiados_deuda:
                    unificado = self._unificar(
                        copiados_deuda,
                        destino_deuda / subcarpeta.name,
                        f"{subcarpeta.name}_DEUDAS_UNIFICADO",
                    )
                    unificados.append(unificado)
                    log.append(f"  [MERGE DEUDA] Generado -> {unificado.name}")
                    self._logger.info(f"[MERGE DEUDA] {unificado}")

        return OrganizacionResult(
            log=tuple(log),
            normales=tuple(normales),
            deudas=tuple(deudas),
            unificados=tuple(unificados),
        )

    def _evaluador(self, caso: CasoUso) -> PdfEvaluador:
        try:
            return self._evaluadores[caso]
        except KeyError as exc:
            raise ConfigError(f"No hay evaluador para el caso {caso.value}") from exc

    def _evaluar(self, evaluador: PdfEvaluador, ruta: Path) -> EvaluacionPdf:
        try:
            pdf = self._lector.leer(ruta)
        except Exception as exc:
            return EvaluacionPdf(
                origen=ruta,
                valido=False,
                categoria=CategoriaPdf.NINGUNA,
                conceptos=(),
                detalle=f"Error de lectura: {exc}",
            )
        return evaluador.evaluar(pdf)

    def _copiar(self, origen: Path, sub_destino: Path, nombre_base: str) -> Path:
        sub_destino.mkdir(parents=True, exist_ok=True)
        nombre_sano = self._renombrador.sanitizar(nombre_base)
        destino = sub_destino / f"{nombre_sano}.pdf"
        contador = 1
        while destino.exists():
            contador += 1
            destino = sub_destino / f"{nombre_sano}_{contador}.pdf"
        shutil.copy(origen, destino)
        return destino

    def _unificar(self, rutas: Sequence[Path], sub_destino: Path, nombre_base: str) -> Path:
        sub_destino.mkdir(parents=True, exist_ok=True)
        nombre = self._renombrador.sanitizar(nombre_base, max_len=60)
        return self._unificador.unificar(rutas, sub_destino / f"{nombre}.pdf")

    def _cabecera(self, subcarpeta: Path, cantidad: int) -> list[str]:
        return [
            "\n=========================================",
            f"SUBCARPETA EVALUADA: [{subcarpeta.name}] ({cantidad} PDFs detectados)",
            f"Ruta completa: {subcarpeta}",
            "=========================================",
        ]