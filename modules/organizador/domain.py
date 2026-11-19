"""Dominio del organizador de PDFs (antes file_organiser.py).

Evalua si un PDF corresponde al caso de uso elegido, lo clasifica en NORMAL /
DEUDA y lo copia con un nombre saneado. La lectura de PDFs entra por el puerto
LectorPdf, de modo que el dominio se puede probar sin disco.
"""

from __future__ import annotations

import re
import unicodedata
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Mapping, Sequence

import pymupdf

from core.config import Settings
from core.documents import PdfDocument, PdfPage
from core.errors import ConfigError, PdfExtractionError

PALABRA_DEUDA = "DEUDA"
EXTENSION_PDF = ".pdf"
NOMBRE_ARCHIVO_POR_DEFECTO = "ARCHIVO_PROCESADO"
LONGITUD_NOMBRE_MAXIMA = 90

TEXTOS_RECIBOS = ("TOTAL DE DEDUCCIONES POR CONCEPTO",)
TEXTOS_APORTES = (
    "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA",
    "FONDO PENSIONES",
    "S.S.O.",
    "PERDIDA INVOLUNTARIA",
    "CAHORMINSAS",
    "CAEMINSA",
)


class TextoNormalizado:
    """Normalizacion de texto compartida por evaluador y renombrador."""

    @staticmethod
    def sin_acentos(texto: str) -> str:
        descompuesto = unicodedata.normalize("NFD", texto)
        return "".join(c for c in descompuesto if unicodedata.category(c) != "Mn")

    @classmethod
    def colapsar(cls, texto: str) -> str:
        """Quita guiones de corte, colapsa espacios y pasa a mayusculas."""
        sin_guiones = texto.replace("-\n", "").replace("- ", "")
        return cls.sin_acentos(" ".join(sin_guiones.split()).upper())


class CasoUso(Enum):
    """Casos de uso que el organizador sabe reconocer."""

    RECIBOS = "RECIBOS"
    APORTES_PATRONALES = "APORTES_PATRONALES"
    CASAS_COMERCIALES = "CASAS_COMERCIALES"


@dataclass(frozen=True)
class ReglaClasificacion:
    """Regla con la que se reconoce un caso de uso dentro de los PDFs."""

    caso: CasoUso
    textos_requeridos: tuple[str, ...]
    mapa_conceptos: Mapping[str, str] = field(default_factory=dict)
    unificar_pdf: bool = False
    incluir_concepto_en_nombre: bool = False
    evaluar_no_ley: bool = False


class ReglasCasos:
    """Catalogo de reglas por caso de uso."""

    def __init__(self, reglas: Sequence[ReglaClasificacion]) -> None:
        self._reglas = {regla.caso: regla for regla in reglas}

    @classmethod
    def por_defecto(cls, mapa_conceptos: Mapping[str, str] | None = None) -> "ReglasCasos":
        mapa = dict(mapa_conceptos or {})
        return cls(
            (
                ReglaClasificacion(caso=CasoUso.RECIBOS, textos_requeridos=TEXTOS_RECIBOS),
                ReglaClasificacion(
                    caso=CasoUso.APORTES_PATRONALES,
                    textos_requeridos=TEXTOS_APORTES,
                    unificar_pdf=True,
                    incluir_concepto_en_nombre=True,
                    evaluar_no_ley=True,
                ),
                ReglaClasificacion(
                    caso=CasoUso.CASAS_COMERCIALES,
                    textos_requeridos=tuple(mapa),
                    mapa_conceptos=mapa,
                    unificar_pdf=True,
                    incluir_concepto_en_nombre=True,
                    evaluar_no_ley=True,
                ),
            )
        )

    @classmethod
    def desde_settings(cls, settings: Settings) -> "ReglasCasos":
        mapa = {
            texto: concepto
            for texto, concepto in settings.mapa_conceptos.items()
            if texto.strip().lower() != "(en blanco)"
        }
        return cls.por_defecto(mapa)

    def obtener(self, caso: CasoUso) -> ReglaClasificacion:
        try:
            return self._reglas[caso]
        except KeyError as exc:
            raise ConfigError(f"No hay regla configurada para el caso {caso.value}") from exc

    def casos(self) -> tuple[CasoUso, ...]:
        return tuple(self._reglas)


class CategoriaPdf(Enum):
    """Resultado de clasificar un PDF."""

    NORMAL = "NORMAL"
    DEUDA = "DEUDA"
    NINGUNA = "NINGUNA"


@dataclass(frozen=True)
class EvaluacionPdf:
    """Resultado detallado de evaluar un PDF contra una regla."""

    origen: Path
    valido: bool
    categoria: CategoriaPdf
    conceptos: tuple[str, ...]
    detalle: str


class PdfEvaluador:
    """Decide si un PDF pertenece al caso de uso y si contiene deuda."""

    def __init__(self, regla: ReglaClasificacion) -> None:
        self._regla = regla
        self._mapa = self._construir_mapa()
        self._deuda = TextoNormalizado.colapsar(PALABRA_DEUDA)

    def _construir_mapa(self) -> dict[str, str]:
        if self._regla.mapa_conceptos:
            return {
                TextoNormalizado.colapsar(texto): concepto
                for texto, concepto in self._regla.mapa_conceptos.items()
                if texto.strip().lower() != "(en blanco)"
            }
        return {
            TextoNormalizado.colapsar(texto): texto
            for texto in self._regla.textos_requeridos
        }

    def evaluar(self, pdf: PdfDocument) -> EvaluacionPdf:
        origen = pdf.origen or Path(pdf.nombre)
        try:
            nombre_normalizado = TextoNormalizado.sin_acentos(pdf.nombre_archivo.upper())
            tiene_no_ley = self._regla.evaluar_no_ley and "NO LEY" in nombre_normalizado

            if not pdf.paginas:
                return EvaluacionPdf(
                    origen=origen,
                    valido=False,
                    categoria=CategoriaPdf.NINGUNA,
                    conceptos=(),
                    detalle="PDF sin páginas o corrupto.",
                )

            indices = [0]
            if len(pdf.paginas) > 1:
                indices.append(len(pdf.paginas) - 1)
            indices.extend(range(1, len(pdf.paginas) - 1))

            tiene_requerido = tiene_no_ley
            tiene_deuda = False
            hallados = ["NO LEY"] if tiene_no_ley else []

            for indice in indices:
                texto_raw = pdf.paginas[indice].texto or ""
                if not texto_raw.strip():
                    continue
                texto_limpio = TextoNormalizado.colapsar(texto_raw)
                for texto_normalizado, concepto in self._mapa.items():
                    if texto_normalizado in texto_limpio:
                        tiene_requerido = True
                        if concepto not in hallados:
                            hallados.append(concepto)
                if self._deuda in texto_limpio:
                    tiene_deuda = True
                if tiene_requerido and tiene_deuda:
                    break

            if not tiene_requerido:
                return EvaluacionPdf(
                    origen=origen,
                    valido=False,
                    categoria=CategoriaPdf.NINGUNA,
                    conceptos=(),
                    detalle=f"No coincide en sus {len(pdf.paginas)} páginas.",
                )

            categoria = CategoriaPdf.DEUDA if tiene_deuda else CategoriaPdf.NORMAL
            detalle = f"Concepto(s): [{', '.join(hallados) if hallados else 'N/A'}]"
            if tiene_deuda:
                detalle += f" Y '{PALABRA_DEUDA}'"
            return EvaluacionPdf(
                origen=origen,
                valido=True,
                categoria=categoria,
                conceptos=tuple(hallados),
                detalle=detalle,
            )
        except Exception as exc:  # pragma: no cover - defensivo
            return EvaluacionPdf(
                origen=origen,
                valido=False,
                categoria=CategoriaPdf.NINGUNA,
                conceptos=(),
                detalle=f"Error de lectura: {exc}",
            )


class RenombradorArchivos:
    """Define el nombre de cada copia y lo sanea para Windows."""

    _CARACTERES_INVALIDOS = re.compile(r'[\\/*?:"<>|]')

    def nombre_base(
        self,
        regla: "ReglaClasificacion",
        evaluacion: "EvaluacionPdf",
        subcarpeta: Path,
        ruta: Path,
    ) -> str:
        """Nombre sin extension segun la regla del caso y los conceptos hallados."""
        es_no_ley = regla.evaluar_no_ley and "NO LEY" in TextoNormalizado.sin_acentos(
            ruta.name.upper()
        )
        if es_no_ley:
            return ruta.stem
        if regla.incluir_concepto_en_nombre and evaluacion.conceptos:
            conceptos = evaluacion.conceptos
            if len(conceptos) > 2:
                sufijo = f"{conceptos[0]}_Y_{len(conceptos) - 1}_MAS"
            else:
                sufijo = "_".join(conceptos)
            return f"{subcarpeta.name}_{sufijo}"
        return subcarpeta.name

    def sanitizar(self, nombre: str, max_len: int = LONGITUD_NOMBRE_MAXIMA) -> str:
        limpio = TextoNormalizado.sin_acentos(nombre)
        limpio = self._CARACTERES_INVALIDOS.sub("_", limpio)
        limpio = limpio.strip(". ")
        if len(limpio) > max_len:
            limpio = limpio[:max_len].strip(". ")
        return limpio or NOMBRE_ARCHIVO_POR_DEFECTO


class BuscadorPdfs:
    """Reglas de descubrimiento de PDFs dentro de una carpeta base.

    Se revisan primero las subcarpetas ORD (salvo ORDINARIA), despues ORDINARIA,
    luego SUELDO y por ultimo la raiz de la subcarpeta, sin duplicar archivos.
    """

    SUFIJO = "*.pdf"

    def subcarpetas(self, base: Path) -> list[Path]:
        return sorted(
            (ruta for ruta in base.iterdir() if ruta.is_dir()), key=lambda ruta: ruta.name
        )

    def pdfs(self, subcarpeta: Path) -> list[Path]:
        hijos = [ruta for ruta in subcarpeta.iterdir() if ruta.is_dir()]
        de_ord = sorted(
            (
                ruta
                for ruta in hijos
                if "ORD" in ruta.name.upper() and "ORDINARIA" not in ruta.name.upper()
            ),
            key=lambda ruta: ruta.name,
        )
        de_ordinaria = sorted(
            (ruta for ruta in hijos if "ORDINARIA" in ruta.name.upper()),
            key=lambda ruta: ruta.name,
        )
        de_sueldo = sorted(
            (ruta for ruta in hijos if "SUELDO" in ruta.name.upper()),
            key=lambda ruta: ruta.name,
        )

        candidatas: list[Path] = []
        for ruta in (*de_ord, *de_ordinaria, *de_sueldo, subcarpeta):
            if ruta not in candidatas:
                candidatas.append(ruta)

        pdfs: list[Path] = []
        vistos: set[Path] = set()
        for ruta in candidatas:
            for pdf in sorted(ruta.rglob(self.SUFIJO), key=lambda elemento: str(elemento)):
                if pdf in vistos:
                    continue
                vistos.add(pdf)
                pdfs.append(pdf)
        return pdfs


class LectorPdf(ABC):
    """Puerto de lectura de PDFs desde disco."""

    @abstractmethod
    def leer(self, ruta: Path) -> PdfDocument:
        """Devuelve el PDF como documento en memoria."""


class LectorPdfPyMuPDF(LectorPdf):
    """Implementacion de LectorPdf basada en PyMuPDF."""

    def leer(self, ruta: Path) -> PdfDocument:
        try:
            with pymupdf.open(ruta) as documento:
                paginas = tuple(
                    PdfPage(numero=indice + 1, texto=documento[indice].get_text() or "")
                    for indice in range(len(documento))
                )
        except Exception as exc:
            raise PdfExtractionError(f"No se pudo leer {ruta.name}: {exc}") from exc
        return PdfDocument(nombre=ruta.stem, paginas=paginas, origen=ruta)


class UnificadorPdfs:
    """Une varios PDFs en uno solo conservando el orden recibido."""

    def unificar(self, rutas: Sequence[Path], salida: Path) -> Path:
        if not rutas:
            raise PdfExtractionError("No hay PDFs para unificar")
        try:
            unificado = pymupdf.open()
            for ruta in rutas:
                with pymupdf.open(ruta) as documento:
                    unificado.insert_pdf(documento)
            salida.parent.mkdir(parents=True, exist_ok=True)
            unificado.save(str(salida))
            unificado.close()
        except Exception as exc:
            raise PdfExtractionError(f"No se pudieron unificar los PDFs: {exc}") from exc
        return salida


@dataclass(frozen=True)
class OrganizacionResult:
    """Resultado de organizar una carpeta base."""

    log: Sequence[str]
    normales: Sequence[Path]
    deudas: Sequence[Path]
    unificados: Sequence[Path]

    def __post_init__(self) -> None:
        object.__setattr__(self, "log", tuple(self.log))
        object.__setattr__(self, "normales", tuple(self.normales))
        object.__setattr__(self, "deudas", tuple(self.deudas))
        object.__setattr__(self, "unificados", tuple(self.unificados))
