"""Dominio de aportes patronales (antes casas_comerciales.py).

Extrae filas de aportes desde PDFs con dos formatos conocidos y construye el
consolidado de Excel sin depender de Excel instalado (se elimino win32com).

Flujo de extraccion: AportesPdfExtractor decide el formato pagina por pagina
(FormatDetector), agrupa paginas consecutivas del mismo formato y delega en la
estrategia PageExtractor correspondiente mediante un flujo plantilla. El estado
del bloque (concepto actual, grupo de nomina, conteo) vive en un objeto local,
nunca en variables globales.
"""

from __future__ import annotations

import re
import unicodedata
from abc import ABC, abstractmethod
from dataclasses import dataclass
from decimal import Decimal
from io import BytesIO
from typing import Any, Mapping, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font

from core.documents import ExcelDocument, PdfDocument, PdfFormat, PdfPage
from core.errors import ReportGenerationError
from core.money import CERO, CommaDotMoneyParser, MoneyParser

CENTRO_DESCONOCIDO = "CENTRO DESCONOCIDO"

ENCABEZADO = (
    "Centro",
    "Grupo",
    "Referencia",
    "Concepto",
    "Cant",
    "Aporte Trabajador",
    "Aporte Empresa",
)

HOJA_BASE = "BASE"
HOJA_TRABAJADOR = "APORTES TRABAJADOR"
HOJA_EMPRESA = "APORTES EMPRESA"
HOJA_GREMIO = "CONSOLIDADO X GREMIO"


@dataclass(frozen=True)
class Centro:
    """Centro de salud al que pertenece un aporte."""

    nombre: str


@dataclass(frozen=True)
class Concepto:
    """Concepto de aporte ya normalizado a su codigo canonico."""

    codigo: str


@dataclass(frozen=True)
class AporteRow:
    """Fila de aporte patronal extraida de un PDF."""

    centro: Centro
    grupo: str
    referencia: str
    concepto: Concepto
    cantidad: int
    trabajador: Decimal
    empresa: Decimal

    def valores(self) -> tuple[Any, ...]:
        """Fila lista para escribir en Excel."""
        return (
            self.centro.nombre,
            self.grupo,
            self.referencia,
            self.concepto.codigo,
            self.cantidad,
            float(self.trabajador),
            float(self.empresa),
        )


class ConceptoNormalizer:
    """Normaliza los nombres de concepto de los listados de aportes."""

    @staticmethod
    def _sin_tildes(texto: str) -> str:
        descompuesto = unicodedata.normalize("NFD", texto)
        return "".join(c for c in descompuesto if unicodedata.category(c) != "Mn")

    def normalize(self, texto: str) -> Concepto:
        sin_tilde = self._sin_tildes(texto.upper())
        mayusculas = texto.upper()
        if "FONDO PENSIONES" in sin_tilde:
            return Concepto("FPJ")
        if "S.S.O." in mayusculas or "4%" in mayusculas:
            return Concepto("SSO")
        if "PERDIDA INVOLUNTARIA" in sin_tilde:
            return Concepto("PIE")
        # Antes que CAHORMINSA: el nombre puede venir truncado ("...CAHORMIN").
        if "SERVICIOS FUNERARIOS" in sin_tilde:
            return Concepto("SERVICIOS FUNERARIOS CAHORMINSAS")
        if "CAHORMINSA" in sin_tilde:
            return Concepto("CAHORMINSAS")
        if "CAJA DE AHORRO CAEMINSA" in sin_tilde or "CAEMINSA" in sin_tilde:
            return Concepto("CAEMINSAS")
        if "CAJA DE AHORRO" in sin_tilde:
            return Concepto("CAJA DE AHORRO")
        if "FONDO DE AHORRO" in sin_tilde or "AHORRO" in sin_tilde:
            return Concepto("FAOV")
        return Concepto(texto)


class CentroResolver:
    """Resuelve el centro a partir del texto del PDF y de las anclas configuradas."""

    _IGNORAR = (
        "NRO CONTROL",
        "MES:",
        "AÑO:",
        "RECIBO:",
        "PÁGINA",
        "FECHA:",
        "PAGINA",
        "LISTADO DE",
    )
    _ENCABEZADOS_FIJOS = (
        "GOBIERNO BOLIVARIANO",
        "GOBERNACION DEL ESTADO",
        "CORPORACION DE SALUD",
        "CORPORACIÓN DE SALUD",
        "TIPO DE RETENCION",
        "GRUPO DE NÓMINA",
        "MOTIVO DE NÓMINA",
        "APORTES",
        "CEDULA",
        "APORTE",
        "RECIBO CONFORME",
    )
    _PATRON_TRANSFERENCIA = re.compile(
        r"gastos de personal de\s+(.+?)\s+del?\s+\(?os?\)?\s*grupo", re.IGNORECASE
    )

    def resolve(self, pdf: PdfDocument, etiquetas: Mapping[str, Any] | None) -> Centro:
        anclas = (etiquetas or {}).get("ancla_centro", [])
        if isinstance(anclas, str):
            anclas = [anclas]
        anclas = list(anclas)
        encabezados = {ancla.upper() for ancla in anclas} | set(self._ENCABEZADOS_FIJOS)

        for pagina in pdf.paginas:
            texto = pagina.texto
            if not texto:
                continue
            lineas = texto.split("\n")

            # Recibos: "Transferencia ... gastos de personal de <CENTRO> del (os) grupo(s)..."
            coincidencia = self._PATRON_TRANSFERENCIA.search(" ".join(texto.split()))
            if coincidencia:
                return Centro(coincidencia.group(1).strip(" .,"))

            for indice, linea in enumerate(lineas):
                actual = linea.strip()
                for ancla in anclas:
                    if ancla not in actual:
                        continue
                    extra = actual.replace(ancla, "").strip(": ").strip()
                    if len(extra) > 3 and not self._es_encabezado(extra, encabezados):
                        return Centro(extra)
                    if indice + 1 < len(lineas):
                        siguiente = lineas[indice + 1].strip()
                        if not self._es_encabezado(siguiente, encabezados):
                            return Centro(siguiente)

            for linea in lineas[:8]:
                candidato = linea.strip()
                if not self._es_encabezado(candidato, encabezados):
                    return Centro(candidato)

        return Centro(CENTRO_DESCONOCIDO)

    def _es_encabezado(self, texto: str, encabezados: set[str]) -> bool:
        normalizado = " ".join(texto.upper().split())
        if not normalizado or len(normalizado) < 4:
            return True
        if any(marca in normalizado for marca in self._IGNORAR):
            return True
        return any(marca in normalizado for marca in encabezados)


class FormatDetector:
    """Decide el formato de una pagina de aportes."""

    MARCA_TABLA = "POR GRUPOS DE NÓMINA"

    def detect(self, pagina: PdfPage) -> PdfFormat:
        texto = (pagina.texto or "").upper()
        if self.MARCA_TABLA in texto:
            return PdfFormat.TABLA_GRUPOS
        return PdfFormat.LISTADO_INDIVIDUAL


@dataclass
class _EstadoExtraccion:
    """Estado mutable del bloque que se esta extrayendo."""

    concepto: str = ""
    grupo: str = "SIN GRUPO"
    concepto_listado: str = ""
    conteo: int = 0
    en_listado: bool = False


class PageExtractor(ABC):
    """Flujo plantilla de extraccion por bloque de paginas."""

    _PATRON_TOTAL = re.compile(r"(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)")
    _PATRON_FILA = re.compile(r"(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2,3})|\b\d+\b)")
    _EXCLUSIONES = (
        "PÁGINA",
        "PAGINA",
        "CONCEPTO:",
        "TRAB.",
        "EMPRESA",
        "RECIBO:",
        "MES:",
        "AÑO:",
        "FECHA:",
    )
    _PATRON_TRABAJADOR = re.compile(r"^\d{5,9}\s+")

    def __init__(self, normalizer: ConceptoNormalizer, money: MoneyParser) -> None:
        self._normalizer = normalizer
        self._money = money

    def extract(self, paginas: Sequence[PdfPage], centro: Centro) -> list[AporteRow]:
        estado = self._estado_inicial()
        filas: list[AporteRow] = []
        for pagina in paginas:
            texto = pagina.texto or ""
            if not texto:
                continue
            for linea in texto.split("\n"):
                filas.extend(self._procesar_linea(linea, estado, centro))
        return filas

    @abstractmethod
    def _estado_inicial(self) -> _EstadoExtraccion:
        """Estado con el que arranca un bloque de paginas."""

    @abstractmethod
    def _procesar_linea(
        self, linea: str, estado: _EstadoExtraccion, centro: Centro
    ) -> list[AporteRow]:
        """Convierte una linea en cero o mas filas."""

    def _es_exclusion(self, linea_up: str) -> bool:
        return any(marca in linea_up for marca in self._EXCLUSIONES)

    def _entero(self, texto: str) -> int:
        limpio = re.sub(r"[.,]", "", texto)
        try:
            return int(limpio)
        except ValueError:
            return 0

    def _fila(
        self,
        centro: Centro,
        grupo: str,
        referencia: str,
        concepto: str,
        cantidad: str,
        trabajador: str,
        empresa: str,
    ) -> AporteRow:
        return AporteRow(
            centro=centro,
            grupo=grupo,
            referencia=referencia,
            concepto=Concepto(concepto),
            cantidad=self._entero(cantidad),
            trabajador=self._money.parse(trabajador),
            empresa=self._money.parse(empresa),
        )


class TablaGruposExtractor(PageExtractor):
    """Formato "listado de deducciones por grupos de nomina"."""

    _PATRON_CONCEPTO = "CONCEPTO:"

    def _estado_inicial(self) -> _EstadoExtraccion:
        return _EstadoExtraccion()

    def _procesar_linea(
        self, linea: str, estado: _EstadoExtraccion, centro: Centro
    ) -> list[AporteRow]:
        linea_up = linea.upper().strip()

        if "LISTADO DE" in linea_up or self._es_exclusion(linea_up):
            if self._PATRON_CONCEPTO in linea_up:
                fragmento = linea_up.split(self._PATRON_CONCEPTO)[-1].strip()
                estado.concepto = self._normalizer.normalize(fragmento).codigo
            return []

        if not estado.concepto:
            return []

        if "TOTAL DEDUCCIONES" in linea_up:
            montos = self._PATRON_TOTAL.findall(linea)
            if len(montos) < 2:
                return []
            cantidad, trabajador, empresa = montos[0], montos[-2], montos[-1]
            return [
                self._fila(
                    centro,
                    "TOTAL GENERAL",
                    f"DEDUCCION {estado.concepto} TOTAL {trabajador} {empresa}",
                    estado.concepto,
                    cantidad,
                    trabajador,
                    empresa,
                )
            ]

        montos = self._PATRON_FILA.findall(linea)
        if len(montos) < 2:
            return []
        nombre_grupo = linea.split(montos[0])[0].strip()
        if len(nombre_grupo) < 4 or "TOTALES" in nombre_grupo.upper():
            return []
        cantidad, trabajador = montos[0], montos[1]
        empresa = montos[2] if len(montos) > 2 else "0,00"
        return [
            self._fila(
                centro,
                nombre_grupo,
                f"DEDUCCION {estado.concepto} {nombre_grupo} {trabajador} {empresa}",
                estado.concepto,
                cantidad,
                trabajador,
                empresa,
            )
        ]


class ListadoIndividualExtractor(PageExtractor):
    """Formato "listado de deducciones" con un trabajador por linea."""

    _PATRON_CONCEPTO_LISTADO = re.compile(r"\s+AL\s+")

    def _estado_inicial(self) -> _EstadoExtraccion:
        return _EstadoExtraccion()

    def _procesar_linea(
        self, linea: str, estado: _EstadoExtraccion, centro: Centro
    ) -> list[AporteRow]:
        linea_up = linea.upper().strip()

        if "GRUPO DE NÓMINA" in linea_up:
            estado.grupo = linea_up.split(":")[-1].strip()
            return []

        if "LISTADO DE" in linea_up:
            estado.en_listado = True
            estado.conteo = 0
            fragmento = linea_up.split("LISTADO DE")[-1]
            nombre_listado = self._PATRON_CONCEPTO_LISTADO.split(fragmento)[0]
            estado.concepto_listado = self._normalizer.normalize(
                nombre_listado.strip(" .")
            ).codigo
            return []

        if estado.en_listado and self._PATRON_TRABAJADOR.match(linea.strip()):
            estado.conteo += 1

        if "TOTAL DEDUCCIONES" not in linea_up:
            return []

        estado.en_listado = False
        montos = self._PATRON_TOTAL.findall(linea)
        if not montos:
            return []
        trabajador = montos[-2] if len(montos) >= 2 else montos[-1]
        empresa = montos[-1] if len(montos) >= 2 else "0,00"
        concepto = estado.concepto_listado or self._normalizer.normalize(
            linea_up.replace("TOTAL DEDUCCIONES", "").split("(")[0].strip()
        ).codigo
        return [
            AporteRow(
                centro=centro,
                grupo=estado.grupo,
                referencia=linea.strip(),
                concepto=Concepto(concepto),
                cantidad=estado.conteo,
                trabajador=self._money.parse(trabajador),
                empresa=self._money.parse(empresa),
            )
        ]


class AportesPdfExtractor:
    """Extrae todas las filas de un PDF mezclando formatos por bloques."""

    def __init__(
        self,
        money: MoneyParser | None = None,
        detector: FormatDetector | None = None,
        resolver: CentroResolver | None = None,
        normalizer: ConceptoNormalizer | None = None,
        extractores: Mapping[PdfFormat, PageExtractor] | None = None,
    ) -> None:
        self._money = money or CommaDotMoneyParser()
        self._detector = detector or FormatDetector()
        self._resolver = resolver or CentroResolver()
        self._normalizer = normalizer or ConceptoNormalizer()
        self._extractores = dict(
            extractores
            or {
                PdfFormat.TABLA_GRUPOS: TablaGruposExtractor(self._normalizer, self._money),
                PdfFormat.LISTADO_INDIVIDUAL: ListadoIndividualExtractor(
                    self._normalizer, self._money
                ),
            }
        )

    def extract(
        self, pdf: PdfDocument, etiquetas: Mapping[str, Any] | None = None
    ) -> list[AporteRow]:
        centro = self._resolver.resolve(pdf, etiquetas)
        filas: list[AporteRow] = []
        bloque: list[PdfPage] = []
        formato: PdfFormat | None = None

        for pagina in pdf.paginas:
            if not (pagina.texto or "").strip():
                continue
            actual = self._detector.detect(pagina)
            if bloque and actual is not formato:
                filas.extend(self._extraer_bloque(bloque, formato, centro))
                bloque = []
            bloque.append(pagina)
            formato = actual

        if bloque:
            filas.extend(self._extraer_bloque(bloque, formato, centro))
        return filas

    def _extraer_bloque(
        self, paginas: Sequence[PdfPage], formato: PdfFormat | None, centro: Centro
    ) -> list[AporteRow]:
        if formato is None or formato not in self._extractores:
            return []
        return self._extractores[formato].extract(paginas, centro)


class ReportBuilder(ABC):
    """Puerto de construccion de libros de Excel."""

    @abstractmethod
    def build(self, filas: Sequence[AporteRow]) -> ExcelDocument:
        """Construye el libro a partir de las filas."""


class AportesExcelBuilder(ReportBuilder):
    """Consolidado de aportes: hoja BASE mas resumenes por centro y por gremio.

    Sustituye las tablas dinamicas que se creaban con win32com (Excel en
    Windows) por tablas resumen escritas con openpyxl.
    """

    NOMBRE_ARCHIVO = "CONSOLIDADO APORTES PATRONALES.xlsx"

    def __init__(self, money: MoneyParser | None = None) -> None:
        self._money = money or CommaDotMoneyParser()

    def build(self, filas: Sequence[AporteRow]) -> ExcelDocument:
        try:
            libro = Workbook()
            base = libro.active
            base.title = HOJA_BASE
            base.append(list(ENCABEZADO))
            for celda in base[1]:
                celda.font = Font(bold=True)
            for fila in filas:
                base.append(list(fila.valores()))

            self._hoja_resumen(libro, HOJA_TRABAJADOR, filas, campo="trabajador")
            self._hoja_resumen(libro, HOJA_EMPRESA, filas, campo="empresa")
            self._hoja_resumen(libro, HOJA_GREMIO, filas, campo="empresa", por_grupo=True)

            buffer = BytesIO()
            libro.save(buffer)
        except ReportGenerationError:
            raise
        except Exception as exc:  # pragma: no cover - defensivo
            raise ReportGenerationError(
                "No se pudo generar el consolidado de aportes patronales"
            ) from exc
        return ExcelDocument(nombre=self.NOMBRE_ARCHIVO, contenido=buffer.getvalue())

    def _numero(self, valor: Decimal) -> float:
        return float(self._money.parse(str(valor)))

    def _clave(self, fila: AporteRow, por_grupo: bool) -> str:
        return fila.grupo if por_grupo else fila.centro.nombre

    def _hoja_resumen(
        self,
        libro: Workbook,
        titulo: str,
        filas: Sequence[AporteRow],
        campo: str,
        por_grupo: bool = False,
    ) -> None:
        hoja = libro.create_sheet(titulo)
        encabezado = "Grupo" if por_grupo else "Centro"
        claves = sorted({self._clave(fila, por_grupo) for fila in filas})
        conceptos = sorted({fila.concepto.codigo for fila in filas})

        hoja.cell(row=1, column=1, value=encabezado)
        for columna, concepto in enumerate(conceptos, start=2):
            hoja.cell(row=1, column=columna, value=concepto)
        for celda in hoja[1]:
            celda.font = Font(bold=True)

        totales = {concepto: CERO for concepto in conceptos}
        fila_excel = 2
        for clave in claves:
            hoja.cell(row=fila_excel, column=1, value=clave)
            for columna, concepto in enumerate(conceptos, start=2):
                suma = sum(
                    (
                        getattr(fila, campo)
                        for fila in filas
                        if self._clave(fila, por_grupo) == clave
                        and fila.concepto.codigo == concepto
                    ),
                    CERO,
                )
                celda = hoja.cell(row=fila_excel, column=columna, value=self._numero(suma))
                celda.number_format = "#,##0.00"
                totales[concepto] += suma
            fila_excel += 1

        hoja.cell(row=fila_excel, column=1, value="TOTAL").font = Font(bold=True)
        for columna, concepto in enumerate(conceptos, start=2):
            celda = hoja.cell(
                row=fila_excel, column=columna, value=self._numero(totales[concepto])
            )
            celda.number_format = "#,##0.00"
            celda.font = Font(bold=True)


@dataclass(frozen=True)
class AportesResult:
    """Resultado de procesar un lote de PDFs de aportes."""

    filas: Sequence[AporteRow]
    documento: ExcelDocument
    registros: int

    def __post_init__(self) -> None:
        object.__setattr__(self, "filas", tuple(self.filas))
