"""Dominio de casas comerciales (antes casas_comercialesf.py).

Modulo autonomo: duplica deliberadamente Centro/Concepto/fila y el detector de
formato de aportes patronales para no importar entre modulos. Genera la
auditoria en Excel con openpyxl, sin tablas dinamicas ni Excel instalado.
"""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import dataclass
from decimal import Decimal
from enum import Enum
from io import BytesIO
from typing import Any, Mapping, Protocol, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font

from core.documents import ExcelDocument, PdfDocument, PdfFormat, PdfPage
from core.errors import ReportGenerationError
from core.money import CommaDotMoneyParser, MoneyParser

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
NOMBRE_ARCHIVO = "_AUDITORIA_CASAS_COMERCIALES.xlsx"
LONGITUD_MAXIMA_HOJA = 31


class GremioCategoria(Enum):
    """Categorias de gremio que resume la auditoria."""

    EMPLEADOS_FIJOS = "Empleados Fijos"
    OBREROS_FIJOS = "Obreros Fijos"
    EMPLEADOS_CONTRATADOS = "Empleados Contratados"
    OBREROS_CONTRATADOS = "Obreros Contratados"


class ClasificadorGremio:
    """Clasifica el grupo de nomina en una categoria de gremio.

    Conserva las reglas historicas: comparaciones de subcadena literal en
    singular y en orden fijo. Los grupos en plural no clasifican (None) y el
    comportamiento queda fijado por tests; revisar en Target 6.
    """

    GREMIOS_JUBILADOS_ADMINISTRATIVOS = frozenset(
        {
            "PERSONAL JUBILADO ADMINISTRATIVO OTROS GREMIOS",
            "PERSONAL JUBILADO MÉDICOS",
            "PERSONAL JUBILADO MEDICOS",
            "PERSONAL PENSIÓN DE SOBREVIVIENTE",
            "PERSONAL PENSION DE SOBREVIVIENTE",
            "PERSONAL PENSIÓN DE INVALIDEZ ADMINISTRATIVOS Y OTROS",
            "PERSONAL PENSION DE INVALIDEZ ADMINISTRATIVOS Y OTROS",
        }
    )

    def clasificar(self, grupo: str | None) -> GremioCategoria | None:
        clasificado = (grupo or "").upper().strip()
        if "EMPLEADO CONTRATADO" in clasificado:
            return GremioCategoria.EMPLEADOS_CONTRATADOS
        if "OBRERO CONTRATADO" in clasificado:
            return GremioCategoria.OBREROS_CONTRATADOS
        if "OBRERO FIJO" in clasificado:
            return GremioCategoria.OBREROS_FIJOS
        if clasificado == "PERSONAL JUBILADO OBRERO":
            return GremioCategoria.OBREROS_FIJOS
        if "EMPLEADO FIJO" in clasificado:
            return GremioCategoria.EMPLEADOS_FIJOS
        if clasificado in self.GREMIOS_JUBILADOS_ADMINISTRATIVOS:
            return GremioCategoria.EMPLEADOS_FIJOS
        if (
            "PENSIÓN" in clasificado
            or "PENSION" in clasificado
            or "JUBILADO" in clasificado
        ) and "OBRERO" not in clasificado:
            return GremioCategoria.EMPLEADOS_FIJOS
        return None


@dataclass(frozen=True)
class Centro:
    """Centro de salud al que pertenece la deduccion."""

    nombre: str


@dataclass(frozen=True)
class Concepto:
    """Concepto de deduccion ya normalizado."""

    codigo: str


@dataclass(frozen=True)
class ResumenCategoria:
    """Cantidad de trabajadores por categoria de gremio."""

    categoria: GremioCategoria
    cantidad: int


class _CentroConNombre(Protocol):
    nombre: str


class _ConceptoConCodigo(Protocol):
    codigo: str


class AporteComun(Protocol):
    """Forma minima de una fila de aportes, sin acoplar el modulo de aportes."""

    centro: _CentroConNombre
    grupo: str
    referencia: str
    concepto: _ConceptoConCodigo
    cantidad: int
    trabajador: Decimal
    empresa: Decimal


@dataclass(frozen=True)
class CasasComercialesRow:
    """Fila de la auditoria de casas comerciales."""

    centro: Centro
    grupo: str
    referencia: str
    concepto: Concepto
    cantidad: int
    trabajador: Decimal
    empresa: Decimal

    @classmethod
    def from_aporte(cls, fila: AporteComun) -> CasasComercialesRow:
        """Convierte una fila de aportes sin importar su modulo.

        La duplicacion de la forma de fila es intencional: los modulos no se
        importan entre si y solo comparten la estructura (Protocol).
        """
        return cls(
            centro=Centro(fila.centro.nombre),
            grupo=fila.grupo,
            referencia=fila.referencia,
            concepto=Concepto(fila.concepto.codigo),
            cantidad=fila.cantidad,
            trabajador=fila.trabajador,
            empresa=fila.empresa,
        )

    def valores(self) -> tuple[Any, ...]:
        return (
            self.centro.nombre,
            self.grupo,
            self.referencia,
            self.concepto.codigo,
            self.cantidad,
            float(self.trabajador),
            float(self.empresa),
        )


class ConceptoCasasNormalizer:
    """Normaliza conceptos con las reglas propias de casas comerciales."""

    def normalize(self, texto: str) -> Concepto:
        normalizado = (texto or "").upper().strip()
        if "FUNERARIOS" in normalizado and "CAHORMINSAS" in normalizado:
            return Concepto("CAHORMINSAS")
        if "CAJA DE AHORRO CAEMINSA" in normalizado or "CAEMINSA" in normalizado:
            return Concepto("CAEMINSA")
        return Concepto(normalizado)


class CentroCasasResolver:
    """Resuelve el centro a partir del texto del PDF y de las anclas configuradas.

    Comparte reglas con CentroResolver de aportes patronales (implementacion
    duplicada a proposito para mantener la autonomia entre modulos):
      - se recorren todas las paginas;
      - se acepta la regla "gastos de personal de ... del (os) grupo(s)";
      - el ancla puede traer el centro en la misma linea o en la siguiente;
      - como respaldo se miran las primeras 8 lineas de cada pagina;
      - se descartan encabezados, metadatos y lineas de gobierno/corporacion.
    """

    _IGNORAR = (
        "NRO CONTROL",
        "MES:",
        "AÑO:",
        "RECIBO:",
        "PÁGINA",
        "FECHA:",
        "PAGINA",
        "LISTADO DE",
        "CORPORACION",
        "GOBERNACION",
    )
    _ENCABEZADOS = (
        "GOBIERNO BOLIVARIANO",
        "TIPO DE RETENCION",
        "GRUPO DE NÓMINA",
        "MOTIVO DE NÓMINA",
        "APORTES",
        "CEDULA",
        "APORTE",
        "RECIBO CONFORME",
        "RECURSOS HUMANOS",
        "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS",
        "DIRECCION DE RECURSOS HUMANOS",
        "DIRECCION REGIONAL DE RECURSOS HUMANOS",
    )
    _PATRON_TRANSFERENCIA = re.compile(
        r"gastos de personal de\s+(.+?)\s+del?\s+\(?os?\)?\s*grupo", re.IGNORECASE
    )

    def resolve(
        self, pdf: PdfDocument, etiquetas: Mapping[str, Any] | None = None
    ) -> Centro:
        anclas = (etiquetas or {}).get("ancla_centro", [])
        if isinstance(anclas, str):
            anclas = [anclas]
        anclas = list(anclas)
        encabezados = {ancla.upper() for ancla in anclas} | set(self._ENCABEZADOS)

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


class _DetectorFormatoCasas:
    """Detecta el formato a nivel de archivo, como el sistema anterior."""

    MARCA_TABLA = "POR GRUPOS DE NÓMINA"

    def detectar(self, pdf: PdfDocument) -> PdfFormat:
        texto = (pdf.texto_primera_pagina or "").upper()
        if self.MARCA_TABLA in texto:
            return PdfFormat.TABLA_GRUPOS
        return PdfFormat.LISTADO_INDIVIDUAL


@dataclass
class _EstadoCasas:
    concepto: str = ""
    grupo: str = "SIN GRUPO"
    conteo: int = 0
    en_listado: bool = False


class _ExtractorCasas(ABC):
    """Flujo plantilla de extraccion de casas comerciales."""

    _PATRON_TOTAL = re.compile(r"(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)")
    _PATRON_FILA = re.compile(r"(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2,3})|\b\d+\b)")
    _PATRON_TRABAJADOR = re.compile(r"^\d{5,9}\s+")
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

    def __init__(self, normalizer: ConceptoCasasNormalizer, money: MoneyParser) -> None:
        self._normalizer = normalizer
        self._money = money

    def extract(self, paginas: Sequence[PdfPage], centro: Centro) -> list[CasasComercialesRow]:
        estado = _EstadoCasas()
        filas: list[CasasComercialesRow] = []
        for pagina in paginas:
            texto = pagina.texto or ""
            if not texto:
                continue
            for linea in texto.split("\n"):
                filas.extend(self._procesar_linea(linea, estado, centro))
        return filas

    @abstractmethod
    def _procesar_linea(
        self, linea: str, estado: _EstadoCasas, centro: Centro
    ) -> list[CasasComercialesRow]:
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
    ) -> CasasComercialesRow:
        return CasasComercialesRow(
            centro=centro,
            grupo=grupo,
            referencia=referencia,
            concepto=Concepto(concepto),
            cantidad=self._entero(cantidad),
            trabajador=self._money.parse(trabajador),
            empresa=self._money.parse(empresa),
        )


class _TablaCasasExtractor(_ExtractorCasas):
    _PATRON_CONCEPTO = "CONCEPTO:"

    def _procesar_linea(
        self, linea: str, estado: _EstadoCasas, centro: Centro
    ) -> list[CasasComercialesRow]:
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


class _ListadoCasasExtractor(_ExtractorCasas):
    def _procesar_linea(
        self, linea: str, estado: _EstadoCasas, centro: Centro
    ) -> list[CasasComercialesRow]:
        linea_up = linea.upper().strip()

        if "GRUPO DE NÓMINA" in linea_up:
            estado.grupo = linea_up.split(":")[-1].strip()
            return []

        if "LISTADO DE" in linea_up:
            estado.en_listado = True
            estado.conteo = 0
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
        concepto = self._normalizer.normalize(
            linea_up.replace("TOTAL DEDUCCIONES", "").split("(")[0].strip()
        ).codigo
        return [
            CasasComercialesRow(
                centro=centro,
                grupo=estado.grupo,
                referencia=linea.strip(),
                concepto=Concepto(concepto),
                cantidad=estado.conteo,
                trabajador=self._money.parse(trabajador),
                empresa=self._money.parse(empresa),
            )
        ]


class CasasPdfExtractor:
    """Extrae filas de auditoria de un PDF de casas comerciales."""

    def __init__(
        self,
        money: MoneyParser | None = None,
        resolver: CentroCasasResolver | None = None,
        normalizer: ConceptoCasasNormalizer | None = None,
        detector: _DetectorFormatoCasas | None = None,
        extractores: Mapping[PdfFormat, _ExtractorCasas] | None = None,
    ) -> None:
        self._money = money or CommaDotMoneyParser()
        self._resolver = resolver or CentroCasasResolver()
        self._normalizer = normalizer or ConceptoCasasNormalizer()
        self._detector = detector or _DetectorFormatoCasas()
        self._extractores = dict(
            extractores
            or {
                PdfFormat.TABLA_GRUPOS: _TablaCasasExtractor(self._normalizer, self._money),
                PdfFormat.LISTADO_INDIVIDUAL: _ListadoCasasExtractor(
                    self._normalizer, self._money
                ),
            }
        )

    def extract(
        self, pdf: PdfDocument, etiquetas: Mapping[str, Any] | None = None
    ) -> list[CasasComercialesRow]:
        if not pdf.paginas:
            return []
        centro = self._resolver.resolve(pdf, etiquetas)
        formato = self._detector.detectar(pdf)
        return self._extractores[formato].extract(pdf.paginas, centro)


class AuditoriaExcelBuilder:
    """Auditoria de casas comerciales: hoja BASE por concepto mas resumen.

    Reemplaza el escritor openpyxl usado dentro del flujo con tkinter. El
    resumen se escribe debajo de los datos con las cuatro categorias de gremio
    y una fila TOTAL. El layout historico pisaba la ultima categoria con el
    TOTAL; aqui se corrige y el TOTAL ocupa la fila siguiente (desviacion
    documentada).
    """

    def __init__(
        self,
        clasificador: ClasificadorGremio | None = None,
        money: MoneyParser | None = None,
    ) -> None:
        self._clasificador = clasificador or ClasificadorGremio()
        self._money = money or CommaDotMoneyParser()

    ORDEN_CATEGORIAS = (
        GremioCategoria.EMPLEADOS_FIJOS,
        GremioCategoria.OBREROS_FIJOS,
        GremioCategoria.EMPLEADOS_CONTRATADOS,
        GremioCategoria.OBREROS_CONTRATADOS,
    )

    def build(self, filas: Sequence[CasasComercialesRow]) -> ExcelDocument:
        try:
            libro = Workbook()
            base = libro.active
            base.title = HOJA_BASE
            base.append(list(ENCABEZADO))
            for celda in base[1]:
                celda.font = Font(bold=True)
            for fila in filas:
                base.append(list(fila.valores()))

            usados: set[str] = {HOJA_BASE}
            for concepto in sorted({fila.concepto.codigo for fila in filas}):
                propias = [fila for fila in filas if fila.concepto.codigo == concepto]
                hoja = libro.create_sheet(self._nombre_hoja(concepto, usados))
                hoja.append(list(ENCABEZADO))
                for fila in propias:
                    hoja.append(list(fila.valores()))
                self._escribir_resumen(hoja, propias)

            buffer = BytesIO()
            libro.save(buffer)
        except ReportGenerationError:
            raise
        except Exception as exc:  # pragma: no cover - defensivo
            raise ReportGenerationError(
                "No se pudo generar la auditoria de casas comerciales"
            ) from exc
        return ExcelDocument(nombre=NOMBRE_ARCHIVO, contenido=buffer.getvalue())

    def _nombre_hoja(self, concepto: str, usados: set[str]) -> str:
        nombre = str(concepto)[:LONGITUD_MAXIMA_HOJA].replace(":", "").replace("/", "").strip()
        nombre = nombre or "CONCEPTO"
        candidato = nombre
        sufijo = 2
        while candidato in usados:
            sufijo_texto = f"_{sufijo}"
            candidato = nombre[: LONGITUD_MAXIMA_HOJA - len(sufijo_texto)] + sufijo_texto
            sufijo += 1
        usados.add(candidato)
        return candidato

    def _escribir_resumen(
        self, hoja, filas: Sequence[CasasComercialesRow]
    ) -> None:
        inicio = len(filas) + 3
        total = 0
        for desplazamiento, categoria in enumerate(self.ORDEN_CATEGORIAS):
            cantidad = sum(
                fila.cantidad
                for fila in filas
                if self._clasificador.clasificar(fila.grupo) is categoria
            )
            total += cantidad
            hoja.cell(row=inicio + desplazamiento, column=1, value=categoria.value)
            hoja.cell(row=inicio + desplazamiento, column=2, value=int(cantidad))
        fila_total = inicio + len(self.ORDEN_CATEGORIAS)
        hoja.cell(row=fila_total, column=1, value="TOTAL").font = Font(bold=True)
        hoja.cell(row=fila_total, column=2, value=int(total)).font = Font(bold=True)

    def resumen_por_categoria(
        self, filas: Sequence[CasasComercialesRow]
    ) -> list[ResumenCategoria]:
        """Resumen de cantidades por categoria (util para reportes y tests)."""
        return [
            ResumenCategoria(
                categoria=categoria,
                cantidad=sum(
                    fila.cantidad
                    for fila in filas
                    if self._clasificador.clasificar(fila.grupo) is categoria
                ),
            )
            for categoria in self.ORDEN_CATEGORIAS
        ]