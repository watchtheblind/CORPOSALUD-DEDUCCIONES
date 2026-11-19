"""Dominio de recibos (antes extractor_nomina.py, procesador_reportes.py y RESUMIDOR_DE_TXT.py).

Reune el parseo de recibos de nomina/deuda, la auditoria en memoria, el
consolidado en Excel y los reportes de retenciones/casas comerciales. No lee ni
escribe archivos: todo entra y sale en memoria (PdfDocument -> ExcelDocument).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from io import BytesIO
from types import MappingProxyType
from typing import Callable, Mapping, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from core.documents import ExcelDocument, PdfDocument
from core.errors import PdfExtractionError, ReportGenerationError
from core.money import LegacyReciboMoneyParser, MoneyParser

CENTRO_DESCONOCIDO = "CENTRO DESCONOCIDO"
PALABRA_DEUDA = "DEUDA"

MESES = (
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
)

MESES_ABREVIADOS: Mapping[str, str] = MappingProxyType(
    {
        "ENERO": "ENE.",
        "FEBRERO": "FEB.",
        "MARZO": "MAR.",
        "ABRIL": "ABR.",
        "MAYO": "MAY.",
        "JUNIO": "JUN.",
        "JULIO": "JUL.",
        "AGOSTO": "AGO.",
        "SEPTIEMBRE": "SEP.",
        "OCTUBRE": "OCT.",
        "NOVIEMBRE": "NOV.",
        "DICIEMBRE": "DIC.",
    }
)

LIMITE_NOMBRE_HOJA = 31

ENCABEZADO_CONSOLIDADO = ("Conceptos", "Entidad", "Recibo", "Montos")
TITULO_NOMINAS = "--- SECCIÓN NÓMINAS ---"
TITULO_DEUDAS = "--- SECCIÓN DEUDAS ---"

CONCEPTOS_POR_DEFECTO: Mapping[str, str] = MappingProxyType(
    {
        # SSO
        "LEY S.S.O. (4%)": "210",
        "S.S.O. (DEUDA)": "210.1.1",
        # COLEGIOS
        "COLEGIO BIOANALISTA": "255",
        "COLEGIO NUTRIC. Y DIET. VZLA": "286",
        "COLEGIO ENFERMERAS": "224",
        "COLEGIO DE ENFERMERAS ESTADO CARABOBO": "967",
        "COLEGIO DE ENFERMERA EDO. MIRANDA": "603",
        # DIAS DESCANSO
        "NO LEY DESC. DIA(S) NO LABORADO(S)": "179",
        "DESC. DIA(S) NO LABORADO(S)": "179",
        # PIE
        "LEY PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.": "244",
        "LEY PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.)": "244",
        "PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.": "244",
        "PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.)": "244",
        "PERDIDA INVOLUNTARIA DEL EMPLEO (DEUDA)": "244.2",
        "LEY PERDIDA INVOLUNTARIA DEL EMPLEO (DEUDA)": "244.2",
        # FAOV (el mapa historico definia dos veces "...PARA LA VIV": gana 245.3)
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIV": "245.3",
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVI": "245",
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA": "245",
        # DELEGACIONES
        "DELEGACIÓN REGIONAL INPRENFERMERA ARAGU": "370",
        "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA": "370",
        "DELEGACIÓN REGIONAL INPREENFERMERA ARAGU": "370",
        "DELEGACIÓN REGIONAL INPREENFERMERA ARAGUA": "370",
        # PENSIONES
        "FONDO PENSIONES (JUBILACION)": "301",
        "FONDO PENSIONES JUBILACION (DEUDA)": "301.2",
        # CAHORMINSAS
        "NO LEY CAHORMINSAS": "212",
        "CAHORMINSAS": "212",
        "SERVICIOS FUNERARIOS CAHORMINSAS": "269",
        # CAJA AHORRO
        "CAJA DE AHORRO CAEMINSA": "238",
        "PRESTAMO CAJA DE AHORRO": "223",
        # SINDICATOS Y SOCIEDADES
        "SINDICATO UNICO DE TRAB.DE LA SALUD Y S": "233",
        "SINDICATO OSBESS ARAGUA": "978",
        "SOCIEDAD ANESTECIOLOGOS": "257",
        "SINBOPROENF": "321",
        "SUNEP-SAS": "213",
        "CUOTA EXTRAORDINARIA SUNEP-SAS": "929",
        "FENASISTRASALUD": "513",
        "SISTRASALUD": "583",
        "SAPTRASEZ": "964",
        "SITRASSS-MIRANDA": "581",
        "ASUNAJUPENSAPROSO": "799",
        # TRIBUNALES
        "TRIBUNAL DE PROTECCION DE NIÑOS, NIÑAS": "807",
        "TRIBUNAL (PERMANENTE)": "215",
    }
)

MAPEO_CENTROS: Mapping[str, tuple[str, ...]] = MappingProxyType(
    {
        "SAANA": ("S.A.A.N.A.", "SAANA", "S.A.A.N.A"),
        "SAGER": ("S.A.G.E.R.", "SAGER", "S.A.G.E.R"),
        "AMBULATORIO DE TURMERO": ("AMB. DE TURMERO", "AMBULATORIO DE TURMERO"),
        "SALA DE PARTO DE TURMERO": (
            "SALA DE PARTO DE TURMERO",
            "SALA DE PARTO",
            "SALA PARTO TURMERO",
        ),
        "SOCIEDAD CIVIL HOSPITAL DEL SUR": (
            "HOSPITAL DEL SUR",
            "SOCIEDAD CIVIL HOSPITAL DEL SUR",
            "SOCIEDAD CIVIL",
        ),
        "D.M.S CAMATAGUA": ("D.M.S. CAMATAGUA", "D.M.S CAMATAGUA-URDANETA", "D.M.S CAMATAGUA-"),
        "HOSPITAL CENTRAL DE MARACAY": ("S.A. HOSPITAL CENTRAL DE MARACAY",),
        "HOSPITAL JOSE MARIA BENITEZ": ("S.A. HOSPITAL JOSE MARIA BENITEZ",),
        "HOSPITAL JOSE MARIA VARGAS": ("HOSPITAL JOSE MARIA VARGAS",),
        "HOSPITAL JOSE RANGEL": ("HOSPITAL JOSE RANGEL",),
        "HOSPITAL NUESTRA SEÑORA DE LA CARIDAD": (
            "HOSPITAL NUESTRA SEÑORA DE LA CARIDAD",
        ),
        "HOSPITAL LAS TEJERIAS": ("HOSPITAL LAS TEJERIAS",),
        "CLINICA PSIQUIATRICA DE MARACAY": ("CLINICA PSIQUIATRICA DE MARACAY",),
        "D.M.S. GIRARDOT": ("D.M.S. GIRARDOT",),
        "D.M.S. MARIÑO": ("D.M.S. MARIÑO",),
        "D.M.S. RIBAS": ("D.M.S. RIBAS",),
        "D.M.S. SUCRE": ("D.M.S. SUCRE",),
        "D.M.S. ZAMORA": ("D.M.S. ZAMORA",),
        "D.M.S. LIBERTADOR": ("D.M.S. LIBERTADOR",),
        "D.M.S. SANTOS MICHELENA": ("D.M.S. SANTOS MICHELENA",),
        "D.M.S. TOVAR": ("D.M.S. TOVAR",),
        "D.M.S. REVENGA": ("D.M.S. REVENGA",),
        "D.M.S. SAN CASIMIRO": ("D.M.S. SAN CASIMIRO",),
        "D.M.S. BOLIVAR": ("D.M.S. BOLIVAR",),
        "D.M.S. FRANCISCO LINARES ALCANTARA": ("D.M.S. FRANCISCO LINARES ALCANTARA",),
        "D.M.S. MARIO BRICEÑO IRAGORRI": ("D.M.S. MARIO BRICEÑO IRAGORRI",),
        "AMB. PALO NEGRO": ("AMB. PALO NEGRO",),
        "AMB. LA CANDELARIA": ("AMB. LA CANDELARIA",),
        "SERVICIOS CENTRALES": ("SERVICIOS CENTRALES",),
        "CORPORACION DE SALUD DEL ESTADO ARAGUA": ("CORPORACION DE SALUD DEL ESTADO ARAGUA",),
    }
)


@dataclass(frozen=True)
class Centro:
    """Entidad o centro de salud al que pertenece un recibo."""

    nombre: str


@dataclass(frozen=True)
class ConceptoCode:
    """Concepto de recibo con su codigo contable."""

    nombre: str
    codigo: str

    @property
    def etiqueta(self) -> str:
        return f"{self.nombre} {self.codigo}"


@dataclass(frozen=True)
class ReciboDocument:
    """Una linea de concepto leida de un recibo."""

    concepto: ConceptoCode
    entidad: Centro
    numero: str
    monto: Decimal


class CatalogoConceptos:
    """Mapa texto -> concepto; gana siempre la clave mas larga presente en la linea."""

    def __init__(self, items: Mapping[str, str] | None = None) -> None:
        self._items = dict(items if items is not None else CONCEPTOS_POR_DEFECTO)
        self._por_longitud = sorted(self._items, key=len, reverse=True)

    def resolver(self, linea: str) -> ConceptoCode | None:
        linea_up = linea.upper()
        for nombre in self._por_longitud:
            if nombre.upper() in linea_up:
                return ConceptoCode(nombre=nombre, codigo=self._items[nombre])
        return None

    def nombres(self) -> tuple[str, ...]:
        return tuple(self._items)


class CatalogoCentros:
    """Normaliza el centro leido en el PDF a su nombre estandar."""

    def __init__(self, alias: Mapping[str, Sequence[str]] | None = None) -> None:
        self._alias = {
            estandar: tuple(variaciones)
            for estandar, variaciones in (alias or MAPEO_CENTROS).items()
        }

    def normalizar(self, crudo: str | None) -> Centro:
        if not crudo:
            return Centro(CENTRO_DESCONOCIDO)
        limpio = crudo.strip().upper()
        for estandar, variaciones in self._alias.items():
            if any(variacion.upper() in limpio for variacion in variaciones):
                return Centro(estandar)
        return Centro(limpio)

    def nombres(self) -> tuple[str, ...]:
        return tuple(self._alias)


@dataclass(frozen=True)
class ReciboParseResult:
    """Resultado de parsear un PDF de recibos."""

    numero: str
    entidad: Centro
    documentos: Sequence[ReciboDocument]

    def __post_init__(self) -> None:
        object.__setattr__(self, "documentos", tuple(self.documentos))


class ReciboParser:
    """Extrae conceptos y montos de un recibo individual."""

    _PATRON_RECIBO = re.compile(r"Recibo\s*N[º°]?\s*(\d+)")
    _MARCA_INICIO = "TOTAL DE DEDUCCIONES POR CONCEPTO"
    _MARCAS_FIN = ("Total Asignaciones", "Total Pago Único")

    def __init__(
        self,
        catalogo: CatalogoConceptos | None = None,
        centros: CatalogoCentros | None = None,
        money: MoneyParser | None = None,
    ) -> None:
        self._catalogo = catalogo or CatalogoConceptos()
        self._centros = centros or CatalogoCentros()
        self._money = money or LegacyReciboMoneyParser()

    def parse(self, pdf: PdfDocument, es_deuda: bool) -> list[ReciboDocument]:
        return list(self.parse_documento(pdf, es_deuda).documentos)

    def nombres_entidades(self) -> tuple[str, ...]:
        return self._centros.nombres()

    def parse_documento(self, pdf: PdfDocument, es_deuda: bool) -> ReciboParseResult:
        if not pdf.paginas:
            raise PdfExtractionError(f"El PDF {pdf.nombre_archivo} no tiene paginas")

        texto_pagina_1 = pdf.texto_primera_pagina.replace("\n", " ")
        entidad = self._centros.normalizar(texto_pagina_1)
        coincidencia = self._PATRON_RECIBO.search(texto_pagina_1)
        numero = coincidencia.group(1) if coincidencia else "S/N"

        activado = bool(es_deuda)
        documentos: list[ReciboDocument] = []
        for linea in pdf.texto_completo().split("\n"):
            if not es_deuda and self._MARCA_INICIO in linea:
                activado = True
                continue
            if any(marca in linea for marca in self._MARCAS_FIN):
                if activado:
                    break
            if not activado:
                continue
            partes = linea.strip().rsplit(" ", 1)
            if len(partes) != 2:
                continue
            concepto = self._catalogo.resolver(partes[0])
            if concepto is None:
                continue
            documentos.append(
                ReciboDocument(
                    concepto=concepto,
                    entidad=entidad,
                    numero=numero,
                    monto=self._money.parse(partes[1]),
                )
            )
        return ReciboParseResult(numero=numero, entidad=entidad, documentos=documentos)


@dataclass(frozen=True)
class AuditoriaEntry:
    """Linea de auditoria en memoria (no se escribe a archivo)."""

    nivel: str
    mensaje: str


class RecibosAuditoria:
    """Acumula entradas de auditoria y las devuelve; no genera archivos."""

    def __init__(self) -> None:
        self._entradas: list[AuditoriaEntry] = []

    def registrar(self, mensaje: str, nivel: str = "INFO") -> None:
        self._entradas.append(AuditoriaEntry(nivel=nivel, mensaje=mensaje))

    def entries(self) -> tuple[AuditoriaEntry, ...]:
        return tuple(self._entradas)

    def limpiar(self) -> None:
        self._entradas.clear()


class NomenclaturaHojas:
    """Reglas de nombres de hoja y codigos de concepto de los reportes."""

    def __init__(self, limite: int = LIMITE_NOMBRE_HOJA) -> None:
        self._limite = limite

    @property
    def limite(self) -> int:
        return self._limite

    def abreviar(self, nombre: str, mes: str) -> str:
        if len(nombre) <= self._limite:
            return nombre
        return nombre.replace(mes, MESES_ABREVIADOS[mes])

    def codigo_concepto(self, texto: object) -> str | None:
        coincidencia = re.search(r"(\d+(?:\.\d+)*)\s*$", str(texto))
        return coincidencia.group(1) if coincidencia else None


class ConsolidadoBuilder:
    """Consolidado de recibos: una hoja BASE con secciones de nominas y deudas."""

    NOMBRE_ARCHIVO = "Consolidado_Final.xlsx"

    def __init__(
        self,
        money: MoneyParser | None = None,
        reloj: Callable[[], datetime] | None = None,
    ) -> None:
        self._money = money or LegacyReciboMoneyParser()
        self._reloj = reloj or datetime.now

    def nombre_hoja(self) -> str:
        hoy = self._reloj()
        return f"BASE {MESES[hoy.month - 1]} {hoy.year}"

    def build(
        self, nominas: Sequence[ReciboDocument], deudas: Sequence[ReciboDocument]
    ) -> ExcelDocument:
        try:
            if not nominas and not deudas:
                raise ReportGenerationError(
                    "No hay recibos para consolidar: el libro quedaria vacio"
                )
            libro = Workbook()
            hoja = libro.active
            hoja.title = self.nombre_hoja()
            fila = 0
            for documentos, titulo in (
                (nominas, TITULO_NOMINAS),
                (deudas, TITULO_DEUDAS),
            ):
                if not documentos:
                    continue
                hoja.merge_cells(
                    start_row=fila + 1, start_column=1, end_row=fila + 1, end_column=4
                )
                celda_titulo = hoja.cell(row=fila + 1, column=1, value=titulo)
                celda_titulo.font = Font(bold=True)
                celda_titulo.fill = PatternFill(fill_type="solid", fgColor="FFEB3B")
                fila += 2
                for grupo in self._agrupar(documentos):
                    for columna, valor in enumerate(ENCABEZADO_CONSOLIDADO, start=1):
                        hoja.cell(row=fila + 1, column=columna, value=valor).font = Font(
                            bold=True
                        )
                    for indice, documento in enumerate(grupo):
                        renglon = fila + indice + 2
                        hoja.cell(
                            row=renglon, column=1, value=documento.concepto.etiqueta
                        )
                        hoja.cell(row=renglon, column=2, value=documento.entidad.nombre)
                        hoja.cell(row=renglon, column=3, value=documento.numero)
                        celda = hoja.cell(
                            row=renglon,
                            column=4,
                            value=float(self._money.parse(str(documento.monto))),
                        )
                        celda.number_format = "0.00"
                    fila += len(grupo) + 4
            for columna, ancho in ((1, 55), (2, 40), (3, 12), (4, 14)):
                hoja.column_dimensions[chr(64 + columna)].width = ancho
            buffer = BytesIO()
            libro.save(buffer)
        except ReportGenerationError:
            raise
        except Exception as exc:  # pragma: no cover - defensivo
            raise ReportGenerationError("No se pudo generar el consolidado de recibos") from exc
        return ExcelDocument(nombre=self.NOMBRE_ARCHIVO, contenido=buffer.getvalue())

    def _agrupar(
        self, documentos: Sequence[ReciboDocument]
    ) -> list[list[ReciboDocument]]:
        """Agrupa recibos consecutivos del mismo PDF (misma entidad y numero)."""
        grupos: list[list[ReciboDocument]] = []
        for documento in documentos:
            if grupos and (
                documento.numero == grupos[-1][0].numero
                and documento.entidad == grupos[-1][0].entidad
            ):
                grupos[-1].append(documento)
            else:
                grupos.append([documento])
        return grupos


@dataclass(frozen=True)
class ReciboResumen:
    """Resumen archivo -> numero de recibo."""

    archivo: str
    numero: str


class ResumidorRecibos:
    """Obtiene el numero de recibo de la primera pagina de cada PDF."""

    NO_ENCONTRADO = "NO ENCONTRADO"
    _PATRON_PRINCIPAL = re.compile(
        r"Recibo\s*(?:[Nn][°ºo]|N\^o|\$N\^\{o\}\$)?\s*(\d+)", re.IGNORECASE
    )
    _PATRON_RESPALDO = re.compile(r"Recibo[^\n\r]*?(\d{5,8})", re.IGNORECASE)

    def resumir(self, documentos: Sequence[PdfDocument]) -> list[ReciboResumen]:
        resumenes: list[ReciboResumen] = []
        for pdf in documentos:
            texto = pdf.texto_primera_pagina
            numero = self._buscar(texto)
            resumenes.append(ReciboResumen(archivo=pdf.nombre, numero=numero))
        return resumenes

    def _buscar(self, texto: str) -> str:
        coincidencia = self._PATRON_PRINCIPAL.search(texto)
        if coincidencia:
            return coincidencia.group(1)
        respaldo = self._PATRON_RESPALDO.search(texto)
        if respaldo:
            return respaldo.group(1)
        return self.NO_ENCONTRADO


@dataclass(frozen=True)
class _GrupoRetencion:
    busqueda: str
    nombre_normal: str
    nombre_deuda: str


@dataclass(frozen=True)
class _FilaConsolidado:
    concepto: str
    entidad: str
    recibo: str
    monto: float


class ReporteConsolidadoBuilder:
    """Agrega las hojas RETENCIONES y CASAS COMERCIALES al consolidado."""

    GRUPOS_RETENCIONES = (
        _GrupoRetencion(
            busqueda="S.S.O.",
            nombre_normal="S.S.O. (4%) 210",
            nombre_deuda="S.S.O. (DEUDA) 210.1.1",
        ),
        _GrupoRetencion(
            busqueda="PERDIDA",
            nombre_normal="PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.) 244",
            nombre_deuda="PERDIDA INVOLUNTARIA DEL EMPLEO (DEUDA) 244.2",
        ),
        _GrupoRetencion(
            busqueda="PENSION",
            nombre_normal="FONDO PENSIONES (JUBILACION) 301",
            nombre_deuda="FONDO PENSIONES JUBILACION (DEUDA) 301.2",
        ),
        _GrupoRetencion(
            busqueda="VIV",
            nombre_normal="FONDO DE AHORRO OBLIGATORIO PARA LA VIV 245",
            nombre_deuda="FONDO DE AHORRO OBLIGATORIO PARA LA VIV (DEUDA) 245.3",
        ),
    )

    CASAS_COMERCIALES = (
        ("COLEGIO BIOANALISTA", "COLEGIO BIOANALISTA 255"),
        ("NUTRIC", "COLEGIO NUTRIC. Y DIET. VZLA 286"),
        ("COLEGIO ENFERMERAS", "COLEGIO ENFERMERAS 224"),
        ("CARABOBO", "COLEGIO DE ENFERMERAS ESTADO CARABOBO 967"),
        ("MIRANDA", "COLEGIO DE ENFERMERA EDO. MIRANDA 603"),
        ("DESC", "DESC. DIA(S) NO LABORADO(S) 179"),
        ("INPRENFERMERA", "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA 370"),
        ("CAHORMINSAS", "CAHORMINSAS 212"),
        ("CAJA DE AHORRO", "CAJA DE AHORRO 907"),
        ("FUNERARIOS CAHORMINSAS", "SERVICIOS FUNERARIOS CAHORMINSAS 269"),
        ("CAEMINSA", "CAJA DE AHORRO CAEMINSA 238"),
        ("PRESTAMO CAJA", "PRESTAMO CAJA DE AHORRO 223"),
        ("UNICO DE TRAB", "SINDICATO UNICO DE TRAB.DE LA SALUD Y S 233"),
        ("OSBESS", "SINDICATO OSBESS ARAGUA 978"),
        ("ANESTECIOLOGOS", "SOCIEDAD ANESTECIOLOGOS 257"),
        ("SINBOPROENF", "SINBOPROENF 321"),
        ("SUNEP-SAS", "SUNEP-SAS 213"),
        ("FENASISTRASALUD", "FENASISTRASALUD 513"),
        ("SISTRASALUD", "SISTRASALUD 583"),
        ("SAPTRASEZ", "SAPTRASEZ 964"),
        ("CUOTA EXTRAORDINARIA SUNEP-SAS", "CUOTA EXTRAORDINARIA SUNEP-SAS 929"),
        ("SITRASSS-MIRANDA", "SITRASSS-MIRANDA 581"),
        ("ASUNAJUPENSAPROSO", "ASUNAJUPENSAPROSO 799"),
        ("TRIBUNAL DE PROTECCION", "TRIBUNAL DE PROTECCION DE NIÑOS, NIÑAS 807"),
        ("TRIBUNAL (PERMANENTE)", "TRIBUNAL (PERMANENTE) 215"),
    )

    def __init__(
        self,
        nomenclatura: NomenclaturaHojas | None = None,
        reloj: Callable[[], datetime] | None = None,
    ) -> None:
        self._nomenclatura = nomenclatura or NomenclaturaHojas()
        self._reloj = reloj or datetime.now

    def build(self, documento: ExcelDocument, hoja_base: str) -> ExcelDocument:
        try:
            libro = load_workbook(BytesIO(documento.contenido))
            if hoja_base not in libro.sheetnames:
                raise ReportGenerationError(
                    f"El consolidado no contiene la hoja {hoja_base}"
                )
            filas = self._leer_filas(libro[hoja_base])
            hoy = self._reloj()
            mes = MESES[hoy.month - 1]
            nombre_retenciones = self._nomenclatura.abreviar(
                f"RETENCIONES {mes} {hoy.year}", mes
            )
            nombre_casas = self._nomenclatura.abreviar(
                f"CASAS COMERCIALES {mes} {hoy.year}", mes
            )
            self._escribir_retenciones(libro, nombre_retenciones, filas)
            self._escribir_casas(libro, nombre_casas, filas)
            buffer = BytesIO()
            libro.save(buffer)
        except ReportGenerationError:
            raise
        except Exception as exc:  # pragma: no cover - defensivo
            raise ReportGenerationError("No se pudieron generar los reportes") from exc
        return ExcelDocument(nombre=documento.nombre, contenido=buffer.getvalue())

    def _leer_filas(self, hoja) -> list[_FilaConsolidado]:
        filas: list[_FilaConsolidado] = []
        for renglon in hoja.iter_rows(min_row=2, values_only=True):
            if renglon is None or len(renglon) < 4:
                continue
            concepto, entidad, recibo, monto = renglon[0], renglon[1], renglon[2], renglon[3]
            try:
                monto_numerico = float(monto)
            except (TypeError, ValueError):
                continue
            filas.append(
                _FilaConsolidado(
                    concepto=str(concepto or ""),
                    entidad=str(entidad or ""),
                    recibo=str(recibo or ""),
                    monto=monto_numerico,
                )
            )
        return filas

    def _escribir_retenciones(
        self, libro, nombre_hoja: str, filas: Sequence[_FilaConsolidado]
    ) -> None:
        hoja = self._hoja_nueva(libro, nombre_hoja)
        fila_actual = 1
        for grupo in self.GRUPOS_RETENCIONES:
            coincidencias = [
                fila for fila in filas if grupo.busqueda.lower() in fila.concepto.lower()
            ]
            if not coincidencias:
                continue
            inicio = fila_actual
            for fila in coincidencias:
                nombre_mostrar = (
                    grupo.nombre_deuda
                    if PALABRA_DEUDA in fila.concepto.upper()
                    else grupo.nombre_normal
                )
                hoja.cell(row=fila_actual, column=1, value=nombre_mostrar)
                hoja.cell(row=fila_actual, column=2, value=fila.entidad)
                hoja.cell(row=fila_actual, column=3, value=fila.monto)
                fila_actual += 1
            hoja.cell(row=fila_actual, column=2, value="TOTAL:")
            hoja.cell(
                row=fila_actual,
                column=3,
                value=f"=SUM(C{inicio}:C{fila_actual - 1})",
            )
            fila_actual += 4

    def _escribir_casas(
        self, libro, nombre_hoja: str, filas: Sequence[_FilaConsolidado]
    ) -> None:
        hoja = self._hoja_nueva(libro, nombre_hoja)
        fila_actual = 1
        for busqueda, nombre_mostrar in self.CASAS_COMERCIALES:
            codigo = self._nomenclatura.codigo_concepto(nombre_mostrar)
            if codigo is None:
                coincidencias = [
                    fila for fila in filas if busqueda.lower() in fila.concepto.lower()
                ]
            else:
                coincidencias = [
                    fila
                    for fila in filas
                    if self._nomenclatura.codigo_concepto(fila.concepto) == codigo
                ]
            if not coincidencias:
                continue
            inicio = fila_actual
            for fila in coincidencias:
                hoja.cell(row=fila_actual, column=1, value=nombre_mostrar)
                hoja.cell(row=fila_actual, column=2, value=fila.entidad)
                hoja.cell(row=fila_actual, column=3, value=fila.monto)
                fila_actual += 1
            hoja.cell(row=fila_actual, column=2, value="TOTAL:")
            hoja.cell(
                row=fila_actual,
                column=3,
                value=f"=SUM(C{inicio}:C{fila_actual - 1})",
            )
            fila_actual += 4

    def _hoja_nueva(self, libro, nombre_hoja: str):
        if nombre_hoja in libro.sheetnames:
            del libro[nombre_hoja]
        return libro.create_sheet(nombre_hoja)
