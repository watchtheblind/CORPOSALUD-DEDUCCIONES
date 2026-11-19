"""Tests del dominio de casas comerciales."""

from __future__ import annotations

from decimal import Decimal
from io import BytesIO

import pytest
from openpyxl import load_workbook

from core.config import ConfigProvider, Settings, StaticConfigProvider
from core.errors import ConfigError, PdfExtractionError
from modules.aportes_patronales import AporteRow as AporteRowAportes
from modules.aportes_patronales import Centro as CentroAportes
from modules.aportes_patronales import Concepto as ConceptoAportes
from modules.casas_comerciales import (
    AuditoriaExcelBuilder,
    CasasComercialesRow,
    CasasComercialesService,
    CasasPdfExtractor,
    Centro,
    CentroCasasResolver,
    ClasificadorGremio,
    Concepto,
    ConceptoCasasNormalizer,
    GremioCategoria,
    ResumenCategoria,
)
from modules.casas_comerciales.domain import ENCABEZADO, NOMBRE_ARCHIVO
from tests.support import documento

CENTRO = "SERVICIOS CENTRALES"
ETIQUETAS = Settings().etiquetas_pdf

PAGINA_TABLA = "\n".join(
    [
        "GOBIERNO BOLIVARIANO DEL ESTADO ARAGUA",
        "LISTADO DE DEDUCCIONES POR GRUPOS DE NÓMINA",
        "CONCEPTO: S.S.O. (4%)",
        "EMPLEADOS FIJOS  12  1.234,56  2.345,67",
        "OBREROS FIJOS  8  500,00  600,00",
        "TOTAL DEDUCCIONES  20  1.734,56  2.945,67",
    ]
)

PAGINA_LISTADO = "\n".join(
    [
        "LISTADO DE DEDUCCIONES S.S.O. (4%) AL 31/10/2026",
        "12345678 PEREZ JUAN  12,34",
        "87654321 GOMEZ ANA  56,78",
        "TOTAL DEDUCCIONES  2  69,12  0,00",
    ]
)

FILAS_TABLA = [
    (
        CENTRO,
        "EMPLEADOS FIJOS",
        "DEDUCCION S.S.O. (4%) EMPLEADOS FIJOS 1.234,56 2.345,67",
        "S.S.O. (4%)",
        12,
        Decimal("1234.56"),
        Decimal("2345.67"),
    ),
    (
        CENTRO,
        "OBREROS FIJOS",
        "DEDUCCION S.S.O. (4%) OBREROS FIJOS 500,00 600,00",
        "S.S.O. (4%)",
        8,
        Decimal("500.00"),
        Decimal("600.00"),
    ),
    (
        CENTRO,
        "TOTAL GENERAL",
        "DEDUCCION S.S.O. (4%) TOTAL 1.734,56 2.945,67",
        "S.S.O. (4%)",
        20,
        Decimal("1734.56"),
        Decimal("2945.67"),
    ),
]


CABECERA_CENTRO = "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS SERVICIOS CENTRALES"


def documento_casas(*textos: str | None):
    """Documento con ancla de centro en la primera pagina."""
    paginas_documento = list(textos)
    paginas_documento[0] = f"{CABECERA_CENTRO}\n{paginas_documento[0] or ''}"
    return documento("c.pdf", *paginas_documento)


def simple(filas) -> list[tuple]:
    return [
        (
            fila.centro.nombre,
            fila.grupo,
            fila.referencia,
            fila.concepto.codigo,
            fila.cantidad,
            fila.trabajador,
            fila.empresa,
        )
        for fila in filas
    ]


CASOS_GREMIO = [
    pytest.param("EMPLEADO FIJO", GremioCategoria.EMPLEADOS_FIJOS, id="empleado-fijo"),
    pytest.param("EMPLEADO FIJO (M)", GremioCategoria.EMPLEADOS_FIJOS, id="con-sufijo"),
    pytest.param("  empleado fijo  ", GremioCategoria.EMPLEADOS_FIJOS, id="normaliza-texto"),
    pytest.param("EMPLEADO CONTRATADO", GremioCategoria.EMPLEADOS_CONTRATADOS, id="contratado"),
    pytest.param("OBRERO CONTRATADO", GremioCategoria.OBREROS_CONTRATADOS, id="obrero-contratado"),
    pytest.param("OBRERO FIJO", GremioCategoria.OBREROS_FIJOS, id="obrero-fijo"),
    pytest.param("PERSONAL JUBILADO OBRERO", GremioCategoria.OBREROS_FIJOS, id="jubilado-obrero"),
    pytest.param("PERSONAL JUBILADO MÉDICOS", GremioCategoria.EMPLEADOS_FIJOS, id="jubilado-medicos"),
    pytest.param(
        "PERSONAL JUBILADO MEDICOS", GremioCategoria.EMPLEADOS_FIJOS, id="jubilado-sin-tilde"
    ),
    pytest.param(
        "PERSONAL JUBILADO ADMINISTRATIVO OTROS GREMIOS",
        GremioCategoria.EMPLEADOS_FIJOS,
        id="jubilado-administrativo",
    ),
    pytest.param(
        "PERSONAL PENSIÓN DE SOBREVIVIENTE",
        GremioCategoria.EMPLEADOS_FIJOS,
        id="pension-sobreviviente",
    ),
    pytest.param(
        "PERSONAL PENSION DE INVALIDEZ ADMINISTRATIVOS Y OTROS",
        GremioCategoria.EMPLEADOS_FIJOS,
        id="pension-invalidez",
    ),
    pytest.param("PENSIONADO", GremioCategoria.EMPLEADOS_FIJOS, id="contiene-pension"),
    pytest.param("JUBILADO CUALQUIERA", GremioCategoria.EMPLEADOS_FIJOS, id="contiene-jubilado"),
    pytest.param(
        "EMPLEADO FIJO CONTRATADO", GremioCategoria.EMPLEADOS_FIJOS, id="subcadena-literal"
    ),
    pytest.param(
        "OBRERO CONTRATADO Y EMPLEADO CONTRATADO",
        GremioCategoria.EMPLEADOS_CONTRATADOS,
        id="orden-empleado-primero",
    ),
    pytest.param("EMPLEADOS FIJOS", None, id="plural-empleados-fijos"),
    pytest.param("OBREROS FIJOS", None, id="plural-obreros-fijos"),
    pytest.param("OBREROS CONTRATADOS", None, id="plural-obreros-contratados"),
    pytest.param("PERSONAL JUBILADO OBREROS", None, id="plural-jubilado-obreros"),
    pytest.param("APRENDICES", None, id="aprendices"),
    pytest.param("", None, id="vacio"),
    pytest.param(None, None, id="none"),
]


class TestClasificadorGremio:
    @pytest.mark.parametrize("grupo,esperado", CASOS_GREMIO)
    def test_clasifica(self, grupo, esperado):
        assert ClasificadorGremio().clasificar(grupo) is esperado

    def test_las_etiquetas_son_las_historicas(self):
        assert [categoria.value for categoria in GremioCategoria] == [
            "Empleados Fijos",
            "Obreros Fijos",
            "Empleados Contratados",
            "Obreros Contratados",
        ]


CASOS_NORMALIZADOR = [
    pytest.param("SERVICIOS FUNERARIOS CAHORMINSAS", "CAHORMINSAS", id="funerarios"),
    pytest.param("servicios funerarios cahorminsas", "CAHORMINSAS", id="minusculas"),
    pytest.param("CAJA DE AHORRO CAEMINSA", "CAEMINSA", id="caeminsa"),
    pytest.param("CAEMINSA", "CAEMINSA", id="caeminsa-solo"),
    pytest.param("S.S.O. (4%)", "S.S.O. (4%)", id="sin-regla-sso"),
    pytest.param(
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA",
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA",
        id="sin-regla-faov",
    ),
    pytest.param("  sindicato   osbess  ", "SINDICATO   OSBESS", id="solo-strip-exterior"),
    pytest.param("", "", id="vacio"),
]


class TestConceptoCasasNormalizer:
    @pytest.mark.parametrize("entrada,esperado", CASOS_NORMALIZADOR)
    def test_normaliza(self, entrada, esperado):
        assert ConceptoCasasNormalizer().normalize(entrada) == Concepto(esperado)


class TestCasasPdfExtractor:
    def test_formato_tabla(self):
        filas = CasasPdfExtractor().extract(documento_casas(PAGINA_TABLA), ETIQUETAS)
        assert simple(filas) == FILAS_TABLA

    def test_formato_listado(self):
        filas = CasasPdfExtractor().extract(documento_casas(PAGINA_LISTADO), ETIQUETAS)
        assert simple(filas) == [
            (
                CENTRO,
                "SIN GRUPO",
                "TOTAL DEDUCCIONES  2  69,12  0,00",
                "2  69,12  0,00",
                2,
                Decimal("69.12"),
                Decimal("0"),
            )
        ]

    def test_el_concepto_del_listado_sale_de_la_linea_de_totales(self):
        pagina = "\n".join(
            [
                "LISTADO DE DEDUCCIONES S.S.O. (4%) AL 31/10/2026",
                "12345678 PEREZ JUAN  12,34",
                "TOTAL DEDUCCIONES S.S.O. (4%)  2  69,12  0,00",
            ]
        )
        filas = CasasPdfExtractor().extract(documento_casas(pagina), ETIQUETAS)
        assert filas[0].concepto == Concepto("S.S.O.")
        assert filas[0].cantidad == 1

    def test_el_formato_se_decide_por_la_primera_pagina(self):
        """A diferencia de aportes, aqui no hay deteccion pagina por pagina."""
        pdf = documento_casas(PAGINA_LISTADO, PAGINA_TABLA)
        filas = CasasPdfExtractor().extract(pdf, ETIQUETAS)
        assert len(filas) == 2
        assert [fila.concepto.codigo for fila in filas] == [
            "2  69,12  0,00",
            "20  1.734,56  2.945,67",
        ]

    def test_documento_sin_paginas(self):
        assert CasasPdfExtractor().extract(documento("c.pdf"), ETIQUETAS) == []

    def test_from_aporte_convierte_sin_importar_el_modulo(self):
        fila_aportes = AporteRowAportes(
            centro=CentroAportes("CENTRO A"),
            grupo="EMPLEADO FIJO",
            referencia="REF",
            concepto=ConceptoAportes("SSO"),
            cantidad=2,
            trabajador=Decimal("10.00"),
            empresa=Decimal("20.00"),
        )
        fila = CasasComercialesRow.from_aporte(fila_aportes)
        assert fila == CasasComercialesRow(
            centro=Centro("CENTRO A"),
            grupo="EMPLEADO FIJO",
            referencia="REF",
            concepto=Concepto("SSO"),
            cantidad=2,
            trabajador=Decimal("10.00"),
            empresa=Decimal("20.00"),
        )


ANCLA = "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS"
CASO_ANCLA_MISMA_LINEA = f"{ANCLA} ZONA NORTE\nMES: OCTUBRE 2026"
CASO_ANCLA_SIGUIENTE_LINEA = f"{ANCLA}\nSERVICIOS CENTRALES\nMES: OCTUBRE 2026"
CASO_SOLO_ENCABEZADOS = "\n".join(
    [
        "GOBIERNO BOLIVARIANO DEL ESTADO ARAGUA",
        "CORPORACION DE SALUD DEL ESTADO ARAGUA",
        "MES: OCTUBRE 2026",
        "PÁGINA 1",
    ]
)
CASO_CENTRO_EN_LA_LINEA_8 = "\n".join(
    [
        "GOBIERNO BOLIVARIANO DEL ESTADO ARAGUA",
        "LISTADO DE DEDUCCIONES",
        "MES: OCTUBRE",
        "AÑO: 2026",
        "PÁGINA 1",
        "RECIBO: 12345",
        "FECHA: 31/10/2026",
        "HOSPITAL LAS TEJERIAS",
    ]
)


class TestCentroCasasResolver:
    def test_ancla_en_la_misma_linea(self):
        pdf = documento("c.pdf", CASO_ANCLA_MISMA_LINEA)
        assert CentroCasasResolver().resolve(pdf, ETIQUETAS) == Centro("ZONA NORTE")

    def test_ancla_y_centro_en_la_siguiente(self):
        pdf = documento("c.pdf", CASO_ANCLA_SIGUIENTE_LINEA)
        assert CentroCasasResolver().resolve(pdf, ETIQUETAS) == Centro("SERVICIOS CENTRALES")

    def test_centro_en_la_linea_8_no_se_encuentra(self):
        pdf = documento("c.pdf", CASO_CENTRO_EN_LA_LINEA_8)
        assert CentroCasasResolver().resolve(pdf, ETIQUETAS) == Centro("CENTRO DESCONOCIDO")

    def test_solo_mira_la_primera_pagina(self):
        pdf = documento(
            "c.pdf", "GOBIERNO BOLIVARIANO\nMES: OCTUBRE", "SERVICIOS CENTRALES"
        )
        assert CentroCasasResolver().resolve(pdf, ETIQUETAS) == Centro(
            "CENTRO DESCONOCIDO"
        )

    def test_solo_encabezados(self):
        pdf = documento("c.pdf", CASO_SOLO_ENCABEZADOS)
        assert CentroCasasResolver().resolve(pdf, ETIQUETAS) == Centro("CENTRO DESCONOCIDO")

    def test_documento_sin_paginas(self):
        assert CentroCasasResolver().resolve(documento("c.pdf"), ETIQUETAS) == Centro(
            "CENTRO DESCONOCIDO"
        )

    def test_divergencia_con_el_resolver_de_aportes(self):
        from modules.aportes_patronales import CentroResolver as ResolverAportes

        pdf = documento("c.pdf", CASO_CENTRO_EN_LA_LINEA_8)
        assert CentroCasasResolver().resolve(pdf, ETIQUETAS).nombre == "CENTRO DESCONOCIDO"
        assert ResolverAportes().resolve(pdf, ETIQUETAS).nombre == "HOSPITAL LAS TEJERIAS"


def _filas_auditoria() -> list[CasasComercialesRow]:
    return [
        CasasComercialesRow(
            Centro("CENTRO A"), "EMPLEADO FIJO", "REF 1", Concepto("CAHORMINSAS"), 2,
            Decimal("10.00"), Decimal("20.00"),
        ),
        CasasComercialesRow(
            Centro("CENTRO B"), "OBRERO CONTRATADO", "REF 2", Concepto("CAHORMINSAS"), 3,
            Decimal("5.00"), Decimal("7.00"),
        ),
        CasasComercialesRow(
            Centro("CENTRO C"), "OBREROS FIJOS", "REF 3", Concepto("CAHORMINSAS"), 5,
            Decimal("1.00"), Decimal("1.00"),
        ),
    ]


class TestAuditoriaExcelBuilder:
    def test_hojas(self):
        libro = load_workbook(BytesIO(AuditoriaExcelBuilder().build(_filas_auditoria()).contenido))
        assert libro.sheetnames == ["BASE", "CAHORMINSAS"]
        assert NOMBRE_ARCHIVO == "_AUDITORIA_CASAS_COMERCIALES.xlsx"

    def test_encabezado_de_la_base(self):
        libro = load_workbook(BytesIO(AuditoriaExcelBuilder().build(_filas_auditoria()).contenido))
        assert [celda.value for celda in libro["BASE"][1]] == list(ENCABEZADO)

    def test_resumen_por_categoria_y_total(self):
        libro = load_workbook(BytesIO(AuditoriaExcelBuilder().build(_filas_auditoria()).contenido))
        hoja = libro["CAHORMINSAS"]
        inicio = 3 + 3
        etiquetas = [hoja.cell(row=inicio + i, column=1).value for i in range(4)]
        cantidades = [hoja.cell(row=inicio + i, column=2).value for i in range(4)]
        assert etiquetas == [
            "Empleados Fijos",
            "Obreros Fijos",
            "Empleados Contratados",
            "Obreros Contratados",
        ]
        assert cantidades == [2, 0, 0, 3]
        assert hoja.cell(row=inicio + 4, column=1).value == "TOTAL"
        assert hoja.cell(row=inicio + 4, column=2).value == 5

    def test_el_total_no_pisa_la_ultima_categoria(self):
        libro = load_workbook(BytesIO(AuditoriaExcelBuilder().build(_filas_auditoria()).contenido))
        hoja = libro["CAHORMINSAS"]
        assert hoja.cell(row=9, column=1).value == "Obreros Contratados"
        assert hoja.cell(row=9, column=2).value == 3

    def test_resumen_por_categoria(self):
        resumen = AuditoriaExcelBuilder().resumen_por_categoria(_filas_auditoria())
        assert resumen == [
            ResumenCategoria(GremioCategoria.EMPLEADOS_FIJOS, 2),
            ResumenCategoria(GremioCategoria.OBREROS_FIJOS, 0),
            ResumenCategoria(GremioCategoria.EMPLEADOS_CONTRATADOS, 0),
            ResumenCategoria(GremioCategoria.OBREROS_CONTRATADOS, 3),
        ]

    def test_nombre_de_hoja_se_recorta_y_sanea(self):
        filas = [
            CasasComercialesRow(
                Centro("C"), "EMPLEADO FIJO", "R",
                Concepto("CONCEPTO: MUY LARGO / CON CARACTERES INVALIDOS XYZ"),
                1, Decimal("1"), Decimal("1"),
            )
        ]
        libro = load_workbook(BytesIO(AuditoriaExcelBuilder().build(filas).contenido))
        hoja = libro.sheetnames[1]
        assert ":" not in hoja and "/" not in hoja
        assert len(hoja) <= 31

    def test_nombres_de_hoja_repetidos_se_desambiguan(self):
        concepto_largo = "X" * 40
        filas = [
            CasasComercialesRow(Centro("C"), "EMPLEADO FIJO", "R", Concepto(concepto_largo + "A"),
                                1, Decimal("1"), Decimal("1")),
            CasasComercialesRow(Centro("C"), "EMPLEADO FIJO", "R", Concepto(concepto_largo + "B"),
                                1, Decimal("1"), Decimal("1")),
        ]
        libro = load_workbook(BytesIO(AuditoriaExcelBuilder().build(filas).contenido))
        assert len(libro.sheetnames) == 3
        assert libro.sheetnames[1] != libro.sheetnames[2]

    def test_documento_con_nombre_y_bytes(self):
        documento_excel = AuditoriaExcelBuilder().build(_filas_auditoria())
        assert documento_excel.nombre == NOMBRE_ARCHIVO
        assert documento_excel.contenido.startswith(b"PK")


class TestCasasComercialesService:
    def test_procesa_y_devuelve_excel(self, logger_falso):
        servicio = CasasComercialesService.crear(logger=logger_falso)
        documento_excel = servicio.procesar(
            [documento_casas(PAGINA_TABLA)], StaticConfigProvider()
        )
        assert documento_excel.nombre == NOMBRE_ARCHIVO
        assert logger_falso.info_llamadas

    def test_sin_documentos_es_error_de_extraccion(self):
        with pytest.raises(PdfExtractionError):
            CasasComercialesService.crear().procesar([], StaticConfigProvider())

    def test_configuracion_rota(self):
        class ConfigRoto(ConfigProvider):
            def settings(self):
                raise RuntimeError("boom")

        with pytest.raises(ConfigError):
            CasasComercialesService.crear().procesar(
                [documento("c.pdf", PAGINA_TABLA)], ConfigRoto()
            )
