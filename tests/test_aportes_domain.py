"""Tests del dominio de aportes patronales."""

from __future__ import annotations

from decimal import Decimal
from io import BytesIO

import pytest
from openpyxl import load_workbook

from core.config import ConfigProvider, Settings, StaticConfigProvider
from core.documents import PdfDocument, PdfFormat
from core.errors import ConfigError, PdfExtractionError, ReportGenerationError
from core.money import CommaDotMoneyParser
from modules.aportes_patronales import (
    AporteRow,
    AportesExcelBuilder,
    AportesPdfExtractor,
    AportesService,
    Centro,
    CentroResolver,
    Concepto,
    ConceptoNormalizer,
    FormatDetector,
    ListadoIndividualExtractor,
    TablaGruposExtractor,
)
from modules.aportes_patronales.domain import ENCABEZADO, HOJA_BASE
from tests.support import documento, paginas

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

PAGINA_TABLA_CONTINUACION = "\n".join(
    [
        "LISTADO DE DEDUCCIONES POR GRUPOS DE NÓMINA",
        "OBREROS CONTRATADOS  5  10,00  20,00",
    ]
)

PAGINA_TABLA_SIN_CONCEPTO = "\n".join(
    [
        "LISTADO DE DEDUCCIONES POR GRUPOS DE NÓMINA",
        "EMPLEADOS CONTRATADOS  3  1,00  2,00",
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


CABECERA_CENTRO = "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS SERVICIOS CENTRALES"


def documento_aporte(*textos: str | None) -> PdfDocument:
    """Documento cuyo primer renglon permite resolver el centro por ancla."""
    paginas_documento = list(textos)
    paginas_documento[0] = f"{CABECERA_CENTRO}\n{paginas_documento[0] or ''}"
    return documento("a.pdf", *paginas_documento)


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


FILAS_TABLA = [
    (
        CENTRO,
        "EMPLEADOS FIJOS",
        "DEDUCCION SSO EMPLEADOS FIJOS 1.234,56 2.345,67",
        "SSO",
        12,
        Decimal("1234.56"),
        Decimal("2345.67"),
    ),
    (
        CENTRO,
        "OBREROS FIJOS",
        "DEDUCCION SSO OBREROS FIJOS 500,00 600,00",
        "SSO",
        8,
        Decimal("500.00"),
        Decimal("600.00"),
    ),
    (
        CENTRO,
        "TOTAL GENERAL",
        "DEDUCCION SSO TOTAL 1.734,56 2.945,67",
        "SSO",
        20,
        Decimal("1734.56"),
        Decimal("2945.67"),
    ),
]

FILAS_LISTADO = [
    (
        CENTRO,
        "SIN GRUPO",
        "TOTAL DEDUCCIONES  2  69,12  0,00",
        "SSO",
        2,
        Decimal("69.12"),
        Decimal("0"),
    ),
]


CASOS_CONCEPTO = [
    pytest.param("FONDO PENSIONES (JUBILACION)", "FPJ", id="fondo-pensiones"),
    pytest.param("fondo pensiones jubilacion", "FPJ", id="fondo-pensiones-minusculas"),
    pytest.param("S.S.O. (4%)", "SSO", id="sso"),
    pytest.param("S.S.O. (DEUDA)", "SSO", id="sso-deuda"),
    pytest.param("RETENCION 4%", "SSO", id="sso-porcentaje"),
    pytest.param("PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.)", "PIE", id="pie"),
    pytest.param("PÉRDIDA INVOLUNTARIA DEL EMPLEO", "PIE", id="pie-con-tilde"),
    pytest.param(
        "SERVICIOS FUNERARIOS CAHORMINSAS",
        "SERVICIOS FUNERARIOS CAHORMINSAS",
        id="servicios-funerarios",
    ),
    pytest.param(
        "SERVICIOS FUNERARIOS CAHORMIN",
        "SERVICIOS FUNERARIOS CAHORMINSAS",
        id="servicios-funerarios-truncado",
    ),
    pytest.param("CAHORMINSAS", "CAHORMINSAS", id="cahorminsas"),
    pytest.param("CAHORMINSA", "CAHORMINSAS", id="cahorminsa"),
    pytest.param("CAJA DE AHORRO CAEMINSA", "CAEMINSAS", id="caeminsa"),
    pytest.param("CAEMINSA", "CAEMINSAS", id="caeminsa-solo"),
    pytest.param("PRESTAMO CAJA DE AHORRO", "CAJA DE AHORRO", id="prestamo"),
    pytest.param("CAJA DE AHORRO", "CAJA DE AHORRO", id="caja-ahorro"),
    pytest.param(
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA", "FAOV", id="faov"
    ),
    pytest.param("AHORRO VOLUNTARIO", "FAOV", id="cualquier-ahorro"),
    pytest.param("SINDICATO OSBESS ARAGUA", "SINDICATO OSBESS ARAGUA", id="sin-regla"),
    pytest.param(
        "  desc. dia(s) no laborado(s) ",
        "  desc. dia(s) no laborado(s) ",
        id="fallback-conserva-original",
    ),
    pytest.param("FONDO PENSIONES 4%", "FPJ", id="prioridad-fpj"),
    pytest.param("PERDIDA INVOLUNTARIA 4%", "SSO", id="prioridad-sso"),
]


class TestConceptoNormalizer:
    @pytest.mark.parametrize("entrada,esperado", CASOS_CONCEPTO)
    def test_normaliza(self, entrada, esperado):
        assert ConceptoNormalizer().normalize(entrada) == Concepto(esperado)

    def test_devuelve_value_object(self):
        assert isinstance(ConceptoNormalizer().normalize("S.S.O."), Concepto)


class TestFormatDetector:
    def test_detecta_tabla(self):
        assert FormatDetector().detect(paginas(PAGINA_TABLA)[0]) is PdfFormat.TABLA_GRUPOS

    def test_detecta_listado(self):
        assert (
            FormatDetector().detect(paginas(PAGINA_LISTADO)[0])
            is PdfFormat.LISTADO_INDIVIDUAL
        )

    def test_pagina_vacia_no_es_tabla(self):
        assert FormatDetector().detect(paginas("")[0]) is PdfFormat.LISTADO_INDIVIDUAL


class TestExtractorTabla:
    def test_filas(self):
        filas = TablaGruposExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(PAGINA_TABLA), Centro(CENTRO))
        assert simple(filas) == FILAS_TABLA

    def test_la_cantidad_es_int(self):
        filas = TablaGruposExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(PAGINA_TABLA), Centro(CENTRO))
        assert isinstance(filas[0].cantidad, int)

    def test_concepto_persiste_entre_paginas_del_bloque(self):
        filas = TablaGruposExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(PAGINA_TABLA, PAGINA_TABLA_CONTINUACION), Centro(CENTRO))
        assert len(filas) == 4
        assert filas[-1].grupo == "OBREROS CONTRATADOS"
        assert filas[-1].concepto == Concepto("SSO")

    def test_sin_concepto_no_emite_filas(self):
        filas = TablaGruposExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(PAGINA_TABLA_SIN_CONCEPTO), Centro(CENTRO))
        assert filas == []

    def test_grupo_corto_y_totales_se_descartan(self):
        pagina = "CONCEPTO: CAJA DE AHORRO\nAB  3  1,00  2,00\nCAJA  4  3,00  4,00\nTOTALES  9  9,00  9,00"
        filas = TablaGruposExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(pagina), Centro(CENTRO))
        assert [fila.grupo for fila in filas] == ["CAJA"]


class TestExtractorListado:
    def test_filas(self):
        filas = ListadoIndividualExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(PAGINA_LISTADO), Centro(CENTRO))
        assert simple(filas) == FILAS_LISTADO

    def test_la_cantidad_es_el_conteo_de_trabajadores(self):
        filas = ListadoIndividualExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(PAGINA_LISTADO), Centro(CENTRO))
        assert filas[0].cantidad == 2

    def test_dos_ultimos_montos_son_trabajador_y_empresa(self):
        pagina = "\n".join(
            [
                "LISTADO DE DEDUCCIONES CAHORMINSAS AL 31/10/2026",
                "12345678 PEREZ JUAN  1,00",
                "TOTAL DEDUCCIONES  1  15,00  25,00",
            ]
        )
        filas = ListadoIndividualExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(pagina), Centro(CENTRO))
        assert filas[0].trabajador == Decimal("15.00")
        assert filas[0].empresa == Decimal("25.00")
        assert filas[0].concepto == Concepto("CAHORMINSAS")

    def test_sin_concepto_en_listado_usa_la_linea_de_totales(self):
        pagina = "\n".join(
            [
                "LISTADO DE  AL 31/10/2026",
                "12345678 PEREZ JUAN  1,00",
                "TOTAL DEDUCCIONES  1  15,00  25,00",
            ]
        )
        filas = ListadoIndividualExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(pagina), Centro(CENTRO))
        assert filas[0].concepto == Concepto("1  15,00  25,00")

    def test_totales_emiten_fila_aunque_no_haya_listado(self):
        pagina = "\n".join(
            [
                "GOBIERNO BOLIVARIANO",
                "12345678 PEREZ JUAN  1,00",
                "TOTAL DEDUCCIONES  1  15,00  25,00",
            ]
        )
        filas = ListadoIndividualExtractor(
            ConceptoNormalizer(), CommaDotMoneyParser()
        ).extract(paginas(pagina), Centro(CENTRO))
        assert len(filas) == 1
        assert filas[0].cantidad == 0


class TestAportesPdfExtractor:
    def test_documento_de_un_formato(self):
        pdf = documento_aporte(PAGINA_TABLA)
        assert simple(AportesPdfExtractor().extract(pdf, ETIQUETAS)) == FILAS_TABLA

    def test_documento_mixto_respeta_el_orden_de_bloques(self):
        pdf = documento_aporte(PAGINA_LISTADO, PAGINA_TABLA)
        assert simple(AportesPdfExtractor().extract(pdf, ETIQUETAS)) == (
            FILAS_LISTADO + FILAS_TABLA
        )

    def test_el_estado_se_reinicia_entre_bloques(self):
        pdf = documento_aporte(PAGINA_LISTADO, PAGINA_TABLA_SIN_CONCEPTO)
        assert simple(AportesPdfExtractor().extract(pdf, ETIQUETAS)) == FILAS_LISTADO

    def test_paginas_sin_texto_no_rompen_el_bloque(self):
        pdf = documento_aporte(PAGINA_TABLA, None, PAGINA_TABLA_CONTINUACION)
        filas = AportesPdfExtractor().extract(pdf, ETIQUETAS)
        assert len(filas) == 4
        assert filas[-1].concepto == Concepto("SSO")

    def test_pdf_vacio(self):
        assert AportesPdfExtractor().extract(documento("a.pdf", ""), ETIQUETAS) == []

    def test_el_centro_sale_del_pdf_y_no_del_llamador(self):
        pdf = documento_aporte(PAGINA_TABLA)
        filas = AportesPdfExtractor().extract(pdf, ETIQUETAS)
        assert {fila.centro for fila in filas} == {Centro(CENTRO)}


ANCLA = "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS"
CASO_ANCLA_MISMA_LINEA = f"{ANCLA} ZONA NORTE\nMES: OCTUBRE 2026"
CASO_ANCLA_SIGUIENTE_LINEA = f"{ANCLA}\nSERVICIOS CENTRALES\nMES: OCTUBRE 2026"
CASO_ANCLA_CON_ENCABEZADO_EN_MEDIO = f"{ANCLA}\nMES: OCTUBRE 2026\nSERVICIOS CENTRALES"
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
CASO_TRANSFERENCIA = (
    "GOBIERNO BOLIVARIANO\n"
    "Transferencia de fondos para gastos de personal de HOSPITAL LAS TEJERIAS "
    "del (os) grupo(s) de nomina"
)


class TestCentroResolver:
    def test_ancla_en_la_misma_linea(self):
        pdf = documento("a.pdf", CASO_ANCLA_MISMA_LINEA)
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("ZONA NORTE")

    def test_ancla_y_centro_en_la_linea_siguiente(self):
        pdf = documento("a.pdf", CASO_ANCLA_SIGUIENTE_LINEA)
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("SERVICIOS CENTRALES")

    def test_ancla_con_encabezado_en_medio(self):
        pdf = documento("a.pdf", CASO_ANCLA_CON_ENCABEZADO_EN_MEDIO)
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("SERVICIOS CENTRALES")

    def test_solo_encabezados(self):
        pdf = documento("a.pdf", CASO_SOLO_ENCABEZADOS)
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("CENTRO DESCONOCIDO")

    def test_centro_en_la_linea_8(self):
        pdf = documento("a.pdf", CASO_CENTRO_EN_LA_LINEA_8)
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("HOSPITAL LAS TEJERIAS")

    def test_regla_de_transferencia_gana(self):
        pdf = documento("a.pdf", CASO_TRANSFERENCIA)
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("HOSPITAL LAS TEJERIAS")

    def test_busca_en_todas_las_paginas(self):
        pdf = documento(
            "a.pdf", "GOBIERNO BOLIVARIANO\nMES: OCTUBRE", "SERVICIOS CENTRALES"
        )
        assert CentroResolver().resolve(pdf, ETIQUETAS) == Centro("SERVICIOS CENTRALES")

    def test_pagina_sin_texto(self):
        assert CentroResolver().resolve(documento("a.pdf", None), ETIQUETAS) == Centro(
            "CENTRO DESCONOCIDO"
        )

    def test_documento_sin_paginas(self):
        assert CentroResolver().resolve(documento("a.pdf"), ETIQUETAS) == Centro(
            "CENTRO DESCONOCIDO"
        )

    def test_ancla_configurada_como_texto(self):
        pdf = documento("a.pdf", CASO_ANCLA_MISMA_LINEA)
        assert CentroResolver().resolve(pdf, {"ancla_centro": ANCLA}) == Centro("ZONA NORTE")

    def test_sin_anclas_usa_la_primera_linea_util(self):
        pdf = documento("a.pdf", "SERVICIOS CENTRALES")
        assert CentroResolver().resolve(pdf, {}) == Centro("SERVICIOS CENTRALES")

    def test_sin_anclas_puede_tomar_la_linea_de_concepto(self):
        """Comportamiento historico: sin anclas, la linea CONCEPTO pasa por centro."""
        pagina = "\n".join(
            [
                "LISTADO DE DEDUCCIONES POR GRUPOS DE NÓMINA",
                "CONCEPTO: S.S.O. (4%)",
                "CAJA  1  1,00  2,00",
            ]
        )
        filas = AportesPdfExtractor().extract(documento("a.pdf", pagina), {})
        assert filas[0].centro == Centro("CONCEPTO: S.S.O. (4%)")


def _filas_ejemplo() -> list[AporteRow]:
    return [
        AporteRow(Centro("CENTRO A"), "EMPLEADO FIJO", "REF A", Concepto("SSO"), 2,
                  Decimal("10.00"), Decimal("20.00")),
        AporteRow(Centro("CENTRO B"), "OBRERO FIJO", "REF B", Concepto("SSO"), 3,
                  Decimal("5.00"), Decimal("7.00")),
    ]


class TestAportesExcelBuilder:
    def test_hojas_del_libro(self):
        libro = load_workbook(BytesIO(AportesExcelBuilder().build(_filas_ejemplo()).contenido))
        assert libro.sheetnames == [
            HOJA_BASE,
            "APORTES TRABAJADOR",
            "APORTES EMPRESA",
            "CONSOLIDADO X GREMIO",
        ]

    def test_hoja_base_con_encabezado_y_datos(self):
        libro = load_workbook(BytesIO(AportesExcelBuilder().build(_filas_ejemplo()).contenido))
        hoja = libro[HOJA_BASE]
        assert [celda.value for celda in hoja[1]] == list(ENCABEZADO)
        assert [celda.value for celda in hoja[2]] == [
            "CENTRO A",
            "EMPLEADO FIJO",
            "REF A",
            "SSO",
            2,
            10.0,
            20.0,
        ]

    def test_resumen_de_aportes_trabajador(self):
        libro = load_workbook(BytesIO(AportesExcelBuilder().build(_filas_ejemplo()).contenido))
        hoja = libro["APORTES TRABAJADOR"]
        assert [hoja.cell(row=1, column=c).value for c in (1, 2)] == ["Centro", "SSO"]
        assert hoja.cell(row=2, column=1).value == "CENTRO A"
        assert hoja.cell(row=2, column=2).value == 10.0
        assert hoja.cell(row=3, column=2).value == 5.0
        assert hoja.cell(row=4, column=1).value == "TOTAL"
        assert hoja.cell(row=4, column=2).value == 15.0

    def test_resumen_de_aportes_empresa(self):
        libro = load_workbook(BytesIO(AportesExcelBuilder().build(_filas_ejemplo()).contenido))
        hoja = libro["APORTES EMPRESA"]
        assert hoja.cell(row=4, column=2).value == 27.0

    def test_resumen_por_gremio_usa_el_grupo(self):
        libro = load_workbook(BytesIO(AportesExcelBuilder().build(_filas_ejemplo()).contenido))
        hoja = libro["CONSOLIDADO X GREMIO"]
        assert hoja.cell(row=1, column=1).value == "Grupo"
        assert [hoja.cell(row=fila, column=1).value for fila in (2, 3)] == [
            "EMPLEADO FIJO",
            "OBRERO FIJO",
        ]
        assert hoja.cell(row=4, column=2).value == 27.0

    def test_documento_con_nombre_y_bytes(self):
        documento_excel = AportesExcelBuilder().build(_filas_ejemplo())
        assert documento_excel.nombre == "CONSOLIDADO APORTES PATRONALES.xlsx"
        assert documento_excel.contenido.startswith(b"PK")

    def test_sin_filas_solo_escribe_el_encabezado(self):
        libro = load_workbook(BytesIO(AportesExcelBuilder().build([]).contenido))
        assert libro[HOJA_BASE].max_row == 1


class TestAportesService:
    def test_procesa_documentos_y_devuelve_resultado(self, logger_falso):
        servicio = AportesService.crear(logger=logger_falso)
        resultado = servicio.procesar(
            [documento_aporte(PAGINA_TABLA)], StaticConfigProvider()
        )
        assert resultado.registros == 3
        assert resultado.documento.nombre == "CONSOLIDADO APORTES PATRONALES.xlsx"
        assert len(resultado.filas) == 3
        assert logger_falso.info_llamadas

    def test_usa_la_configuracion_inyectada(self):
        config = StaticConfigProvider(
            Settings(etiquetas_pdf={"ancla_centro": ["CENTRO X"]})
        )
        servicio = AportesService.crear()
        pdf = documento("a.pdf", "CENTRO X ZONA NORTE\n" + PAGINA_TABLA)
        resultado = servicio.procesar([pdf], config)
        assert resultado.filas[0].centro == Centro("ZONA NORTE")

    def test_sin_documentos_es_error_de_extraccion(self):
        with pytest.raises(PdfExtractionError):
            AportesService.crear().procesar([], StaticConfigProvider())

    def test_configuracion_rota_es_config_error(self):
        class ConfigRoto(ConfigProvider):
            def settings(self):
                raise RuntimeError("boom")

        with pytest.raises(ConfigError):
            AportesService.crear().procesar([documento("a.pdf", PAGINA_TABLA)], ConfigRoto())

    def test_configuracion_nula_es_config_error(self):
        class ConfigNula(ConfigProvider):
            def settings(self):
                return None

        with pytest.raises(ConfigError):
            AportesService.crear().procesar([documento("a.pdf", PAGINA_TABLA)], ConfigNula())

    def test_fallo_de_extraccion_se_traduce(self):
        class ExtractorRoto:
            def extract(self, pdf, etiquetas):
                raise RuntimeError("boom")

        servicio = AportesService(extractor=ExtractorRoto(), builder=AportesExcelBuilder())
        with pytest.raises(PdfExtractionError):
            servicio.procesar([documento("a.pdf", PAGINA_TABLA)], StaticConfigProvider())

    def test_fallo_del_builder_se_propaga_como_domain_error(self):
        class BuilderRoto:
            def build(self, filas):
                raise ReportGenerationError("boom")

        servicio = AportesService(extractor=AportesPdfExtractor(), builder=BuilderRoto())
        with pytest.raises(ReportGenerationError):
            servicio.procesar([documento("a.pdf", PAGINA_TABLA)], StaticConfigProvider())
