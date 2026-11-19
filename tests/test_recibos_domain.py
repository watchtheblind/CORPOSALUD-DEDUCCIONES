"""Tests del dominio de recibos."""

from __future__ import annotations

from datetime import datetime
from decimal import Decimal
from io import BytesIO

import pytest
from openpyxl import load_workbook

from core.errors import PdfExtractionError, ReportGenerationError
from modules.recibos import (
    AuditoriaEntry,
    CatalogoCentros,
    CatalogoConceptos,
    Centro,
    ConceptoCode,
    ConsolidadoBuilder,
    NomenclaturaHojas,
    ReciboDocument,
    ReciboParser,
    ReciboResumen,
    RecibosAuditoria,
    RecibosService,
    ReporteConsolidadoBuilder,
    ResumidorRecibos,
)
from modules.recibos.domain import (
    ENCABEZADO_CONSOLIDADO,
    TITULO_DEUDAS,
    TITULO_NOMINAS,
)
from tests.support import documento

HOJA_BASE = "BASE OCTUBRE 2026"


def reloj_fijo() -> datetime:
    return datetime(2026, 10, 31, 8, 0, 0)


def recibo(
    concepto: str,
    codigo: str,
    entidad: str = "SAANA",
    numero: str = "12345",
    monto: str = "100.00",
) -> ReciboDocument:
    return ReciboDocument(
        concepto=ConceptoCode(nombre=concepto, codigo=codigo),
        entidad=Centro(entidad),
        numero=numero,
        monto=Decimal(monto),
    )


PAGINA_NOMINA = "\n".join(
    [
        "Recibo Nº 12345",
        "S.A.A.N.A.",
        "TOTAL DE DEDUCCIONES POR CONCEPTO",
        "SISTRASALUD  1.234,56",
        "LEY S.S.O. (4%)  210,00",
        "Total Asignaciones  5.000,00",
    ]
)

PAGINA_DEUDA = "\n".join(
    [
        "Recibo Nº 777",
        "HOSPITAL LAS TEJERIAS",
        "S.S.O. (DEUDA)  50,00",
        "Total Asignaciones  0,00",
    ]
)

CASOS_CONCEPTO = [
    pytest.param("LEY S.S.O. (4%)  210,00", ("LEY S.S.O. (4%)", "210"), id="sso-ley"),
    pytest.param("S.S.O. (DEUDA)  50,00", ("S.S.O. (DEUDA)", "210.1.1"), id="sso-deuda"),
    pytest.param("FENASISTRASALUD  1.234,00", ("FENASISTRASALUD", "513"), id="fena"),
    pytest.param("SISTRASALUD  1.234,00", ("SISTRASALUD", "583"), id="sistrasalud"),
    pytest.param("sistrasalud  1.234,00", ("SISTRASALUD", "583"), id="minusculas"),
    pytest.param(
        "CUOTA EXTRAORDINARIA SUNEP-SAS  10,00",
        ("CUOTA EXTRAORDINARIA SUNEP-SAS", "929"),
        id="cuota",
    ),
    pytest.param("SUNEP-SAS  10,00", ("SUNEP-SAS", "213"), id="sunep"),
    pytest.param(
        "NO LEY CAHORMINSAS  20,00", ("NO LEY CAHORMINSAS", "212"), id="no-ley"
    ),
    pytest.param(
        "SERVICIOS FUNERARIOS CAHORMINSAS  3,00",
        ("SERVICIOS FUNERARIOS CAHORMINSAS", "269"),
        id="funerarios",
    ),
    pytest.param(
        "FONDO PENSIONES (JUBILACION)  45,00",
        ("FONDO PENSIONES (JUBILACION)", "301"),
        id="fondo-pensiones",
    ),
    pytest.param(
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIV  45,00",
        ("FONDO DE AHORRO OBLIGATORIO PARA LA VIV", "245.3"),
        id="faov-clave-duplicada",
    ),
    pytest.param(
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA  45,00",
        ("FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA", "245"),
        id="faov-completo",
    ),
    pytest.param(
        "DESC. DIA(S) NO LABORADO(S)  15,00",
        ("DESC. DIA(S) NO LABORADO(S)", "179"),
        id="descanso",
    ),
    pytest.param(
        "TRIBUNAL DE PROTECCION DE NIÑOS, NIÑAS Y ADOLESCENTES  5,00",
        ("TRIBUNAL DE PROTECCION DE NIÑOS, NIÑAS", "807"),
        id="tribunal",
    ),
    pytest.param("SUELDO BASICO  1.000,00", None, id="sin-concepto"),
    pytest.param("", None, id="vacio"),
]


class TestCatalogoConceptos:
    @pytest.mark.parametrize("linea,esperado", CASOS_CONCEPTO)
    def test_resuelve(self, linea, esperado):
        resultado = CatalogoConceptos().resolver(linea)
        if esperado is None:
            assert resultado is None
        else:
            assert resultado == ConceptoCode(nombre=esperado[0], codigo=esperado[1])

    def test_expone_la_etiqueta_que_se_escribe_en_excel(self):
        assert CatalogoConceptos().resolver("SISTRASALUD  1,00").etiqueta == "SISTRASALUD 583"

    def test_la_linea_debe_contener_la_clave(self):
        assert CatalogoConceptos().resolver("SINDICATO UNICO DE TRAB.DE LA SALUD  9,00") is None

    def test_catalogo_personalizado(self):
        catalogo = CatalogoConceptos({"CONCEPTO X": "999"})
        assert catalogo.resolver("CONCEPTO X  1,00") == ConceptoCode("CONCEPTO X", "999")
        assert catalogo.resolver("SISTRASALUD") is None


CASOS_CENTRO = [
    pytest.param(None, "CENTRO DESCONOCIDO", id="none"),
    pytest.param("", "CENTRO DESCONOCIDO", id="vacio"),
    pytest.param("  s.a.a.n.a.  ", "SAANA", id="saana"),
    pytest.param("s.a.g.e.r.", "SAGER", id="sager"),
    pytest.param("AMB. DE TURMERO", "AMBULATORIO DE TURMERO", id="ambulatorio"),
    pytest.param("HOSPITAL DEL SUR", "SOCIEDAD CIVIL HOSPITAL DEL SUR", id="hospital-sur"),
    pytest.param(
        "SOCIEDAD CIVIL CUALQUIERA",
        "SOCIEDAD CIVIL HOSPITAL DEL SUR",
        id="sociedad-civil",
    ),
    pytest.param(
        "S.A. HOSPITAL CENTRAL DE MARACAY", "HOSPITAL CENTRAL DE MARACAY", id="central"
    ),
    pytest.param("d.m.s. girardot", "D.M.S. GIRARDOT", id="girardot"),
    pytest.param("D.M.S CAMATAGUA-URDANETA", "D.M.S CAMATAGUA", id="camatagua"),
    pytest.param("HOSPITAL DESCONOCIDO", "HOSPITAL DESCONOCIDO", id="fallback"),
    pytest.param("  hospital las tejerias  ", "HOSPITAL LAS TEJERIAS", id="tejerias"),
]


class TestCatalogoCentros:
    @pytest.mark.parametrize("entrada,esperado", CASOS_CENTRO)
    def test_normaliza(self, entrada, esperado):
        assert CatalogoCentros().normalizar(entrada) == Centro(esperado)

    def test_catalogo_personalizado(self):
        catalogo = CatalogoCentros({"CENTRO X": ("ALIAS X",)})
        assert catalogo.normalizar("ALIAS X ZONA") == Centro("CENTRO X")
        assert catalogo.nombres() == ("CENTRO X",)


class TestReciboParser:
    def test_extrae_conceptos_de_nomina(self):
        parser = ReciboParser()
        resultado = parser.parse_documento(documento("r.pdf", PAGINA_NOMINA), es_deuda=False)
        assert resultado.numero == "12345"
        assert resultado.entidad == Centro("SAANA")
        assert [docu.concepto.codigo for docu in resultado.documentos] == ["583", "210"]
        assert resultado.documentos[0].monto == Decimal("1234.56")

    def test_parse_es_el_atajo_de_parse_documento(self):
        parser = ReciboParser()
        assert parser.parse(documento("r.pdf", PAGINA_NOMINA), False) == list(
            parser.parse_documento(documento("r.pdf", PAGINA_NOMINA), False).documentos
        )

    def test_se_detiene_al_llegar_a_asignaciones(self):
        pagina = "\n".join(
            [
                "TOTAL DE DEDUCCIONES POR CONCEPTO",
                "SISTRASALUD  1,00",
                "Total Asignaciones  5,00",
                "SISTRASALUD  9,00",
            ]
        )
        documentos = ReciboParser().parse(documento("r.pdf", pagina), es_deuda=False)
        assert [docu.monto for docu in documentos] == [Decimal("1.00")]

    def test_modo_deuda_no_necesita_marca_de_inicio(self):
        documentos = ReciboParser().parse(documento("r.pdf", PAGINA_DEUDA), es_deuda=True)
        assert [docu.concepto.codigo for docu in documentos] == ["210.1.1"]
        assert documentos[0].numero == "777"
        assert documentos[0].entidad == Centro("HOSPITAL LAS TEJERIAS")

    def test_modo_nomina_antes_de_la_marca_no_extrae(self):
        pagina = "SISTRASALUD  1,00\nTOTAL DE DEDUCCIONES POR CONCEPTO\nSISTRASALUD  2,00"
        documentos = ReciboParser().parse(documento("r.pdf", pagina), es_deuda=False)
        assert [docu.monto for docu in documentos] == [Decimal("2.00")]

    def test_sin_numero_de_recibo(self):
        pagina = "TOTAL DE DEDUCCIONES POR CONCEPTO\nSISTRASALUD  1,00"
        resultado = ReciboParser().parse_documento(documento("r.pdf", pagina), es_deuda=False)
        assert resultado.numero == "S/N"

    def test_sin_conceptos_devuelve_lista_vacia(self):
        resultado = ReciboParser().parse_documento(
            documento("r.pdf", "Recibo Nº 5\nSAANA"), es_deuda=False
        )
        assert resultado.documentos == ()
        assert resultado.numero == "5"

    def test_pdf_sin_paginas(self):
        with pytest.raises(PdfExtractionError):
            ReciboParser().parse_documento(documento("r.pdf"), es_deuda=False)

    def test_el_monto_usa_el_parser_de_recibos(self):
        pagina = "TOTAL DE DEDUCCIONES POR CONCEPTO\nSISTRASALUD  -1.234,56"
        documentos = ReciboParser().parse(documento("r.pdf", pagina), es_deuda=False)
        assert documentos[0].monto == Decimal("1234.56")

    def test_nombres_de_entidades(self):
        assert "SAANA" in ReciboParser().nombres_entidades()


class TestRecibosAuditoria:
    def test_acumula_entradas_en_memoria(self):
        auditoria = RecibosAuditoria()
        auditoria.registrar("uno")
        auditoria.registrar("dos", "ERROR")
        assert auditoria.entries() == (
            AuditoriaEntry(nivel="INFO", mensaje="uno"),
            AuditoriaEntry(nivel="ERROR", mensaje="dos"),
        )

    def test_se_puede_limpiar(self):
        auditoria = RecibosAuditoria()
        auditoria.registrar("uno")
        auditoria.limpiar()
        assert auditoria.entries() == ()


class TestNomenclaturaHojas:
    def test_limite_de_excel(self):
        assert NomenclaturaHojas().limite == 31

    @pytest.mark.parametrize(
        "texto,esperado",
        [
            pytest.param("SISTRASALUD 583", "583", id="simple"),
            pytest.param("S.S.O. (DEUDA) 210.1.1", "210.1.1", id="compuesto"),
            pytest.param("PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.) 244", "244", id="pie"),
            pytest.param("Concepto 12", "12", id="con-espacio"),
            pytest.param("CONCEPTO SIN CODIGO", None, id="sin-codigo"),
            pytest.param("", None, id="vacio"),
            pytest.param(None, None, id="none"),
        ],
    )
    def test_codigo_concepto(self, texto, esperado):
        assert NomenclaturaHojas().codigo_concepto(texto) == esperado

    def test_nombre_corto_no_cambia(self):
        assert (
            NomenclaturaHojas().abreviar("RETENCIONES ENERO 2026", "ENERO")
            == "RETENCIONES ENERO 2026"
        )

    def test_abrevia_el_mes_cuando_excede(self):
        assert (
            NomenclaturaHojas().abreviar("CASAS COMERCIALES SEPTIEMBRE 2026", "SEPTIEMBRE")
            == "CASAS COMERCIALES SEP. 2026"
        )

    def test_nombre_largo_sin_el_mes_queda_igual(self):
        assert NomenclaturaHojas().abreviar("Z" * 40, "ENERO") == "Z" * 40

    def test_mes_desconocido_provoca_keyerror(self):
        with pytest.raises(KeyError):
            NomenclaturaHojas().abreviar("X" * 40, "MES INEXISTENTE")


class TestConsolidadoBuilder:
    def test_nombre_de_hoja_con_reloj_inyectado(self):
        assert ConsolidadoBuilder(reloj=reloj_fijo).nombre_hoja() == HOJA_BASE

    def test_layout_de_secciones(self):
        nominas = [
            recibo("SISTRASALUD", "583"),
            recibo("LEY S.S.O. (4%)", "210"),
            recibo("CAHORMINSAS", "212", numero="999"),
        ]
        deudas = [recibo("S.S.O. (DEUDA)", "210.1.1", numero="777", monto="50.00")]
        documento_excel = ConsolidadoBuilder(reloj=reloj_fijo).build(nominas, deudas)
        libro = load_workbook(BytesIO(documento_excel.contenido))
        hoja = libro[HOJA_BASE]
        assert hoja["A1"].value == TITULO_NOMINAS
        assert [hoja.cell(row=3, column=c).value for c in range(1, 5)] == list(
            ENCABEZADO_CONSOLIDADO
        )
        assert hoja["A4"].value == "SISTRASALUD 583"
        assert hoja["B4"].value == "SAANA"
        assert hoja["C4"].value == "12345"
        assert hoja["D4"].value == 100.0
        assert hoja["A5"].value == "LEY S.S.O. (4%) 210"
        assert hoja["A9"].value == "Conceptos"
        assert hoja["A10"].value == "CAHORMINSAS 212"
        assert hoja["A14"].value == TITULO_DEUDAS
        assert hoja["A17"].value == "S.S.O. (DEUDA) 210.1.1"
        assert hoja["D17"].value == 50.0

    def test_agrupa_recibos_consecutivos_del_mismo_pdf(self):
        nominas = [
            recibo("SISTRASALUD", "583"),
            recibo("LEY S.S.O. (4%)", "210"),
            recibo("CAHORMINSAS", "212", numero="999"),
        ]
        documento_excel = ConsolidadoBuilder(reloj=reloj_fijo).build(nominas, [])
        libro = load_workbook(BytesIO(documento_excel.contenido))
        hoja = libro[HOJA_BASE]
        encabezados = [fila for fila in range(1, hoja.max_row + 1) if hoja.cell(row=fila, column=1).value == "Conceptos"]
        assert encabezados == [3, 9]

    def test_sin_recibos_no_genera_libro_vacio(self):
        with pytest.raises(ReportGenerationError):
            ConsolidadoBuilder(reloj=reloj_fijo).build([], [])

    def test_documento_con_nombre_y_bytes(self):
        documento_excel = ConsolidadoBuilder(reloj=reloj_fijo).build([recibo("X", "1")], [])
        assert documento_excel.nombre == "Consolidado_Final.xlsx"
        assert documento_excel.contenido.startswith(b"PK")

    def test_solo_nominas(self):
        documento_excel = ConsolidadoBuilder(reloj=reloj_fijo).build([recibo("X", "1")], [])
        libro = load_workbook(BytesIO(documento_excel.contenido))
        hoja = libro[HOJA_BASE]
        assert not any(
            hoja.cell(row=fila, column=1).value == TITULO_DEUDAS
            for fila in range(1, hoja.max_row + 1)
        )


class TestReporteConsolidadoBuilder:
    def _libro_final(self):
        nominas = [
            recibo("S.S.O. (4%)", "210", monto="210.00"),
            recibo("SISTRASALUD", "583", monto="1234.56"),
            recibo("FENASISTRASALUD", "513", monto="10.00"),
            recibo("FONDO DE AHORRO OBLIGATORIO PARA LA VIV", "245.3", monto="45.00"),
        ]
        deudas = [recibo("S.S.O. (DEUDA)", "210.1.1", monto="50.00")]
        consolidado = ConsolidadoBuilder(reloj=reloj_fijo).build(nominas, deudas)
        final = ReporteConsolidadoBuilder(reloj=reloj_fijo).build(consolidado, HOJA_BASE)
        return load_workbook(BytesIO(final.contenido)), final

    def test_hojas_del_reporte(self):
        libro, documento_excel = self._libro_final()
        assert libro.sheetnames == [
            HOJA_BASE,
            "RETENCIONES OCTUBRE 2026",
            "CASAS COMERCIALES OCTUBRE 2026",
        ]
        assert documento_excel.nombre == "Consolidado_Final.xlsx"

    def test_retenciones_normales_y_deuda(self):
        libro, _ = self._libro_final()
        hoja = libro["RETENCIONES OCTUBRE 2026"]
        assert hoja["A1"].value == "S.S.O. (4%) 210"
        assert hoja["C1"].value == 210.0
        assert hoja["A2"].value == "S.S.O. (DEUDA) 210.1.1"
        assert hoja["C2"].value == 50.0
        assert hoja["B3"].value == "TOTAL:"
        assert hoja["C3"].value == "=SUM(C1:C2)"

    def test_faov_se_reporta_como_normal(self):
        libro, _ = self._libro_final()
        hoja = libro["RETENCIONES OCTUBRE 2026"]
        etiquetas = [fila[0] for fila in hoja.iter_rows(values_only=True) if fila[0]]
        assert "FONDO DE AHORRO OBLIGATORIO PARA LA VIV 245" in etiquetas

    def test_casas_comerciales_por_codigo_sin_colisiones(self):
        libro, _ = self._libro_final()
        hoja = libro["CASAS COMERCIALES OCTUBRE 2026"]
        filas = {fila[0]: fila[2] for fila in hoja.iter_rows(values_only=True) if fila[0]}
        assert filas["SISTRASALUD 583"] == 1234.56
        assert filas["FENASISTRASALUD 513"] == 10.0

    def test_hoja_base_inexistente(self):
        consolidado = ConsolidadoBuilder(reloj=reloj_fijo).build([recibo("X", "1")], [])
        with pytest.raises(ReportGenerationError):
            ReporteConsolidadoBuilder(reloj=reloj_fijo).build(consolidado, "NO EXISTE")


class TestResumidorRecibos:
    @pytest.mark.parametrize(
        "texto,esperado",
        [
            pytest.param("Recibo Nº 12345678", "12345678", id="numero"),
            pytest.param("Recibo N° 998877", "998877", id="grado"),
            pytest.param("Recibo 1234567", "1234567", id="sin-simbolo"),
            pytest.param("Recibo: 1234567", "1234567", id="dos-puntos"),
            pytest.param("Sin recibo", "NO ENCONTRADO", id="sin-recibo"),
            pytest.param("", "NO ENCONTRADO", id="vacio"),
        ],
    )
    def test_numero_desde_la_primera_pagina(self, texto, esperado):
        resumen = ResumidorRecibos().resumir([documento("archivo.pdf", texto)])
        assert resumen == [ReciboResumen(archivo="archivo.pdf", numero=esperado)]

    def test_usa_respaldo_con_ruido_entre_medio(self):
        pagina = "Recibo (formato raro) 7654321"
        resumen = ResumidorRecibos().resumir([documento("a.pdf", pagina)])
        assert resumen[0].numero == "7654321"

    def test_total_recibos(self):
        resumen = ResumidorRecibos().resumir(
            [documento("a.pdf", "Recibo Nº 1"), documento("b.pdf", "nada")]
        )
        assert [item.archivo for item in resumen] == ["a.pdf", "b.pdf"]


class TestRecibosService:
    def test_extraer_nominas_y_auditar(self, logger_falso):
        servicio = RecibosService.crear(logger=logger_falso)
        documentos = servicio.extraer([documento("r.pdf", PAGINA_NOMINA)], es_deuda=False)
        assert [docu.concepto.codigo for docu in documentos] == ["583", "210"]
        assert servicio.auditoria()[0].mensaje.startswith("--- INICIO ESCANEO CARPETA NOMINAS")
        assert logger_falso.info_llamadas

    def test_extraer_deudas(self):
        servicio = RecibosService.crear()
        documentos = servicio.extraer([documento("r.pdf", PAGINA_DEUDA)], es_deuda=True)
        assert [docu.concepto.codigo for docu in documentos] == ["210.1.1"]

    def test_un_pdf_roto_no_aborta_el_lote(self):
        servicio = RecibosService.crear()
        documentos = servicio.extraer(
            [documento("roto.pdf"), documento("bueno.pdf", PAGINA_NOMINA)], es_deuda=False
        )
        assert len(documentos) == 2
        assert any(
            entrada.nivel == "ERROR" and "roto.pdf" in entrada.mensaje
            for entrada in servicio.auditoria()
        )

    def test_aviso_sin_conceptos(self):
        servicio = RecibosService.crear()
        servicio.extraer([documento("r.pdf", "Recibo Nº 42\nSAANA")], es_deuda=False)
        assert any(
            entrada.nivel == "AVISO" and "Recibo 42" in entrada.mensaje
            for entrada in servicio.auditoria()
        )

    def test_generar_reportes_end_to_end(self):
        servicio = RecibosService.crear()
        nominas = servicio.extraer([documento("r.pdf", PAGINA_NOMINA)], es_deuda=False)
        documento_excel = servicio.generar_reportes(nominas, [])
        libro = load_workbook(BytesIO(documento_excel.contenido))
        assert libro.sheetnames[0].startswith("BASE ")
        assert len(libro.sheetnames) == 3

    def test_generar_reportes_sin_recibos(self):
        with pytest.raises(ReportGenerationError):
            RecibosService.crear().generar_reportes([], [])

    def test_resumir(self):
        servicio = RecibosService.crear()
        assert servicio.resumir([documento("r.pdf", "Recibo Nº 12345")]) == [
            ReciboResumen(archivo="r.pdf", numero="12345")
        ]

    def test_auditar_cobertura_completa(self):
        parser = ReciboParser(centros=CatalogoCentros({"CENTRO X": ("CENTRO X",)}))
        servicio = RecibosService(parser=parser, consolidado=ConsolidadoBuilder(),
                                  reportes=ReporteConsolidadoBuilder(),
                                  resumidor=ResumidorRecibos())
        entradas = servicio.auditar_cobertura([Centro("CENTRO X")])
        assert entradas[-1].mensaje == "PROCESO: Todas las entidades fueron detectadas"

    def test_auditar_cobertura_reporta_faltantes(self):
        parser = ReciboParser(centros=CatalogoCentros({"CENTRO X": ("CENTRO X",)}))
        servicio = RecibosService(parser=parser, consolidado=ConsolidadoBuilder(),
                                  reportes=ReporteConsolidadoBuilder(),
                                  resumidor=ResumidorRecibos())
        entradas = servicio.auditar_cobertura([])
        mensajes = [entrada.mensaje for entrada in entradas]
        assert "--- ENTIDADES NO ENCONTRADAS ---" in mensajes
        assert "AUSENTE: CENTRO X" in mensajes
