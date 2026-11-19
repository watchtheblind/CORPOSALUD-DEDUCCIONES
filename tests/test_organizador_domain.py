"""Tests del dominio y del servicio del organizador."""

from __future__ import annotations

from pathlib import Path

import pymupdf
import pytest

from core.config import Settings
from core.errors import ConfigError, PdfExtractionError
from modules.organizador import (
    BuscadorPdfs,
    CasoUso,
    CategoriaPdf,
    EvaluacionPdf,
    LectorPdfPyMuPDF,
    OrganizadorService,
    PdfEvaluador,
    ReglaClasificacion,
    ReglasCasos,
    RenombradorArchivos,
    TextoNormalizado,
    UnificadorPdfs,
)
from modules.organizador.domain import (
    PALABRA_DEUDA,
    TEXTOS_APORTES,
    TEXTOS_RECIBOS,
)
from tests.support import crear_pdf, documento

MAPA = {"CAHORMINSAS": "CAHORMINSAS", "CAJA DE AHORRO CAEMINSA": "CAEMINSAS"}


def evaluador(mapa=MAPA, **extra) -> PdfEvaluador:
    regla = ReglaClasificacion(
        caso=CasoUso.CASAS_COMERCIALES,
        textos_requeridos=tuple(mapa),
        mapa_conceptos=dict(mapa),
        **extra,
    )
    return PdfEvaluador(regla)


class TestTextoNormalizado:
    @pytest.mark.parametrize(
        "entrada,esperado",
        [
            pytest.param("PÉRDIDA", "PERDIDA", id="acento"),
            pytest.param("NIÑA", "NINA", id="enye"),
            pytest.param("ÜBER", "UBER", id="dieresis"),
            pytest.param("ç", "c", id="cedilla"),
            pytest.param("ÁÉÍÓÚ", "AEIOU", id="vocales"),
            pytest.param("", "", id="vacio"),
        ],
    )
    def test_sin_acentos(self, entrada, esperado):
        assert TextoNormalizado.sin_acentos(entrada) == esperado

    def test_colapsar_quita_guiones_de_corte(self):
        assert TextoNormalizado.colapsar("CAJA DE AHO-\nRRO") == "CAJA DE AHORRO"

    def test_colapsar_quita_guiones_con_espacio(self):
        assert TextoNormalizado.colapsar("PERDIDA - INVOLUNTARIA") == "PERDIDA INVOLUNTARIA"

    def test_colapsar_mayusculas_y_espacios(self):
        assert TextoNormalizado.colapsar("  caixa   de ahorro ") == "CAIXA DE AHORRO"


CASOS_SANITIZAR = [
    pytest.param("SERVICIOS CENTRALES.pdf", 90, "SERVICIOS CENTRALES.pdf", id="sin-cambios"),
    pytest.param("ÑANDÚ.pdf", 90, "NANDU.pdf", id="quita-acentos"),
    pytest.param('a/b\\c:d*e?f"g<h>i|j', 90, "a_b_c_d_e_f_g_h_i_j", id="invalidos"),
    pytest.param("nombre.", 90, "nombre", id="punto-final"),
    pytest.param("...", 90, "ARCHIVO_PROCESADO", id="solo-puntos"),
    pytest.param("", 90, "ARCHIVO_PROCESADO", id="vacio"),
    pytest.param("  nombre  ", 90, "nombre", id="espacios"),
    pytest.param(".oculto.pdf", 90, "oculto.pdf", id="punto-inicial"),
    pytest.param("X" * 100, 90, "X" * 90, id="recorta"),
    pytest.param("X" * 20, 10, "X" * 10, id="max_len"),
    pytest.param("B" * 89 + " CCCC", 90, "B" * 89, id="strip-tras-recortar"),
    pytest.param("///", 90, "___", id="guiones-bajos"),
]


class TestRenombradorArchivos:
    @pytest.mark.parametrize("entrada,max_len,esperado", CASOS_SANITIZAR)
    def test_sanitiza(self, entrada, max_len, esperado):
        assert RenombradorArchivos().sanitizar(entrada, max_len) == esperado

    def test_maximo_por_defecto(self):
        assert len(RenombradorArchivos().sanitizar("Y" * 200)) == 90


class TestPdfEvaluador:
    def test_concepto_valido_normal(self):
        evaluacion = evaluador().evaluar(
            documento("x.pdf", "TOTAL DE DEDUCCIONES", "CAHORMINSAS  10,00")
        )
        assert evaluacion.valido is True
        assert evaluacion.categoria is CategoriaPdf.NORMAL
        assert evaluacion.conceptos == ("CAHORMINSAS",)
        assert evaluacion.detalle == "Concepto(s): [CAHORMINSAS]"

    def test_con_deuda(self):
        evaluacion = evaluador().evaluar(documento("x.pdf", "CAHORMINSAS EN DEUDA"))
        assert evaluacion.categoria is CategoriaPdf.DEUDA
        assert evaluacion.detalle == "Concepto(s): [CAHORMINSAS] Y 'DEUDA'"

    def test_sin_coincidencias(self):
        evaluacion = evaluador().evaluar(documento("x.pdf", "SUELDO BASICO"))
        assert evaluacion.valido is False
        assert evaluacion.categoria is CategoriaPdf.NINGUNA
        assert evaluacion.conceptos == ()
        assert evaluacion.detalle == "No coincide en sus 1 páginas."

    def test_ordena_por_orden_del_mapa(self):
        mapa = {"CAHORMINSAS": "CONC_B", "CAJA DE AHORRO CAEMINSA": "CONC_A"}
        evaluacion = evaluador(mapa).evaluar(
            documento("x.pdf", "CAJA DE AHORRO CAEMINSA Y CAHORMINSAS")
        )
        assert evaluacion.conceptos == ("CONC_B", "CONC_A")

    def test_conceptos_repetidos_una_sola_vez(self):
        mapa = {"CAHORMINSAS": "CAEMINSAS", "CAJA DE AHORRO CAEMINSA": "CAEMINSAS"}
        evaluacion = evaluador(mapa).evaluar(
            documento("x.pdf", "CAHORMINSAS CAJA DE AHORRO CAEMINSA")
        )
        assert evaluacion.conceptos == ("CAEMINSAS",)

    def test_normaliza_el_texto_del_pdf(self):
        evaluacion = evaluador().evaluar(
            documento("x.pdf", "caja de aho-\nrro caeminsa")
        )
        assert evaluacion.conceptos == ("CAEMINSAS",)

    def test_normaliza_el_mapa_con_acentos_y_espacios(self):
        mapa = {"  Pérdida   Involuntaria ": "PIE"}
        evaluacion = evaluador(mapa).evaluar(documento("x.pdf", "PÉRDIDA INVOLUNTARIA"))
        assert evaluacion.conceptos == ("PIE",)

    def test_regla_no_ley_por_nombre(self):
        evaluacion = evaluador(evaluar_no_ley=True).evaluar(
            documento("ALGO NO LEY.pdf", "SIN CONCEPTOS")
        )
        assert evaluacion.valido is True
        assert evaluacion.categoria is CategoriaPdf.NORMAL
        assert evaluacion.conceptos == ("NO LEY",)

    def test_regla_no_ley_desactivada(self):
        evaluacion = evaluador(evaluar_no_ley=False).evaluar(
            documento("ALGO NO LEY.pdf", "SIN CONCEPTOS")
        )
        assert evaluacion.valido is False

    def test_no_ley_con_deuda(self):
        evaluacion = evaluador(evaluar_no_ley=True).evaluar(
            documento("X NO LEY.pdf", "MONTO EN DEUDA")
        )
        assert evaluacion.categoria is CategoriaPdf.DEUDA
        assert evaluacion.detalle == "Concepto(s): [NO LEY] Y 'DEUDA'"

    def test_busca_en_la_ultima_pagina(self):
        evaluacion = evaluador().evaluar(
            documento("x.pdf", "PORTADA", "CAHORMINSAS")
        )
        assert evaluacion.valido is True

    def test_pdf_sin_paginas(self):
        evaluacion = evaluador().evaluar(documento("x.pdf"))
        assert evaluacion.valido is False
        assert evaluacion.detalle == "PDF sin páginas o corrupto."

    def test_usa_el_origen_del_documento(self):
        evaluacion = evaluador().evaluar(
            documento("logico.pdf", "CAHORMINSAS", origen=Path("/tmp/real.pdf"))
        )
        assert evaluacion.origen == Path("/tmp/real.pdf")

    def test_la_palabra_deuda_es_configurable_por_constante(self):
        assert PALABRA_DEUDA == "DEUDA"


class TestBuscadorPdfs:
    def test_subcarpetas_ordenadas(self, tmp_path):
        base = tmp_path / "base"
        (base / "B").mkdir(parents=True)
        (base / "A").mkdir()
        (base / "suelto.txt").write_text("x", encoding="utf-8")
        assert [ruta.name for ruta in BuscadorPdfs().subcarpetas(base)] == ["A", "B"]

    def test_orden_de_exploracion(self, tmp_path):
        base = tmp_path / "SAANA"
        (base / "ORD NOMINA").mkdir(parents=True)
        (base / "ORDINARIA").mkdir()
        (base / "SUELDO").mkdir()
        crear_pdf(base / "ORD NOMINA" / "ord.pdf", "X")
        crear_pdf(base / "ORDINARIA" / "ordinaria.pdf", "X")
        crear_pdf(base / "SUELDO" / "sueldo.pdf", "X")
        crear_pdf(base / "raiz.pdf", "X")
        assert [ruta.name for ruta in BuscadorPdfs().pdfs(base)] == [
            "ord.pdf",
            "ordinaria.pdf",
            "sueldo.pdf",
            "raiz.pdf",
        ]

    def test_no_duplica_la_raiz_revisada_dos_veces(self, tmp_path):
        base = tmp_path / "SUELDO"
        base.mkdir()
        crear_pdf(base / "sueldo.pdf", "X")
        assert [ruta.name for ruta in BuscadorPdfs().pdfs(base)] == ["sueldo.pdf"]

    def test_subcarpeta_sin_pdfs(self, tmp_path):
        base = tmp_path / "VACIA"
        base.mkdir()
        assert BuscadorPdfs().pdfs(base) == []


class TestPoliticaDeNombre:
    def _regla(self, **extra) -> ReglaClasificacion:
        datos = {
            "caso": CasoUso.CASAS_COMERCIALES,
            "textos_requeridos": (),
            "incluir_concepto_en_nombre": True,
            "evaluar_no_ley": True,
        }
        datos.update(extra)
        return ReglaClasificacion(**datos)

    def _evaluacion(self, conceptos=()) -> EvaluacionPdf:
        return EvaluacionPdf(
            origen=Path("/tmp/a.pdf"),
            valido=True,
            categoria=CategoriaPdf.NORMAL,
            conceptos=tuple(conceptos),
            detalle="",
        )

    def test_sin_concepto_en_el_nombre(self):
        regla = self._regla(incluir_concepto_en_nombre=False)
        assert (
            RenombradorArchivos().nombre_base(
                regla, self._evaluacion(("A",)), Path("SAANA"), Path("a.pdf")
            )
            == "SAANA"
        )

    def test_un_concepto(self):
        assert (
            RenombradorArchivos().nombre_base(
                self._regla(), self._evaluacion(("CAHORMINSAS",)), Path("SAANA"), Path("a.pdf")
            )
            == "SAANA_CAHORMINSAS"
        )

    def test_dos_conceptos(self):
        assert (
            RenombradorArchivos().nombre_base(
                self._regla(),
                self._evaluacion(("CAHORMINSAS", "S.S.O.")),
                Path("SAANA"),
                Path("a.pdf"),
            )
            == "SAANA_CAHORMINSAS_S.S.O."
        )

    def test_mas_de_dos_conceptos(self):
        assert (
            RenombradorArchivos().nombre_base(
                self._regla(),
                self._evaluacion(("FONDO PENSIONES", "S.S.O.", "CAHORMINSAS")),
                Path("SAANA"),
                Path("a.pdf"),
            )
            == "SAANA_FONDO PENSIONES_Y_2_MAS"
        )

    def test_no_ley_conserva_el_nombre_original(self):
        assert (
            RenombradorArchivos().nombre_base(
                self._regla(),
                self._evaluacion(("CAHORMINSAS",)),
                Path("SAANA"),
                Path("EXTRA NO LEY.pdf"),
            )
            == "EXTRA NO LEY"
        )

    def test_no_ley_desactivado_usa_la_politica_normal(self):
        regla = self._regla(evaluar_no_ley=False)
        assert (
            RenombradorArchivos().nombre_base(
                regla, self._evaluacion(("CAHORMINSAS",)), Path("SAANA"), Path("x NO LEY.pdf")
            )
            == "SAANA_CAHORMINSAS"
        )


class TestReglasCasos:
    def test_reglas_por_defecto(self):
        reglas = ReglasCasos.desde_settings(Settings())
        assert set(reglas.casos()) == {
            CasoUso.RECIBOS,
            CasoUso.APORTES_PATRONALES,
            CasoUso.CASAS_COMERCIALES,
        }
        recibos = reglas.obtener(CasoUso.RECIBOS)
        assert recibos.textos_requeridos == TEXTOS_RECIBOS
        assert recibos.unificar_pdf is False
        assert recibos.evaluar_no_ley is False

    def test_regla_de_aportes(self):
        regla = ReglasCasos.desde_settings(Settings()).obtener(CasoUso.APORTES_PATRONALES)
        assert regla.textos_requeridos == TEXTOS_APORTES
        assert regla.unificar_pdf is True
        assert regla.incluir_concepto_en_nombre is True
        assert regla.evaluar_no_ley is True

    def test_regla_de_casas_excluye_la_fila_en_blanco(self):
        regla = ReglasCasos.desde_settings(Settings()).obtener(CasoUso.CASAS_COMERCIALES)
        assert "(en blanco)" not in regla.mapa_conceptos
        assert regla.mapa_conceptos["CAJA DE AHORRO CAEMINSA"] == "CAEMINSAS"
        assert regla.textos_requeridos == tuple(regla.mapa_conceptos)

    def test_caso_sin_regla(self):
        reglas = ReglasCasos([])
        with pytest.raises(ConfigError):
            reglas.obtener(CasoUso.RECIBOS)


class TestLectorPdf:
    def test_lee_un_pdf_real(self, tmp_path):
        ruta = crear_pdf(tmp_path / "recibo.pdf", "Recibo Nº 123", "SAANA")
        pdf = LectorPdfPyMuPDF().leer(ruta)
        assert pdf.nombre == "recibo"
        assert pdf.origen == ruta
        assert len(pdf.paginas) == 1
        assert "Recibo" in pdf.texto_primera_pagina

    def test_archivo_corrupto(self, tmp_path):
        ruta = tmp_path / "roto.pdf"
        ruta.write_bytes(b"esto no es un pdf")
        with pytest.raises(PdfExtractionError):
            LectorPdfPyMuPDF().leer(ruta)


class TestUnificadorPdfs:
    def test_une_dos_pdfs(self, tmp_path):
        uno = crear_pdf(tmp_path / "uno.pdf", "PAGINA UNO")
        dos = crear_pdf(tmp_path / "dos.pdf", "PAGINA DOS")
        salida = tmp_path / "salida" / "unificado.pdf"
        resultado = UnificadorPdfs().unificar([uno, dos], salida)
        assert resultado == salida
        doc = pymupdf.open(str(salida))
        try:
            assert len(doc) == 2
            assert "PAGINA UNO" in doc[0].get_text()
            assert "PAGINA DOS" in doc[1].get_text()
        finally:
            doc.close()

    def test_sin_rutas(self, tmp_path):
        with pytest.raises(PdfExtractionError):
            UnificadorPdfs().unificar([], tmp_path / "x.pdf")


def _servicio(**extra) -> OrganizadorService:
    return OrganizadorService.crear(Settings(), **extra)


def _base(tmp_path: Path) -> tuple[Path, Path]:
    """Crea la raiz y su subcarpeta SAANA (el servicio itera subcarpetas)."""
    raiz = tmp_path / "base"
    subcarpeta = raiz / "SAANA"
    subcarpeta.mkdir(parents=True)
    return raiz, subcarpeta


class TestOrganizadorService:
    def test_organiza_por_gremio_y_deuda(self, tmp_path, logger_falso):
        raiz, subcarpeta = _base(tmp_path)
        (subcarpeta / "ORD NOMINA").mkdir()
        crear_pdf(
            subcarpeta / "ORD NOMINA" / "normal.pdf", "FONDO PENSIONES (JUBILACION) 100,00"
        )
        crear_pdf(subcarpeta / "ORD NOMINA" / "deuda.pdf", "FONDO PENSIONES EN DEUDA 100,00")
        crear_pdf(subcarpeta / "raiz.pdf", "SIN COINCIDENCIAS")

        resultado = _servicio(logger=logger_falso).organizar(
            raiz, CasoUso.APORTES_PATRONALES, tmp_path / "normal", tmp_path / "deuda"
        )

        assert [ruta.name for ruta in resultado.normales] == ["SAANA_FONDO PENSIONES.pdf"]
        assert [ruta.name for ruta in resultado.deudas] == ["SAANA_FONDO PENSIONES.pdf"]
        assert resultado.normales[0].parent == tmp_path / "normal" / "SAANA"
        assert any("[ACEPTADO - NORMAL] normal.pdf" in linea for linea in resultado.log)
        assert any("[ACEPTADO - DEUDA] deuda.pdf" in linea for linea in resultado.log)
        assert any("[RECHAZADO] raiz.pdf" in linea for linea in resultado.log)
        assert logger_falso.error_llamadas

    def test_unifica_normales_y_deudas(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        crear_pdf(subcarpeta / "a.pdf", "FONDO PENSIONES 1,00")
        crear_pdf(subcarpeta / "b.pdf", "CAHORMINSAS 2,00")
        crear_pdf(subcarpeta / "c.pdf", "S.S.O. EN DEUDA 3,00")

        resultado = _servicio().organizar(
            raiz, CasoUso.APORTES_PATRONALES, tmp_path / "normal", tmp_path / "deuda"
        )

        assert [ruta.name for ruta in resultado.unificados] == [
            "SAANA_UNIFICADO.pdf",
            "SAANA_DEUDAS_UNIFICADO.pdf",
        ]
        unificado = pymupdf.open(str(resultado.unificados[0]))
        try:
            assert len(unificado) == 2
        finally:
            unificado.close()

    def test_sin_unificacion_para_recibos(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        crear_pdf(subcarpeta / "a.pdf", "TOTAL DE DEDUCCIONES POR CONCEPTO 1,00")
        resultado = _servicio().organizar(
            raiz, CasoUso.RECIBOS, tmp_path / "normal", tmp_path / "deuda"
        )
        assert [ruta.name for ruta in resultado.normales] == ["SAANA.pdf"]
        assert resultado.unificados == ()

    def test_nombre_no_ley_conserva_el_original(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        crear_pdf(subcarpeta / "EXTRA NO LEY.pdf", "FONDO PENSIONES 1,00")
        resultado = _servicio().organizar(
            raiz, CasoUso.APORTES_PATRONALES, tmp_path / "normal", tmp_path / "deuda"
        )
        assert [ruta.name for ruta in resultado.normales] == ["EXTRA NO LEY.pdf"]

    def test_mas_de_dos_conceptos_resume_el_nombre(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        crear_pdf(
            subcarpeta / "multi.pdf",
            "FONDO PENSIONES 1,00",
            "S.S.O. 2,00",
            "CAHORMINSAS 3,00",
        )
        resultado = _servicio().organizar(
            raiz, CasoUso.APORTES_PATRONALES, tmp_path / "normal", tmp_path / "deuda"
        )
        assert [ruta.name for ruta in resultado.normales] == [
            "SAANA_FONDO PENSIONES_Y_2_MAS.pdf"
        ]

    def test_nombres_repetidos_se_desambiguan(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        crear_pdf(subcarpeta / "a.pdf", "FONDO PENSIONES 1,00")
        crear_pdf(subcarpeta / "b.pdf", "FONDO PENSIONES 2,00")
        resultado = _servicio().organizar(
            raiz, CasoUso.APORTES_PATRONALES, tmp_path / "normal", tmp_path / "deuda"
        )
        assert [ruta.name for ruta in resultado.normales] == [
            "SAANA_FONDO PENSIONES.pdf",
            "SAANA_FONDO PENSIONES_2.pdf",
        ]

    def test_explora_ord_ordinaria_sueldo_y_la_raiz(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        (subcarpeta / "ORD NOMINA").mkdir()
        (subcarpeta / "ORDINARIA").mkdir()
        (subcarpeta / "SUELDO").mkdir()
        crear_pdf(subcarpeta / "ORD NOMINA" / "ord.pdf", "FONDO PENSIONES 1,00")
        crear_pdf(subcarpeta / "ORDINARIA" / "ordinaria.pdf", "SIN CONCEPTOS")
        crear_pdf(subcarpeta / "SUELDO" / "sueldo.pdf", "CAHORMINSAS 2,00")
        crear_pdf(subcarpeta / "raiz.pdf", "S.S.O. 3,00")

        resultado = _servicio().organizar(
            raiz, CasoUso.APORTES_PATRONALES, tmp_path / "normal", tmp_path / "deuda"
        )

        assert len(resultado.normales) == 3
        assert any(
            "SUBCARPETA EVALUADA: [SAANA] (4 PDFs detectados)" in linea
            for linea in resultado.log
        )
        assert any("[RECHAZADO] ordinaria.pdf" in linea for linea in resultado.log)

    def test_subcarpeta_sin_pdfs(self, tmp_path):
        raiz = tmp_path / "base"
        (raiz / "VACIA").mkdir(parents=True)
        resultado = _servicio().organizar(
            raiz, CasoUso.RECIBOS, tmp_path / "normal", tmp_path / "deuda"
        )
        assert resultado.normales == ()
        assert any(
            "[VACÍO] No se encontraron archivos .pdf." in linea for linea in resultado.log
        )

    def test_pdf_que_no_se_puede_leer(self, tmp_path):
        raiz, subcarpeta = _base(tmp_path)
        (subcarpeta / "roto.pdf").write_bytes(b"no es un pdf")
        resultado = _servicio().organizar(
            raiz, CasoUso.RECIBOS, tmp_path / "normal", tmp_path / "deuda"
        )
        assert resultado.normales == ()
        assert any("Razón: Error de lectura:" in linea for linea in resultado.log)

    def test_base_inexistente(self, tmp_path):
        with pytest.raises(ConfigError):
            _servicio().organizar(
                tmp_path / "no-existe",
                CasoUso.RECIBOS,
                tmp_path / "normal",
                tmp_path / "deuda",
            )

    def test_caso_sin_evaluador(self, tmp_path):
        raiz, _ = _base(tmp_path)
        servicio = OrganizadorService.crear(Settings())
        servicio._evaluadores = {}
        with pytest.raises(ConfigError):
            servicio.organizar(raiz, CasoUso.RECIBOS, tmp_path / "n", tmp_path / "d")
