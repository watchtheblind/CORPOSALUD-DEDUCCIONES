"""Tests de core: parsers de montos, configuracion y documentos."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest

from core.config import (
    ETIQUETAS_PDF,
    MAPA_CONCEPTOS,
    ConfigProvider,
    Settings,
    StaticConfigProvider,
)
from core.documents import ExcelDocument, PdfDocument, PdfFormat, PdfPage
from core.errors import ConfigError, DomainError, PdfExtractionError, ReportGenerationError
from core.logging import AuditLogger, NullAuditLogger
from core.money import CommaDotMoneyParser, LegacyReciboMoneyParser, MoneyParser

CASOS_MONTO_CANONICO = [
    pytest.param(None, Decimal("0"), id="none"),
    pytest.param("", Decimal("0"), id="cadena-vacia"),
    pytest.param("0,00", Decimal("0.00"), id="cero"),
    pytest.param("1234", Decimal("1234"), id="entero-plano"),
    pytest.param("1.234,56", Decimal("1234.56"), id="miles-punto-decimal-coma"),
    pytest.param("1,234.56", Decimal("1234.56"), id="miles-coma-decimal-punto"),
    pytest.param("1.234.567,89", Decimal("1234567.89"), id="millon-punto-decimal-coma"),
    pytest.param("1,234,567.89", Decimal("1234567.89"), id="millon-coma-decimal-punto"),
    pytest.param("$ 12.50", Decimal("12.50"), id="prefijo-dolar"),
    pytest.param("12,50 Bs", Decimal("12.50"), id="sufijo-bs"),
    pytest.param('"1.000,00"', Decimal("1000.00"), id="comillas-envolventes"),
    pytest.param("-1.234,56", Decimal("-1234.56"), id="negativo-conserva-signo"),
    pytest.param("1.234,56\u00a0", Decimal("1234.56"), id="nbsp-final"),
    pytest.param(" 7 ", Decimal("7"), id="espacios-exteriores"),
    pytest.param("abc", Decimal("0"), id="texto-no-numerico"),
    pytest.param("1 234", Decimal("0"), id="espacio-interno-invalido"),
    pytest.param("S/N", Decimal("0"), id="sin-numero"),
    pytest.param("1.234", Decimal("1.234"), id="ambiguo-punto-unico-es-decimal"),
    pytest.param("1,234", Decimal("1.234"), id="ambiguo-coma-unica-es-decimal"),
    pytest.param("1.234.567", Decimal("0"), id="miles-punto-sin-decimales-falla"),
    pytest.param("1,234,567", Decimal("0"), id="miles-coma-sin-decimales-falla"),
]

CASOS_MONTO_RECIBO = [
    pytest.param(None, Decimal("0"), id="none"),
    pytest.param("", Decimal("0"), id="cadena-vacia"),
    pytest.param("12", Decimal("12"), id="entero-plano"),
    pytest.param("1.234,56", Decimal("1234.56"), id="miles-punto-decimal-coma"),
    pytest.param("1,234.56", Decimal("1234.56"), id="miles-coma-decimal-punto"),
    pytest.param("$1.234,56", Decimal("1234.56"), id="prefijo-dolar"),
    pytest.param("Bs 12,50", Decimal("12.50"), id="prefijo-bs"),
    pytest.param("(1.234,56)", Decimal("1234.56"), id="parentesis-de-negativo"),
    pytest.param("-1.234,56", Decimal("1234.56"), id="signo-inicial-descartado"),
    pytest.param("1.234,56-", Decimal("1234.56"), id="signo-final-descartado"),
    pytest.param("1.234", Decimal("1.234"), id="ambiguo-punto-unico"),
    pytest.param("1,234", Decimal("1.234"), id="ambiguo-coma-unica"),
    pytest.param("1.234.567", Decimal("0"), id="miles-sin-decimales-falla"),
    pytest.param("abc", Decimal("0"), id="texto-no-numerico"),
    pytest.param("S/N", Decimal("0"), id="sin-numero"),
]


class TestCommaDotMoneyParser:
    @pytest.mark.parametrize("entrada,esperado", CASOS_MONTO_CANONICO)
    def test_parsea(self, entrada, esperado):
        assert CommaDotMoneyParser().parse(entrada) == esperado

    def test_devuelve_decimal(self):
        assert isinstance(CommaDotMoneyParser().parse("1,00"), Decimal)

    def test_el_texto_invalido_no_lanza(self):
        assert CommaDotMoneyParser().parse("no es un monto") == Decimal("0")


class TestLegacyReciboMoneyParser:
    @pytest.mark.parametrize("entrada,esperado", CASOS_MONTO_RECIBO)
    def test_parsea(self, entrada, esperado):
        assert LegacyReciboMoneyParser().parse(entrada) == esperado

    def test_pierde_el_signo_negativo(self):
        """Comportamiento historico: descarta todo caracter no numerico."""
        assert LegacyReciboMoneyParser().parse("-1.234,56") == Decimal("1234.56")

    def test_ambos_parsers_fallan_igual_con_miles_sin_decimales(self):
        assert CommaDotMoneyParser().parse("1.234.567") == LegacyReciboMoneyParser().parse(
            "1.234.567"
        ) == Decimal("0")


class TestPuertosDeMoney:
    def test_money_parser_es_abstracto(self):
        with pytest.raises(TypeError):
            MoneyParser()


class TestSettings:
    def test_los_valores_por_defecto_son_los_de_config_json(self):
        settings = Settings()
        assert settings.mapa_conceptos == dict(MAPA_CONCEPTOS)
        assert settings.etiquetas_pdf == dict(ETIQUETAS_PDF)
        assert len(settings.etiquetas_pdf["ancla_centro"]) == 4
        assert "(en blanco)" in settings.mapa_conceptos

    def test_regex_datos_encaja_con_una_fila_de_aporte(self):
        import re

        coincidencia = re.match(Settings().regex_datos, "EMPLEADOS FIJOS 12 1.234,56 2.345,67")
        assert coincidencia is not None
        assert coincidencia.group(2) == "12"

    def test_es_inmutable(self):
        with pytest.raises(Exception):
            Settings().regex_datos = "otro"


class TestStaticConfigProvider:
    def test_devuelve_la_configuracion_inyectada(self):
        settings = Settings(regex_datos="^x$")
        provider = StaticConfigProvider(settings)
        assert provider.settings() is settings

    def test_sin_argumentos_usa_los_valores_por_defecto(self):
        assert StaticConfigProvider().settings() == Settings()

    def test_implementa_el_puerto(self):
        assert isinstance(StaticConfigProvider(), ConfigProvider)


class TestDocumentos:
    def test_pdf_document_congela_las_paginas(self):
        pdf = PdfDocument(nombre="x.pdf", paginas=[PdfPage(numero=1, texto="hola")])
        assert pdf.paginas == (PdfPage(numero=1, texto="hola"),)
        with pytest.raises(Exception):
            pdf.nombre = "otro"

    def test_nombre_archivo_prefiere_el_origen(self):
        pdf = PdfDocument(
            nombre="logico", paginas=(), origen=Path("/tmp/real.pdf")
        )
        assert pdf.nombre_archivo == "real.pdf"

    def test_accesos_comodos(self):
        pdf = PdfDocument(
            nombre="x.pdf",
            paginas=[PdfPage(numero=1, texto="uno"), PdfPage(numero=2, texto="  ")],
        )
        assert pdf.primera_pagina == PdfPage(numero=1, texto="uno")
        assert pdf.texto_primera_pagina == "uno"
        assert pdf.paginas_con_texto() == (PdfPage(numero=1, texto="uno"),)
        assert pdf.texto_completo() == "uno\n  "

    def test_pdf_sin_paginas(self):
        pdf = PdfDocument(nombre="vacio.pdf", paginas=())
        assert pdf.primera_pagina is None
        assert pdf.texto_primera_pagina == ""

    def test_formatos_y_libro_excel(self):
        assert {formato.name for formato in PdfFormat} == {
            "TABLA_GRUPOS",
            "LISTADO_INDIVIDUAL",
        }
        libro = ExcelDocument(nombre="x.xlsx", contenido=b"PK")
        assert libro.contenido == b"PK"
        assert "PK" not in repr(libro)


class TestErroresYAuditoria:
    def test_jerarquia_de_errores(self):
        for clase in (PdfExtractionError, ReportGenerationError, ConfigError):
            assert issubclass(clase, DomainError)
        assert issubclass(DomainError, Exception)

    def test_audit_logger_es_abstracto(self):
        with pytest.raises(TypeError):
            AuditLogger()

    def test_null_logger_no_hace_nada(self):
        NullAuditLogger().info("x")
        NullAuditLogger().error("y")
