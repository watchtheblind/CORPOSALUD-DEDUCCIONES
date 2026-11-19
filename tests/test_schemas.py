"""Tests de los esquemas Pydantic estrictos de los cuatro modulos."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path

import pytest
from pydantic import ValidationError

from modules.aportes_patronales import (
    AporteRow,
    AporteRowSchema,
    AportesResultSchema,
    AportesService,
    Centro as CentroAportes,
    Concepto as ConceptoAportes,
)
from core.config import StaticConfigProvider
from modules.casas_comerciales import (
    CasasComercialesRow,
    CasasComercialesRowSchema,
    CasasComercialesResultSchema,
    Centro as CentroCasas,
    Concepto as ConceptoCasas,
)
from modules.organizador import (
    CasoUso,
    CategoriaPdf,
    EvaluacionPdf,
    EvaluacionPdfSchema,
    OrganizacionRequestSchema,
    OrganizacionResult,
    OrganizacionResultSchema,
)
from modules.recibos import (
    AuditoriaEntry,
    AuditoriaEntrySchema,
    Centro as CentroRecibos,
    ConceptoCode,
    ReciboDocument,
    ReciboDocumentSchema,
    ReciboResumen,
    ReciboResumenSchema,
    RecibosExtraccionSchema,
)
from tests.support import documento

EXTRA = {"campo_extra": 1}

ESQUEMAS_VALIDOS = [
    ("CentroSchema", {"nombre": "CENTRO"}),
    ("ConceptoSchema", {"codigo": "SSO"}),
    (
        "AporteRowSchema",
        {
            "centro": "CENTRO",
            "grupo": "EMPLEADO FIJO",
            "referencia": "REF",
            "concepto": "SSO",
            "cantidad": 1,
            "trabajador": "10.00",
            "empresa": "20.00",
        },
    ),
    (
        "CasasComercialesRowSchema",
        {
            "centro": "CENTRO",
            "grupo": "EMPLEADO FIJO",
            "referencia": "REF",
            "concepto": "CAHORMINSAS",
            "cantidad": 1,
            "trabajador": "10.00",
            "empresa": "20.00",
        },
    ),
    (
        "ReciboDocumentSchema",
        {
            "concepto": "SISTRASALUD",
            "codigo": "583",
            "entidad": "SAANA",
            "numero": "12345",
            "monto": "10.00",
        },
    ),
    ("ReciboResumenSchema", {"archivo": "a.pdf", "numero": "12345"}),
    ("AuditoriaEntrySchema", {"nivel": "INFO", "mensaje": "uno"}),
    (
        "OrganizacionRequestSchema",
        {
            "base": "/tmp/base",
            "caso": "RECIBOS",
            "destino_normal": "/tmp/normal",
            "destino_deuda": "/tmp/deuda",
        },
    ),
]


def _todos_los_esquemas() -> dict[str, type]:
    from modules.aportes_patronales import schemas as schemas_aportes
    from modules.casas_comerciales import schemas as schemas_casas
    from modules.organizador import schemas as schemas_organizador
    from modules.recibos import schemas as schemas_recibos

    clases: dict[str, type] = {}
    for modulo in (schemas_aportes, schemas_casas, schemas_recibos, schemas_organizador):
        for nombre in dir(modulo):
            objeto = getattr(modulo, nombre)
            if (
                isinstance(objeto, type)
                and hasattr(objeto, "model_fields")
                and not nombre.startswith("_")
                and objeto.__module__ == modulo.__name__
            ):
                clases[nombre] = objeto
    return clases


class TestEsquemasEstrictos:
    def test_todos_prohiben_campos_extra(self):
        for nombre, clase in _todos_los_esquemas().items():
            assert clase.model_config.get("extra") == "forbid", nombre

    def test_nombres_exportados_estan_cubiertos(self):
        clases = _todos_los_esquemas()
        for nombre, _ in ESQUEMAS_VALIDOS:
            assert nombre in clases

    def test_extra_prohibido_en_cada_esquema(self):
        clases = _todos_los_esquemas()
        for nombre, datos in ESQUEMAS_VALIDOS:
            clase = clases[nombre]
            with pytest.raises(ValidationError):
                clase(**datos, **EXTRA)


class TestEsquemasAportes:
    def _fila(self) -> AporteRow:
        return AporteRow(
            centro=CentroAportes("CENTRO A"),
            grupo="EMPLEADO FIJO",
            referencia="REF",
            concepto=ConceptoAportes("SSO"),
            cantidad=2,
            trabajador=Decimal("10.00"),
            empresa=Decimal("20.00"),
        )

    def test_ida_y_vuelta(self):
        esquema = AporteRowSchema.from_domain(self._fila())
        assert esquema.centro == "CENTRO A"
        assert esquema.trabajador == Decimal("10.00")
        assert esquema.to_domain() == self._fila()

    def test_cantidad_negativa_invalida(self):
        with pytest.raises(ValidationError):
            AporteRowSchema(
                centro="C", grupo="G", referencia="R", concepto="X", cantidad=-1,
                trabajador="1", empresa="1",
            )

    def test_centro_vacio_invalido(self):
        with pytest.raises(ValidationError):
            AporteRowSchema(
                centro="", grupo="G", referencia="R", concepto="X", cantidad=1,
                trabajador="1", empresa="1",
            )

    def test_result_schema(self):
        resultado = AportesService.crear().procesar(
            [documento("a.pdf", "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS SERVICIOS CENTRALES\n"
                        "LISTADO DE DEDUCCIONES POR GRUPOS DE NÓMINA\nCONCEPTO: S.S.O.\n"
                        "EMPLEADO FIJO  1  1,00  2,00")],
            StaticConfigProvider(),
        )
        esquema = AportesResultSchema.from_domain(resultado)
        assert esquema.registros == 1
        assert esquema.documento == "CONSOLIDADO APORTES PATRONALES.xlsx"
        assert esquema.filas[0].centro == "SERVICIOS CENTRALES"


class TestEsquemasCasas:
    def test_ida_y_vuelta(self):
        fila = CasasComercialesRow(
            centro=CentroCasas("CENTRO A"),
            grupo="EMPLEADO FIJO",
            referencia="REF",
            concepto=ConceptoCasas("CAHORMINSAS"),
            cantidad=2,
            trabajador=Decimal("10.00"),
            empresa=Decimal("20.00"),
        )
        assert CasasComercialesRowSchema.from_domain(fila).to_domain() == fila

    def test_result_schema(self):
        esquema = CasasComercialesResultSchema(documento="X.xlsx", registros=0, filas=[])
        assert esquema.model_dump() == {"documento": "X.xlsx", "registros": 0, "filas": []}


class TestEsquemasRecibos:
    def test_documento(self):
        documento_recibo = ReciboDocument(
            concepto=ConceptoCode("SISTRASALUD", "583"),
            entidad=CentroRecibos("SAANA"),
            numero="12345",
            monto=Decimal("10.00"),
        )
        esquema = ReciboDocumentSchema.from_domain(documento_recibo)
        assert esquema.model_dump() == {
            "concepto": "SISTRASALUD",
            "codigo": "583",
            "entidad": "SAANA",
            "numero": "12345",
            "monto": Decimal("10.00"),
        }

    def test_resumen_y_auditoria(self):
        assert ReciboResumenSchema.from_domain(
            ReciboResumen(archivo="a.pdf", numero="1")
        ).archivo == "a.pdf"
        assert AuditoriaEntrySchema.from_domain(
            AuditoriaEntry(nivel="AVISO", mensaje="mensaje")
        ).nivel == "AVISO"

    def test_extraccion_compuesta(self):
        esquema = RecibosExtraccionSchema(documentos=[], auditoria=[])
        assert esquema.documentos == []


class TestEsquemasOrganizador:
    def test_request_convierte_rutas_y_enum(self):
        esquema = OrganizacionRequestSchema(
            base="/tmp/base",
            caso="CASAS_COMERCIALES",
            destino_normal="/tmp/normal",
            destino_deuda="/tmp/deuda",
        )
        assert esquema.base == Path("/tmp/base")
        assert esquema.caso is CasoUso.CASAS_COMERCIALES

    def test_caso_invalido(self):
        with pytest.raises(ValidationError):
            OrganizacionRequestSchema(
                base="/tmp/b", caso="OTRO", destino_normal="/tmp/n", destino_deuda="/tmp/d"
            )

    def test_evaluacion(self):
        evaluacion = EvaluacionPdf(
            origen=Path("/tmp/a.pdf"),
            valido=True,
            categoria=CategoriaPdf.DEUDA,
            conceptos=("CAHORMINSAS",),
            detalle="Concepto(s): [CAHORMINSAS] Y 'DEUDA'",
        )
        esquema = EvaluacionPdfSchema.from_domain(evaluacion)
        assert esquema.model_dump() == {
            "origen": Path("/tmp/a.pdf"),
            "valido": True,
            "categoria": CategoriaPdf.DEUDA,
            "conceptos": ["CAHORMINSAS"],
            "detalle": "Concepto(s): [CAHORMINSAS] Y 'DEUDA'",
        }

    def test_resultado(self):
        resultado = OrganizacionResult(
            log=("uno",), normales=(Path("/tmp/a.pdf"),), deudas=(), unificados=()
        )
        esquema = OrganizacionResultSchema.from_domain(resultado)
        assert esquema.total_normales == 1
        assert esquema.normales == ["/tmp/a.pdf"]
        assert esquema.log == ["uno"]
