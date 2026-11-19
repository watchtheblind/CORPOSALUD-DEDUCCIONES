# Suite de tests (Targets 0-4 del plan)

Tests de la arquitectura nueva: core compartido y los cuatro modulos del
monolito modular (aportes_patronales, casas_comerciales, recibos, organizador).
El comportamiento heredado se congelo en Target 0 y estas pruebas lo re-apuntan
a las clases nuevas sin cambiar las expectativas, salvo las desviaciones
documentadas en MIGRACION.md.

## Como correr

    python3 -m venv .venv
    .venv/bin/python -m pip install -r requirements-dev.txt
    .venv/bin/python -m pytest

Son ~385 casos y corre en menos de dos segundos. Los tests escriben solo en los
directorios temporales de pytest.

## Archivos

| Archivo | Cubre |
|---|---|
| test_core.py | parsers de montos (Decimal), Settings/ConfigProvider, PdfDocument/ExcelDocument, errores y logger |
| test_aportes_domain.py | ConceptoNormalizer, CentroResolver, FormatDetector, extractores tabla/listado, deteccion por bloques, AportesExcelBuilder y AportesService |
| test_casas_domain.py | ClasificadorGremio, ConceptoCasasNormalizer, CentroCasasResolver, CasasPdfExtractor, AuditoriaExcelBuilder y CasasComercialesService |
| test_recibos_domain.py | catalogos de conceptos y centros, ReciboParser, auditoria en memoria, consolidado, reporte RETENCIONES/CASAS y RecibosService |
| test_organizador_domain.py | normalizacion de texto, PdfEvaluador, reglas por caso, lector/unificador de PDFs y OrganizadorService con PDFs reales |
| test_schemas.py | esquemas Pydantic: extra="forbid", validaciones, ida y vuelta con el dominio |
| test_architecture.py | guardas del plan: archivos viejos borrados, sin tkinter/win32com/print/__file__/Desktop, modulos autonomos, servicios sin logica |

## Reglas que la suite protege

- core no importa modules; ningun modulo importa a otro modulo.
- service.py solo orquesta: no define funciones ni logica de negocio.
- Los servicios no dependen de FastAPI y reciben configuracion por constructor
  o por parametro (ConfigProvider).
- No hay GUI, ni print, ni rutas magicas, ni Excel nativo: los libros de Excel
  se generan en memoria con openpyxl y se devuelven como ExcelDocument.
