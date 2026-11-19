# Migracion: CORPOSALUD-DEDUCCIONES -> monolito modular

Documento de la migracion ejecutada en los Targets 1 a 4 del plan.md.

## Mapa de archivos

| Antes | Ahora |
|---|---|
| utils.py | core/money.py (CommaDotMoneyParser) |
| casas_comerciales.py | modules/aportes_patronales/{domain,schemas,service}.py |
| casas_comercialesf.py | modules/casas_comerciales/{domain,schemas,service}.py |
| recibos/extractor_nomina.py | modules/recibos/domain.py (ReciboParser, LegacyReciboMoneyParser) + service.py |
| recibos/procesador_reportes.py | modules/recibos/domain.py (ReporteConsolidadoBuilder, NomenclaturaHojas) |
| RESUMIDOR_DE_TXT.py | modules/recibos/domain.py (ResumidorRecibos) |
| file_organiser.py | modules/organizador/{domain,service}.py |
| config.json | core/config.py (Settings) |
| patronales_masivo.py, recibos/main.py, requirements.txt | eliminados |

## Archivos nuevos de core

El plan declara core/config.py, core/errors.py y core/logging.py como
"existing" en el repo destino, pero ese repo no esta en este checkout. Se crean
aqui versiones minimas de esos puertos para que los modulos importen y los tests
corran; el repo destino debe conservar las suyas. Ademas se agrega
core/documents.py con los tipos compartidos del diagrama (PdfDocument, PdfPage,
PdfFormat, ExcelDocument), que no tenian archivo asignado.

## Decisiones y desviaciones

1. Dinero en Decimal. core/money.py devuelve Decimal. Se conservan las rarezas
   historicas fijadas por tests: el texto invalido vale 0, un unico separador se
   interpreta como decimal y LegacyReciboMoneyParser pierde el signo negativo.
2. Mapa de conceptos de recibos. El diccionario historico definia dos veces
   "FONDO DE AHORRO OBLIGATORIO PARA LA VIV"; gana 245.3. Se migro como una sola
   clave con ese valor y quedo documentado en el codigo.
3. Autonomia de modulos. CasasComercialesRow.from_aporte acepta un Protocol con
   la forma de la fila de aportes en lugar de importar el modulo de aportes,
   cumpliendo "los modulos no se importan entre si".
4. Evaluadores por caso. OrganizadorService recibe un mapa CasoUso -> PdfEvaluador
   en lugar de un unico evaluador: organizar(caso, ...) necesita la regla del
   caso y no se admite estado mutable.
5. Anclas inyectadas. CasasPdfExtractor.extract y CentroCasasResolver.resolve
   aceptan etiquetas opcionales para usar los Settings inyectados; el servicio
   las pasa siempre. En aportes la firma ya las recibia.
6. Resumen de auditoria sin colision (correccion). El layout historico escribia
   el TOTAL en la misma fila que "Obreros Contratados" (bug de _escribir_resumen).
   Ahora las cuatro categorias ocupan filas consecutivas y el TOTAL va en la
   siguiente. Los totales no cambian.
7. Nombres de hoja. Se mantiene el recorte a 31 caracteres y la limpieza de ":" y
   "/"; si dos conceptos colisionan tras el recorte se agrega un sufijo numerico
   (antes el escritor fallaba).
8. Sin tablas dinamicas de Excel. AportesExcelBuilder genera BASE mas tres tablas
   resumen (APORTES TRABAJADOR, APORTES EMPRESA, CONSOLIDADO X GREMIO) con
   openpyxl: centro/gremio en filas, concepto en columnas y fila TOTAL.
9. Organizador determinista. El barrido de carpetas es secuencial y ordenado
   (antes usaba ThreadPoolExecutor y el orden de iterdir()), para que el log y
   los nombres repetidos sean reproducibles.
10. Auditoria de cobertura explicita. RecibosService.auditar_cobertura()
    reemplaza el reporte automatico del final de ejecutar_sistema(); el router
    del repo destino decide cuando invocarlo con las entidades vistas.
11. Log del organizador en memoria. Ya no se escriben log_*.txt; las lineas se
    devuelven en OrganizacionResult.log y se emiten por AuditLogger.
12. Errores con tipo. Los servicios lanzan PdfExtractionError,
    ReportGenerationError o ConfigError de core/errors.py. Un lote sin registros
    ahora es PdfExtractionError en lugar de terminar en silencio.
13. Dependencias. Ya no se usan pandas, xlsxwriter, pdfplumber ni pywin32.
    requirements.txt desaparece (lo pide el plan) y la suite usa
    requirements-dev.txt (pytest, pydantic, openpyxl, pymupdf).

## Fuera de alcance (sin cambios)

- router.py, endpoints, auth, rate limiting, CORS/JWT y scoping por departamento:
  Target 5 en el repo destino.
- pymupdf, openpyxl y pydantic deben declararse en el repo destino (Target 6).
