# CORPOSALUD-DEDUCCIONES

Migracion del sistema de escritorio (tkinter + Excel nativo) a un monolito
modular compatible con FastAPI, segun plan.md.

## Estructura

    core/            puertos y tipos compartidos
      config.py      Settings + ConfigProvider (reemplaza config.json)
      documents.py   PdfDocument, PdfPage, PdfFormat, ExcelDocument
      errors.py      DomainError, PdfExtractionError, ReportGenerationError, ConfigError
      logging.py     AuditLogger + NullAuditLogger
      money.py       MoneyParser, CommaDotMoneyParser, LegacyReciboMoneyParser

    modules/
      aportes_patronales/  dominio + esquemas + servicio (antes casas_comerciales.py)
      casas_comerciales/   dominio + esquemas + servicio (antes casas_comercialesf.py)
      recibos/             dominio + esquemas + servicio (extractor y reportes)
      organizador/         dominio + esquemas + servicio (antes file_organiser.py)

Cada modulo declara domain.py (logica), schemas.py (Pydantic estricto) y
service.py (fachada que solo orquesta). Los modulos no se importan entre si.

## Tests

    python3 -m venv .venv
    .venv/bin/python -m pip install -r requirements-dev.txt
    .venv/bin/python -m pytest

## Documentacion

- plan.md: objetivos y contrato del repo destino.
- MIGRACION.md: mapa de archivos y desviaciones respecto al comportamiento viejo.
- tests/README.md: que cubre cada archivo de tests.
