"""Nucleo compartido por los modulos: configuracion, errores, auditoria, dinero y documentos."""

from core.config import ConfigProvider, Settings, StaticConfigProvider
from core.documents import ExcelDocument, PdfDocument, PdfFormat, PdfPage
from core.errors import ConfigError, DomainError, PdfExtractionError, ReportGenerationError
from core.logging import AuditLogger, NullAuditLogger
from core.money import CommaDotMoneyParser, LegacyReciboMoneyParser, MoneyParser

__all__ = [
    "AuditLogger",
    "CommaDotMoneyParser",
    "ConfigError",
    "ConfigProvider",
    "DomainError",
    "ExcelDocument",
    "LegacyReciboMoneyParser",
    "MoneyParser",
    "NullAuditLogger",
    "PdfDocument",
    "PdfExtractionError",
    "PdfFormat",
    "PdfPage",
    "ReportGenerationError",
    "Settings",
    "StaticConfigProvider",
]
