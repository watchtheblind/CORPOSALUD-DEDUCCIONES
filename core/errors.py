"""Errores de dominio compartidos por todos los modulos."""

from __future__ import annotations


class DomainError(Exception):
    """Base de todos los errores de negocio del sistema."""


class PdfExtractionError(DomainError):
    """El PDF no se pudo leer o no tiene la estructura esperada."""


class ReportGenerationError(DomainError):
    """No se pudo construir un libro de Excel a partir de los datos extraidos."""


class ConfigError(DomainError):
    """La configuracion inyectada es invalida o esta incompleta."""
