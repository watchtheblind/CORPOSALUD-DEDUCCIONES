"""Configuracion inyectable que reemplaza config.json.

Los valores por defecto son exactamente los del antiguo config.json. Ningun
modulo lee archivos ni usa __file__/cwd: la configuracion entra por constructor
a traves del puerto ConfigProvider.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import Any, Mapping

ANCLAS_CENTRO = (
    "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS",
    "DIRECCIoN REGIONAL DE RECURSOS HUMANOS",
    "DIRECCION DE RECURSOS HUMANOS",
    "DIRECCIÓN DE RECURSOS HUMANOS",
)

ETIQUETAS_PDF: Mapping[str, Any] = MappingProxyType(
    {
        "ancla_centro": ANCLAS_CENTRO,
        "ancla_recibo": r"Recibo:\s*(\d+)",
        "ancla_concepto": "CONCEPTO:",
    }
)

MAPA_CONCEPTOS: Mapping[str, str] = MappingProxyType(
    {
        "ASUNAJUPENSAPROSO": "ASUNAJUPENSAPROSO",
        "CAHORMINSAS": "CAHORMINSAS",
        "CAJA DE AHORRO": "CAJA DE AHORRO",
        "CAJA DE AHORRO CAEMINSA": "CAEMINSAS",
        "COLEGIO BIOANALISTA": "COLEGIO BIOANALISTA",
        "COLEGIO DE ENFERMERA EDO. MIRANDA": "COLEGIO DE ENFERMERA EDO. MIRANDA",
        "COLEGIO DE ENFERMERAS ESTADO CARABOBO": "COLEGIO DE ENFERMERAS ESTADO CARABOBO",
        "COLEGIO ENFERMERAS": "COLEGIO ENFERMERAS",
        "COLEGIO NUTRIC. Y DIET. VZLA": "COLEGIO NUTRIC. Y DIET. VZLA",
        "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA": "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA",
        "DESC. DIA(S) NO LABORADO(S)": "DESC. DIA(S) NO LABORADO(S)",
        "FENASISTRASALUD": "FENASISTRASALUD",
        "JUZGADO PRIMERO DE MENORES": "JUZGADO PRIMERO DE MENORES",
        "PENSIÓN DE ALIMENTACIÓN": "PENSIÓN DE ALIMENTACIÓN",
        "PRESTAMO CAJA DE AHORRO": "PRESTAMO CAJA DE AHORRO",
        "SAPTRASEZ": "SAPTRASEZ",
        "SERVICIOS FUNERARIOS CAHORMINSAS": "SERVICIOS FUNERARIOS CAHORMINSAS",
        "SINBOPROENF": "SINBOPROENF",
        "SINDICATO OSBESS ARAGUA": "SINDICATO OSBESS ARAGUA",
        "SINDICATO UNICO DE TRAB.DE LA SALUD Y SUS SIMILARES DEL EDO.ARAGUA": (
            "SINDICATO UNICO DE TRAB.DE LA SALUD Y SUS SIMILARES DEL EDO.ARAGUA"
        ),
        "SISTRASALUD": "SISTRASALUD",
        "SITRASSS-MIRANDA": "SITRASSS-MIRANDA",
        "SOCIEDAD ANESTECIOLOGOS": "SOCIEDAD ANESTECIOLOGOS",
        "SUNEP-SAS": "SUNEP-SAS",
        "CUOTA EXTRAORDINARIA SUNEP-SAS": "CUOTA EXTRAORDINARIA SUNEP-SAS",
        "TRIBUNAL (PERMANENTE)": "TRIBUNAL (PERMANENTE)",
        "TRIBUNAL DE PROTECCION DE NIÑOS. NIÑAS Y ADOLESCENTES": (
            "TRIBUNAL DE PROTECCION DE NIÑOS. NIÑAS Y ADOLESCENTES"
        ),
        "(en blanco)": "SIN CATEGORIA",
    }
)

REGEX_DATOS = r"^(.+?)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)"

NOMBRES_COLUMNAS = ("CENTRO", "GREMIO", "TRABAJADOR", "EMPRESA", "CONCEPTO", "NRO_RECIBO")


@dataclass(frozen=True)
class Settings:
    """Configuracion inmutable que consumen los modulos."""

    etiquetas_pdf: Mapping[str, Any] = field(default_factory=lambda: dict(ETIQUETAS_PDF))
    mapa_conceptos: Mapping[str, str] = field(default_factory=lambda: dict(MAPA_CONCEPTOS))
    regex_datos: str = REGEX_DATOS
    nombres_columnas: tuple[str, ...] = NOMBRES_COLUMNAS


class ConfigProvider(ABC):
    """Puerto de configuracion."""

    @abstractmethod
    def settings(self) -> Settings:
        """Devuelve la configuracion vigente."""


class StaticConfigProvider(ConfigProvider):
    """Proveedor de configuracion fija (tests, CLI y composicion por defecto)."""

    def __init__(self, settings: Settings | None = None) -> None:
        self._settings = settings if settings is not None else Settings()

    def settings(self) -> Settings:
        return self._settings
