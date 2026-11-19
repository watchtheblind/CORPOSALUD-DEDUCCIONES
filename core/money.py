"""Parsers de montos monetarios.

Se conserva el comportamiento del sistema anterior: los separadores se deciden
con la ultima coma/punto y el texto que no se puede convertir vale cero. La
diferencia es que ahora se devuelve Decimal y que cada variante es una clase con
el mismo puerto MoneyParser.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from decimal import Decimal, InvalidOperation

CERO = Decimal("0")


class MoneyParser(ABC):
    """Puerto de conversion de texto monetario a Decimal."""

    @abstractmethod
    def parse(self, texto: str | None) -> Decimal:
        """Convierte el texto a Decimal; el texto invalido vale CERO."""


class CommaDotMoneyParser(MoneyParser):
    """Parser de los listados de aportes (equivalente a utils.limpiar_monto)."""

    def parse(self, texto: str | None) -> Decimal:
        if not texto:
            return CERO
        limpio = texto.strip().replace("$", "").replace("Bs", "").replace('"', "")
        pos_coma = limpio.rfind(",")
        pos_punto = limpio.rfind(".")
        if pos_coma > pos_punto:
            limpio = limpio.replace(".", "").replace(",", ".")
        elif pos_punto > pos_coma:
            limpio = limpio.replace(",", "")
        try:
            return Decimal(limpio)
        except (InvalidOperation, ValueError):
            return CERO


class LegacyReciboMoneyParser(MoneyParser):
    """Parser de los recibos de nomina (equivalente a limpiar_monto_estandar).

    Descarta todo caracter que no sea digito o punto, por lo que el signo
    negativo se pierde. Es el comportamiento historico y esta fijado por tests.
    """

    def parse(self, texto: str | None) -> Decimal:
        if not texto:
            return CERO
        limpio = texto.strip()
        ultima_coma = limpio.rfind(",")
        ultimo_punto = limpio.rfind(".")
        if ultima_coma > ultimo_punto:
            limpio = limpio.replace(".", "").replace(",", ".")
        elif ultimo_punto > ultima_coma:
            limpio = limpio.replace(",", "")
        else:
            limpio = limpio.replace(",", ".")
        limpio = "".join(c for c in limpio if c.isdigit() or c == ".")
        try:
            return Decimal(limpio)
        except (InvalidOperation, ValueError):
            return CERO
