"""Documentos de entrada y salida compartidos por los modulos."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Sequence


class PdfFormat(Enum):
    """Formatos de listado que se reconocen dentro de un PDF de aportes."""

    TABLA_GRUPOS = "TABLA_GRUPOS"
    LISTADO_INDIVIDUAL = "LISTADO_INDIVIDUAL"


@dataclass(frozen=True)
class PdfPage:
    """Una pagina de PDF ya extraida a texto."""

    numero: int
    texto: str


@dataclass(frozen=True)
class PdfDocument:
    """PDF en memoria: nombre logico, paginas y origen opcional en disco."""

    nombre: str
    paginas: Sequence[PdfPage]
    origen: Path | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "paginas", tuple(self.paginas))

    @property
    def nombre_archivo(self) -> str:
        return self.origen.name if self.origen is not None else self.nombre

    @property
    def primera_pagina(self) -> PdfPage | None:
        return self.paginas[0] if self.paginas else None

    @property
    def texto_primera_pagina(self) -> str:
        pagina = self.primera_pagina
        return (pagina.texto if pagina is not None else "") or ""

    def paginas_con_texto(self) -> tuple[PdfPage, ...]:
        return tuple(p for p in self.paginas if (p.texto or "").strip())

    def texto_completo(self) -> str:
        return "\n".join(p.texto or "" for p in self.paginas)


@dataclass(frozen=True)
class ExcelDocument:
    """Libro de Excel generado en memoria, listo para entregar al router."""

    nombre: str
    contenido: bytes = field(repr=False)
