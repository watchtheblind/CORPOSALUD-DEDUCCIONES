"""Utilidades compartidas por la suite de la arquitectura nueva.

Construye documentos en memoria (PdfDocument) y PDFs reales en disco para los
pocos casos que necesitan PyMuPDF, y ofrece un doble de AuditLogger.
"""

from __future__ import annotations

from pathlib import Path

import pymupdf

from core.documents import PdfDocument, PdfPage
from core.logging import AuditLogger

RAIZ = Path(__file__).resolve().parents[1]


def paginas(*textos: str | None) -> tuple[PdfPage, ...]:
    return tuple(
        PdfPage(numero=indice + 1, texto=texto if texto is not None else "")
        for indice, texto in enumerate(textos)
    )


def documento(nombre: str, *textos: str | None, origen: Path | None = None) -> PdfDocument:
    """PDF en memoria a partir del texto de cada pagina."""
    return PdfDocument(nombre=nombre, paginas=paginas(*textos), origen=origen)


def crear_pdf(ruta: Path, *lineas: str) -> Path:
    """Crea un PDF de una pagina con las lineas indicadas."""
    documento_pdf = pymupdf.open()
    pagina = documento_pdf.new_page()
    y = 72.0
    for linea in lineas:
        pagina.insert_text((72.0, y), linea, fontsize=11)
        y += 18.0
    documento_pdf.save(str(ruta))
    documento_pdf.close()
    return ruta


class RegistroAuditoriaFalso(AuditLogger):
    """Doble de AuditLogger que guarda lo emitido."""

    def __init__(self) -> None:
        self.info_llamadas: list[str] = []
        self.error_llamadas: list[str] = []

    def info(self, mensaje: str) -> None:
        self.info_llamadas.append(mensaje)

    def error(self, mensaje: str) -> None:
        self.error_llamadas.append(mensaje)
