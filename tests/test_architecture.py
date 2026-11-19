"""Guardas de arquitectura del monolito modular.

Verifican las reglas del plan: modulos autonomos, sin GUI, sin Excel nativo, sin
print ni rutas magicas, servicios que solo orquestan y puertos abstractos.
"""

from __future__ import annotations

import ast
import dataclasses
from pathlib import Path

import pytest

from core.config import ConfigProvider
from core.documents import PdfFormat
from core.logging import AuditLogger
from core.money import MoneyParser
from modules.aportes_patronales import AporteRow, AportesResult, ReportBuilder, PageExtractor
from modules.casas_comerciales import CasasComercialesRow, GremioCategoria
from modules.organizador import CategoriaPdf, CasoUso, EvaluacionPdf, LectorPdf
from modules.recibos import ReciboDocument, ReciboResumen

RAIZ = Path(__file__).resolve().parents[1]
PAQUETES = [RAIZ / "core", *sorted((RAIZ / "modules").iterdir())]

ARCHIVOS_PROHIBIDOS = [
    "casas_comerciales.py",
    "casas_comercialesf.py",
    "file_organiser.py",
    "RESUMIDOR_DE_TXT.py",
    "utils.py",
    "config.json",
    "requirements.txt",
    "recibos",
    "patronales_masivo.py",
]

TOKENS_PROHIBIDOS = {
    "tkinter": "GUI eliminada",
    "filedialog": "GUI eliminada",
    "messagebox": "GUI eliminada",
    "win32com": "Excel nativo eliminado",
    "pywin32": "Excel nativo eliminado",
    "os.startfile": "apertura de carpetas eliminada",
    "pandas": "reemplazado por openpyxl",
    "pdfplumber": "la lectura de PDFs entra por PdfDocument/LectorPdf",
    "xlsxwriter": "reemplazado por openpyxl",
    "__file__": "dependencia de ruta del modulo",
    "Desktop": "ruta de escritorio hardcodeada",
}


def _archivos_python() -> list[Path]:
    return sorted(
        archivo
        for paquete in PAQUETES
        if paquete.exists()
        for archivo in paquete.rglob("*.py")
    )


class TestArchivosMigrados:
    @pytest.mark.parametrize("nombre", ARCHIVOS_PROHIBIDOS)
    def test_los_archivos_viejos_ya_no_existen(self, nombre):
        assert not (RAIZ / nombre).exists()

    def test_la_estructura_objetivo_existe(self):
        for modulo in (
            "aportes_patronales",
            "casas_comerciales",
            "recibos",
            "organizador",
        ):
            for archivo in ("__init__.py", "schemas.py", "service.py", "domain.py"):
                assert (RAIZ / "modules" / modulo / archivo).is_file()
        for archivo in ("config.py", "errors.py", "logging.py", "money.py"):
            assert (RAIZ / "core" / archivo).is_file()

    def test_no_hay_router_py(self):
        """Target 5: los routers viven en el repo destino, no aqui."""
        assert not list((RAIZ / "modules").rglob("router.py"))


def _codigo_efectivo(archivo: Path) -> str:
    """Fuente sin docstrings ni comentarios: los textos no son acoplamientos."""
    arbol = ast.parse(archivo.read_text(encoding="utf-8"))
    for nodo in ast.walk(arbol):
        if not isinstance(
            nodo, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)
        ):
            continue
        cuerpo = nodo.body
        if (
            cuerpo
            and isinstance(cuerpo[0], ast.Expr)
            and isinstance(cuerpo[0].value, ast.Constant)
            and isinstance(cuerpo[0].value.value, str)
        ):
            nodo.body = cuerpo[1:] or [ast.Pass()]
    return ast.unparse(arbol)


class TestTokensProhibidos:
    @pytest.mark.parametrize("token,razon", sorted(TOKENS_PROHIBIDOS.items()))
    def test_token_ausente(self, token, razon):
        for archivo in _archivos_python():
            assert token not in _codigo_efectivo(archivo), f"{token} en {archivo}: {razon}"

    def test_sin_print_en_produccion(self):
        for archivo in _archivos_python():
            assert "print(" not in _codigo_efectivo(archivo), f"print en {archivo}"


class TestDependenciasEntrePaquetes:
    def _imports(self, archivo: Path) -> set[str]:
        arbol = ast.parse(archivo.read_text(encoding="utf-8"))
        modulos: set[str] = set()
        for nodo in ast.walk(arbol):
            if isinstance(nodo, ast.Import):
                modulos.update(alias.name for alias in nodo.names)
            elif isinstance(nodo, ast.ImportFrom) and nodo.module:
                modulos.add(nodo.module)
        return modulos

    def test_los_modulos_no_se_importan_entre_si(self):
        for modulo_dir in sorted((RAIZ / "modules").iterdir()):
            if not modulo_dir.is_dir():
                continue
            propio = f"modules.{modulo_dir.name}"
            for archivo in modulo_dir.rglob("*.py"):
                for importado in self._imports(archivo):
                    if not importado.startswith("modules."):
                        continue
                    assert importado.startswith(propio), (
                        f"{archivo} importa {importado}: los modulos no se acoplan"
                    )

    def test_core_no_conoce_los_modulos(self):
        for archivo in (RAIZ / "core").rglob("*.py"):
            for importado in self._imports(archivo):
                assert not importado.startswith("modules"), f"{archivo} importa {importado}"

    def test_los_servicios_solo_orquestan(self):
        """service.py no define funciones sueltas ni logica de negocio."""
        for servicio in (RAIZ / "modules").rglob("service.py"):
            arbol = ast.parse(servicio.read_text(encoding="utf-8"))
            funciones = [
                nodo.name for nodo in arbol.body if isinstance(nodo, ast.FunctionDef)
            ]
            assert funciones == [], f"{servicio} define funciones de nivel de modulo"
            clases = [nodo for nodo in arbol.body if isinstance(nodo, ast.ClassDef)]
            assert len(clases) == 1, f"{servicio} debe declarar exactamente una fachada"


class TestModelosDeDominio:
    @pytest.mark.parametrize(
        "clase",
        [
            AporteRow,
            AportesResult,
            CasasComercialesRow,
            ReciboDocument,
            ReciboResumen,
            EvaluacionPdf,
        ],
    )
    def test_los_value_objects_son_frozen(self, clase):
        assert dataclasses.is_dataclass(clase)
        assert clase.__dataclass_params__.frozen

    @pytest.mark.parametrize(
        "enumeracion",
        [PdfFormat, GremioCategoria, CasoUso, CategoriaPdf],
    )
    def test_los_conjuntos_cerrados_son_enums(self, enumeracion):
        from enum import Enum

        assert issubclass(enumeracion, Enum)

    @pytest.mark.parametrize(
        "puerto",
        [MoneyParser, AuditLogger, ConfigProvider, ReportBuilder, PageExtractor, LectorPdf],
    )
    def test_los_puertos_no_se_instancian(self, puerto):
        with pytest.raises(TypeError):
            puerto()
