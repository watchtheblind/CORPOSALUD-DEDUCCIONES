"""Fixtures de la suite de la arquitectura nueva."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))

from tests.support import RegistroAuditoriaFalso  # noqa: E402


@pytest.fixture
def logger_falso() -> RegistroAuditoriaFalso:
    return RegistroAuditoriaFalso()
