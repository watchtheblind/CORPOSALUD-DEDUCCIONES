"""Puerto de auditoria: los modulos no escriben archivos ni usan print."""

from __future__ import annotations

from abc import ABC, abstractmethod


class AuditLogger(ABC):
    """Puerto de auditoria que el repo destino implementa sobre su logging."""

    @abstractmethod
    def info(self, mensaje: str) -> None:
        """Registra un evento informativo del proceso."""

    @abstractmethod
    def error(self, mensaje: str) -> None:
        """Registra un evento de error recuperable."""


class NullAuditLogger(AuditLogger):
    """Implementacion no-op util para tests y para composiciones minimas."""

    def info(self, mensaje: str) -> None:
        return None

    def error(self, mensaje: str) -> None:
        return None
