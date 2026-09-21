# DOMINIO/control_cancelacion.py
from __future__ import annotations

import threading
import time

_local = threading.local()


class ProcesamientoCancelado(InterruptedError):
    """Cancelación cooperativa solicitada por el usuario."""


def establecer_evento_cancelacion(evento) -> None:
    _local.cancel_event = evento


def limpiar_evento_cancelacion() -> None:
    if hasattr(_local, "cancel_event"):
        delattr(_local, "cancel_event")


def obtener_evento_cancelacion():
    return getattr(_local, "cancel_event", None)


def cancelacion_solicitada() -> bool:
    evento = obtener_evento_cancelacion()
    return bool(evento is not None and evento.is_set())


def verificar_cancelacion() -> None:
    if cancelacion_solicitada():
        raise ProcesamientoCancelado("Procesamiento cancelado por el usuario.")


def esperar_cancelable(segundos: float) -> None:
    evento = obtener_evento_cancelacion()
    if evento is None:
        time.sleep(segundos)
        return
    if evento.wait(timeout=segundos):
        raise ProcesamientoCancelado("Procesamiento cancelado por el usuario.")
