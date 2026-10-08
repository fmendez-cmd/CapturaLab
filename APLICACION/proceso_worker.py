from __future__ import annotations

import traceback


def ejecutar_lote_en_proceso(carpetas_dict, cola_progreso, cola_resultado):
    """Ejecuta un lote fuera del proceso del servidor NiceGUI.

    El proceso padre puede terminar este proceso sin apagar el servidor.
    """
    try:
        from APLICACION.workflow_service import WorkflowService

        servicio = WorkflowService()
        servicio.preparar_procesamiento()

        def progreso(datos):
            try:
                cola_progreso.put(datos)
            except Exception:
                pass

        resultado = servicio.procesar_lote_carpetas(
            carpetas_dict,
            progreso,
        )

        cola_resultado.put({
            "tipo": "resultado",
            "resultado": resultado,
        })

    except BaseException as error:
        try:
            cola_resultado.put({
                "tipo": "error",
                "error": f"{type(error).__name__}: {error}",
                "traceback": traceback.format_exc(),
            })
        except Exception:
            pass
