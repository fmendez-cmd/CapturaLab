# APLICACION/workflow_service.py

import os
import re
import traceback
import threading

from DOMINIO.generador_excel import GeneradorExcel
from INFRAESTRUCTURA.config_rutas import DOCUMENTOS_DIR
from DOMINIO.control_cancelacion import (
    establecer_evento_cancelacion,
    limpiar_evento_cancelacion,
    ProcesamientoCancelado,
)


class WorkflowService:

    def __init__(self):

        # =====================================================
        # CONTROL DE CANCELACIÓN
        # =====================================================

        self._cancel_event = threading.Event()

        # Callback opcional para informar progreso a la interfaz
        self._progress_callback = None


    # =========================================================
    # CANCELACIÓN
    # =========================================================

    def solicitar_cancelacion(self):
        """
        Solicita una cancelación segura.

        NO mata el hilo.
        NO interrumpe Excel a la fuerza.

        El archivo que se está procesando termina y después
        WorkflowService elimina las salidas generadas durante
        este lote y deja de procesar los siguientes ensayes.
        """

        print()
        print("=" * 80)
        print("SOLICITUD DE CANCELACIÓN RECIBIDA")
        print("El ensaye actual terminará de forma segura.")
        print("=" * 80)

        self._cancel_event.set()


    def cancelacion_solicitada(self):
        return self._cancel_event.is_set()


    def reiniciar_cancelacion(self):
        self._cancel_event.clear()

    def preparar_procesamiento(self):
        """Llamar desde la interfaz antes de lanzar el worker."""
        self.reiniciar_cancelacion()


    # =========================================================
    # PROGRESO
    # =========================================================

    def _notificar_progreso(
        self,
        actual,
        total,
        archivo="",
        gst="",
        etapa="Procesando...",
    ):

        if not self._progress_callback:
            return

        try:

            self._progress_callback({
                "actual": actual,
                "total": total,
                "archivo": archivo,
                "gst": gst,
                "etapa": etapa,
            })

        except Exception:
            # El progreso nunca debe detener el procesamiento.
            pass


    # =========================================================
    # ELIMINAR SALIDAS DEL LOTE CANCELADO
    # =========================================================

    @staticmethod
    def _eliminar_salidas_generadas(rutas_generadas):
        """
        Elimina únicamente los archivos Excel creados durante
        ESTA ejecución.

        Los mapas NO se tocan.
        Las plantillas NO se tocan.
        Los archivos ERP originales NO se tocan.
        """

        print()
        print("=" * 80)
        print("LIMPIANDO SALIDAS DEL LOTE CANCELADO")
        print("=" * 80)

        for ruta in rutas_generadas:

            if not ruta:
                continue

            try:

                ruta = os.path.abspath(
                    str(ruta)
                )

                if os.path.isfile(ruta):

                    os.remove(ruta)

                    print(
                        f"🗑️ Excel eliminado: {ruta}"
                    )

            except Exception as error:

                print(
                    "⚠️ No se pudo eliminar:"
                )

                print(ruta)

                print(
                    f"Motivo: {error}"
                )

        print("=" * 80)


    # =========================================================
    # OT Y CARPETA DE SALIDA
    # =========================================================

    @staticmethod
    def _obtener_numero_ot(item, nombre_archivo, nombre_carpeta=""):
        # Primero respetar un dato explícito si ya viene del parser/UI.
        for clave in ("ot", "numero_ot", "orden_trabajo"):
            valor = item.get(clave)
            if valor:
                m = re.search(r"\d+", str(valor))
                if m:
                    return m.group(0)

        textos = [
            str(nombre_archivo or ""),
            str(nombre_carpeta or ""),
        ]

        # Formato explícito OT-12345 / OT_12345 / OT 12345.
        for texto in textos:
            m = re.search(r"(?i)(?:^|[^A-Z0-9])OT[-_\s]*(\d+)", texto)
            if m:
                return m.group(1)

        # Formato actual:
        # E-1524_GST-042_17215_17215-1276_PCA-19.xls
        for texto in textos:
            m = re.search(r"(?i)GST-\d+_(\d+)(?:_|$)", texto)
            if m:
                return m.group(1)

        raise ValueError(
            "No se pudo identificar la OT del archivo: "
            f"{nombre_archivo}"
        )

    @staticmethod
    def _carpeta_salida_ot(numero_ot):
        ruta = os.path.join(
            DOCUMENTOS_DIR,
            f"OT-{numero_ot}",
        )
        os.makedirs(ruta, exist_ok=True)
        return ruta


    # =========================================================
    # PROCESAR LOTE
    # =========================================================

    def procesar_lote_carpetas(
        self,
        carpetas_dict: dict,
        progress_callback=None,
    ) -> dict:

        # -----------------------------------------------------
        # NUEVA EJECUCIÓN
        # -----------------------------------------------------

        # La señal se limpia en preparar_procesamiento(), ANTES del worker.
        # No se limpia aquí para no perder una cancelación temprana.
        self._progress_callback = progress_callback

        # Hace visible el Event a toda la cadena ejecutada en este worker.
        establecer_evento_cancelacion(self._cancel_event)

        # -----------------------------------------------------
        # VALIDAR
        # -----------------------------------------------------

        if not carpetas_dict:

            return {
                "status": "error",
                "mensaje": (
                    "No hay carpetas cargadas para procesar."
                ),
                "total_carpetas": 0,
                "total_archivos": 0,
                "excel_generados": 0,
                "pdfs_generados": 0,
                "errores": 0,
                "requisiciones": [],
                "expediente_final": None,
            }

        # =====================================================
        # CONTAR ARCHIVOS ANTES DE EMPEZAR
        # =====================================================

        total_archivos_lote = sum(
            len(lista_archivos)
            for lista_archivos
            in carpetas_dict.values()
        )

        archivos_procesados = 0
        excels_generados = 0
        errores = 0

        requisiciones_encontradas = set()

        archivos_exitosos = []
        archivos_con_error = []

        # IMPORTANTE:
        # Aquí registramos solamente los Excel creados
        # durante ESTA ejecución.
        #
        # Si se cancela, solamente éstos serán eliminados.
        salidas_generadas_lote = []

        # =====================================================
        # INICIO
        # =====================================================

        self._notificar_progreso(
            actual=0,
            total=total_archivos_lote,
            etapa="Preparando procesamiento...",
        )

        # =====================================================
        # RECORRER CARPETAS Y ARCHIVOS
        # =====================================================

        for nombre_carpeta, lista_archivos in (
            carpetas_dict.items()
        ):

            for item in lista_archivos:

                # =================================================
                # PUNTO SEGURO DE CANCELACIÓN
                #
                # Antes de comenzar el siguiente ensaye.
                # =================================================

                if self.cancelacion_solicitada():

                    self._eliminar_salidas_generadas(
                        salidas_generadas_lote
                    )

                    self._notificar_progreso(
                        actual=archivos_procesados,
                        total=total_archivos_lote,
                        etapa="Procesamiento cancelado.",
                    )

                    return {
                        "status": "cancelado",
                        "mensaje": (
                            "Procesamiento cancelado. "
                            "Los mapas generados fueron conservados "
                            "y los Excel de este lote fueron eliminados."
                        ),
                        "total_carpetas": len(
                            carpetas_dict
                        ),
                        "total_archivos": (
                            total_archivos_lote
                        ),
                        "archivos_procesados": (
                            archivos_procesados
                        ),
                        "excel_generados": 0,
                        "pdfs_generados": 0,
                        "errores": errores,
                        "requisiciones": sorted(
                            requisiciones_encontradas
                        ),
                        "archivos_exitosos": [],
                        "archivos_con_error": (
                            archivos_con_error
                        ),
                        "expediente_final": None,
                        "carpeta_salida": os.path.abspath(
                            DOCUMENTOS_DIR
                        ),
                    }

                # =================================================
                # ENSAYE ACTUAL
                # =================================================

                numero_actual = (
                    archivos_procesados + 1
                )

                req = item.get("req")

                if req:
                    requisiciones_encontradas.add(
                        str(req)
                    )

                ruta_archivo_origen = item.get(
                    "ruta"
                )

                codigo_gst = item.get(
                    "gst"
                )

                nombre_archivo = item.get(
                    "archivo",
                    "archivo_sin_nombre",
                )

                # =================================================
                # INFORMAR A LA INTERFAZ
                # =================================================

                self._notificar_progreso(
                    actual=numero_actual,
                    total=total_archivos_lote,
                    archivo=nombre_archivo,
                    gst=codigo_gst or "",
                    etapa="Preparando ensaye...",
                )

                print()
                print("=" * 80)
                print(
                    f"PROCESANDO ENSAYE "
                    f"{numero_actual} DE "
                    f"{total_archivos_lote}"
                )
                print(
                    f"{nombre_archivo} "
                    f"({codigo_gst or 'SIN GST'})"
                )
                print("=" * 80)

                # =================================================
                # VALIDACIONES BÁSICAS
                # =================================================

                if not ruta_archivo_origen:

                    errores += 1
                    archivos_procesados += 1

                    archivos_con_error.append({
                        "archivo": nombre_archivo,
                        "error": (
                            "No se recibió la ruta "
                            "del archivo."
                        ),
                    })

                    print(
                        f"❌ No se recibió ruta para "
                        f"{nombre_archivo}"
                    )

                    continue

                if not os.path.isfile(
                    ruta_archivo_origen
                ):

                    errores += 1
                    archivos_procesados += 1

                    archivos_con_error.append({
                        "archivo": nombre_archivo,
                        "error": (
                            "El archivo de origen no existe: "
                            f"{ruta_archivo_origen}"
                        ),
                    })

                    print(
                        "❌ No existe el archivo:"
                    )

                    print(
                        ruta_archivo_origen
                    )

                    continue

                if not codigo_gst:

                    errores += 1
                    archivos_procesados += 1

                    archivos_con_error.append({
                        "archivo": nombre_archivo,
                        "error": (
                            "No se detectó el código GST."
                        ),
                    })

                    print(
                        "❌ El archivo no tiene código GST:"
                    )

                    print(
                        nombre_archivo
                    )

                    continue

                # =================================================
                # PROCESAR ENSAYE
                # =================================================

                try:

                    # =============================================
                    # 1. PREPARAR NOMBRE DE SALIDA
                    # =============================================

                    nombre_base = os.path.splitext(
                        nombre_archivo
                    )[0]

                    numero_ot = self._obtener_numero_ot(
                        item,
                        nombre_archivo,
                        nombre_carpeta,
                    )

                    carpeta_salida_ot = self._carpeta_salida_ot(
                        numero_ot
                    )

                    ruta_xlsx_out = os.path.join(
                        carpeta_salida_ot,
                        f"{nombre_base}_procesado.xlsx",
                    )

                    print(
                        f"Carpeta de salida: {carpeta_salida_ot}"
                    )

                    print()
                    print(
                        "Archivo ERP original:"
                    )
                    print(
                        ruta_archivo_origen
                    )

                    print()
                    print(
                        f"Código GST: {codigo_gst}"
                    )

                    # =============================================
                    # 2. GENERAR EXCEL
                    # =============================================

                    self._notificar_progreso(
                        actual=numero_actual,
                        total=total_archivos_lote,
                        archivo=nombre_archivo,
                        gst=codigo_gst,
                        etapa=(
                            "Generando mapa y "
                            "procesando Excel..."
                        ),
                    )

                    # GeneradorExcel recibe directamente
                    # el archivo original.
                    #
                    # Aquí puede:
                    # - reutilizar un mapa V2.2
                    # - crear un mapa V2.2 con Gemini
                    # - validar el mapa
                    # - generar el Excel
                    #
                    # NO interrumpimos esta llamada.
                    #
                    # Si el usuario pulsa CANCELAR mientras
                    # está aquí, esperamos a que termine.

                    resultado_excel = (
                        GeneradorExcel
                        .inyectar_datos_dinamicos(
                            codigo_gst,
                            ruta_archivo_origen,
                            ruta_xlsx_out,
                        )
                    )

                    # =============================================
                    # 3. VALIDAR RESULTADO
                    # =============================================

                    if not resultado_excel:

                        raise RuntimeError(
                            "GeneradorExcel no devolvió "
                            "una ruta de salida."
                        )

                    resultado_excel = os.path.abspath(
                        str(resultado_excel)
                    )

                    if not os.path.isfile(
                        resultado_excel
                    ):

                        raise FileNotFoundError(
                            "GeneradorExcel indicó que terminó, "
                            "pero el archivo no existe:\n"
                            f"{resultado_excel}"
                        )

                    # =============================================
                    # REGISTRAR EXCEL COMO SALIDA DEL LOTE
                    # =============================================

                    salidas_generadas_lote.append(
                        resultado_excel
                    )

                    excels_generados += 1
                    archivos_procesados += 1

                    archivos_exitosos.append({
                        "archivo": nombre_archivo,
                        "gst": codigo_gst,
                        "ruta_excel": resultado_excel,
                    })

                    print()
                    print(
                        "✅ Excel generado y validado:"
                    )

                    print(
                        resultado_excel
                    )

                    # =============================================
                    # CANCELACIÓN SOLICITADA DURANTE ESTE ENSAYE
                    # =============================================

                    if self.cancelacion_solicitada():

                        print()
                        print(
                            "⛔ Cancelación detectada "
                            "después de cerrar el ensaye."
                        )

                        self._notificar_progreso(
                            actual=numero_actual,
                            total=total_archivos_lote,
                            archivo=nombre_archivo,
                            gst=codigo_gst,
                            etapa=(
                                "Eliminando archivos "
                                "del lote cancelado..."
                            ),
                        )

                        # El mapa NO está dentro de esta lista.
                        # Por tanto se conserva.
                        self._eliminar_salidas_generadas(
                            salidas_generadas_lote
                        )

                        return {
                            "status": "cancelado",
                            "mensaje": (
                                "Procesamiento cancelado. "
                                "El mapeo fue conservado y "
                                "los Excel generados durante "
                                "este lote fueron eliminados."
                            ),
                            "total_carpetas": len(
                                carpetas_dict
                            ),
                            "total_archivos": (
                                total_archivos_lote
                            ),
                            "archivos_procesados": (
                                archivos_procesados
                            ),
                            "excel_generados": 0,
                            "pdfs_generados": 0,
                            "errores": errores,
                            "requisiciones": sorted(
                                requisiciones_encontradas
                            ),
                            "archivos_exitosos": [],
                            "archivos_con_error": (
                                archivos_con_error
                            ),
                            "expediente_final": None,
                            "carpeta_salida": os.path.abspath(
                                DOCUMENTOS_DIR
                            ),
                        }

                    # =============================================
                    # ENSAYE TERMINADO
                    # =============================================

                    self._notificar_progreso(
                        actual=numero_actual,
                        total=total_archivos_lote,
                        archivo=nombre_archivo,
                        gst=codigo_gst,
                        etapa="Ensaye terminado.",
                    )

                except Exception as error:

                    errores += 1
                    archivos_procesados += 1

                    archivos_con_error.append({
                        "archivo": nombre_archivo,
                        "gst": codigo_gst,
                        "error": str(error),
                    })

                    print()
                    print(
                        f"❌ ERROR AL PROCESAR "
                        f"{nombre_archivo}:"
                    )

                    traceback.print_exc()

                    # =============================================
                    # SI CANCELARON MIENTRAS OCURRIÓ EL ERROR
                    # =============================================

                    if self.cancelacion_solicitada():

                        self._eliminar_salidas_generadas(
                            salidas_generadas_lote
                        )

                        return {
                            "status": "cancelado",
                            "mensaje": (
                                "Procesamiento cancelado. "
                                "Los mapas generados fueron "
                                "conservados y los Excel del "
                                "lote fueron eliminados."
                            ),
                            "total_carpetas": len(
                                carpetas_dict
                            ),
                            "total_archivos": (
                                total_archivos_lote
                            ),
                            "archivos_procesados": (
                                archivos_procesados
                            ),
                            "excel_generados": 0,
                            "pdfs_generados": 0,
                            "errores": errores,
                            "requisiciones": sorted(
                                requisiciones_encontradas
                            ),
                            "archivos_exitosos": [],
                            "archivos_con_error": (
                                archivos_con_error
                            ),
                            "expediente_final": None,
                            "carpeta_salida": os.path.abspath(
                                DOCUMENTOS_DIR
                            ),
                        }

                    # Si un archivo falla,
                    # continúa con el siguiente.
                    continue

        # =====================================================
        # DETERMINAR ESTADO FINAL
        # =====================================================

        if archivos_procesados == 0:

            estado = "error"

            mensaje = (
                "No se encontraron archivos "
                "para procesar."
            )

        elif (
            excels_generados
            ==
            archivos_procesados
        ):

            estado = "exito"

            mensaje = (
                "Proceso terminado correctamente. "
                f"Se generaron {excels_generados} "
                f"de {archivos_procesados} "
                "archivos Excel."
            )

        elif excels_generados > 0:

            estado = "parcial"

            mensaje = (
                "Proceso terminado parcialmente. "
                f"Se generaron {excels_generados} "
                f"de {archivos_procesados} "
                "archivos Excel. "
                f"Archivos con error: {errores}."
            )

        else:

            estado = "error"

            mensaje = (
                "No se pudo generar ningún archivo Excel. "
                f"Archivos procesados: "
                f"{archivos_procesados}. "
                f"Errores: {errores}. "
                "Revisa la consola para conocer "
                "los detalles."
            )

        # =====================================================
        # RESUMEN
        # =====================================================

        print()
        print("=" * 80)
        print(
            "RESUMEN DEL PROCESAMIENTO"
        )
        print("=" * 80)

        print(
            f"Carpetas recibidas: "
            f"{len(carpetas_dict)}"
        )

        print(
            f"Archivos procesados: "
            f"{archivos_procesados}"
        )

        print(
            f"Excel generados: "
            f"{excels_generados}"
        )

        print(
            f"Errores: {errores}"
        )

        print("=" * 80)

        # =====================================================
        # RESPUESTA
        # =====================================================

        return {
            "status": estado,
            "total_carpetas": len(
                carpetas_dict
            ),
            "total_archivos": archivos_procesados,
            "excel_generados": excels_generados,
            "pdfs_generados": 0,
            "expediente_final": None,
            "errores": errores,
            "requisiciones": sorted(
                requisiciones_encontradas
            ),
            "archivos_exitosos": (
                archivos_exitosos
            ),
            "archivos_con_error": (
                archivos_con_error
            ),
            "carpeta_salida": os.path.abspath(
                DOCUMENTOS_DIR
            ),
            "mensaje": mensaje,
        }