# PRESENTACION/resumen_procesamiento.py

from __future__ import annotations

import os
import re
import zipfile
import tempfile
import uuid

from nicegui import ui


class ResumenProcesamientoUI:
    """Pantalla final de resultados del procesamiento."""

    # ============================================================
    # MENSAJES AMIGABLES
    # ============================================================

    @staticmethod
    def _mensaje_amigable(
        error: str,
        gst: str = "",
    ) -> str:

        texto = str(error or "")
        bajo = texto.lower()
        gst_txt = str(gst or "").strip()

        # --------------------------------------------------------
        # PLANTILLA NO ENCONTRADA
        # --------------------------------------------------------

        if (
            "no se encontró ninguna plantilla" in bajo
            or "no se encontro ninguna plantilla" in bajo
            or (
                "no existe" in bajo
                and "plantilla" in bajo
            )
        ):

            return (
                f"No se encontró una plantilla oficial "
                f"disponible para "
                f"{gst_txt or 'este tipo de ensaye'}."
            )

        # --------------------------------------------------------
        # ÁREA ERP INCOMPATIBLE
        # --------------------------------------------------------

        if (
            (
                "área" in bajo
                and (
                    "erp" in bajo
                    or "pegado" in bajo
                )
            )
            or (
                "area" in bajo
                and (
                    "erp" in bajo
                    or "pegado" in bajo
                )
            )
            or "zona segura" in bajo
            or "zona de captura" in bajo
            or "hoja erp" in bajo
            or (
                "no contiene" in bajo
                and "erp" in bajo
            )
        ):

            return (
                "La plantilla oficial no contiene un área "
                "compatible para colocar los resultados "
                "del ERP."
            )

        # --------------------------------------------------------
        # GST
        # --------------------------------------------------------

        if (
            "no se detectó el código gst" in bajo
            or
            "no se detecto el código gst" in bajo
        ):

            return (
                "No fue posible identificar el tipo de "
                "ensaye (código GST)."
            )

        # --------------------------------------------------------
        # RUTA
        # --------------------------------------------------------

        if (
            "no se recibió la ruta" in bajo
            or
            "no se recibio la ruta" in bajo
        ):

            return (
                "No fue posible localizar el archivo "
                "seleccionado para procesarlo."
            )

        # --------------------------------------------------------
        # ARCHIVO ORIGINAL
        # --------------------------------------------------------

        if (
            "archivo de origen no existe" in bajo
            or
            "el archivo de origen no existe" in bajo
        ):

            return (
                "El archivo original ya no está disponible "
                "para procesarlo."
            )

        # --------------------------------------------------------
        # OT
        # --------------------------------------------------------

        if "no se pudo identificar la ot" in bajo:

            return (
                "No fue posible identificar la Orden "
                "de Trabajo del archivo."
            )

        # --------------------------------------------------------
        # MAPEO
        # --------------------------------------------------------

        if (
            "mapa" in bajo
            and (
                "incompatible" in bajo
                or "valid" in bajo
                or "estructura" in bajo
                or "reconstru" in bajo
            )
        ):

            return (
                "La estructura del ensaye no es compatible "
                "con el formato de captura de la plantilla "
                "oficial."
            )

        # --------------------------------------------------------
        # FÓRMULAS
        # --------------------------------------------------------

        if (
            "formula" in bajo
            or
            "fórmula" in bajo
        ):

            return (
                "La plantilla contiene una fórmula protegida "
                "en una zona destinada a los datos del ERP. "
                "El archivo no fue modificado."
            )

        # --------------------------------------------------------
        # PERMISOS
        # --------------------------------------------------------

        if (
            "permission" in bajo
            or "permiso" in bajo
            or "access is denied" in bajo
        ):

            return (
                "No fue posible guardar temporalmente el "
                "archivo porque el servidor no concedió "
                "acceso a la ubicación."
            )

        # --------------------------------------------------------
        # ERROR GENÉRICO
        # --------------------------------------------------------

        return (
            "No fue posible generar este reporte. "
            "El archivo original se conserva sin cambios. "
            "Si el problema continúa, contacta al "
            "responsable del sistema."
        )

    # ============================================================
    # OBTENER OT
    # ============================================================

    @staticmethod
    def _obtener_ot(
        resultado: dict,
    ) -> str:

        # --------------------------------------------------------
        # PRIMERA OPCIÓN:
        # CARPETA REAL DE UN EXCEL GENERADO
        # --------------------------------------------------------

        for item in (
            resultado.get(
                "archivos_exitosos",
                [],
            )
            or []
        ):

            ruta = str(
                item.get(
                    "ruta_excel",
                    "",
                )
            )

            m = re.search(
                r"(?i)(?:^|[\\/])OT-(\d+)(?:[\\/]|$)",
                ruta,
            )

            if m:

                return (
                    f"OT-{m.group(1)}"
                )

        # --------------------------------------------------------
        # RESPALDO:
        # NOMBRE DEL ARCHIVO ERP
        # --------------------------------------------------------

        items = (
            list(
                resultado.get(
                    "archivos_exitosos",
                    [],
                )
                or []
            )
            +
            list(
                resultado.get(
                    "archivos_con_error",
                    [],
                )
                or []
            )
        )

        for item in items:

            nombre = str(
                item.get(
                    "archivo",
                    "",
                )
            )

            # Ejemplo:
            # 17215_17215-1276

            m = re.search(
                r"(?<!\d)(\d{3,})_\1-\d+",
                nombre,
            )

            if m:

                return (
                    f"OT-{m.group(1)}"
                )

            # Ejemplo:
            # OT-17215

            m = re.search(
                r"(?i)(?:^|[^A-Z0-9])"
                r"OT[-_\s]*(\d+)",
                nombre,
            )

            if m:

                return (
                    f"OT-{m.group(1)}"
                )

        return "Orden_de_Trabajo"

    # ============================================================
    # CREAR ZIP
    # ============================================================

    @staticmethod
    def _crear_zip_resultados(
        resultado: dict,
    ) -> str | None:
        """
        Crea un ZIP temporal en el servidor con todos
        los Excel generados correctamente.

        El ZIP posteriormente será enviado al navegador
        del usuario mediante ui.download().
        """

        exitosos = list(
            resultado.get(
                "archivos_exitosos",
                [],
            )
            or []
        )

        if not exitosos:

            return None

        # --------------------------------------------------------
        # NOMBRE DEL ZIP
        # --------------------------------------------------------

        ot = (
            ResumenProcesamientoUI
            ._obtener_ot(
                resultado
            )
        )

        nombre_seguro = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            ot,
        )

        identificador = (
            uuid.uuid4().hex[:8]
        )

        nombre_zip = (
            f"{nombre_seguro}_"
            f"Resultados_"
            f"{identificador}.zip"
        )

        ruta_zip = os.path.join(
            tempfile.gettempdir(),
            nombre_zip,
        )

        # --------------------------------------------------------
        # CREAR ZIP
        # --------------------------------------------------------

        archivos_agregados = 0

        with zipfile.ZipFile(
            ruta_zip,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
        ) as zipf:

            for item in exitosos:

                ruta_excel = str(
                    item.get(
                        "ruta_excel",
                        "",
                    )
                    or ""
                )

                if not ruta_excel:
                    continue

                ruta_excel = os.path.abspath(
                    ruta_excel
                )

                if not os.path.isfile(
                    ruta_excel
                ):
                    continue

                # --------------------------------------------
                # Nombre dentro del ZIP
                # --------------------------------------------

                nombre_excel = os.path.basename(
                    ruta_excel
                )

                zipf.write(
                    ruta_excel,
                    arcname=nombre_excel,
                )

                archivos_agregados += 1

        # --------------------------------------------------------
        # SI NO SE AGREGÓ NADA
        # --------------------------------------------------------

        if archivos_agregados == 0:

            try:

                if os.path.isfile(
                    ruta_zip
                ):

                    os.remove(
                        ruta_zip
                    )

            except Exception:

                pass

            return None

        return ruta_zip

    # ============================================================
    # DESCARGAR RESULTADOS
    # ============================================================

    @staticmethod
    def _descargar_resultados(
        resultado: dict,
    ):
        """
        Crea el ZIP temporalmente en el servidor y
        NiceGUI lo envía al navegador del usuario.

        El navegador cliente decide dónde guardar
        el archivo.
        """

        try:

            # ----------------------------------------------------
            # CREAR ZIP
            # ----------------------------------------------------

            ruta_zip = (
                ResumenProcesamientoUI
                ._crear_zip_resultados(
                    resultado
                )
            )

            if not ruta_zip:

                ui.notify(
                    "No hay archivos disponibles "
                    "para descargar.",
                    type="warning",
                )

                return

            # ----------------------------------------------------
            # NOMBRE PARA EL CLIENTE
            # ----------------------------------------------------

            ot = (
                ResumenProcesamientoUI
                ._obtener_ot(
                    resultado
                )
            )

            nombre_descarga = (
                f"{ot}_Resultados.zip"
            )

            # ----------------------------------------------------
            # DESCARGA HTTP AL NAVEGADOR
            # ----------------------------------------------------

            ui.download(
                ruta_zip,
                filename=nombre_descarga,
            )

            ui.notify(
                "La descarga de los resultados "
                "ha comenzado.",
                type="positive",
            )

        except Exception as error:

            print()
            print("=" * 80)
            print(
                "ERROR PREPARANDO DESCARGA"
            )
            print("=" * 80)
            print(error)
            print("=" * 80)

            ui.notify(
                "No fue posible preparar la descarga "
                "de los resultados.",
                type="negative",
                duration=7000,
            )

    # ============================================================
    # MOSTRAR RESUMEN
    # ============================================================

    @staticmethod
    def mostrar(
        resultado: dict,
        total_carpetas: int | None = None,
        on_close=None,
    ):

        # --------------------------------------------------------
        # RESULTADOS
        # --------------------------------------------------------

        exitosos = list(
            resultado.get(
                "archivos_exitosos",
                [],
            )
            or []
        )

        errores = list(
            resultado.get(
                "archivos_con_error",
                [],
            )
            or []
        )

        # --------------------------------------------------------
        # MÉTRICAS
        # --------------------------------------------------------

        procesados = int(
            resultado.get(
                "total_archivos",
                len(exitosos)
                +
                len(errores),
            )
            or 0
        )

        excel_generados = int(
            resultado.get(
                "excel_generados",
                len(exitosos),
            )
            or 0
        )

        cantidad_errores = int(
            resultado.get(
                "errores",
                len(errores),
            )
            or 0
        )

        carpetas = (
            int(total_carpetas)
            if total_carpetas is not None
            else int(
                resultado.get(
                    "total_carpetas",
                    0,
                )
                or 0
            )
        )

        # --------------------------------------------------------
        # OT Y ESTADO
        # --------------------------------------------------------

        ot = (
            ResumenProcesamientoUI
            ._obtener_ot(
                resultado
            )
        )

        estado = str(
            resultado.get(
                "status",
                "",
            )
        ).lower()

        # ========================================================
        # DIÁLOGO
        # ========================================================

        dialogo = (
            ui.dialog()
            .props(
                "persistent"
            )
        )

        with dialogo:

            with ui.card().classes(
                "w-[900px] "
                "max-w-[96vw] "
                "max-h-[92vh] "
                "p-0 "
                "overflow-hidden "
                "geotest-card"
            ):

                # =================================================
                # CABECERA
                # =================================================

                with ui.column().classes(
                    "w-full relative overflow-hidden bg-white "
                    "px-7 py-6 gap-1 border-b border-[#DFE4EF]"
                ):
                    # Franja del ERP: tres bloques verticales.
                    with ui.element("div").style(
                        "position:absolute;top:0;bottom:0;left:0;width:5px;"
                        "display:flex;flex-direction:column;pointer-events:none;"
                    ):
                        ui.element("div").style("height:34%;background:#3045B4;")
                        ui.element("div").style("height:33%;background:#95A9EF;")
                        ui.element("div").style("height:33%;background:#E23742;")

                    ui.label("LABORATORIO · GEOTEST").classes(
                        "text-xs font-bold tracking-[0.18em] text-[#253B83]"
                    )
                    ui.label("Resumen del procesamiento").classes(
                        "text-2xl font-bold text-[#202938]"
                    )
                    ui.label(ot).classes(
                        "text-sm font-semibold text-[#64748B]"
                    )

                # =================================================
                # CONTENIDO
                # =================================================

                with ui.scroll_area().classes(
                    "w-full h-[68vh]"
                ):

                    with ui.column().classes(
                        "w-full p-6 gap-5 bg-[#F8FAFF]"
                    ):

                        # =========================================
                        # MÉTRICAS
                        # =========================================

                        with ui.row().classes(
                            "w-full "
                            "grid "
                            "grid-cols-2 "
                            "md:grid-cols-4 "
                            "gap-3"
                        ):

                            (
                                ResumenProcesamientoUI
                                ._tarjeta_metrica(
                                    "Carpetas recibidas",
                                    carpetas,
                                    "folder",
                                )
                            )

                            (
                                ResumenProcesamientoUI
                                ._tarjeta_metrica(
                                    "Archivos procesados",
                                    procesados,
                                    "description",
                                )
                            )

                            (
                                ResumenProcesamientoUI
                                ._tarjeta_metrica(
                                    "Excel generados",
                                    excel_generados,
                                    "check_circle",
                                )
                            )

                            (
                                ResumenProcesamientoUI
                                ._tarjeta_metrica(
                                    "Incidencias",
                                    cantidad_errores,
                                    "warning",
                                )
                            )

                        # =========================================
                        # ESTADO
                        # =========================================

                        if estado == "exito":

                            ui.label(
                                "El procesamiento terminó "
                                "correctamente."
                            ).classes(
                                "text-sm "
                                "text-green-800 "
                                "bg-green-50 "
                                "border "
                                "border-green-200 "
                                "rounded-lg "
                                "p-3 "
                                "w-full"
                            )

                        elif estado == "parcial":

                            ui.label(
                                "El procesamiento terminó "
                                "con algunas incidencias. "
                                "Puedes revisar el detalle "
                                "a continuación."
                            ).classes(
                                "text-sm "
                                "text-amber-800 "
                                "bg-amber-50 "
                                "border "
                                "border-amber-200 "
                                "rounded-lg "
                                "p-3 "
                                "w-full"
                            )

                        else:

                            ui.label(
                                "El procesamiento terminó "
                                "con incidencias. "
                                "Revisa el detalle "
                                "a continuación."
                            ).classes(
                                "text-sm "
                                "text-red-800 "
                                "bg-red-50 "
                                "border "
                                "border-red-200 "
                                "rounded-lg "
                                "p-3 "
                                "w-full"
                            )

                        # =========================================
                        # EXCEL GENERADOS
                        # =========================================

                        with ui.expansion(
                            (
                                f"Excel generados "
                                f"({len(exitosos)})"
                            ),
                            icon="check_circle",
                            value=bool(
                                exitosos
                                and
                                not errores
                            ),
                        ).classes(
                            "w-full "
                            "border "
                            "border-[#DFE4EF] "
                            "rounded-xl "
                            "bg-white "
                            "text-[#253B83] "
                            "font-semibold"
                        ):

                            with ui.column().classes(
                                "w-full "
                                "bg-white "
                                "p-3 "
                                "gap-2"
                            ):

                                if not exitosos:

                                    ui.label(
                                        "No se generaron "
                                        "archivos Excel."
                                    ).classes(
                                        "text-sm "
                                        "text-gray-500"
                                    )

                                for item in exitosos:

                                    (
                                        ResumenProcesamientoUI
                                        ._fila_exito(
                                            item
                                        )
                                    )

                        # =========================================
                        # INCIDENCIAS
                        # =========================================

                        with ui.expansion(
                            (
                                f"Incidencias "
                                f"({len(errores)})"
                            ),
                            icon="warning",
                            value=bool(
                                errores
                            ),
                        ).classes(
                            "w-full "
                            "border "
                            "border-[#DFE4EF] "
                            "rounded-xl "
                            "bg-white "
                            "text-[#253B83] "
                            "font-semibold"
                        ):

                            with ui.column().classes(
                                "w-full "
                                "bg-white "
                                "p-3 "
                                "gap-2"
                            ):

                                if not errores:

                                    ui.label(
                                        "No se registraron "
                                        "incidencias."
                                    ).classes(
                                        "text-sm "
                                        "text-gray-500"
                                    )

                                for item in errores:

                                    (
                                        ResumenProcesamientoUI
                                        ._fila_error(
                                            item
                                        )
                                    )

                # =================================================
                # CERRAR
                # =================================================

                def cerrar_resumen():

                    dialogo.close()

                    if on_close:

                        try:

                            on_close()

                        except Exception as error:

                            print(
                                "⚠️ No se pudo limpiar "
                                "la vista después del "
                                f"resumen: {error}"
                            )

                # =================================================
                # BOTONES INFERIORES
                # =================================================

                with ui.row().classes(
                    "w-full "
                    "justify-between "
                    "items-center "
                    "border-t "
                    "border-gray-200 "
                    "px-6 py-4"
                ):

                    # ---------------------------------------------
                    # DESCARGAR RESULTADOS
                    # ---------------------------------------------

                    if exitosos:

                        ui.button(
                            "DESCARGAR RESULTADOS",
                            icon="download",
                            on_click=lambda: (
                                ResumenProcesamientoUI
                                ._descargar_resultados(
                                    resultado
                                )
                            ),
                        ).classes(
                            "text-white font-bold px-6 py-2 rounded-xl"
                        ).style(
                            "background-color:#253B83 !important;"
                        )

                    else:

                        # Mantiene CERRAR alineado a la derecha
                        ui.space()

                    # ---------------------------------------------
                    # CERRAR
                    # ---------------------------------------------

                    ui.button(
                        "CERRAR",
                        on_click=cerrar_resumen,
                        icon="close",
                    ).props(
                        "outline color=secondary no-caps"
                    ).classes(
                        "px-6 py-2 rounded-xl"
                    )

        # ========================================================
        # ABRIR
        # ========================================================

        dialogo.open()

        return dialogo

    # ============================================================
    # TARJETA MÉTRICA
    # ============================================================

    @staticmethod
    def _tarjeta_metrica(
        titulo: str,
        valor: int,
        icono: str,
    ):

        with ui.card().classes(
            "min-w-0 "
            "p-4 "
            "shadow-none "
            "border "
            "border-[#DFE4EF] "
            "rounded-xl bg-white"
        ):

            with ui.row().classes(
                "items-center "
                "gap-3 "
                "no-wrap"
            ):

                ui.icon(
                    icono
                ).classes(
                    "text-2xl "
                    "text-[#253B83]"
                )

                with ui.column().classes(
                    "gap-0 min-w-0"
                ):

                    ui.label(
                        str(valor)
                    ).classes(
                        "text-2xl "
                        "font-bold "
                        "text-gray-900"
                    )

                    ui.label(
                        titulo
                    ).classes(
                        "text-xs "
                        "text-gray-500"
                    )

    # ============================================================
    # FILA DE ÉXITO
    # ============================================================

    @staticmethod
    def _fila_exito(
        item: dict,
    ):

        nombre = str(
            item.get(
                "archivo",
                "Archivo procesado",
            )
        )

        gst = str(
            item.get(
                "gst",
                "",
            )
            or ""
        )

        with ui.row().classes(
            "w-full "
            "items-center "
            "justify-between "
            "gap-3 "
            "border "
            "border-[#DFE4EF] "
            "rounded-xl bg-white "
            "p-3"
        ):

            with ui.column().classes(
                "gap-0 "
                "flex-1 "
                "min-w-0"
            ):

                ui.label(
                    nombre
                ).classes(
                    "text-sm "
                    "font-medium "
                    "text-gray-800 "
                    "break-all"
                )

                if gst:

                    ui.label(
                        gst
                    ).classes(
                        "text-xs "
                        "text-gray-500"
                    )

            # IMPORTANTE:
            #
            # Ya NO existe el botón "Abrir".
            #
            # Antes ese botón ejecutaba os.startfile()
            # y por eso abría Excel en el servidor.
            #
            # Ahora todos los resultados se descargan
            # mediante el ZIP al navegador del cliente.

    # ============================================================
    # FILA DE ERROR
    # ============================================================

    @staticmethod
    def _fila_error(
        item: dict,
    ):

        nombre = str(
            item.get(
                "archivo",
                "Archivo",
            )
        )

        gst = str(
            item.get(
                "gst",
                "",
            )
            or ""
        )

        mensaje = (
            ResumenProcesamientoUI
            ._mensaje_amigable(
                item.get(
                    "error",
                    "",
                ),
                gst,
            )
        )

        with ui.column().classes(
            "w-full "
            "gap-1 "
            "border "
            "border-amber-200 "
            "rounded-lg "
            "p-3 "
            "bg-amber-50"
        ):

            with ui.row().classes(
                "w-full "
                "items-center "
                "gap-2"
            ):

                ui.icon(
                    "warning"
                ).classes(
                    "text-amber-700"
                )

                ui.label(
                    nombre
                ).classes(
                    "text-sm "
                    "font-semibold "
                    "text-gray-800 "
                    "break-all"
                )

            if gst:

                ui.label(
                    gst
                ).classes(
                    "text-xs "
                    "font-medium "
                    "text-gray-500 "
                    "ml-7"
                )

            ui.label(
                mensaje
            ).classes(
                "text-sm "
                "text-gray-700 "
                "ml-7"
            )
