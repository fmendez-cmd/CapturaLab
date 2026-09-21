# PRESENTACION/componentes_ui.py
from nicegui import ui
import os


class ComponentesUI:
    @staticmethod
    def renderizar_acordeon_carpetas(carpetas_dict: dict, callback_eliminar):
        if not carpetas_dict:
            ui.label(
                "No hay carpetas cargadas. Arrastra carpetas arriba o usa el botón de selección."
            ).classes("text-gray-500 italic p-4")
            return

        for nombre_carpeta, lista_archivos in carpetas_dict.items():
            num_archivos = len(lista_archivos)

            with ui.expansion(
                f" {nombre_carpeta} ({num_archivos} archivos)", icon="folder"
            ).classes(
                "w-full bg-blue-50 border border-blue-200 rounded-lg mb-2 text-blue-900 font-semibold"
            ):

                with ui.column().classes("w-full p-2 bg-white gap-1"):
                    columnas = [
                        {
                            "name": "archivo",
                            "label": "Nombre del Archivo",
                            "field": "archivo",
                            "align": "left",
                        },
                        {
                            "name": "req",
                            "label": "Requisición",
                            "field": "req",
                            "align": "center",
                        },
                        {
                            "name": "gst",
                            "label": "Código GST",
                            "field": "gst",
                            "align": "center",
                        },
                        {
                            "name": "ensaye",
                            "label": "Prueba / Ensaye",
                            "field": "ensaye",
                            "align": "left",
                        },
                        {
                            "name": "formato",
                            "label": "Plantilla Oficial",
                            "field": "formato",
                            "align": "center",
                        },
                    ]

                    ui.table(
                        columns=columnas, rows=lista_archivos, row_key="archivo"
                    ).classes("w-full border")

                    with ui.row().classes("w-full justify-end mt-2"):
                        ui.button(
                            "Quitar Carpeta",
                            on_click=lambda c=nombre_carpeta: callback_eliminar(c),
                            icon="delete",
                        ).props("outline color=red size=sm")
