import os
import sys
import base64
import asyncio

from nicegui import ui, run

import tkinter as tk
from tkinter import filedialog


# ============================================================
# IMPORTACIONES DEL PROYECTO
# ============================================================

sys.path.append(
    os.path.dirname(
        os.path.dirname(
            os.path.abspath(__file__)
        )
    )
)

from DOMINIO.parser_erp import ParserERP
from APLICACION.workflow_service import WorkflowService
from PRESENTACION.componentes_ui import ComponentesUI


# ============================================================
# ESTADO GLOBAL
# ============================================================

carpetas_dict = {}

workflow_service = WorkflowService()

procesando = False


# ============================================================
# CARPETA DE ENTRADA
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.dirname(
        os.path.abspath(__file__)
    )
)

ENTRADA_DIR = os.path.join(
    BASE_DIR,
    "ARCHIVOS_ENTRADA"
)

os.makedirs(
    ENTRADA_DIR,
    exist_ok=True
)


# ============================================================
# RECIBIR ARCHIVO DESDE DRAG & DROP
# ============================================================

def recibir_un_archivo(e):

    try:

        if procesando:
            return

        data_item = e.args

        nombre_carpeta = data_item.get(
            "folder",
            "Carpeta Importada"
        )

        filename = data_item.get(
            "name",
            ""
        )

        contenido_b64 = data_item.get(
            "content",
            ""
        )

        if not filename.lower().endswith(
            (".xls", ".xlsx")
        ):
            return

        dir_destino = os.path.join(
            ENTRADA_DIR,
            nombre_carpeta
        )

        os.makedirs(
            dir_destino,
            exist_ok=True
        )

        ruta_fisica_real = os.path.join(
            dir_destino,
            filename
        )

        if contenido_b64:

            header, data = (
                contenido_b64.split(",", 1)
                if "," in contenido_b64
                else ("", contenido_b64)
            )

            with open(
                ruta_fisica_real,
                "wb"
            ) as f:

                f.write(
                    base64.b64decode(data)
                )

        if nombre_carpeta not in carpetas_dict:

            carpetas_dict[
                nombre_carpeta
            ] = []

        info = ParserERP.extraer_metadatos_archivo(
            filename,
            ruta_fisica_real
        )

        existe = any(
            archivo["archivo"] == filename
            for archivo
            in carpetas_dict[nombre_carpeta]
        )

        if not existe:

            carpetas_dict[
                nombre_carpeta
            ].append(info)

        actualizar_vista_carpetas()

    except Exception as error:

        print(
            "\nERROR AL RECIBIR ARCHIVO:"
        )

        print(error)

        ui.notify(
            f"Error al cargar archivo: {error}",
            type="negative",
        )


# ============================================================
# SELECCIONAR ARCHIVOS
# ============================================================

def abrir_dialogo_archivos():

    if procesando:
        return

    root = tk.Tk()

    root.withdraw()

    root.attributes(
        "-topmost",
        True
    )

    archivos = filedialog.askopenfilenames(
        title="Seleccionar archivos del ERP",
        filetypes=[
            (
                "Archivos ERP",
                "*.xls *.xlsx"
            ),
            (
                "Todos",
                "*.*"
            ),
        ],
    )

    root.destroy()

    if not archivos:
        return

    for ruta_file in archivos:

        carpeta_path = os.path.dirname(
            os.path.normpath(
                ruta_file
            )
        )

        nombre_carpeta = os.path.basename(
            carpeta_path
        )

        if nombre_carpeta not in carpetas_dict:

            carpetas_dict[
                nombre_carpeta
            ] = []

        # Se mantiene el comportamiento actual:
        # al seleccionar un archivo se cargan los Excel
        # contenidos en esa carpeta.

        for archivo in os.listdir(
            carpeta_path
        ):

            if not archivo.lower().endswith(
                (".xls", ".xlsx")
            ):
                continue

            ruta_completa = os.path.join(
                carpeta_path,
                archivo
            )

            info = ParserERP.extraer_metadatos_archivo(
                archivo,
                ruta_completa
            )

            existe = any(
                item["archivo"] == archivo
                for item
                in carpetas_dict[nombre_carpeta]
            )

            if not existe:

                carpetas_dict[
                    nombre_carpeta
                ].append(info)

    actualizar_vista_carpetas()

    ui.notify(
        "Carpeta(s) cargada(s) con éxito",
        type="positive",
    )


# ============================================================
# ACTUALIZAR CARPETAS
# ============================================================

def actualizar_vista_carpetas():

    contenedor_carpetas.clear()

    with contenedor_carpetas:

        ComponentesUI.renderizar_acordeon_carpetas(
            carpetas_dict,
            eliminar_carpeta
        )


# ============================================================
# ELIMINAR CARPETA
# ============================================================

def eliminar_carpeta(
    nombre_carpeta
):

    if procesando:
        return

    if nombre_carpeta in carpetas_dict:

        del carpetas_dict[
            nombre_carpeta
        ]

        actualizar_vista_carpetas()

        ui.notify(
            "Carpeta removida",
            type="info",
        )


# ============================================================
# LIMPIAR
# ============================================================

def limpiar_todo():

    if procesando:
        return

    carpetas_dict.clear()

    actualizar_vista_carpetas()

    ui.notify(
        "Pantalla limpiada",
        type="info",
    )


def limpiar_sin_notificacion():

    carpetas_dict.clear()

    actualizar_vista_carpetas()


# ============================================================
# CANCELAR
# ============================================================

def cancelar_procesamiento():

    if not procesando:
        return

    # No matamos el hilo.
    # Sólo avisamos a WorkflowService.

    workflow_service.solicitar_cancelacion()

    titulo_modal.set_text(
        "Cancelando procesamiento..."
    )

    etapa_modal.set_text(
        "Esperando a que el ensaye actual "
        "termine de forma segura."
    )

    detalle_modal.set_text(
        "Los mapas generados se conservarán. "
        "Los Excel de este lote serán eliminados."
    )

    boton_cancelar.disable()

    spinner_modal.set_visibility(
        True
    )


# ============================================================
# PROGRESO DESDE WORKFLOW
# ============================================================

def crear_callback_progreso(loop):

    def callback(datos):

        # Este callback se ejecuta desde el worker thread.
        #
        # NO debemos modificar NiceGUI directamente desde
        # ese hilo.
        #
        # Mandamos la actualización al event loop principal.

        try:

            asyncio.run_coroutine_threadsafe(
                actualizar_modal_progreso(
                    datos
                ),
                loop,
            )

        except Exception:
            pass

    return callback


async def actualizar_modal_progreso(
    datos
):

    actual = datos.get(
        "actual",
        0
    )

    total = datos.get(
        "total",
        0
    )

    archivo = datos.get(
        "archivo",
        ""
    )

    gst = datos.get(
        "gst",
        ""
    )

    etapa = datos.get(
        "etapa",
        "Procesando..."
    )

    # --------------------------------------------------------
    # TEXTO PRINCIPAL
    # --------------------------------------------------------

    if actual > 0 and total > 0:

        titulo_modal.set_text(
            f"Procesando ensaye "
            f"{actual} de {total}"
        )

    else:

        titulo_modal.set_text(
            "Preparando procesamiento..."
        )

    # --------------------------------------------------------
    # GST
    # --------------------------------------------------------

    if gst:

        gst_modal.set_text(
            gst
        )

    else:

        gst_modal.set_text(
            ""
        )

    # --------------------------------------------------------
    # ARCHIVO
    # --------------------------------------------------------

    if archivo:

        archivo_modal.set_text(
            archivo
        )

    else:

        archivo_modal.set_text(
            ""
        )

    # --------------------------------------------------------
    # ETAPA
    # --------------------------------------------------------

    etapa_modal.set_text(
        etapa
    )

    # --------------------------------------------------------
    # PROGRESO
    # --------------------------------------------------------

    if total > 0:

        if actual <= 0:

            progreso = 0

        else:

            # El ensaye actual representa el progreso
            # visual dentro del lote.

            progreso = (
                (actual - 1)
                /
                total
            )

        barra_progreso.set_value(
            max(
                0,
                min(
                    progreso,
                    1
                )
            )
        )

        porcentaje = int(
            progreso * 100
        )

        porcentaje_modal.set_text(
            f"{porcentaje}%"
        )

    else:

        barra_progreso.set_value(
            0
        )

        porcentaje_modal.set_text(
            "0%"
        )


# ============================================================
# PROCESAR
# ============================================================

async def ejecutar_procesamiento():

    global procesando

    if procesando:
        return

    if not carpetas_dict:

        ui.notify(
            "Por favor importa al menos una carpeta "
            "antes de procesar.",
            type="warning",
        )

        return

    # ========================================================
    # CALCULAR TOTAL
    # ========================================================

    total = sum(
        len(archivos)
        for archivos
        in carpetas_dict.values()
    )

    if total <= 0:

        ui.notify(
            "No hay archivos para procesar.",
            type="warning",
        )

        return

    procesando = True

    # Limpiar la señal antes de lanzar el worker.
    workflow_service.preparar_procesamiento()

    # ========================================================
    # PREPARAR MODAL
    # ========================================================

    titulo_modal.set_text(
        f"Preparando {total} ensaye(s)..."
    )

    gst_modal.set_text(
        ""
    )

    archivo_modal.set_text(
        ""
    )

    etapa_modal.set_text(
        "Preparando procesamiento..."
    )

    detalle_modal.set_text(
        "No cierres la aplicación mientras "
        "se procesan los archivos."
    )

    porcentaje_modal.set_text(
        "0%"
    )

    barra_progreso.set_value(
        0
    )

    spinner_modal.set_visibility(
        True
    )

    boton_cancelar.enable()

    boton_procesar.disable()

    dialogo_proceso.open()

    # ========================================================
    # EVENT LOOP
    # ========================================================

    loop = asyncio.get_running_loop()

    callback_progreso = (
        crear_callback_progreso(
            loop
        )
    )

    try:

        # ====================================================
        # EJECUTAR EN WORKER THREAD
        # ====================================================

        resultado = await run.io_bound(
            workflow_service.procesar_lote_carpetas,
            carpetas_dict,
            callback_progreso,
        )

        estado = resultado.get(
            "status"
        )

        # ====================================================
        # ÉXITO
        # ====================================================

        if estado == "exito":

            barra_progreso.set_value(
                1
            )

            porcentaje_modal.set_text(
                "100%"
            )

            titulo_modal.set_text(
                "Procesamiento terminado"
            )

            etapa_modal.set_text(
                "Todos los ensayes fueron "
                "procesados correctamente."
            )

            detalle_modal.set_text(
                resultado.get(
                    "mensaje",
                    ""
                )
            )

            spinner_modal.set_visibility(
                False
            )

            boton_cancelar.disable()

            # Esperar un momento para que el usuario
            # alcance a ver el 100%.

            await asyncio.sleep(
                1.2
            )

            dialogo_proceso.close()

            # ================================================
            # MAIN LIMPIO
            # ================================================

            limpiar_sin_notificacion()

            ui.notify(
                resultado.get(
                    "mensaje",
                    "Proceso terminado correctamente."
                ),
                type="positive",
                duration=7000,
            )

            requisiciones = resultado.get(
                "requisiciones",
                []
            )

            if requisiciones:

                ui.notify(
                    "Requisición(es): "
                    + ", ".join(
                        requisiciones
                    ),
                    type="info",
                    duration=7000,
                )

        # ====================================================
        # CANCELADO
        # ====================================================

        elif estado == "cancelado":

            titulo_modal.set_text(
                "Procesamiento cancelado"
            )

            etapa_modal.set_text(
                "La cancelación terminó "
                "de forma segura."
            )

            detalle_modal.set_text(
                "Los mapas fueron conservados y "
                "los Excel generados durante este "
                "lote fueron eliminados."
            )

            spinner_modal.set_visibility(
                False
            )

            barra_progreso.set_value(
                0
            )

            porcentaje_modal.set_text(
                ""
            )

            await asyncio.sleep(
                1.5
            )

            dialogo_proceso.close()

            # IMPORTANTE:
            # NO limpiamos carpetas_dict.
            #
            # El usuario conserva sus archivos
            # seleccionados.

            ui.notify(
                "Procesamiento cancelado. "
                "Puedes volver a procesar los "
                "archivos cuando quieras.",
                type="warning",
                duration=7000,
            )

        # ====================================================
        # ERROR / PARCIAL
        # ====================================================

        else:

            dialogo_proceso.close()

            ui.notify(
                resultado.get(
                    "mensaje",
                    "Error al procesar el lote."
                ),
                type="negative",
                duration=10000,
            )

            # Tampoco limpiamos los archivos.

    except Exception as error:

        print()
        print("=" * 80)
        print(
            "ERROR DESDE LA INTERFAZ"
        )
        print("=" * 80)

        print(error)

        print("=" * 80)

        dialogo_proceso.close()

        ui.notify(
            f"Error durante el procesamiento: "
            f"{error}",
            type="negative",
            duration=10000,
        )

    finally:

        procesando = False

        boton_procesar.enable()

        boton_cancelar.disable()


# ============================================================
# DISEÑO GENERAL
# ============================================================

ui.query(
    "body"
).style(
    "background-color: #f4f6f9; "
    "font-family: Arial, sans-serif;"
)


# ============================================================
# MODAL DE PROCESAMIENTO
# ============================================================

with ui.dialog().props(
    "persistent"
) as dialogo_proceso:

    with ui.card().classes(
        "w-[520px] max-w-[92vw] p-7"
    ):

        # ----------------------------------------------------
        # CABECERA
        # ----------------------------------------------------

        with ui.column().classes(
            "w-full items-center gap-2"
        ):

            spinner_modal = ui.spinner(
                size="45px"
            ).classes(
                "text-blue-800 mb-2"
            )

            titulo_modal = ui.label(
                "Preparando procesamiento..."
            ).classes(
                "text-xl font-bold "
                "text-blue-900 text-center"
            )

            gst_modal = ui.label(
                ""
            ).classes(
                "text-base font-bold "
                "text-gray-700"
            )

        # ----------------------------------------------------
        # ARCHIVO
        # ----------------------------------------------------

        archivo_modal = ui.label(
            ""
        ).classes(
            "w-full text-xs text-gray-500 "
            "text-center break-all mt-2"
        )

        # ----------------------------------------------------
        # ETAPA
        # ----------------------------------------------------

        etapa_modal = ui.label(
            "Preparando..."
        ).classes(
            "w-full text-sm font-medium "
            "text-gray-700 text-center mt-3"
        )

        # ----------------------------------------------------
        # BARRA
        # ----------------------------------------------------

        with ui.row().classes(
            "w-full items-center gap-3 mt-4"
        ):

            barra_progreso = (
                ui.linear_progress(
                    value=0
                )
                .classes(
                    "flex-grow"
                )
            )

            porcentaje_modal = ui.label(
                "0%"
            ).classes(
                "text-sm font-bold "
                "text-blue-900 w-12 text-right"
            )

        # ----------------------------------------------------
        # DETALLE
        # ----------------------------------------------------

        detalle_modal = ui.label(
            "No cierres la aplicación mientras "
            "se procesan los archivos."
        ).classes(
            "w-full text-xs text-gray-500 "
            "text-center mt-3"
        )

        ui.separator().classes(
            "my-4"
        )

        # ----------------------------------------------------
        # CANCELAR
        # ----------------------------------------------------

        with ui.row().classes(
            "w-full justify-center"
        ):

            boton_cancelar = ui.button(
                "CANCELAR",
                on_click=cancelar_procesamiento,
                icon="close",
            ).classes(
                "bg-red-600 hover:bg-red-700 "
                "text-white font-bold px-8"
            )


# ============================================================
# HEADER
# ============================================================

with ui.header().classes(
    "bg-blue-900 text-white p-4 "
    "justify-between items-center"
):

    ui.label(
        "GEOTEST - Sistema de Captura de Laboratorio"
    ).classes(
        "text-xl font-bold"
    )

    ui.button(
        "Limpiar Todo",
        on_click=limpiar_todo,
        icon="delete",
    ).classes(
        "bg-red-600 text-white"
    )


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

with ui.column().classes(
    "w-full max-w-6xl mx-auto my-6 gap-6"
):

    # ========================================================
    # PASO 1
    # ========================================================

    with ui.card().classes(
        "w-full p-6 shadow-md rounded-lg"
    ):

        ui.label(
            "PASO 1: Importar Carpetas del ERP"
        ).classes(
            "text-lg font-bold text-gray-800 mb-2"
        )

        ui.label(
            "Arrastra carpetas dentro de la caja azul "
            "O usa el botón de exploración:"
        ).classes(
            "text-sm text-gray-600 mb-4"
        )

        ui.html(
            """
            <div
                id="drop_zone"
                style="
                    width: 100%;
                    border: 2px dashed #3b82f6;
                    background-color: #eff6ff;
                    padding: 30px;
                    text-align: center;
                    border-radius: 8px;
                    cursor: pointer;
                "
            >
                <p
                    style="
                        font-size: 16px;
                        font-weight: bold;
                        color: #1e40af;
                        margin-bottom: 5px;
                    "
                >
                    📂 Arrastra y suelta aquí tus CARPETAS del ERP
                </p>

                <p
                    style="
                        font-size: 12px;
                        color: #6b7280;
                        margin: 0;
                    "
                >
                    Se leerán automáticamente los archivos .xls contenidos
                </p>
            </div>
            """
        ).classes(
            "w-full"
        )

        with ui.row().classes(
            "w-full justify-center items-center mt-6"
        ):

            ui.button(
                "SELECCIONAR ARCHIVOS",
                on_click=abrir_dialogo_archivos,
                icon="attach_file",
            ).classes(
                "bg-blue-800 hover:bg-blue-900 "
                "text-white font-bold h-12 px-6 "
                "text-sm shadow-md rounded-lg"
            )


    # ========================================================
    # PASO 2
    # ========================================================

    with ui.card().classes(
        "w-full p-6 shadow-md rounded-lg"
    ):

        ui.label(
            "PASO 2: Carpetas y Archivos Importados"
        ).classes(
            "text-lg font-bold text-gray-800 mb-2"
        )

        contenedor_carpetas = (
            ui.column()
            .classes(
                "w-full gap-2"
            )
        )

        actualizar_vista_carpetas()


    # ========================================================
    # PROCESAR
    # ========================================================

    with ui.row().classes(
        "w-full justify-end my-2"
    ):

        boton_procesar = ui.button(
            "PROCESAR",
            on_click=ejecutar_procesamiento,
        ).classes(
            "bg-green-700 hover:bg-green-800 "
            "text-white font-bold py-3 px-6 "
            "rounded-lg text-base shadow-lg"
        )


# ============================================================
# DRAG & DROP
# ============================================================

ui.add_body_html(
    """
    <script>

        setTimeout(() => {

            const dropZone =
                document.getElementById('drop_zone');

            if (!dropZone)
                return;


            dropZone.addEventListener(
                'dragover',
                (e) => {

                    e.preventDefault();

                    dropZone.style.backgroundColor =
                        '#dbeafe';

                }
            );


            dropZone.addEventListener(
                'dragleave',
                () => {

                    dropZone.style.backgroundColor =
                        '#eff6ff';

                }
            );


            dropZone.addEventListener(
                'drop',
                async (e) => {

                    e.preventDefault();

                    dropZone.style.backgroundColor =
                        '#eff6ff';

                    const items =
                        e.dataTransfer.items;


                    function readFileAsBase64(
                        file
                    ) {

                        return new Promise(
                            (resolve) => {

                                const reader =
                                    new FileReader();

                                reader.onload =
                                    () => resolve(
                                        reader.result
                                    );

                                reader.readAsDataURL(
                                    file
                                );

                            }
                        );

                    }


                    async function processFile(
                        file,
                        folderName
                    ) {

                        const nombre =
                            file.name.toLowerCase();

                        if (
                            nombre.endsWith('.xls')
                            ||
                            nombre.endsWith('.xlsx')
                        ) {

                            const base64Content =
                                await readFileAsBase64(
                                    file
                                );

                            emitEvent(
                                'arrastrar_un_archivo',
                                {
                                    name:
                                        file.name,

                                    folder:
                                        folderName
                                        ||
                                        'Carpeta Importada',

                                    content:
                                        base64Content
                                }
                            );

                        }

                    }


                    async function scanFiles(
                        item,
                        folderName = ''
                    ) {

                        if (item.isFile) {

                            const file =
                                await new Promise(
                                    (resolve) =>
                                        item.file(
                                            resolve
                                        )
                                );

                            await processFile(
                                file,
                                folderName
                            );

                        }

                        else if (
                            item.isDirectory
                        ) {

                            const dirReader =
                                item.createReader();

                            const entries =
                                await new Promise(
                                    (resolve) =>
                                        dirReader.readEntries(
                                            resolve
                                        )
                                );

                            for (
                                let i = 0;
                                i < entries.length;
                                i++
                            ) {

                                await scanFiles(
                                    entries[i],
                                    folderName
                                    ||
                                    item.name
                                );

                            }

                        }

                    }


                    for (
                        let i = 0;
                        i < items.length;
                        i++
                    ) {

                        const item =
                            items[i]
                            .webkitGetAsEntry();

                        if (item) {

                            await scanFiles(
                                item
                            );

                        }

                    }

                }
            );

        }, 1000);

    </script>
    """
)


# ============================================================
# EVENTO
# ============================================================

ui.on(
    "arrastrar_un_archivo",
    recibir_un_archivo
)


# ============================================================
# INICIAR
# ============================================================

if __name__ in {
    "__main__",
    "__mp_main__",
}:

    ui.run(
    title="LabCaptura Geotest",
    host="0.0.0.0",
    port=8080,
    reload=False,
)