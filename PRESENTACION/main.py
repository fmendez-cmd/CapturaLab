import os
import sys
import base64
import asyncio
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

from nicegui import app, ui, run

CARPETA_STATIC = Path(__file__).resolve().parent.parent / "MEDIA"
app.add_static_files("/MEDIA", str(CARPETA_STATIC))

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
from PRESENTACION.estilo_geotest import aplicar_estilo_geotest
from PRESENTACION.resumen_procesamiento import ResumenProcesamientoUI
from PRESENTACION.historial_resultados import HistorialResultados


def fecha_encabezado():
    """Fecha local de Geotest escrita en español."""
    ahora = datetime.now(ZoneInfo("America/Mexico_City"))
    dias = ("LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁBADO", "DOMINGO")
    meses = ("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
             "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE")
    return f"{dias[ahora.weekday()]}, {ahora.day} DE {meses[ahora.month - 1]} DE {ahora.year}"


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

# El historial queda fuera de ARCHIVOS_ENTRADA, para conservarlo al limpiar la vista.
historial_resultados = HistorialResultados(Path(BASE_DIR) / "RESULTADOS_HISTORIAL")


# ============================================================
# UTILIDAD: REGISTRAR ARCHIVO
# ============================================================

def registrar_archivo(
    nombre_carpeta,
    filename,
    ruta_fisica_real,
):
    """
    Registra un archivo ya guardado físicamente en el servidor
    dentro de carpetas_dict y extrae sus metadatos.
    """

    if nombre_carpeta not in carpetas_dict:
        carpetas_dict[nombre_carpeta] = []

    info = ParserERP.extraer_metadatos_archivo(
        filename,
        ruta_fisica_real,
    )

    existe = any(
        archivo["archivo"] == filename
        for archivo in carpetas_dict[nombre_carpeta]
    )

    if not existe:
        carpetas_dict[nombre_carpeta].append(info)


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

        if not filename:
            return

        if not filename.lower().endswith(
            (".xls", ".xlsx")
        ):
            return

        # Evitar nombres de carpeta problemáticos
        nombre_carpeta = (
            nombre_carpeta
            or
            "Carpeta Importada"
        )

        nombre_carpeta = os.path.basename(
            nombre_carpeta
        )

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

            _, data = (
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

        registrar_archivo(
            nombre_carpeta,
            filename,
            ruta_fisica_real,
        )

        actualizar_vista_carpetas()

    except Exception as error:

        print()
        print("=" * 80)
        print("ERROR AL RECIBIR ARCHIVO POR DRAG & DROP")
        print(error)
        print("=" * 80)

        ui.notify(
            f"Error al cargar archivo: {error}",
            type="negative",
            duration=8000,
        )


# ============================================================
# RECIBIR ARCHIVO DESDE SELECTOR DEL NAVEGADOR
# ============================================================

async def recibir_archivo_upload(e):
    """
    Este evento se ejecuta cuando el usuario selecciona
    archivos desde SU navegador.

    El explorador que se abre pertenece a la computadora
    cliente, NO al servidor.
    """

    try:

        if procesando:

            ui.notify(
                "Hay un procesamiento en curso.",
                type="warning",
            )

            return

        # ----------------------------------------------------
        # OBTENER NOMBRE
        # ----------------------------------------------------

        filename = getattr(
            e.file,
            "name",
            ""
        )

        if not filename:

            ui.notify(
                "No fue posible obtener el nombre del archivo.",
                type="negative",
            )

            return

        if not filename.lower().endswith(
            (".xls", ".xlsx")
        ):

            ui.notify(
                f"Archivo no permitido: {filename}",
                type="warning",
            )

            return

        # ----------------------------------------------------
        # CARPETA LÓGICA
        # ----------------------------------------------------
        #
        # Un navegador NO entrega la ruta C:\... del usuario.
        # Eso es intencional por seguridad.
        #
        # Los archivos seleccionados manualmente se agrupan
        # aquí.
        # ----------------------------------------------------

        nombre_carpeta = "Archivos seleccionados"

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

        # ----------------------------------------------------
        # GUARDAR ARCHIVO EN EL SERVIDOR
        # ----------------------------------------------------

        await e.file.save(
            ruta_fisica_real
        )

        print(
            f"[UPLOAD] Archivo recibido: {filename}"
        )

        print(
            f"[UPLOAD] Guardado en: {ruta_fisica_real}"
        )

        # ----------------------------------------------------
        # REGISTRAR
        # ----------------------------------------------------

        registrar_archivo(
            nombre_carpeta,
            filename,
            ruta_fisica_real,
        )

        actualizar_vista_carpetas()

    except Exception as error:

        print()
        print("=" * 80)
        print(
            "ERROR AL RECIBIR ARCHIVO "
            "DESDE EL NAVEGADOR"
        )
        print(error)
        print("=" * 80)

        ui.notify(
            f"Error al cargar archivo: {error}",
            type="negative",
            duration=8000,
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

    total_carpetas_lote = len(
        carpetas_dict
    )

    if total <= 0:

        ui.notify(
            "No hay archivos para procesar.",
            type="warning",
        )

        return

    procesando = True

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
        "No cierres esta pestaña mientras "
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

            await asyncio.sleep(
                0.8
            )

            dialogo_proceso.close()

            try:
                await run.io_bound(historial_resultados.guardar, resultado)
            except Exception as error:
                print(f"No se pudo guardar el historial de resultados: {error}")
                ui.notify("No se pudo guardar este lote en Resultados.", type="negative")

            ResumenProcesamientoUI.mostrar(
                resultado,
                total_carpetas=total_carpetas_lote,
                on_close=limpiar_sin_notificacion,
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

            try:
                await run.io_bound(historial_resultados.guardar, resultado)
            except Exception as error:
                print(f"No se pudo guardar el historial de resultados: {error}")
                ui.notify("No se pudo guardar este lote en Resultados.", type="negative")

            ResumenProcesamientoUI.mostrar(
                resultado,
                total_carpetas=total_carpetas_lote,
                on_close=limpiar_sin_notificacion,
            )

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
            "Ocurrió un problema inesperado "
            "en la interfaz. "
            "Revisa la consola técnica o contacta "
            "al responsable del sistema.",
            type="negative",
            duration=10000,
        )

    finally:

        procesando = False

        boton_procesar.enable()

        boton_cancelar.disable()


def descargar_resultado_guardado(registro: dict):
    ruta = historial_resultados.ruta_descarga(registro["id"])
    if ruta is None:
        ui.notify("La descarga de esta OT ya no está disponible.", type="warning")
        return
    fecha = datetime.fromisoformat(registro["fecha"]).strftime("%Y%m%d_%H%M")
    ui.download(str(ruta), filename=f"{registro['ot']}_Resultados_{fecha}.zip")


def actualizar_historial():
    contenedor_resultados.clear()
    with contenedor_resultados:
        try:
            registros = historial_resultados.listar()
        except Exception as error:
            print(f"No se pudo leer el historial: {error}")
            ui.label("No se pudieron cargar los resultados.").classes("text-red-700")
            return
        if not registros:
            ui.label("Todavía no hay procesamientos guardados.").classes(
                "text-sm text-gray-500 p-5"
            )
            return
        ot_actual = None
        for registro in registros:
            if registro["ot"] != ot_actual:
                ot_actual = registro["ot"]
                ui.label(ot_actual).classes(
                    "text-lg font-bold text-[#253B83] mt-4 first:mt-0"
                )
            with ui.card().classes(
                "geotest-card w-full px-5 py-4 shadow-none"
            ):
                with ui.row().classes(
                    "w-full items-center justify-between gap-4 flex-wrap"
                ):
                    with ui.column().classes("gap-1 min-w-0"):
                        fecha = datetime.fromisoformat(registro["fecha"])
                        ui.label(fecha.strftime("%d/%m/%Y · %H:%M")).classes(
                            "text-sm font-semibold text-[#202938]"
                        )
                        with ui.row().classes("gap-4 text-sm"):
                            ui.label(
                                f"{registro['exitosos']} ensayes exitosos"
                            ).classes("text-green-800")
                            ui.label(
                                f"{registro['incidencias']} incidencias"
                            ).classes("text-amber-800")
                    if registro["zip_nombre"]:
                        with ui.button(
                            on_click=lambda r=registro: descargar_resultado_guardado(r)
                        ).props("flat round").classes("hover:bg-blue-50"):
                            ui.image("/MEDIA/descargar.png").classes("w-7 h-7 object-contain")
                            ui.tooltip("Descargar Excel de esta OT")
                    else:
                        ui.label("Sin archivos para descargar").classes(
                            "text-xs text-gray-500"
                        )


def mostrar_captura():
    vista_resultados.set_visibility(False)
    vista_captura.set_visibility(True)


def mostrar_resultados():
    actualizar_historial()
    vista_captura.set_visibility(False)
    vista_resultados.set_visibility(True)


# ============================================================
# DISEÑO GENERAL
# ============================================================

aplicar_estilo_geotest()


# ============================================================
# MODAL DE PROCESAMIENTO
# ============================================================

with ui.dialog().props("persistent") as dialogo_proceso:
    with ui.card().classes(
        "geotest-card w-[600px] max-w-[94vw] max-h-[92vh] p-0 overflow-hidden"
    ):
        # Cabecera clara, con la franja de tres colores del ERP.
        with ui.column().classes(
            "w-full relative bg-white px-7 py-6 gap-2 border-b border-[#DFE4EF]"
        ):
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
            with ui.row().classes("w-full items-center gap-3 no-wrap"):
                spinner_modal = ui.spinner(size="28px").classes(
                    "text-[#5979E6] shrink-0"
                )
                titulo_modal = ui.label("Preparando procesamiento...").classes(
                    "text-xl md:text-2xl font-bold text-[#202938]"
                )
            gst_modal = ui.label("").classes(
                "text-sm font-semibold text-[#253B83]"
            )

        with ui.column().classes("w-full bg-[#F8FAFF] px-7 py-6 gap-4"):
            ui.label("ARCHIVO EN PROCESO").classes(
                "text-xs font-bold tracking-[0.14em] text-[#64748B]"
            )
            archivo_modal = ui.label("").classes(
                "w-full text-sm text-[#39465C] break-all"
            )
            etapa_modal = ui.label("Preparando...").classes(
                "w-full text-sm font-medium text-[#253B83]"
            )

            with ui.row().classes("w-full items-center gap-3 no-wrap"):
                barra_progreso = ui.linear_progress(
                    value=0, show_value=False,
                ).props("color=secondary track-color=grey-3 rounded").classes(
                    "flex-grow"
                )
                porcentaje_modal = ui.label("0%").classes(
                    "text-sm font-bold text-[#253B83] w-12 text-right"
                )

            detalle_modal = ui.label(
                "No cierres esta pestaña mientras se procesan los archivos."
            ).classes("w-full text-xs text-[#64748B]")

        with ui.row().classes(
            "w-full justify-end items-center bg-white "
            "border-t border-[#DFE4EF] px-7 py-4"
        ):
            boton_cancelar = ui.button(
                "Cancelar",
                on_click=cancelar_procesamiento,
                icon="close",
            ).props("outline color=negative no-caps").classes(
                "rounded-xl px-5 py-2 font-semibold"
            )


# ============================================================
# HEADER
# ============================================================

with ui.header().classes(
    "geotest-header h-[76px] px-5 items-center justify-between"
):

    with ui.row().classes("items-center gap-3"):
        ui.button(
            icon="menu",
            on_click=lambda: menu_lateral.toggle(),
        ).props("flat round color=secondary aria-label='Abrir o cerrar menú'")
        ui.separator().props("vertical").classes("h-8")

        with ui.avatar().classes("bg-[#000000]"):
            ui.image("/MEDIA/logo.jpg").classes("w-full h-full object-contain")

        with ui.column().classes("gap-0"):
            ui.label("GEOTEST").classes("text-base font-bold tracking-[0.16em]")
            ui.label("LABORATORIO · CAPTURA").classes(
                "text-[10px] tracking-[0.18em] text-gray-500"
            )

    ui.button(
        "Limpiar todo",
        on_click=limpiar_todo,
        icon="delete_outline",
    ).props("outline color=negative no-caps")


# Navegación visual de LabCaptura. Las acciones usan elementos existentes.
with ui.left_drawer(value=True).props("show-if-above breakpoint=768").classes(
    "geotest-sidebar w-[250px] p-0"
) as menu_lateral:
    ui.label("OPERACIÓN").classes(
        "text-xs tracking-[0.17em] text-blue-200 mt-5 ml-5 mb-2"
    )
    ui.item("Captura de ensayes", on_click=mostrar_captura).props(
        "clickable"
    ).classes("text-white")
    ui.item("Resultados", on_click=mostrar_resultados).props(
        "clickable"
    ).classes("text-white")
    ui.label("Geotest Ingeniería").classes(
        "absolute bottom-6 left-5 text-sm font-semibold text-gray-300"
    )


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

with ui.column().classes(
    "w-full max-w-6xl mx-auto my-6 gap-6 px-4"
) as vista_captura:
    with ui.card().classes("geotest-card geotest-banner w-full p-6"):
        ui.label("SISTEMA DE CAPTURA - LABORATORIO · GEOTEST").classes(
            "text-lg tracking-[0.18em] font-bold text-[#253B83]"
        )
        ui.label(fecha_encabezado()).classes(
            "text-xs tracking-[0.18em] font-semibold text-[#39465C]"
        )
        ui.label("Importa archivos del ERP y genera los Excel en sus plantillas oficiales.").classes(
            "text-sm text-gray-600"
        )

    # ========================================================
    # PASO 1
    # ========================================================

    with ui.card().classes(
        "geotest-card w-full p-6"
    ):

        ui.label(
            "PASO 1: Importar Carpetas del ERP"
        ).classes(
            "text-lg font-bold "
            "text-gray-800 mb-2"
        )


        # ----------------------------------------------------
        # DROP ZONE
        # ----------------------------------------------------

        ui.html(
            """
            <div id="drop_zone" class="geotest-drop-zone">
                <img
                    src="/MEDIA/carpeta-abierta.png"
                    alt=""
                    class="geotest-drop-icon"
                >
                <div class="geotest-drop-text">
                    <div class="geotest-drop-title">
                        Arrastra aquí tu carpeta del ERP
                    </div>
                    <div class="geotest-drop-description">
                        Se cargarán los archivos .xls y .xlsx que contenga
                    </div>
                </div>
            </div>
            """
        ).classes("w-full")

        # ----------------------------------------------------
        # SELECTOR WEB
        # ----------------------------------------------------

        with ui.column().classes(
            "w-full justify-center items-center mt-6"
        ):

            ui.label(
                "O selecciona los archivos desde tu computadora"
            ).classes(
                "text-sm text-gray-500 mb-2"
            )

            uploader = ui.upload(
                label="SELECCIONAR ARCHIVOS",
                on_upload=recibir_archivo_upload,
                multiple=True,
                auto_upload=True,
            ).props(
                'accept=".xls,.xlsx" '
                'flat '
                'color="primary" '
                'hide-upload-btn'
            ).classes(
               "selector-archivos"
            )


    # ========================================================
    # PASO 2
    # ========================================================

    with ui.card().classes(
        "geotest-card w-full p-6"
    ):

        ui.label(
            "PASO 2: Carpetas y Archivos Importados"
        ).classes(
            "text-lg font-bold "
            "text-gray-800 mb-2"
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

    with ui.row().classes("w-full justify-end my-2"):
        boton_procesar = ui.button(
            "PROCESAR",
            on_click=ejecutar_procesamiento,
        ).props("no-caps").classes(
            "rounded-lg px-7 py-2.5 font-bold shadow-md"
        ).style(
            "background-color: #253B83 !important; color: white !important;"
        )


with ui.column().classes(
    "w-full max-w-6xl mx-auto my-6 gap-5 px-4"
) as vista_resultados:
    with ui.card().classes("geotest-card geotest-banner w-full p-6"):
        ui.label("RESULTADOS · LABORATORIO").classes(
            "text-lg tracking-[0.18em] font-bold text-[#253B83]"
        )
        ui.label("Procesamientos agrupados por Orden de Trabajo").classes(
            "text-sm text-gray-600"
        )
    with ui.row().classes("w-full justify-between items-center"):
        ui.label("Historial de resultados").classes(
            "text-xl font-bold text-[#202938]"
        )
        ui.button("Actualizar", on_click=actualizar_historial, icon="refresh").props(
            "flat color=secondary no-caps"
        )
    contenedor_resultados = ui.column().classes("w-full gap-3")

vista_resultados.set_visibility(False)


ui.add_head_html(
    """
    <style>
    
    body {
        user-select: none;
        -webkit-user-select: none;
    }

        /* ==========================================
           UPLOADER COMO BOTÓN NORMAL
           ========================================== */

        .selector-archivos {
            width: 245px !important;
            min-width: 245px !important;
            height: 50px !important;

            box-shadow: none !important;
            border: none !important;
            background: transparent !important;

            position: relative !important;
            overflow: hidden !important;
        }


        /* No mostrar lista ni preview de archivos */
        .selector-archivos .q-uploader__list {
            display: none !important;
        }


        /* Botón azul */
        .selector-archivos .q-uploader__header {

            width: 245px !important;
            height: 50px !important;
            min-height: 50px !important;

            padding: 0 !important;

            border-radius: 8px !important;

            background: #5979E6 !important;

            box-shadow:
                0 4px 8px rgba(0, 0, 0, 0.15) !important;

            position: relative !important;

            overflow: hidden !important;
        }


        /* Contenedor */
        .selector-archivos .q-uploader__header-content {

            width: 100% !important;
            height: 100% !important;

            display: flex !important;

            align-items: center !important;
            justify-content: center !important;

            padding: 0 !important;

            position: relative !important;
        }


        /* Texto */
        .selector-archivos .q-uploader__title {

            display: block !important;

            color: white !important;

            font-size: 14px !important;
            font-weight: 700 !important;

            text-align: center !important;

            white-space: nowrap !important;

            position: relative !important;

            z-index: 1 !important;

            pointer-events: none !important;
        }


        /* Ocultar 0B / 0% */
        .selector-archivos .q-uploader__subtitle {
            display: none !important;
        }


        /* ==========================================
           BOTÓN + REAL
           
           NO LO ELIMINAMOS.
           Lo hacemos transparente y ocupa TODO
           el botón azul.
           ========================================== */

        .selector-archivos .q-uploader__header .q-btn {

            display: block !important;

            position: absolute !important;

            top: 0 !important;
            left: 0 !important;

            width: 100% !important;
            height: 100% !important;

            min-width: 100% !important;
            min-height: 100% !important;

            padding: 0 !important;
            margin: 0 !important;

            opacity: 0 !important;

            z-index: 10 !important;

            cursor: pointer !important;
        }


        /* Hover */
        .selector-archivos .q-uploader__header:hover {
            background: #4869D2 !important;
            cursor: pointer !important;
        }

    </style>
    """
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
                            (resolve, reject) => {

                                const reader =
                                    new FileReader();

                                reader.onload =
                                    () => resolve(
                                        reader.result
                                    );

                                reader.onerror =
                                    () => reject(
                                        reader.error
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


                    async function readAllEntries(
                        directoryReader
                    ) {

                        const entries = [];

                        while (true) {

                            const batch =
                                await new Promise(
                                    (resolve) => {

                                        directoryReader
                                            .readEntries(
                                                resolve
                                            );

                                    }
                                );

                            if (!batch.length)
                                break;

                            entries.push(
                                ...batch
                            );

                        }

                        return entries;

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

                            const carpetaRaiz =
                                folderName
                                ||
                                item.name;

                            const dirReader =
                                item.createReader();

                            const entries =
                                await readAllEntries(
                                    dirReader
                                );

                            for (
                                let i = 0;
                                i < entries.length;
                                i++
                            ) {

                                await scanFiles(
                                    entries[i],
                                    carpetaRaiz
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
# EVENTO DRAG & DROP
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
    FAVICON_GEOTEST = (
        "data:image/jpeg;base64,"
        + base64.b64encode((CARPETA_STATIC / "logo.jpg").read_bytes()).decode("ascii")
    )

    ui.run(
        title="Captura - Geotest",
        favicon=FAVICON_GEOTEST,
        host="0.0.0.0",
        port=8086,
        reload=False,
    )
