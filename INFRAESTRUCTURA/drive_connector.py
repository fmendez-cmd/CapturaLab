# INFRAESTRUCTURA/drive_connector.py

import io
import json
import os
import re
import time

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import (
    MediaIoBaseDownload,
    MediaIoBaseUpload,
)

from INFRAESTRUCTURA.config_rutas import (
    DRIVE_PLANTILLAS_URL,
    DRIVE_MAPAS_URL,
    TEMP_PLANTILLAS_DIR,
)


class DriveConnector:

    # Caché VOLÁTIL de plantillas. Solo existe durante un lote.
    _cache_lote_activo = False
    _cache_plantillas_lote = {}

    @staticmethod
    def iniciar_cache_lote():
        DriveConnector._cache_lote_activo = True
        DriveConnector._cache_plantillas_lote = {}
        print("⚡ Caché temporal de plantillas iniciado.")

    @staticmethod
    def limpiar_cache_lote():
        # Las plantillas descargadas son temporales. Se eliminan al cerrar el lote.
        rutas = set(DriveConnector._cache_plantillas_lote.values())
        DriveConnector._cache_plantillas_lote = {}
        DriveConnector._cache_lote_activo = False

        for ruta in rutas:
            try:
                if ruta and os.path.isfile(ruta):
                    os.remove(ruta)
            except Exception as error:
                print(f"⚠️ No se pudo borrar plantilla temporal '{ruta}': {error}")

        print("🧹 Caché temporal de plantillas limpiado.")

    # ==========================================================
    # IMPORTANTE
    # ==========================================================
    # Ya no usamos readonly porque necesitamos:
    #
    # - descargar plantillas
    # - descargar mapas
    # - crear mapas
    # - actualizar mapas
    # - crear locks
    # - borrar locks
    #
    # La cuenta de servicio debe tener permiso de EDITOR
    # en la carpeta MAPAS_GST.
    # ==========================================================

    SCOPES = [
        "https://www.googleapis.com/auth/drive"
    ]

    # ==========================================================
    # SERVICIO DRIVE
    # ==========================================================

    @staticmethod
    def _obtener_servicio_drive():

        creds_path = os.path.join(
            os.path.dirname(
                os.path.dirname(
                    os.path.abspath(__file__)
                )
            ),
            "service_account.json",
        )

        if not os.path.exists(creds_path):

            raise FileNotFoundError(
                "No se encontró el archivo "
                "'service_account.json' "
                "en la raíz del proyecto."
            )

        creds = (
            service_account
            .Credentials
            .from_service_account_file(
                creds_path,
                scopes=DriveConnector.SCOPES,
            )
        )

        return build(
            "drive",
            "v3",
            credentials=creds,
            cache_discovery=False,
        )

    # ==========================================================
    # OBTENER ID DESDE URL DRIVE
    # ==========================================================

    @staticmethod
    def _extraer_id_drive(
        url_o_id: str
    ) -> str:

        if not url_o_id:

            raise ValueError(
                "No se recibió una URL "
                "o ID de Google Drive."
            )

        texto = str(
            url_o_id
        ).strip()

        patrones = [
            r"/folders/([a-zA-Z0-9_-]+)",
            r"/d/([a-zA-Z0-9_-]+)",
            r"[?&]id=([a-zA-Z0-9_-]+)",
        ]

        for patron in patrones:

            match = re.search(
                patron,
                texto
            )

            if match:
                return match.group(1)

        # Si no es URL, asumimos que ya es ID.
        return texto

    # ==========================================================
    # ESCAPAR TEXTO PARA QUERY DRIVE
    # ==========================================================

    @staticmethod
    def _escapar_query(
        texto: str
    ) -> str:

        return (
            str(texto)
            .replace("\\", "\\\\")
            .replace("'", "\\'")
        )

    # ==========================================================
    # CARPETA PLANTILLAS
    # ==========================================================

    @staticmethod
    def obtener_id_carpeta_plantillas():

        if not DRIVE_PLANTILLAS_URL:

            raise FileNotFoundError(
                "La variable DRIVE_PLANTILLAS_URL "
                "no está definida en .env"
            )

        return (
            DriveConnector
            ._extraer_id_drive(
                DRIVE_PLANTILLAS_URL
            )
        )

    # ==========================================================
    # CARPETA MAPAS
    # ==========================================================

    @staticmethod
    def obtener_id_carpeta_mapas():

        if not DRIVE_MAPAS_URL:

            raise FileNotFoundError(
                "La variable DRIVE_MAPAS_URL "
                "no está definida en .env"
            )

        return (
            DriveConnector
            ._extraer_id_drive(
                DRIVE_MAPAS_URL
            )
        )

    # ==========================================================
    # DESCARGAR PLANTILLA POR GST
    # ==========================================================

    @staticmethod
    def obtener_plantilla_por_gst(
        codigo_gst: str
    ) -> str:

        clave_cache = str(codigo_gst).upper().strip()
        if DriveConnector._cache_lote_activo:
            ruta_cache = DriveConnector._cache_plantillas_lote.get(clave_cache)
            if ruta_cache and os.path.isfile(ruta_cache):
                print(f"⚡ Plantilla {codigo_gst} recuperada del caché del lote.")
                return ruta_cache

        codigo_limpio = str(codigo_gst).upper().strip()
        patron_gst = r"(?i)\bGST[\s_-]*0*(\d+)(?!\d)"

        match_codigo = re.search(patron_gst, codigo_limpio)
        if not match_codigo:
            raise ValueError(f"Código GST inválido: '{codigo_gst}'")

        numero_gst = int(match_codigo.group(1))

        # Drive sigue siendo siempre la fuente oficial.
        os.makedirs(TEMP_PLANTILLAS_DIR, exist_ok=True)
        raiz_id = DriveConnector.obtener_id_carpeta_plantillas()
        service = DriveConnector._obtener_servicio_drive()

        mime_carpeta = "application/vnd.google-apps.folder"
        carpetas_excluidas = {
            "REPORTE FOTOGRAFICO",
            "Z-DOCTOS.INFORMATIVOS",
            "SEDENA-SICT",
            "OBSOLETOS",
        }

        def nombre_normalizado(nombre: str) -> str:
            return " ".join(str(nombre).upper().strip().split())

        def coincide_gst(nombre: str) -> bool:
            if not nombre.lower().endswith((".xlsx", ".xlsm")):
                return False
            coincidencias = re.findall(patron_gst, nombre)
            return any(int(n) == numero_gst for n in coincidencias)

        def listar_hijos(parent_id: str):
            items = []
            page_token = None
            while True:
                resultado = service.files().list(
                    q=f"'{parent_id}' in parents and trashed = false",
                    pageSize=1000,
                    fields="nextPageToken,files(id,name,mimeType,modifiedTime)",
                    pageToken=page_token,
                    supportsAllDrives=True,
                    includeItemsFromAllDrives=True,
                ).execute()
                items.extend(resultado.get("files", []))
                page_token = resultado.get("nextPageToken")
                if not page_token:
                    break
            return items

        # 1) Rápido: solo archivos directamente dentro de la raíz.
        print(f"🔎 Buscando plantilla {codigo_gst} en nivel principal de Drive...")
        items_raiz = listar_hijos(raiz_id)

        candidatos_superficiales = [
            item for item in items_raiz
            if str(item.get("mimeType", "")) != mime_carpeta
            and coincide_gst(str(item.get("name", "")))
        ]

        if candidatos_superficiales:
            candidatos_superficiales.sort(
                key=lambda x: str(x.get("modifiedTime", "")),
                reverse=True,
            )
            item = candidatos_superficiales[0]
            print(
                f"⚡ Plantilla encontrada en nivel principal para {codigo_gst}: "
                f"'{item['name']}'"
            )
        else:
            # 2) Respaldo: recorrer subcarpetas, excepto ramas excluidas.
            print(
                f"🔎 No apareció {codigo_gst} en el nivel principal. "
                "Buscando en subcarpetas permitidas..."
            )

            candidatos = []
            carpetas_pendientes = []

            for hijo in items_raiz:
                if str(hijo.get("mimeType", "")) != mime_carpeta:
                    continue
                nombre = nombre_normalizado(str(hijo.get("name", "")))
                if nombre in carpetas_excluidas:
                    print(f"⛔ Carpeta excluida: '{hijo.get('name', '')}'")
                    continue
                carpetas_pendientes.append(hijo["id"])

            visitadas = set()
            while carpetas_pendientes:
                parent_id = carpetas_pendientes.pop()
                if parent_id in visitadas:
                    continue
                visitadas.add(parent_id)

                for hijo in listar_hijos(parent_id):
                    mime_type = str(hijo.get("mimeType", ""))
                    nombre = str(hijo.get("name", ""))

                    if mime_type == mime_carpeta:
                        if nombre_normalizado(nombre) in carpetas_excluidas:
                            print(f"⛔ Carpeta excluida: '{nombre}'")
                            continue
                        carpetas_pendientes.append(hijo["id"])
                        continue

                    if coincide_gst(nombre):
                        candidatos.append(hijo)

            if not candidatos:
                raise FileNotFoundError(
                    f"No se encontró ninguna plantilla para '{codigo_gst}' "
                    "ni en el nivel principal ni en las subcarpetas permitidas."
                )

            candidatos.sort(
                key=lambda x: str(x.get("modifiedTime", "")),
                reverse=True,
            )
            item = candidatos[0]
            print(
                f"📌 Plantilla más reciente encontrada en profundidad para "
                f"{codigo_gst}: '{item['name']}' "
                f"({item.get('modifiedTime', 'sin fecha')})"
            )

        file_id = item["id"]
        file_name = item["name"]
        ruta_destino = os.path.join(TEMP_PLANTILLAS_DIR, file_name)

        print(f"📥 Descargando plantilla oficial '{file_name}'...")
        request = service.files().get_media(
            fileId=file_id,
            supportsAllDrives=True,
        )

        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

        with open(ruta_destino, "wb") as archivo:
            archivo.write(fh.getvalue())

        if DriveConnector._cache_lote_activo:
            DriveConnector._cache_plantillas_lote[clave_cache] = ruta_destino
            print(f"⚡ Plantilla {codigo_gst} guardada en caché hasta terminar el lote.")

        return ruta_destino

    # ==========================================================
    # BUSCAR SUBCARPETA
    # ==========================================================

    @staticmethod
    def buscar_carpeta(
        parent_id: str,
        nombre: str
    ):

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        nombre_query = (
            DriveConnector
            ._escapar_query(
                nombre
            )
        )

        query = (
            f"'{parent_id}' in parents "
            f"and name = '{nombre_query}' "
            "and mimeType = "
            "'application/vnd.google-apps.folder' "
            "and trashed = false"
        )

        resultado = (
            service
            .files()
            .list(
                q=query,
                pageSize=1,
                fields="files(id,name)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )

        archivos = resultado.get(
            "files",
            []
        )

        if not archivos:
            return None

        return archivos[0][
            "id"
        ]

    # ==========================================================
    # CREAR SUBCARPETA SI NO EXISTE
    # ==========================================================

    @staticmethod
    def obtener_o_crear_carpeta(
        parent_id: str,
        nombre: str
    ) -> str:

        existente = (
            DriveConnector
            .buscar_carpeta(
                parent_id,
                nombre
            )
        )

        if existente:
            return existente

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        metadata = {
            "name": nombre,

            "mimeType":
                "application/vnd.google-apps.folder",

            "parents": [
                parent_id
            ],
        }

        carpeta = (
            service
            .files()
            .create(
                body=metadata,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )

        print(
            "📁 Carpeta creada en Drive: "
            f"{nombre}"
        )

        return carpeta[
            "id"
        ]

    # ==========================================================
    # BUSCAR ARCHIVO EXACTO
    # ==========================================================

    @staticmethod
    def buscar_archivo(
        parent_id: str,
        nombre_archivo: str
    ):

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        nombre_query = (
            DriveConnector
            ._escapar_query(
                nombre_archivo
            )
        )

        query = (
            f"'{parent_id}' in parents "
            f"and name = '{nombre_query}' "
            "and trashed = false"
        )

        resultado = (
            service
            .files()
            .list(
                q=query,
                pageSize=10,
                fields=(
                    "files("
                    "id,"
                    "name,"
                    "modifiedTime"
                    ")"
                ),
                orderBy="modifiedTime desc",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )

        archivos = resultado.get(
            "files",
            []
        )

        if not archivos:
            return None

        return archivos[0]

    # ==========================================================
    # DESCARGAR JSON
    # ==========================================================

    @staticmethod
    def descargar_json(
        file_id: str
    ) -> dict:

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        request = (
            service
            .files()
            .get_media(
                fileId=file_id,
                supportsAllDrives=True,
            )
        )

        fh = io.BytesIO()

        downloader = (
            MediaIoBaseDownload(
                fh,
                request
            )
        )

        done = False

        while not done:

            _, done = (
                downloader
                .next_chunk()
            )

        contenido = (
            fh.getvalue()
            .decode(
                "utf-8"
            )
        )

        return json.loads(
            contenido
        )

    # ==========================================================
    # SUBIR JSON
    # ==========================================================

    @staticmethod
    def subir_json(
        parent_id: str,
        nombre_archivo: str,
        datos: dict
    ) -> str:

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        contenido = json.dumps(
            datos,
            ensure_ascii=False,
            indent=2
        ).encode(
            "utf-8"
        )

        media = MediaIoBaseUpload(
            io.BytesIO(
                contenido
            ),
            mimetype="application/json",
            resumable=False,
        )

        existente = (
            DriveConnector
            .buscar_archivo(
                parent_id,
                nombre_archivo
            )
        )

        # ------------------------------------------------------
        # ACTUALIZAR
        # ------------------------------------------------------

        if existente:

            file_id = existente[
                "id"
            ]

            (
                service
                .files()
                .update(
                    fileId=file_id,
                    media_body=media,
                    fields="id",
                    supportsAllDrives=True,
                )
                .execute()
            )

            print(
                "🔄 Mapa actualizado en Drive: "
                f"{nombre_archivo}"
            )

            return file_id

        # ------------------------------------------------------
        # CREAR
        # ------------------------------------------------------

        metadata = {
            "name":
                nombre_archivo,

            "parents": [
                parent_id
            ],

            "mimeType":
                "application/json",
        }

        archivo = (
            service
            .files()
            .create(
                body=metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )

        print(
            "☁️ Mapa subido a Drive: "
            f"{nombre_archivo}"
        )

        return archivo[
            "id"
        ]

    # ==========================================================
    # CREAR ARCHIVO DE BLOQUEO
    # ==========================================================

    @staticmethod
    def crear_lock(
        parent_id: str,
        nombre_lock: str,
        contenido: dict
    ):

        # Primero verificar si ya existe.
        existente = (
            DriveConnector
            .buscar_archivo(
                parent_id,
                nombre_lock
            )
        )

        if existente:

            return {
                "creado": False,
                "file_id": existente[
                    "id"
                ]
            }

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        datos = json.dumps(
            contenido,
            ensure_ascii=False,
            indent=2
        ).encode(
            "utf-8"
        )

        media = MediaIoBaseUpload(
            io.BytesIO(
                datos
            ),
            mimetype="application/json",
            resumable=False,
        )

        metadata = {
            "name":
                nombre_lock,

            "parents": [
                parent_id
            ],

            "mimeType":
                "application/json",
        }

        archivo = (
            service
            .files()
            .create(
                body=metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True,
            )
            .execute()
        )

        return {
            "creado": True,
            "file_id": archivo[
                "id"
            ]
        }

    # ==========================================================
    # ELIMINAR ARCHIVO
    # ==========================================================

    @staticmethod
    def eliminar_archivo(
        file_id: str
    ):

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        try:

            (
                service
                .files()
                .delete(
                    fileId=file_id,
                    supportsAllDrives=True,
                )
                .execute()
            )

        except Exception as error:

            print(
                "⚠️ No se pudo eliminar "
                f"archivo Drive: {error}"
            )