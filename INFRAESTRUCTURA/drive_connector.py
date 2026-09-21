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

        codigo_limpio = (
            codigo_gst
            .upper()
            .strip()
        )

        # No existe caché local de plantillas.
        # Drive es siempre la fuente oficial.
        os.makedirs(TEMP_PLANTILLAS_DIR, exist_ok=True)

        # ------------------------------------------------------
        # DRIVE: FUENTE OFICIAL
        # ------------------------------------------------------

        folder_id = (
            DriveConnector
            .obtener_id_carpeta_plantillas()
        )

        service = (
            DriveConnector
            ._obtener_servicio_drive()
        )

        codigo_query = (
            DriveConnector
            ._escapar_query(
                codigo_limpio
            )
        )

        query = (
            f"'{folder_id}' in parents "
            f"and name contains '{codigo_query}' "
            f"and trashed = false"
        )

        results = (
            service
            .files()
            .list(
                q=query,
                pageSize=20,
                fields=(
                    "files("
                    "id,"
                    "name,"
                    "mimeType,"
                    "modifiedTime"
                    ")"
                ),
                orderBy="modifiedTime desc",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            )
            .execute()
        )

        items = results.get(
            "files",
            []
        )

        # Solo Excel
        items = [
            item
            for item in items
            if item[
                "name"
            ].lower().endswith(
                (
                    ".xlsx",
                    ".xlsm",
                )
            )
        ]

        if not items:

            raise FileNotFoundError(
                "No se encontró ninguna "
                "plantilla para el código "
                f"'{codigo_gst}' "
                "en Google Drive."
            )

        # La más recientemente modificada.
        item = items[0]

        file_id = item[
            "id"
        ]

        file_name = item[
            "name"
        ]

        ruta_destino = os.path.join(
            TEMP_PLANTILLAS_DIR,
            file_name
        )

        print(
            "📥 Descargando plantilla "
            f"oficial '{file_name}'..."
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

        with open(
            ruta_destino,
            "wb"
        ) as archivo:

            archivo.write(
                fh.getvalue()
            )

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