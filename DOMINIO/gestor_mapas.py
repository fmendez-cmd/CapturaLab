from __future__ import annotations

import hashlib
import copy
from pathlib import Path

from DOMINIO.mapeador_ia import MapeadorIA
from DOMINIO.validador_mapa import ValidadorMapa
from INFRAESTRUCTURA.drive_connector import DriveConnector
from DOMINIO.control_cancelacion import verificar_cancelacion
import time


class GestorMapas:
    # Caché VOLÁTIL de mapas. Solo existe durante un lote.
    _cache_lote_activo = False
    _cache_mapas_lote = {}

    @staticmethod
    def iniciar_cache_lote():
        GestorMapas._cache_lote_activo = True
        GestorMapas._cache_mapas_lote = {}
        print("⚡ Caché temporal de mapas iniciado.")

    @staticmethod
    def limpiar_cache_lote():
        GestorMapas._cache_mapas_lote = {}
        GestorMapas._cache_lote_activo = False
        print("🧹 Caché temporal de mapas limpiado.")

    @staticmethod
    def _clave_cache(codigo_gst, nombre_mapa):
        return (str(codigo_gst).upper().strip(), str(nombre_mapa))

    @staticmethod
    def _guardar_cache(codigo_gst, nombre_mapa, mapa):
        if not GestorMapas._cache_lote_activo:
            return
        copia = copy.deepcopy(mapa)
        copia.pop("_reutilizado_desde_drive", None)
        GestorMapas._cache_mapas_lote[
            GestorMapas._clave_cache(codigo_gst, nombre_mapa)
        ] = copia

    """
    Gestor V2.2 sin caché local.

    Drive es la única fuente persistente de mapas.
    La firma de la plantilla y la firma estructural del ERP (incluido el patrón de ocupación) forman
    el nombre exacto del mapa; si cambia cualquiera, el mapa anterior
    deja de ser candidato automáticamente.
    """

    @staticmethod
    def _sha256(ruta):
        h = hashlib.sha256()
        with open(ruta, "rb") as f:
            for bloque in iter(lambda: f.read(1024 * 1024), b""):
                h.update(bloque)
        return h.hexdigest()

    @staticmethod
    def _nombre_mapa(firma_plantilla, firma_origen):
        return (
            f"v22_{firma_plantilla[:16]}_"
            f"{firma_origen[:16]}.json"
        )

    @staticmethod
    def _descargar_oficial_drive(
        codigo_gst,
        nombre_mapa,
        ruta_erp,
        ruta_plantilla,
    ):
        try:
            raiz = DriveConnector.obtener_id_carpeta_mapas()
            carpeta_gst = DriveConnector.buscar_carpeta(
                raiz, str(codigo_gst)
            )
            if not carpeta_gst:
                return None

            remoto = DriveConnector.buscar_archivo(
                carpeta_gst, nombre_mapa
            )
            if not remoto:
                return None

            print(f"☁️ Mapa encontrado en Drive: {nombre_mapa}")
            mapa = DriveConnector.descargar_json(remoto["id"])

            # El nombre buscado se calculó con las firmas ACTUALES
            # de plantilla + estructura ERP. Por tanto, encontrar
            # exactamente este archivo ya acredita la compatibilidad
            # estructural para reutilización.
            #
            # Conservamos la validación del esquema del JSON para no
            # aceptar un archivo corrupto o incompleto.
            ValidadorMapa.validar_esquema(mapa)

            # Marca SOLO de ejecución. No se persiste en Drive.
            mapa["_reutilizado_desde_drive"] = True
            return mapa

        except Exception as error:
            print(
                "⚠️ No se pudo consultar Drive para mapas: "
                f"{error}"
            )
            return None

    @staticmethod
    def _subir_oficial_drive(
        codigo_gst,
        nombre_mapa,
        mapa,
    ):
        raiz = DriveConnector.obtener_id_carpeta_mapas()
        carpeta_gst = DriveConnector.obtener_o_crear_carpeta(
            raiz, str(codigo_gst)
        )
        DriveConnector.subir_json(
            carpeta_gst, nombre_mapa, mapa
        )
        print("☁️ Mapa oficial V2.2 guardado en Drive.")

    @staticmethod
    def obtener_o_crear(
        codigo_gst,
        ruta_erp,
        ruta_plantilla,
        forzar_nuevo=False,
        cancel_event=None,
    ):
        _t_mapa_total = time.perf_counter()
        verificar_cancelacion()

        _t = time.perf_counter()
        firma_plantilla = GestorMapas._sha256(ruta_plantilla)
        print(f"⏱️ Firma SHA plantilla: {time.perf_counter() - _t:.3f} s")
        firma_origen = MapeadorIA.firma_estructura_origen(ruta_erp)
        nombre_mapa = GestorMapas._nombre_mapa(
            firma_plantilla,
            firma_origen,
        )

        if not forzar_nuevo and GestorMapas._cache_lote_activo:
            clave_cache = GestorMapas._clave_cache(codigo_gst, nombre_mapa)
            mapa_cache = GestorMapas._cache_mapas_lote.get(clave_cache)
            if mapa_cache is not None:
                mapa_cache = copy.deepcopy(mapa_cache)
                # Ya fue validado/recuperado o creado y subido durante este mismo lote.
                mapa_cache["_reutilizado_desde_drive"] = True
                print(f"⚡ Mapa recuperado del caché del lote: {nombre_mapa}")
                print("🚫 Drive/Gemini NO serán llamados para este mapa.")
                print(
                    f"⏱️ TOTAL GestorMapas.obtener_o_crear: "
                    f"{time.perf_counter() - _t_mapa_total:.3f} s"
                )
                return mapa_cache

        print("\n" + "=" * 78)
        print("MAPA ESTRUCTURAL V2.2")
        print("=" * 78)
        print(f"GST: {codigo_gst}")
        print(f"Mapa oficial Drive: {nombre_mapa}")
        print("=" * 78)

        verificar_cancelacion()

        if not forzar_nuevo:
            _t = time.perf_counter()
            mapa_drive = GestorMapas._descargar_oficial_drive(
                codigo_gst,
                nombre_mapa,
                ruta_erp,
                ruta_plantilla,
            )
            print(f"⏱️ Consulta + descarga mapa Drive: {time.perf_counter() - _t:.3f} s")
            if mapa_drive is not None:
                GestorMapas._guardar_cache(codigo_gst, nombre_mapa, mapa_drive)
                print("⚡ Mapa guardado en caché hasta terminar el lote.")
                print("🚫 Gemini NO será llamado.")
                print(
                    f"⏱️ TOTAL GestorMapas.obtener_o_crear: "
                    f"{time.perf_counter() - _t_mapa_total:.3f} s"
                )
                return mapa_drive

        verificar_cancelacion()

        print(
            "⚠️ No existe en Drive un mapa compatible para "
            "esta plantilla y estructura. Se llamará a Gemini."
        )

        mapa = MapeadorIA.generar_mapa(
            codigo_gst,
            ruta_erp,
            ruta_plantilla,
            cancel_event=cancel_event,
        )

        # La reconstrucción V2.2 sigue siendo obligatoria.
        mapa["firma_plantilla"] = firma_plantilla
        mapa["firma_estructura_origen"] = firma_origen

        mapa = ValidadorMapa.validar_y_reconstruir_con_archivos(
            mapa,
            ruta_erp,
            ruta_plantilla,
        )

        # El mapa validado se persiste únicamente en Drive.
        GestorMapas._subir_oficial_drive(
            codigo_gst,
            nombre_mapa,
            mapa,
        )

        # Solo DESPUÉS de persistirlo correctamente en Drive entra al caché.
        GestorMapas._guardar_cache(codigo_gst, nombre_mapa, mapa)
        print("⚡ Mapa nuevo guardado en caché hasta terminar el lote.")

        # Este mapa acaba de ser creado y validado por primera vez.
        # El GeneradorExcel hará la ruta completa de seguridad en esta
        # primera ejecución.
        mapa["_reutilizado_desde_drive"] = False

        verificar_cancelacion()
        return mapa
