from __future__ import annotations

"""
00_preparar_casos_desde_labcaptura.py

PREPARADOR MASIVO ESTRICTO:
reproduce las decisiones del flujo ACTUAL de LabCaptura sin volver a
interpretar GST ni elegir plantillas por cuenta propia.

Para cada ERP:
    ParserERP.extraer_metadatos_archivo(...)
        -> codigo_gst
    DriveConnector.obtener_plantilla_por_gst(codigo_gst)
        -> plantilla oficial que LabCaptura usaría
    MapeadorIA.firma_estructura_origen(...)
        -> firma ERP
    GestorMapas._sha256(...)
        -> firma plantilla
    GestorMapas._nombre_mapa(...)
        -> nombre exacto del mapa V2.2

NO llama a Gemini.
NO modifica LabCaptura.
NO decide variantes.
NO busca plantillas en carpetas locales.
"""

import csv
import json
import shutil
import sys
from collections import OrderedDict
from pathlib import Path

# ---------------------------------------------------------------------
# Localizar la raíz real del proyecto
# ---------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent
candidatos = [SCRIPT_DIR, Path.cwd().resolve(), SCRIPT_DIR.parent]

PROJECT_ROOT = None
for candidato in candidatos:
    if (
        (candidato / "DOMINIO" / "parser_erp.py").is_file()
        and (candidato / "DOMINIO" / "mapeador_ia.py").is_file()
        and (candidato / "DOMINIO" / "gestor_mapas.py").is_file()
        and (candidato / "INFRAESTRUCTURA" / "drive_connector.py").is_file()
    ):
        PROJECT_ROOT = candidato
        break

if PROJECT_ROOT is None:
    raise SystemExit(
        "\nERROR: No se encontró la raíz actual de LabCaptura.\n"
        "Coloca este script en C:\\Users\\Gerencia TI\\Documents\\LabCaptura "
        "y ejecútalo desde ahí.\n"
    )

sys.path.insert(0, str(PROJECT_ROOT))

from DOMINIO.parser_erp import ParserERP
from DOMINIO.mapeador_ia import MapeadorIA
from DOMINIO.gestor_mapas import GestorMapas
from INFRAESTRUCTURA.drive_connector import DriveConnector

# La interfaz actual de LabCaptura admite .xls y .xlsx.
EXT_ERP = {".xls", ".xlsx"}

BASE = Path.home() / "Downloads" / "MAPEO_MASIVO_GEOTEST"
SALIDA = BASE / "CASOS_PARA_MAPEAR"
RESUMEN = BASE / "resumen_casos.csv"
INCIDENCIAS = BASE / "incidencias.csv"
AUDITORIA = BASE / "auditoria_flujo_labcaptura.csv"


def _limpiar_salida():
    if SALIDA.exists():
        shutil.rmtree(SALIDA)
    SALIDA.mkdir(parents=True, exist_ok=True)


def _copiar_sin_colision(origen: Path, destino: Path):
    if destino.exists():
        destino.unlink()
    shutil.copy2(origen, destino)


def main():
    print("\n" + "=" * 82)
    print("PREPARADOR MASIVO - FLUJO REAL DE LABCAPTURA")
    print("=" * 82)
    print(f"Proyecto LabCaptura: {PROJECT_ROOT}")

    entrada = input("Carpeta que contiene TODOS los ERP: ").strip().strip('"')
    carpeta_erp = Path(entrada).expanduser().resolve()

    if not carpeta_erp.is_dir():
        raise SystemExit(f"\nERROR: No existe la carpeta:\n{carpeta_erp}\n")

    erps = sorted(
        (
            p for p in carpeta_erp.rglob("*")
            if p.is_file()
            and not p.name.startswith("~$")
            and p.suffix.lower() in EXT_ERP
        ),
        key=lambda p: str(p).lower(),
    )

    _limpiar_salida()

    print(f"ERP encontrados : {len(erps)}")
    print(f"Salida          : {SALIDA}")
    print("=" * 82)

    # Igual que el lote real: activar caché temporal de plantillas.
    cache_iniciado = False
    if hasattr(DriveConnector, "iniciar_cache_lote"):
        DriveConnector.iniciar_cache_lote()
        cache_iniciado = True

    # clave REAL de identidad de mapa:
    # (GST recibido por LabCaptura, firma plantilla, firma ERP)
    grupos = OrderedDict()
    incidencias = []
    auditoria = []

    try:
        for n, ruta_erp in enumerate(erps, 1):
            print(f"\n[{n}/{len(erps)}] {ruta_erp.name}")

            try:
                # -------------------------------------------------
                # 1. EXACTAMENTE ParserERP
                # -------------------------------------------------
                meta = ParserERP.extraer_metadatos_archivo(
                    ruta_erp.name,
                    str(ruta_erp),
                )
                codigo_gst = str(meta.get("gst", "")).strip()

                print(f"  ParserERP GST : {codigo_gst}")

                if not codigo_gst or codigo_gst == "Desconocido":
                    raise ValueError(
                        "ParserERP de LabCaptura devolvió GST 'Desconocido'."
                    )

                # -------------------------------------------------
                # 2. EXACTAMENTE DriveConnector que usa GeneradorExcel
                # -------------------------------------------------
                ruta_plantilla = Path(
                    DriveConnector.obtener_plantilla_por_gst(codigo_gst)
                ).resolve()

                if not ruta_plantilla.is_file():
                    raise FileNotFoundError(
                        "DriveConnector devolvió una plantilla inexistente: "
                        f"{ruta_plantilla}"
                    )

                print(f"  Plantilla REAL: {ruta_plantilla.name}")

                # -------------------------------------------------
                # 3. EXACTAMENTE las firmas del mapa V2.2
                # -------------------------------------------------
                firma_plantilla = GestorMapas._sha256(str(ruta_plantilla))
                firma_erp = MapeadorIA.firma_estructura_origen(str(ruta_erp))
                nombre_mapa = GestorMapas._nombre_mapa(
                    firma_plantilla,
                    firma_erp,
                )

                print(f"  Firma plantilla: {firma_plantilla[:16]}...")
                print(f"  Firma ERP      : {firma_erp[:16]}...")
                print(f"  Mapa esperado  : {nombre_mapa}")

                clave = (
                    codigo_gst.upper().strip(),
                    firma_plantilla,
                    firma_erp,
                )

                if clave not in grupos:
                    grupos[clave] = {
                        "gst": codigo_gst,
                        "plantilla": ruta_plantilla,
                        "firma_plantilla": firma_plantilla,
                        "firma_erp": firma_erp,
                        "nombre_mapa": nombre_mapa,
                        "erps": [],
                    }

                grupos[clave]["erps"].append(ruta_erp)

                auditoria.append({
                    "archivo_erp": ruta_erp.name,
                    "ruta_erp": str(ruta_erp),
                    "gst_parser_erp": codigo_gst,
                    "plantilla_driveconnector": ruta_plantilla.name,
                    "ruta_plantilla_temporal": str(ruta_plantilla),
                    "firma_plantilla": firma_plantilla,
                    "firma_erp": firma_erp,
                    "nombre_mapa": nombre_mapa,
                    "estado": "OK",
                    "error": "",
                })

            except Exception as exc:
                error = f"{type(exc).__name__}: {exc}"
                print(f"  ❌ {error}")

                incidencias.append({
                    "archivo_erp": ruta_erp.name,
                    "ruta_erp": str(ruta_erp),
                    "error": error,
                })

                auditoria.append({
                    "archivo_erp": ruta_erp.name,
                    "ruta_erp": str(ruta_erp),
                    "gst_parser_erp": locals().get("codigo_gst", ""),
                    "plantilla_driveconnector": "",
                    "ruta_plantilla_temporal": "",
                    "firma_plantilla": "",
                    "firma_erp": "",
                    "nombre_mapa": "",
                    "estado": "ERROR",
                    "error": error,
                })

    finally:
        if cache_iniciado and hasattr(DriveConnector, "limpiar_cache_lote"):
            DriveConnector.limpiar_cache_lote()

    # -----------------------------------------------------------------
    # Crear UN caso por identidad exacta de mapa.
    # Copiamos la plantilla que DriveConnector realmente devolvió.
    # -----------------------------------------------------------------
    filas_resumen = []

    for indice, grupo in enumerate(grupos.values(), 1):
        gst_seguro = grupo["gst"].replace("/", "_").replace("\\", "_")
        nombre_caso = (
            f"CASO_{indice:03d}_{gst_seguro}_"
            f"{grupo['firma_plantilla'][:8]}_"
            f"{grupo['firma_erp'][:8]}"
        )
        carpeta = SALIDA / nombre_caso
        carpeta.mkdir(parents=True, exist_ok=True)

        erp_representativo = grupo["erps"][0]
        plantilla = grupo["plantilla"]

        destino_erp = carpeta / f"ERP__{erp_representativo.name}"
        destino_plantilla = carpeta / f"PLANTILLA__{plantilla.name}"

        _copiar_sin_colision(erp_representativo, destino_erp)
        _copiar_sin_colision(plantilla, destino_plantilla)

        manifiesto = {
            "origen_decisiones": "FLUJO_REAL_LABCAPTURA",
            "gst_labcaptura": grupo["gst"],
            "cantidad_erp_equivalentes": len(grupo["erps"]),
            "erp_representativo": erp_representativo.name,
            "plantilla_devuelta_por_driveconnector": plantilla.name,
            "firma_plantilla": grupo["firma_plantilla"],
            "firma_estructura_origen": grupo["firma_erp"],
            "nombre_mapa_esperado": grupo["nombre_mapa"],
            "carpeta_drive_mapa": grupo["gst"],
            "erp_equivalentes": [p.name for p in grupo["erps"]],
            "flujo": [
                "ParserERP.extraer_metadatos_archivo",
                "DriveConnector.obtener_plantilla_por_gst",
                "GestorMapas._sha256",
                "MapeadorIA.firma_estructura_origen",
                "GestorMapas._nombre_mapa",
            ],
        }

        (carpeta / "MANIFIESTO.json").write_text(
            json.dumps(manifiesto, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        (carpeta / "NOMBRE_MAPA_ESPERADO.txt").write_text(
            grupo["nombre_mapa"] + "\n",
            encoding="utf-8",
        )

        filas_resumen.append({
            "caso": nombre_caso,
            "gst_labcaptura": grupo["gst"],
            "cantidad_erp_equivalentes": len(grupo["erps"]),
            "erp_representativo": erp_representativo.name,
            "plantilla_driveconnector": plantilla.name,
            "firma_plantilla": grupo["firma_plantilla"],
            "firma_erp": grupo["firma_erp"],
            "nombre_mapa_esperado": grupo["nombre_mapa"],
        })

    def escribir_csv(ruta: Path, filas: list[dict], campos: list[str]):
        with ruta.open("w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=campos)
            w.writeheader()
            w.writerows(filas)

    escribir_csv(
        RESUMEN,
        filas_resumen,
        [
            "caso", "gst_labcaptura", "cantidad_erp_equivalentes",
            "erp_representativo", "plantilla_driveconnector",
            "firma_plantilla", "firma_erp", "nombre_mapa_esperado",
        ],
    )

    escribir_csv(
        INCIDENCIAS,
        incidencias,
        ["archivo_erp", "ruta_erp", "error"],
    )

    escribir_csv(
        AUDITORIA,
        auditoria,
        [
            "archivo_erp", "ruta_erp", "gst_parser_erp",
            "plantilla_driveconnector", "ruta_plantilla_temporal",
            "firma_plantilla", "firma_erp", "nombre_mapa",
            "estado", "error",
        ],
    )

    print("\n" + "=" * 82)
    print("RESUMEN")
    print("=" * 82)
    print(f"ERP encontrados             : {len(erps)}")
    print(f"ERP procesados correctamente: {len(erps) - len(incidencias)}")
    print(f"Incidencias                 : {len(incidencias)}")
    print(f"Casos/mapas únicos          : {len(grupos)}")
    print(f"Casos                       : {SALIDA}")
    print(f"Auditoría flujo             : {AUDITORIA}")
    print(f"Resumen casos               : {RESUMEN}")
    print(f"Incidencias                 : {INCIDENCIAS}")
    print("=" * 82)
    print("\nIMPORTANTE: revisa primero auditoria_flujo_labcaptura.csv.")
    print("La columna 'plantilla_driveconnector' muestra la plantilla que")
    print("LabCaptura ACTUAL realmente selecciona para cada ERP.\n")


if __name__ == "__main__":
    main()
