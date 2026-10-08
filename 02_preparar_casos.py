from __future__ import annotations

"""
02_preparar_casos_FINAL.py
==========================

Prepara casos únicos para mapeo manual con Gemini reproduciendo la identidad
de mapas usada por LabCaptura V2.2.

IMPORTANTE:
- NO llama a Gemini.
- NO modifica los ERP.
- NO modifica las plantillas.
- NO sube mapas a Drive.
- Usa ParserERP.extraer_metadatos_archivo() del proyecto para detectar GST.
- Usa MapeadorIA.firma_estructura_origen() para la firma ERP.
- Usa GestorMapas._sha256() y GestorMapas._nombre_mapa() para obtener
  exactamente la identidad esperada del mapa.

Estructura esperada:
    ~/Downloads/MAPEO_MASIVO_GEOTEST/
        PLANTILLAS/
            GST-003/
                plantilla.xlsx
            GST-024/
                plantilla.xlsx
            ...
        CASOS_PARA_MAPEAR/       <- se reconstruye en cada ejecución

Cada caso único se agrupa por:
    GST de producción + SHA256 plantilla + firma estructural ERP
"""

import csv
import json
import os
import shutil
import sys
from collections import defaultdict
from pathlib import Path

# ---------------------------------------------------------------------------
# Localizar raíz de LabCaptura
# ---------------------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR

# Permite ejecutar el script desde la raíz del proyecto.
if not (PROJECT_ROOT / "DOMINIO").is_dir():
    cwd = Path.cwd().resolve()
    if (cwd / "DOMINIO").is_dir():
        PROJECT_ROOT = cwd

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    from DOMINIO.parser_erp import ParserERP
    from DOMINIO.mapeador_ia import MapeadorIA
    from DOMINIO.gestor_mapas import GestorMapas
except Exception as exc:
    raise SystemExit(
        "\nERROR: Este script debe ejecutarse desde la raíz de LabCaptura "
        "(donde existen DOMINIO, APLICACION e INFRAESTRUCTURA).\n"
        f"Detalle: {type(exc).__name__}: {exc}\n"
    )

BASE_MASIVO = Path.home() / "Downloads" / "MAPEO_MASIVO_GEOTEST"
PLANTILLAS_DIR = BASE_MASIVO / "PLANTILLAS"
SALIDA_DIR = BASE_MASIVO / "CASOS_PARA_MAPEAR"
RESUMEN_CSV = BASE_MASIVO / "resumen_casos.csv"
INCIDENCIAS_CSV = BASE_MASIVO / "incidencias.csv"

EXT_ERP = {".xls", ".xlsx", ".xlsm"}
EXT_PLANTILLA = {".xlsx", ".xlsm"}


def limpiar_texto(valor) -> str:
    return str(valor or "").strip()


def es_temporal_excel(path: Path) -> bool:
    # Excel crea archivos de bloqueo como "~$archivo.xlsx".
    return path.name.startswith("~$")


def listar_erp(carpeta: Path) -> list[Path]:
    archivos = []
    for p in carpeta.rglob("*"):
        if (
            p.is_file()
            and p.suffix.lower() in EXT_ERP
            and not es_temporal_excel(p)
        ):
            archivos.append(p)
    return sorted(archivos, key=lambda p: str(p).lower())


def normalizar_gst_para_carpeta(gst: str) -> str:
    """
    ParserERP devuelve, por ejemplo, GST-024 incluso si el nombre original
    contenía GST-024B. No reinterpretamos el nombre: usamos el resultado
    real del parser de LabCaptura.
    """
    gst = limpiar_texto(gst)
    if gst.lower() == "desconocido" or not gst:
        return ""
    return gst.upper()


def candidatos_plantilla(codigo_gst: str) -> list[Path]:
    """
    Paso 01 ya descargó la plantilla seleccionada por LabCaptura dentro de
    PLANTILLAS/GST-###. Por eso esa carpeta es la autoridad local.

    Se ignoran archivos temporales "~$".
    """
    carpeta = PLANTILLAS_DIR / codigo_gst

    if not carpeta.is_dir():
        return []

    candidatos = [
        p
        for p in carpeta.iterdir()
        if (
            p.is_file()
            and p.suffix.lower() in EXT_PLANTILLA
            and not es_temporal_excel(p)
        )
    ]
    return sorted(candidatos, key=lambda p: p.name.lower())


def resolver_plantilla(codigo_gst: str):
    candidatos = candidatos_plantilla(codigo_gst)

    if not candidatos:
        return None, "SIN_PLANTILLA", (
            f"No hay plantilla válida en {PLANTILLAS_DIR / codigo_gst}"
        )

    if len(candidatos) > 1:
        nombres = " | ".join(p.name for p in candidatos)
        return None, "PLANTILLA_AMBIGUA", (
            f"Hay {len(candidatos)} plantillas no temporales en "
            f"{PLANTILLAS_DIR / codigo_gst}: {nombres}"
        )

    return candidatos[0], "OK", ""


def ruta_relativa_segura(path: Path, raiz: Path) -> str:
    try:
        return str(path.relative_to(raiz))
    except Exception:
        return str(path)


def copiar_si_no_existe(origen: Path, destino: Path):
    destino.parent.mkdir(parents=True, exist_ok=True)
    if not destino.exists():
        shutil.copy2(origen, destino)


def escribir_csv(path: Path, filas: list[dict], campos: list[str]):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        for fila in filas:
            writer.writerow({campo: fila.get(campo, "") for campo in campos})


def main():
    print()
    carpeta_txt = input("Carpeta que contiene TODOS los ERP: ").strip().strip('"')
    carpeta_erp = Path(carpeta_txt).expanduser().resolve()

    if not carpeta_erp.is_dir():
        raise SystemExit(f"\nERROR: No existe la carpeta:\n{carpeta_erp}\n")

    if not PLANTILLAS_DIR.is_dir():
        raise SystemExit(
            "\nERROR: No existe la carpeta de plantillas descargadas:\n"
            f"{PLANTILLAS_DIR}\n"
            "Ejecuta primero 01_descargar_plantillas.py.\n"
        )

    erps = listar_erp(carpeta_erp)

    # Solo se reconstruyen productos de este script.
    if SALIDA_DIR.exists():
        shutil.rmtree(SALIDA_DIR)
    SALIDA_DIR.mkdir(parents=True, exist_ok=True)

    print()
    print("=" * 78)
    print("PREPARADOR DE CASOS ÚNICOS FINAL - GEOTEST / LABCAPTURA V2.2")
    print("=" * 78)
    print(f"ERP encontrados : {len(erps)}")
    print(f"Plantillas      : {PLANTILLAS_DIR}")
    print(f"Salida          : {SALIDA_DIR}")
    print("=" * 78)

    resumen = []
    incidencias = []
    grupos = defaultdict(list)

    for indice, ruta_erp in enumerate(erps, start=1):
        print()
        print(f"[{indice}/{len(erps)}] {ruta_erp.name}")

        try:
            # MISMO parser usado por la interfaz de LabCaptura.
            meta = ParserERP.extraer_metadatos_archivo(
                ruta_erp.name,
                str(ruta_erp),
            )
            codigo_gst = normalizar_gst_para_carpeta(meta.get("gst"))

            if not codigo_gst:
                mensaje = (
                    "ParserERP de LabCaptura devolvió 'Desconocido'. "
                    "El archivo no se fuerza ni se reinterpreta."
                )
                print("  ⚠️ GST no reconocido por LabCaptura.")
                incidencias.append({
                    "archivo": ruta_erp.name,
                    "ruta": str(ruta_erp),
                    "gst": "Desconocido",
                    "tipo": "GST_NO_RECONOCIDO",
                    "detalle": mensaje,
                })
                continue

            print(f"  GST LabCaptura : {codigo_gst}")

            plantilla, estado, detalle = resolver_plantilla(codigo_gst)
            if plantilla is None:
                print(f"  ❌ {estado}: {detalle}")
                incidencias.append({
                    "archivo": ruta_erp.name,
                    "ruta": str(ruta_erp),
                    "gst": codigo_gst,
                    "tipo": estado,
                    "detalle": detalle,
                })
                continue

            print(f"  Plantilla      : {plantilla.name}")

            firma_plantilla = GestorMapas._sha256(str(plantilla))
            firma_erp = MapeadorIA.firma_estructura_origen(str(ruta_erp))
            nombre_mapa = GestorMapas._nombre_mapa(
                firma_plantilla,
                firma_erp,
            )

            # La clave reproduce el ámbito real del mapa:
            # carpeta GST + nombre calculado por firmas.
            clave = (codigo_gst, firma_plantilla, firma_erp)

            registro = {
                "gst": codigo_gst,
                "archivo_erp": ruta_erp.name,
                "ruta_erp": str(ruta_erp),
                "archivo_plantilla": plantilla.name,
                "ruta_plantilla": str(plantilla),
                "firma_plantilla": firma_plantilla,
                "firma_estructura_origen": firma_erp,
                "nombre_mapa": nombre_mapa,
            }
            grupos[clave].append(registro)

            print(f"  Firma plantilla: {firma_plantilla[:16]}...")
            print(f"  Firma ERP      : {firma_erp[:16]}...")
            print(f"  Mapa esperado  : {nombre_mapa}")

        except Exception as exc:
            detalle = f"{type(exc).__name__}: {exc}"
            print(f"  ❌ ERROR: {detalle}")
            incidencias.append({
                "archivo": ruta_erp.name,
                "ruta": str(ruta_erp),
                "gst": "",
                "tipo": "ERROR_ANALISIS",
                "detalle": detalle,
            })

    # -----------------------------------------------------------------------
    # Materializar UN caso por identidad de mapa.
    # -----------------------------------------------------------------------
    claves = sorted(
        grupos.keys(),
        key=lambda x: (x[0], x[1], x[2]),
    )

    for numero, clave in enumerate(claves, start=1):
        registros = grupos[clave]
        primero = registros[0]

        codigo_gst = primero["gst"]
        nombre_mapa = primero["nombre_mapa"]

        nombre_caso = (
            f"CASO_{numero:03d}_{codigo_gst}_"
            f"{primero['firma_plantilla'][:8]}_"
            f"{primero['firma_estructura_origen'][:8]}"
        )
        carpeta_caso = SALIDA_DIR / nombre_caso
        carpeta_caso.mkdir(parents=True, exist_ok=True)

        origen_erp = Path(primero["ruta_erp"])
        origen_plantilla = Path(primero["ruta_plantilla"])

        # Nombres simples para que sea cómodo adjuntarlos a Gemini.
        destino_erp = carpeta_caso / f"ERP__{origen_erp.name}"
        destino_plantilla = carpeta_caso / f"PLANTILLA__{origen_plantilla.name}"

        copiar_si_no_existe(origen_erp, destino_erp)
        copiar_si_no_existe(origen_plantilla, destino_plantilla)

        manifiesto = {
            "caso": nombre_caso,
            "gst_labcaptura": codigo_gst,
            "cantidad_erp_equivalentes": len(registros),
            "erp_representativo": origen_erp.name,
            "plantilla": origen_plantilla.name,
            "firma_plantilla": primero["firma_plantilla"],
            "firma_estructura_origen": primero["firma_estructura_origen"],
            "nombre_mapa_esperado": nombre_mapa,
            "carpeta_drive_mapa": codigo_gst,
            "archivos_erp_equivalentes": [
                r["archivo_erp"] for r in registros
            ],
            "nota": (
                "Todos los ERP listados comparten la misma identidad estructural "
                "para esta plantilla. Solo se necesita crear un mapa para este caso."
            ),
        }

        with (carpeta_caso / "MANIFIESTO.json").open(
            "w", encoding="utf-8"
        ) as f:
            json.dump(manifiesto, f, ensure_ascii=False, indent=2)

        with (carpeta_caso / "NOMBRE_MAPA_ESPERADO.txt").open(
            "w", encoding="utf-8"
        ) as f:
            f.write(nombre_mapa + "\n")

        resumen.append({
            "caso": nombre_caso,
            "gst": codigo_gst,
            "cantidad_erp_equivalentes": len(registros),
            "erp_representativo": origen_erp.name,
            "plantilla": origen_plantilla.name,
            "firma_plantilla": primero["firma_plantilla"],
            "firma_estructura_origen": primero["firma_estructura_origen"],
            "nombre_mapa": nombre_mapa,
            "carpeta_caso": str(carpeta_caso),
            "otros_erp": " | ".join(
                r["archivo_erp"] for r in registros[1:]
            ),
        })

    escribir_csv(
        RESUMEN_CSV,
        resumen,
        [
            "caso",
            "gst",
            "cantidad_erp_equivalentes",
            "erp_representativo",
            "plantilla",
            "firma_plantilla",
            "firma_estructura_origen",
            "nombre_mapa",
            "carpeta_caso",
            "otros_erp",
        ],
    )

    escribir_csv(
        INCIDENCIAS_CSV,
        incidencias,
        ["archivo", "ruta", "gst", "tipo", "detalle"],
    )

    print()
    print("=" * 78)
    print("RESUMEN FINAL")
    print("=" * 78)
    print(f"ERP analizados                  : {len(erps)}")
    print(f"ERP con incidencia              : {len(incidencias)}")
    print(f"Identidades de mapa únicas      : {len(resumen)}")
    print(f"Mapas manuales máximos a crear  : {len(resumen)}")
    print(f"Resumen CSV                     : {RESUMEN_CSV}")
    print(f"Incidencias CSV                 : {INCIDENCIAS_CSV}")
    print(f"Casos para Gemini               : {SALIDA_DIR}")
    print("=" * 78)

    if resumen:
        print()
        print("SIGUIENTE PASO:")
        print(
            "Abre CASOS_PARA_MAPEAR y trabaja un CASO a la vez. "
            "Cada carpeta contiene el ERP representativo, la plantilla y "
            "el nombre exacto que deberá tener el mapa final."
        )


if __name__ == "__main__":
    main()
