from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path

# Permite ejecutar este archivo desde la raíz del proyecto o desde una subcarpeta.
SCRIPT_DIR = Path(__file__).resolve().parent
CANDIDATOS_RAIZ = [SCRIPT_DIR, SCRIPT_DIR.parent]

for candidato in CANDIDATOS_RAIZ:
    if (candidato / "INFRAESTRUCTURA" / "drive_connector.py").exists():
        sys.path.insert(0, str(candidato))
        break
else:
    raise RuntimeError(
        "No se encontró INFRAESTRUCTURA/drive_connector.py. "
        "Coloca este script dentro de la raíz de LabCaptura "
        "o en una subcarpeta directa del proyecto."
    )

from INFRAESTRUCTURA.drive_connector import DriveConnector


PATRON_GST = re.compile(r"GST-(\d+)", re.IGNORECASE)
EXTENSIONES_ERP = {".xls", ".xlsx", ".xlsm"}


def detectar_gst_en_nombre(nombre: str) -> str | None:
    match = PATRON_GST.search(nombre)
    if not match:
        return None
    return f"GST-{int(match.group(1)):03d}"


def buscar_gst_en_carpeta(carpeta_erp: Path) -> tuple[list[str], list[Path]]:
    gst = set()
    no_reconocidos = []

    for ruta in carpeta_erp.rglob("*"):
        if not ruta.is_file() or ruta.suffix.lower() not in EXTENSIONES_ERP:
            continue

        codigo = detectar_gst_en_nombre(ruta.name)
        if codigo:
            gst.add(codigo)
        else:
            no_reconocidos.append(ruta)

    return sorted(gst, key=lambda x: int(x.split("-")[1])), no_reconocidos


def copiar_plantilla(codigo_gst: str, destino_raiz: Path) -> tuple[bool, str]:
    ruta_temporal = None
    try:
        # IMPORTANTE:
        # Esta es exactamente la misma función que usa LabCaptura.
        ruta_temporal = Path(
            DriveConnector.obtener_plantilla_por_gst(codigo_gst)
        ).resolve()

        if not ruta_temporal.is_file():
            raise FileNotFoundError(
                f"DriveConnector devolvió una ruta inexistente: {ruta_temporal}"
            )

        carpeta_gst = destino_raiz / codigo_gst
        carpeta_gst.mkdir(parents=True, exist_ok=True)

        destino = carpeta_gst / ruta_temporal.name
        shutil.copy2(ruta_temporal, destino)

        return True, str(destino)

    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Detecta los GST presentes en una carpeta de ERP y descarga "
            "sus plantillas usando exactamente DriveConnector de LabCaptura."
        )
    )
    parser.add_argument(
        "carpeta_erp",
        nargs="?",
        help="Carpeta que contiene los ERP (.xls/.xlsx/.xlsm).",
    )
    parser.add_argument(
        "--salida",
        help=(
            "Carpeta destino. Por defecto: "
            "Descargas/MAPEO_MASIVO_GEOTEST/PLANTILLAS"
        ),
    )
    args = parser.parse_args()

    if args.carpeta_erp:
        carpeta_erp = Path(args.carpeta_erp).expanduser().resolve()
    else:
        entrada = input("Carpeta que contiene los ERP: ").strip().strip('"')
        carpeta_erp = Path(entrada).expanduser().resolve()

    if not carpeta_erp.is_dir():
        raise NotADirectoryError(f"No existe la carpeta: {carpeta_erp}")

    if args.salida:
        destino_raiz = Path(args.salida).expanduser().resolve()
    else:
        destino_raiz = (
            Path.home()
            / "Downloads"
            / "MAPEO_MASIVO_GEOTEST"
            / "PLANTILLAS"
        )

    destino_raiz.mkdir(parents=True, exist_ok=True)

    print("\n" + "=" * 72)
    print("PREPARADOR MASIVO DE PLANTILLAS - GEOTEST")
    print("=" * 72)
    print(f"ERP origen : {carpeta_erp}")
    print(f"Salida     : {destino_raiz}")

    gst_encontrados, no_reconocidos = buscar_gst_en_carpeta(carpeta_erp)

    print(f"\nGST únicos encontrados: {len(gst_encontrados)}")
    if gst_encontrados:
        print(", ".join(gst_encontrados))

    if no_reconocidos:
        print(
            f"\nAviso: {len(no_reconocidos)} archivo(s) Excel no contenían "
            "'GST-###' en el nombre y se ignoraron."
        )

    if not gst_encontrados:
        print("\nNo se encontraron GST para procesar.")
        return

    # LabCaptura usa caché temporal por lote. La usamos también aquí para
    # mantener el mismo comportamiento y evitar descargas repetidas.
    DriveConnector.iniciar_cache_lote()

    exitosos = []
    errores = []

    try:
        total = len(gst_encontrados)

        for indice, codigo_gst in enumerate(gst_encontrados, start=1):
            print(f"\n[{indice}/{total}] {codigo_gst}")
            print("-" * 72)

            ok, detalle = copiar_plantilla(codigo_gst, destino_raiz)

            if ok:
                exitosos.append((codigo_gst, detalle))
                print(f"OK: {detalle}")
            else:
                errores.append((codigo_gst, detalle))
                print(f"ERROR: {detalle}")

    finally:
        # El conector puede borrar su copia TEMPORAL.
        # Nuestras copias de Descargas no se ven afectadas.
        DriveConnector.limpiar_cache_lote()

    print("\n" + "=" * 72)
    print("RESUMEN")
    print("=" * 72)
    print(f"GST detectados             : {len(gst_encontrados)}")
    print(f"Plantillas copiadas        : {len(exitosos)}")
    print(f"Plantillas con incidencia  : {len(errores)}")
    print(f"Destino                    : {destino_raiz}")

    if errores:
        print("\nINCIDENCIAS:")
        for gst, error in errores:
            print(f"  - {gst}: {error}")

    print("\nProceso terminado.")


if __name__ == "__main__":
    main()
