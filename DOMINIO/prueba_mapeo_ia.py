# DOMINIO/prueba_mapeo_ia.py
#
# Cambia únicamente NOMBRE_ERP para probar cualquier GST.

from pathlib import Path
import re

from DOMINIO.generador_excel import GeneradorExcel


BASE = Path(__file__).resolve().parent.parent

NOMBRE_ERP = (
    "E-1524_GST-166_17215_17215-1273_PCA-19.xls"
)

RUTA_ERP = (
    BASE
    / "ARCHIVOS_ENTRADA"
    / "ERP"
    / NOMBRE_ERP
)


def extraer_gst(nombre: str) -> str:
    match = re.search(
        r"GST[-_ ]?(\d{3})",
        nombre.upper(),
    )

    if not match:
        raise ValueError(
            f"No pude detectar GST en: {nombre}"
        )

    return f"GST-{match.group(1)}"


def main():
    if not RUTA_ERP.is_file():
        raise FileNotFoundError(
            f"No existe:\n{RUTA_ERP}"
        )

    codigo_gst = extraer_gst(
        RUTA_ERP.name
    )

    salida = (
        BASE
        / "SALIDAS_PRUEBA"
        / f"{RUTA_ERP.stem}_MAPA_IA_V2.xlsx"
    )

    print()
    print("=" * 80)
    print("PRUEBA MAPEO ESTRUCTURAL IA V2")
    print("=" * 80)
    print(f"ERP: {RUTA_ERP}")
    print(f"GST: {codigo_gst}")
    print(f"Salida: {salida}")

    resultado = (
        GeneradorExcel.inyectar_datos_dinamicos(
            codigo_gst,
            str(RUTA_ERP),
            str(salida),
        )
    )

    print()
    print("=" * 80)
    print("RESULTADO")
    print("=" * 80)
    print(resultado)


if __name__ == "__main__":
    main()
