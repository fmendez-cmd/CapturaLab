from __future__ import annotations

"""
03_generar_prompts_gemini.py

Genera, para cada carpeta CASO_* creada por 02_preparar_casos_FINAL.py:
- SYSTEM_INSTRUCTION.txt: exactamente MapeadorIA._system_instruction()
- PROMPT_GEMINI.txt: exactamente el `contents` que LabCaptura manda a Gemini
- CONTEXTO_LABCAPTURA.json: contexto estructural producido por construir_contexto()
- INSTRUCCIONES_MANUALES.txt: guía para reproducir la llamada manual

NO llama a Gemini y NO modifica los Excel.
"""

import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR
if not (PROJECT_ROOT / "DOMINIO").is_dir():
    cwd = Path.cwd().resolve()
    if (cwd / "DOMINIO").is_dir():
        PROJECT_ROOT = cwd

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

try:
    from DOMINIO.mapeador_ia import MapeadorIA
except Exception as exc:
    raise SystemExit(
        "\nERROR: Ejecuta este script desde la raíz actual de LabCaptura.\n"
        f"Detalle: {type(exc).__name__}: {exc}\n"
    )

BASE = Path.home() / "Downloads" / "MAPEO_MASIVO_GEOTEST"
CASOS_DIR = BASE / "CASOS_PARA_MAPEAR"


def unico(carpeta: Path, patron: str) -> Path:
    archivos = sorted(
        [p for p in carpeta.glob(patron) if p.is_file() and not p.name.startswith("~$")],
        key=lambda p: p.name.lower(),
    )
    if len(archivos) != 1:
        raise RuntimeError(
            f"Se esperaba exactamente 1 archivo '{patron}' en {carpeta.name}; "
            f"se encontraron {len(archivos)}."
        )
    return archivos[0]


def main():
    if not CASOS_DIR.is_dir():
        raise SystemExit(
            f"\nERROR: No existe:\n{CASOS_DIR}\n"
            "Ejecuta primero 02_preparar_casos_FINAL.py.\n"
        )

    casos = sorted([p for p in CASOS_DIR.iterdir() if p.is_dir() and p.name.startswith("CASO_")])
    print("\n" + "=" * 78)
    print("GENERADOR DE PROMPTS GEMINI - ESPEJO DE LABCAPTURA V2.2")
    print("=" * 78)
    print(f"Casos encontrados: {len(casos)}")
    print("=" * 78)

    ok = 0
    errores = []

    for i, carpeta in enumerate(casos, 1):
        print(f"\n[{i}/{len(casos)}] {carpeta.name}")
        try:
            manifiesto_path = carpeta / "MANIFIESTO.json"
            if not manifiesto_path.is_file():
                raise RuntimeError("Falta MANIFIESTO.json")

            manifiesto = json.loads(manifiesto_path.read_text(encoding="utf-8"))
            gst = str(manifiesto["gst_labcaptura"]).strip()
            erp = unico(carpeta, "ERP__*")
            plantilla = unico(carpeta, "PLANTILLA__*")

            # EXACTAMENTE la misma construcción usada por generar_mapa().
            contexto = MapeadorIA.construir_contexto(gst, str(erp), str(plantilla))
            system_instruction = MapeadorIA._system_instruction()
            prompt = (
                "Genera el mapa estructural completo CELDA->CELDA.\n\n"
                + json.dumps(
                    contexto,
                    ensure_ascii=False,
                    separators=(",", ":"),
                    default=str,
                )
            )

            (carpeta / "SYSTEM_INSTRUCTION.txt").write_text(
                system_instruction, encoding="utf-8"
            )
            (carpeta / "PROMPT_GEMINI.txt").write_text(prompt, encoding="utf-8")
            (carpeta / "CONTEXTO_LABCAPTURA.json").write_text(
                json.dumps(contexto, ensure_ascii=False, indent=2, default=str),
                encoding="utf-8",
            )

            guia = f"""CASO: {carpeta.name}
GST: {gst}
ERP: {erp.name}
PLANTILLA: {plantilla.name}
MAPA FINAL ESPERADO: {manifiesto.get('nombre_mapa_esperado', '')}

OBJETIVO
Reproducir manualmente la entrada que LabCaptura envía a Gemini.

IMPORTANTE
1. SYSTEM_INSTRUCTION.txt contiene la instrucción de sistema exacta de MapeadorIA.
2. PROMPT_GEMINI.txt contiene el contents exacto construido por LabCaptura.
3. CONTEXTO_LABCAPTURA.json es una copia legible del contexto; NO sustituye al prompt exacto.
4. Los Excel se conservan como referencia del caso, pero el prompt ya contiene el contexto
   estructural que LabCaptura extrajo de ellos.
5. Guarda la respuesta de Gemini como RESPUESTA_GEMINI.json dentro de esta misma carpeta.
6. No renombres todavía la respuesta al nombre final del mapa. El siguiente script validará,
   completará metadatos confiables y generará el JSON oficial.
"""
            (carpeta / "INSTRUCCIONES_MANUALES.txt").write_text(guia, encoding="utf-8")

            print(f"  ✓ GST: {gst}")
            print(f"  ✓ Prompt: {len(prompt):,} caracteres")
            print(f"  ✓ Zona: {contexto['destino']['zona']['rango']}")
            print(f"  ✓ Origen no vacío: {len(contexto['origen']['celdas_no_vacias'])}")
            print(f"  ✓ Dependencias: {len(contexto['dependencias_formulas'])}")
            ok += 1

        except Exception as exc:
            msg = f"{type(exc).__name__}: {exc}"
            print(f"  ✗ {msg}")
            errores.append((carpeta.name, msg))

    print("\n" + "=" * 78)
    print("RESUMEN")
    print("=" * 78)
    print(f"Casos preparados: {ok}")
    print(f"Casos con error : {len(errores)}")
    print(f"Carpeta         : {CASOS_DIR}")
    if errores:
        print("\nERRORES:")
        for caso, msg in errores:
            print(f"- {caso}: {msg}")
    print("=" * 78)


if __name__ == "__main__":
    main()
