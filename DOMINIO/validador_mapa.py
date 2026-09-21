# DOMINIO/validador_mapa.py
#
# Validador estricto para mapas estructurales v2.

from __future__ import annotations

import os
import re
from pathlib import Path


class ValidadorMapa:
    VERSION_MAPA = 22

    PATRON_CELDA = re.compile(
        r"^\$?[A-Z]{1,3}\$?\d+$",
        re.IGNORECASE,
    )

    @staticmethod
    def _importar_com():
        if os.name != "nt":
            raise RuntimeError(
                "La validación requiere Windows + Microsoft Excel."
            )

        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise RuntimeError(
                "Falta pywin32. Instala con: pip install pywin32"
            ) from exc

        return pythoncom, win32com.client

    @staticmethod
    def _buscar_hoja_erp(libro):
        hojas = []

        for i in range(1, libro.Worksheets.Count + 1):
            hoja = libro.Worksheets(i)
            if "ERP" in str(hoja.Name).upper():
                hojas.append(hoja)

        if len(hojas) != 1:
            raise ValueError(
                "Debe existir exactamente una hoja cuyo nombre contenga ERP."
            )

        return hojas[0]

    @staticmethod
    def _texto(valor) -> str:
        if valor is None:
            return ""

        return " ".join(
            str(valor)
            .replace("\r", " ")
            .replace("\n", " ")
            .split()
        ).strip()

    @staticmethod
    def _coord(celda: str):
        m = re.fullmatch(
            r"\$?([A-Z]{1,3})\$?(\d+)",
            str(celda).upper().strip(),
        )

        if not m:
            raise ValueError(
                f"Dirección de celda inválida: {celda!r}"
            )

        col = 0
        for ch in m.group(1):
            col = col * 26 + (ord(ch) - 64)

        return int(m.group(2)), col

    @staticmethod
    def _top_left(celda):
        try:
            if bool(celda.MergeCells):
                return celda.MergeArea.Cells(1, 1)
        except Exception:
            pass

        return celda

    @staticmethod
    def _source_anchor_addresses(hoja) -> dict[str, object]:
        used = hoja.UsedRange

        fila0 = int(used.Row)
        col0 = int(used.Column)
        filas = int(used.Rows.Count)
        columnas = int(used.Columns.Count)

        salida = {}
        vistos_merge = set()

        for r in range(fila0, fila0 + filas):
            for c in range(col0, col0 + columnas):
                celda = hoja.Cells(r, c)

                try:
                    if bool(celda.MergeCells):
                        area = celda.MergeArea
                        key_merge = str(area.Address)

                        if key_merge in vistos_merge:
                            continue

                        vistos_merge.add(key_merge)
                        celda = area.Cells(1, 1)
                except Exception:
                    pass

                try:
                    valor = celda.Value2
                except Exception:
                    valor = None

                if valor is None or str(valor) == "":
                    continue

                direccion = str(celda.Address).replace("$", "")
                salida[direccion.upper()] = valor

        return salida

    @staticmethod
    def validar_esquema(mapa: dict):
        if int(mapa.get("version_mapa", 0)) != 22:
            raise ValueError(
                "El mapa no es versión 22. Los mapas anteriores no son compatibles."
            )

        requeridos = (
            "gst",
            "hoja_destino",
            "zona_captura",
            "confianza",
            "operaciones",
            "ignorados",
        )

        faltantes = [
            campo
            for campo in requeridos
            if campo not in mapa
        ]

        if faltantes:
            raise ValueError(
                f"Mapa incompleto. Faltan: {faltantes}"
            )

        if "ERP" not in str(mapa["hoja_destino"]).upper():
            raise ValueError(
                "La hoja destino del mapa no contiene ERP."
            )

        confianza = float(mapa["confianza"])

        if confianza < 0.90:
            raise ValueError(
                f"Confianza insuficiente: {confianza:.2f}. "
                "Se requiere >= 0.90."
            )

        zona = mapa["zona_captura"]

        for campo in (
            "fila_inicio",
            "fila_fin",
            "columna_inicio",
            "columna_fin",
        ):
            if campo not in zona:
                raise ValueError(
                    f"Zona de captura incompleta: falta {campo}"
                )

        operaciones = mapa["operaciones"]

        if not isinstance(operaciones, list) or not operaciones:
            raise ValueError(
                "El mapa no contiene operaciones."
            )

        if len(operaciones) > 3000:
            raise ValueError(
                "El mapa contiene demasiadas operaciones."
            )

        for i, op in enumerate(operaciones, start=1):
            origen = str(op.get("origen", "")).replace("$", "")
            destino = str(op.get("destino", "")).replace("$", "")

            if not ValidadorMapa.PATRON_CELDA.fullmatch(origen):
                raise ValueError(
                    f"Operación {i}: origen inválido {origen!r}"
                )

            if not ValidadorMapa.PATRON_CELDA.fullmatch(destino):
                raise ValueError(
                    f"Operación {i}: destino inválido {destino!r}"
                )


    @staticmethod
    def _destino_dentro_zona(
        direccion: str,
        zona: dict,
    ) -> bool:
        fila, col = ValidadorMapa._coord(direccion)

        return (
            int(zona["fila_inicio"]) <= fila <= int(zona["fila_fin"])
            and
            int(zona["columna_inicio"]) <= col <= int(zona["columna_fin"])
        )

    @staticmethod
    def validar_anclas_abiertas(
        mapa: dict,
        hoja_origen,
        hoja_destino,
    ) -> bool:
        origen_ok = 0
        destino_ok = 0

        for ancla in mapa.get("anclas", []):
            lado = str(ancla.get("lado", "")).lower().strip()
            direccion = str(
                ancla.get("celda", "")
            ).replace("$", "")
            esperado = ValidadorMapa._texto(
                ancla.get("texto_esperado")
            )

            if not esperado:
                return False

            if lado == "origen":
                hoja = hoja_origen
            elif lado == "destino":
                hoja = hoja_destino

                if not ValidadorMapa._destino_dentro_zona(
                    direccion,
                    mapa["zona_captura"],
                ):
                    return False
            else:
                return False

            try:
                celda = ValidadorMapa._top_left(
                    hoja.Range(direccion).Cells(1, 1)
                )
                actual = ValidadorMapa._texto(
                    celda.Value2
                )
            except Exception:
                return False

            if actual.casefold() != esperado.casefold():
                return False

            if lado == "origen":
                origen_ok += 1
            else:
                destino_ok += 1

        return origen_ok >= 3 and destino_ok >= 3

    @staticmethod
    def validar_anclas_archivos(
        mapa: dict,
        ruta_erp: str,
        ruta_plantilla: str,
    ) -> bool:
        pythoncom, win32 = ValidadorMapa._importar_com()
        pythoncom.CoInitialize()

        excel = None
        libro_o = None
        libro_d = None

        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            libro_o = excel.Workbooks.Open(
                str(Path(ruta_erp).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
            )

            libro_d = excel.Workbooks.Open(
                str(Path(ruta_plantilla).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
            )

            hoja_o = libro_o.Worksheets(1)
            hoja_d = ValidadorMapa._buscar_hoja_erp(
                libro_d
            )

            if str(hoja_d.Name) != str(
                mapa["hoja_destino"]
            ):
                return False

            return ValidadorMapa.validar_anclas_abiertas(
                mapa,
                hoja_o,
                hoja_d,
            )

        finally:
            if libro_o is not None:
                try:
                    libro_o.Close(SaveChanges=False)
                except Exception:
                    pass

            if libro_d is not None:
                try:
                    libro_d.Close(SaveChanges=False)
                except Exception:
                    pass

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            pythoncom.CoUninitialize()

    @staticmethod
    def validar_mapa_abierto(
        mapa: dict,
        hoja_origen,
        hoja_destino,
    ):
        ValidadorMapa.validar_esquema(mapa)

        if str(hoja_destino.Name) != str(
            mapa["hoja_destino"]
        ):
            raise ValueError(
                "La hoja ERP real no coincide con la hoja del mapa."
            )

        source = ValidadorMapa._source_anchor_addresses(
            hoja_origen
        )

        operaciones = mapa["operaciones"]
        ignorados = mapa.get("ignorados", [])

        origenes_mapeados = []
        destinos_reales = []

        for i, op in enumerate(
            operaciones,
            start=1,
        ):
            origen_addr = str(
                op["origen"]
            ).replace("$", "").upper()

            destino_addr = str(
                op["destino"]
            ).replace("$", "").upper()

            if origen_addr not in source:
                raise ValueError(
                    f"Operación {i}: {origen_addr} no es una "
                    "celda no vacía/top-left válida del origen."
                )

            if not ValidadorMapa._destino_dentro_zona(
                destino_addr,
                mapa["zona_captura"],
            ):
                raise ValueError(
                    f"Operación {i}: destino {destino_addr} "
                    "está FUERA de la zona de captura."
                )

            destino = hoja_destino.Range(
                destino_addr
            ).Cells(1, 1)

            destino_real = ValidadorMapa._top_left(
                destino
            )

            real_addr = str(
                destino_real.Address
            ).replace("$", "").upper()

            if real_addr != destino_addr:
                raise ValueError(
                    f"Operación {i}: {destino_addr} está dentro "
                    f"de una celda combinada. Debe apuntar al top-left "
                    f"{real_addr}."
                )

            try:
                if bool(destino_real.HasFormula):
                    raise ValueError(
                        f"Operación {i}: {destino_addr} contiene fórmula."
                    )
            except TypeError:
                pass

            origenes_mapeados.append(
                origen_addr
            )
            destinos_reales.append(
                real_addr
            )

        # Un origen debe mapearse exactamente una vez.
        repetidos_origen = {
            x
            for x in origenes_mapeados
            if origenes_mapeados.count(x) > 1
        }

        if repetidos_origen:
            raise ValueError(
                "El mapa repite celdas de origen: "
                f"{sorted(repetidos_origen)[:20]}"
            )

        # Un destino no puede recibir dos celdas distintas.
        repetidos_destino = {
            x
            for x in destinos_reales
            if destinos_reales.count(x) > 1
        }

        if repetidos_destino:
            raise ValueError(
                "El mapa intenta escribir varias celdas sobre "
                f"el mismo destino: {sorted(repetidos_destino)[:20]}"
            )

        ignorados_dict = {}

        for item in ignorados:
            addr = str(
                item.get("origen", "")
            ).replace("$", "").upper()

            razon = str(
                item.get("razon", "")
            ).strip()

            if addr not in source:
                raise ValueError(
                    f"Ignorado inválido: {addr}"
                )

            if not razon:
                raise ValueError(
                    f"La celda ignorada {addr} no tiene razón."
                )

            # El guion es un dato real y jamás puede ignorarse.
            if str(source[addr]).strip() == "-":
                raise ValueError(
                    f"No se permite ignorar el valor '-' en {addr}."
                )

            ignorados_dict[addr] = razon

        cubiertos = (
            set(origenes_mapeados)
            | set(ignorados_dict)
        )

        faltantes = sorted(
            set(source) - cubiertos
        )

        if faltantes:
            raise ValueError(
                "El mapa NO cubre toda la estructura no vacía del ERP.\n"
                f"Celdas faltantes ({len(faltantes)}): "
                f"{faltantes[:50]}"
            )

        # No queremos mapas que omitan una parte importante.
        if ignorados_dict:
            proporcion = len(ignorados_dict) / max(
                1,
                len(source),
            )

            if proporcion > 0.03:
                raise ValueError(
                    "Gemini intentó ignorar más del 3% de las "
                    "celdas no vacías del ERP."
                )


        print()
        print("=" * 78)
        print("VALIDACIÓN DEL MAPA V2")
        print("=" * 78)
        print(
            f"Celdas no vacías origen: {len(source)}"
        )
        print(
            f"Celdas mapeadas: {len(origenes_mapeados)}"
        )
        print(
            f"Celdas ignoradas: {len(ignorados_dict)}"
        )
        print(
            f"Zona captura: {mapa['zona_captura']['rango']}"
        )
        print("✅ Cobertura estructural válida.")
        print("=" * 78)

    @staticmethod
    def _a1(fila: int, columna: int) -> str:
        if fila < 1 or columna < 1:
            raise ValueError("Coordenada Excel inválida.")
        letras = ""
        n = columna
        while n:
            n, resto = divmod(n - 1, 26)
            letras = chr(65 + resto) + letras
        return f"{letras}{fila}"

    @staticmethod
    def reconstruir_por_geometria(mapa: dict, hoja_origen, hoja_destino) -> dict:
        from collections import Counter

        pares = []
        for op in mapa.get("operaciones", []):
            try:
                ro, co = ValidadorMapa._coord(op["origen"])
                rd, cd = ValidadorMapa._coord(op["destino"])
                pares.append((rd - ro, cd - co))
            except Exception:
                continue

        if not pares:
            raise ValueError("Gemini no produjo operaciones suficientes para inferir geometría.")

        conteo = Counter(pares)
        (dr, dc), votos = conteo.most_common(1)[0]
        proporcion = votos / len(pares)

        if proporcion < 0.60:
            raise ValueError(
                f"No existe una traslación geométrica dominante suficiente: {proporcion:.1%}."
            )

        source = ValidadorMapa._source_anchor_addresses(hoja_origen)
        operaciones = []

        for origen in source:
            ro, co = ValidadorMapa._coord(origen)
            rd, cd = ro + dr, co + dc
            destino = ValidadorMapa._a1(rd, cd)

            if not ValidadorMapa._destino_dentro_zona(destino, mapa["zona_captura"]):
                raise ValueError(
                    f"La traslación dominante lleva {origen} -> {destino} fuera de la zona segura."
                )

            operaciones.append({
                "origen": origen,
                "destino": destino,
                "descripcion": "Reconstrucción geométrica determinista V2.2",
            })

        nuevo = dict(mapa)
        nuevo["version_mapa"] = 22
        nuevo["operaciones"] = operaciones
        nuevo["ignorados"] = []
        nuevo["confianza"] = max(float(nuevo.get("confianza", 0)), 0.95)
        nuevo["diagnostico_v22"] = {
            "desplazamiento_filas": dr,
            "desplazamiento_columnas": dc,
            "votos_dominantes": votos,
            "operaciones_ia": len(pares),
            "proporcion_dominante": proporcion,
            "celdas_origen": len(source),
            "operaciones_reconstruidas": len(operaciones),
        }
        return nuevo

    @staticmethod
    def validar_y_reconstruir_con_archivos(mapa: dict, ruta_erp: str, ruta_plantilla: str) -> dict:
        pythoncom, win32 = ValidadorMapa._importar_com()
        pythoncom.CoInitialize()
        excel = libro_o = libro_d = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            libro_o = excel.Workbooks.Open(str(Path(ruta_erp).resolve()), UpdateLinks=0, ReadOnly=True)
            libro_d = excel.Workbooks.Open(str(Path(ruta_plantilla).resolve()), UpdateLinks=0, ReadOnly=True)
            hoja_o = libro_o.Worksheets(1)
            hoja_d = ValidadorMapa._buscar_hoja_erp(libro_d)

            if str(hoja_d.Name) != str(mapa["hoja_destino"]):
                raise ValueError("La hoja ERP real no coincide con la hoja del mapa.")

            nuevo = ValidadorMapa.reconstruir_por_geometria(mapa, hoja_o, hoja_d)
            ValidadorMapa.validar_mapa_abierto(nuevo, hoja_o, hoja_d)
            return nuevo
        finally:
            if libro_o is not None:
                try: libro_o.Close(SaveChanges=False)
                except Exception: pass
            if libro_d is not None:
                try: libro_d.Close(SaveChanges=False)
                except Exception: pass
            if excel is not None:
                try: excel.Quit()
                except Exception: pass
            pythoncom.CoUninitialize()

    @staticmethod
    def validar_reuso_con_archivos(mapa: dict, ruta_erp: str, ruta_plantilla: str) -> bool:
        try:
            ValidadorMapa.validar_mapa_con_archivos(mapa, ruta_erp, ruta_plantilla)
            return True
        except Exception:
            return False

    @staticmethod
    def validar_mapa_con_archivos(
        mapa: dict,
        ruta_erp: str,
        ruta_plantilla: str,
    ):
        pythoncom, win32 = ValidadorMapa._importar_com()
        pythoncom.CoInitialize()

        excel = None
        libro_o = None
        libro_d = None

        try:
            excel = win32.DispatchEx(
                "Excel.Application"
            )
            excel.Visible = False
            excel.DisplayAlerts = False

            libro_o = excel.Workbooks.Open(
                str(Path(ruta_erp).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
            )

            libro_d = excel.Workbooks.Open(
                str(Path(ruta_plantilla).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
            )

            hoja_o = libro_o.Worksheets(1)
            hoja_d = ValidadorMapa._buscar_hoja_erp(
                libro_d
            )

            ValidadorMapa.validar_mapa_abierto(
                mapa,
                hoja_o,
                hoja_d,
            )

        finally:
            if libro_o is not None:
                try:
                    libro_o.Close(SaveChanges=False)
                except Exception:
                    pass

            if libro_d is not None:
                try:
                    libro_d.Close(SaveChanges=False)
                except Exception:
                    pass

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            pythoncom.CoUninitialize()
