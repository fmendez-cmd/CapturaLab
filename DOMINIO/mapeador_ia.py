# DOMINIO/mapeador_ia.py
#
# Mapeador estructural IA v2
# --------------------------
# La IA NO llena Excel.
# La IA compara:
#   1) la estructura completa del ensayo ERP,
#   2) SOLO la hoja cuyo nombre contiene "ERP" de la plantilla,
#   3) las referencias de las fórmulas del reporte hacia la zona de captura.
#
# El programa determina primero una ZONA DE CAPTURA segura.
# Gemini únicamente puede proponer destinos dentro de esa zona.
#
# Requisitos:
#   pip install -U google-genai pydantic pywin32
#
# Variable:
#   GEMINI_API_KEY

from __future__ import annotations

import hashlib
import json
import os
import re
import time
import random
import threading
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field
from DOMINIO.control_cancelacion import verificar_cancelacion, esperar_cancelable


class OperacionCelda(BaseModel):
    origen: str = Field(description="Celda A1 no vacía/top-left del ensayo ERP")
    destino: str = Field(description="Celda A1 top-left dentro de la zona segura ERP")
    descripcion: str = ""


class CeldaIgnorada(BaseModel):
    origen: str
    razon: str


class RespuestaMapaIA(BaseModel):
    confianza: float = Field(ge=0.0, le=1.0)
    estrategia: str
    operaciones: list[OperacionCelda]
    ignorados: list[CeldaIgnorada] = []
    advertencias: list[str] = []


class MapeadorIA:
    VERSION_MAPA = 22
    MODELO = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")

    # ============================================================
    # COM
    # ============================================================
    @staticmethod
    def _importar_com():
        if os.name != "nt":
            raise RuntimeError(
                "El mapeador requiere Windows + Microsoft Excel."
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
    def _abrir_excel():
        pythoncom, win32 = MapeadorIA._importar_com()
        pythoncom.CoInitialize()

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        excel.ScreenUpdating = False

        return pythoncom, excel

    # ============================================================
    # HELPERS
    # ============================================================
    @staticmethod
    def _normalizar_nombre_hoja(nombre: str) -> str:
        """
        Normaliza el nombre únicamente para decidir si una hoja es
        de validación. No altera el nombre real de Excel.
        """
        import unicodedata

        texto = unicodedata.normalize("NFKD", str(nombre))
        texto = "".join(
            ch for ch in texto
            if not unicodedata.combining(ch)
        )
        return " ".join(texto.upper().split())

    @staticmethod
    def _buscar_hoja_erp(libro):
        hojas_erp = []
        hojas_validacion = []

        for i in range(1, libro.Worksheets.Count + 1):
            hoja = libro.Worksheets(i)
            nombre_real = str(hoja.Name)
            nombre = MapeadorIA._normalizar_nombre_hoja(nombre_real)

            if "ERP" not in nombre:
                continue

            # Las hojas de VALIDACION ERP son auxiliares y nunca son
            # el área de captura que debe utilizar el mapeador.
            if "VALIDACION" in nombre:
                hojas_validacion.append(nombre_real)
                continue

            hojas_erp.append(hoja)

        if not hojas_erp:
            detalle = (
                f" Se ignoraron hojas auxiliares de validación: "
                f"{hojas_validacion}."
                if hojas_validacion
                else ""
            )
            raise ValueError(
                "La plantilla no contiene una hoja ERP de captura válida."
                + detalle
            )

        if len(hojas_erp) > 1:
            raise ValueError(
                "La plantilla contiene más de una hoja ERP de captura "
                "posible. No es seguro elegir automáticamente. "
                f"Coincidencias: {[str(h.Name) for h in hojas_erp]}"
            )

        if hojas_validacion:
            print(
                "ℹ️ Hojas ERP auxiliares de validación ignoradas: "
                f"{hojas_validacion}"
            )

        print(
            f"📄 Hoja ERP de captura seleccionada: "
            f"'{str(hojas_erp[0].Name)}'"
        )

        return hojas_erp[0]

    @staticmethod
    def _to_matrix(valor: Any) -> list[list[Any]]:
        if isinstance(valor, tuple):
            if valor and isinstance(valor[0], tuple):
                return [list(fila) for fila in valor]
            return [list(valor)]

        return [[valor]]

    @staticmethod
    def _texto(valor: Any) -> str:
        if valor is None:
            return ""

        return " ".join(
            str(valor)
            .replace("\r", " ")
            .replace("\n", " ")
            .split()
        ).strip()

    @staticmethod
    def _col_letras(numero: int) -> str:
        resultado = ""

        while numero:
            numero, resto = divmod(numero - 1, 26)
            resultado = chr(65 + resto) + resultado

        return resultado

    @staticmethod
    def _direccion(fila: int, columna: int) -> str:
        return f"{MapeadorIA._col_letras(columna)}{fila}"

    @staticmethod
    def _coord_desde_a1(celda: str):
        m = re.fullmatch(
            r"\$?([A-Z]{1,3})\$?(\d+)",
            str(celda).upper().strip(),
        )

        if not m:
            return None

        letras = m.group(1)
        fila = int(m.group(2))

        columna = 0
        for ch in letras:
            columna = columna * 26 + (ord(ch) - 64)

        return fila, columna

    @staticmethod
    def _es_no_vacio(valor) -> bool:
        return valor is not None and str(valor) != ""

    # ============================================================
    # FIRMA DE ESTRUCTURA DEL ERP
    # ============================================================
    @staticmethod
    def firma_estructura_origen(ruta_erp: str) -> str:
        """
        La firma NO depende de cliente, fechas ni resultados.
        Se basa en:
          - dimensiones del UsedRange,
          - celda inicial,
          - rangos combinados.

        La compatibilidad real se confirma después mediante anclas.
        """

        ruta_erp = str(Path(ruta_erp).resolve())

        pythoncom = None
        excel = None
        libro = None

        try:
            pythoncom, excel = MapeadorIA._abrir_excel()

            libro = excel.Workbooks.Open(
                ruta_erp,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )

            hoja = libro.Worksheets(1)
            used = hoja.UsedRange

            fila0 = int(used.Row)
            col0 = int(used.Column)
            filas = int(used.Rows.Count)
            columnas = int(used.Columns.Count)

            merges = set()

            for r in range(fila0, fila0 + filas):
                for c in range(col0, col0 + columnas):
                    celda = hoja.Cells(r, c)

                    try:
                        if bool(celda.MergeCells):
                            merges.add(str(celda.MergeArea.Address))
                    except Exception:
                        pass

            payload = {
                "fila_inicio": fila0,
                "columna_inicio": col0,
                "filas": filas,
                "columnas": columnas,
                "merges": sorted(merges),
            }

            raw = json.dumps(
                payload,
                sort_keys=True,
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode("utf-8")

            return hashlib.sha256(raw).hexdigest()

        finally:
            if libro is not None:
                try:
                    libro.Close(SaveChanges=False)
                except Exception:
                    pass

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # ============================================================
    # SOURCE CELLS / STRUCTURE
    # ============================================================
    @staticmethod
    def _source_anchor_cells(hoja) -> list[dict]:
        """
        Devuelve TODAS las celdas no vacías del origen.
        Si una celda pertenece a un merge, solamente devuelve el top-left.
        Incluye etiquetas y valores: NO hace interpretación semántica.
        """

        used = hoja.UsedRange

        fila0 = int(used.Row)
        col0 = int(used.Column)
        filas = int(used.Rows.Count)
        columnas = int(used.Columns.Count)

        vistos_merge = set()
        salida = []

        for r in range(fila0, fila0 + filas):
            for c in range(col0, col0 + columnas):
                celda = hoja.Cells(r, c)

                try:
                    if bool(celda.MergeCells):
                        area = celda.MergeArea
                        direccion_merge = str(area.Address)

                        if direccion_merge in vistos_merge:
                            continue

                        vistos_merge.add(direccion_merge)
                        celda_real = area.Cells(1, 1)
                        valor = celda_real.Value2
                        direccion = str(celda_real.Address).replace("$", "")
                        merge = direccion_merge.replace("$", "")
                    else:
                        valor = celda.Value2
                        direccion = str(celda.Address).replace("$", "")
                        merge = None

                except Exception:
                    valor = celda.Value2
                    direccion = str(celda.Address).replace("$", "")
                    merge = None

                if not MapeadorIA._es_no_vacio(valor):
                    continue

                salida.append({
                    "celda": direccion,
                    "valor": valor,
                    "texto": MapeadorIA._texto(valor),
                    "merge": merge,
                    "fila": r,
                    "columna": c,
                })

        return salida

    @staticmethod
    def _merges_en_rango(
        hoja,
        fila_inicio: int,
        fila_fin: int,
        col_inicio: int,
        col_fin: int,
    ) -> list[str]:
        merges = set()

        for r in range(fila_inicio, fila_fin + 1):
            for c in range(col_inicio, col_fin + 1):
                celda = hoja.Cells(r, c)

                try:
                    if bool(celda.MergeCells):
                        merges.add(
                            str(celda.MergeArea.Address).replace("$", "")
                        )
                except Exception:
                    pass

        return sorted(merges)

    # ============================================================
    # FORMULAS / REFERENCIAS
    # ============================================================
    @staticmethod
    def _ultima_fila_formula(hoja) -> int:
        used = hoja.UsedRange
        fila0 = int(used.Row)
        col0 = int(used.Column)
        filas = int(used.Rows.Count)
        columnas = int(used.Columns.Count)

        ultima = 0

        for r in range(fila0, fila0 + filas):
            for c in range(col0, col0 + columnas):
                celda = hoja.Cells(r, c)

                try:
                    if bool(celda.HasFormula):
                        ultima = max(ultima, r)
                except Exception:
                    pass

        return ultima

    @staticmethod
    def _refs_formula_local(formula: str) -> list[str]:
        """
        Extrae referencias A1 sencillas/rangos del texto de una fórmula.
        Es deliberadamente conservador.
        """

        if not formula:
            return []

        formula = str(formula).upper()

        # Ignorar referencias a otras hojas: se capturan igualmente las A1,
        # pero luego solo consideramos filas por debajo de la última fórmula,
        # que es la zona de staging del propio ERP.
        refs = re.findall(
            r"(?<![A-Z0-9_])\$?([A-Z]{1,3})\$?(\d+)",
            formula,
        )

        salida = []
        for col, fila in refs:
            salida.append(f"{col}{int(fila)}")

        return salida

    @staticmethod
    def _contexto_vecino_formula(hoja, fila: int, columna: int) -> list[dict]:
        """
        Textos constantes cerca de una fórmula del reporte.
        Sirven a Gemini para entender qué dato consume esa fórmula.
        """

        salida = []

        for r in range(max(1, fila - 1), fila + 2):
            for c in range(max(1, columna - 3), columna + 4):
                celda = hoja.Cells(r, c)

                try:
                    if bool(celda.HasFormula):
                        continue
                except Exception:
                    pass

                try:
                    valor = celda.Value2
                except Exception:
                    valor = None

                texto = MapeadorIA._texto(valor)

                if not texto:
                    continue

                if len(texto) > 120:
                    texto = texto[:120]

                salida.append({
                    "celda": str(celda.Address).replace("$", ""),
                    "texto": texto,
                })

        return salida

    @staticmethod
    def _analizar_dependencias_y_zona(hoja) -> tuple[list[dict], dict]:
        """
        Determina la zona de captura usando referencias REALES de las fórmulas.

        Regla:
        - localizamos la última fila que contiene fórmulas;
        - tomamos referencias A1 cuya fila esté por debajo;
        - la zona segura se expande 5 filas arriba/abajo y 3 columnas
          izquierda/derecha para incluir título, blancos y estructura.
        """

        ultima_formula = MapeadorIA._ultima_fila_formula(hoja)

        used = hoja.UsedRange
        fila0 = int(used.Row)
        col0 = int(used.Column)
        filas = int(used.Rows.Count)
        columnas = int(used.Columns.Count)

        dependencias = []
        referencias_inferiores = []

        for r in range(fila0, fila0 + filas):
            for c in range(col0, col0 + columnas):
                celda = hoja.Cells(r, c)

                try:
                    tiene_formula = bool(celda.HasFormula)
                except Exception:
                    tiene_formula = False

                if not tiene_formula:
                    continue

                try:
                    formula = str(celda.Formula)
                except Exception:
                    continue

                refs = MapeadorIA._refs_formula_local(formula)
                refs_inferiores = []

                for ref in refs:
                    coord = MapeadorIA._coord_desde_a1(ref)

                    if not coord:
                        continue

                    fila_ref, col_ref = coord

                    if fila_ref > ultima_formula:
                        referencias_inferiores.append(
                            (fila_ref, col_ref, ref)
                        )
                        refs_inferiores.append(ref)

                if refs_inferiores:
                    dependencias.append({
                        "formula_en": str(celda.Address).replace("$", ""),
                        "formula": formula,
                        "referencias_captura": refs_inferiores,
                        "contexto_reporte": MapeadorIA._contexto_vecino_formula(
                            hoja,
                            r,
                            c,
                        ),
                    })

        if not referencias_inferiores:
            raise ValueError(
                "No fue posible detectar una zona de captura mediante "
                "referencias de fórmulas hacia filas inferiores. "
                "No se permitirá que la IA invente una zona."
            )

        filas_ref = [x[0] for x in referencias_inferiores]
        cols_ref = [x[1] for x in referencias_inferiores]

        fila_min_ref = min(filas_ref)
        fila_max_ref = max(filas_ref)
        col_min_ref = min(cols_ref)
        col_max_ref = max(cols_ref)

        # Padding deliberado para abarcar encabezado, blancos y firmas.
        fila_inicio = max(
            ultima_formula + 1,
            fila_min_ref - 5,
        )

        fila_fin = max(
            fila_max_ref + 5,
            fila_inicio,
        )

        # Usamos las columnas del UsedRange como límite máximo.
        ultima_col_used = col0 + columnas - 1

        col_inicio = max(
            col0,
            col_min_ref - 3,
        )

        col_fin = min(
            max(ultima_col_used, col_max_ref),
            col_max_ref + 6,
        )

        zona = {
            "ultima_fila_formula": ultima_formula,
            "primera_fila_referenciada": fila_min_ref,
            "ultima_fila_referenciada": fila_max_ref,
            "fila_inicio": fila_inicio,
            "fila_fin": fila_fin,
            "columna_inicio": col_inicio,
            "columna_fin": col_fin,
            "rango": (
                f"{MapeadorIA._direccion(fila_inicio, col_inicio)}:"
                f"{MapeadorIA._direccion(fila_fin, col_fin)}"
            ),
        }

        return dependencias, zona

    # ============================================================
    # SERIALIZAR ZONA DESTINO
    # ============================================================
    @staticmethod
    def _serializar_zona_destino(hoja, zona: dict) -> dict:
        filas_json = []

        for r in range(
            zona["fila_inicio"],
            zona["fila_fin"] + 1,
        ):
            celdas = []

            for c in range(
                zona["columna_inicio"],
                zona["columna_fin"] + 1,
            ):
                celda = hoja.Cells(r, c)

                try:
                    valor = celda.Value2
                except Exception:
                    valor = None

                try:
                    formula = (
                        str(celda.Formula)
                        if bool(celda.HasFormula)
                        else None
                    )
                except Exception:
                    formula = None

                try:
                    merge = (
                        str(celda.MergeArea.Address).replace("$", "")
                        if bool(celda.MergeCells)
                        else None
                    )
                except Exception:
                    merge = None

                try:
                    alineacion = int(celda.HorizontalAlignment)
                except Exception:
                    alineacion = None

                try:
                    number_format = str(celda.NumberFormat)
                except Exception:
                    number_format = None

                bordes = {}
                # Excel: 7=left, 8=top, 9=bottom, 10=right
                for nombre, idx in (
                    ("left", 7),
                    ("top", 8),
                    ("bottom", 9),
                    ("right", 10),
                ):
                    try:
                        bordes[nombre] = int(celda.Borders(idx).LineStyle)
                    except Exception:
                        bordes[nombre] = None

                celdas.append({
                    "celda": MapeadorIA._direccion(r, c),
                    "valor_actual": valor,
                    "formula": formula,
                    "merge": merge,
                    "alineacion": alineacion,
                    "formato_numero": number_format,
                    "bordes": bordes,
                })

            try:
                alto = hoja.Rows(r).RowHeight
            except Exception:
                alto = None

            filas_json.append({
                "fila": r,
                "alto": alto,
                "celdas": celdas,
            })

        columnas_json = []

        for c in range(
            zona["columna_inicio"],
            zona["columna_fin"] + 1,
        ):
            try:
                ancho = hoja.Columns(c).ColumnWidth
            except Exception:
                ancho = None

            columnas_json.append({
                "columna": MapeadorIA._col_letras(c),
                "numero": c,
                "ancho": ancho,
            })

        return {
            "hoja": str(hoja.Name),
            "zona": zona,
            "filas": filas_json,
            "columnas": columnas_json,
            "merges": MapeadorIA._merges_en_rango(
                hoja,
                zona["fila_inicio"],
                zona["fila_fin"],
                zona["columna_inicio"],
                zona["columna_fin"],
            ),
        }

    # ============================================================
    # CONTEXTO COMPLETO
    # ============================================================
    @staticmethod
    def construir_contexto(
        codigo_gst: str,
        ruta_erp: str,
        ruta_plantilla: str,
    ) -> dict:
        pythoncom = None
        excel = None
        libro_origen = None
        libro_destino = None

        try:
            pythoncom, excel = MapeadorIA._abrir_excel()

            libro_origen = excel.Workbooks.Open(
                str(Path(ruta_erp).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )

            libro_destino = excel.Workbooks.Open(
                str(Path(ruta_plantilla).resolve()),
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )

            hoja_origen = libro_origen.Worksheets(1)
            hoja_destino = MapeadorIA._buscar_hoja_erp(
                libro_destino
            )

            used_origen = hoja_origen.UsedRange

            dependencias, zona = (
                MapeadorIA._analizar_dependencias_y_zona(
                    hoja_destino
                )
            )

            origen = {
                "hoja": str(hoja_origen.Name),
                "rango_usado": str(used_origen.Address).replace("$", ""),
                "fila_inicio": int(used_origen.Row),
                "columna_inicio": int(used_origen.Column),
                "filas": int(used_origen.Rows.Count),
                "columnas": int(used_origen.Columns.Count),
                "celdas_no_vacias": MapeadorIA._source_anchor_cells(
                    hoja_origen
                ),
                "merges": MapeadorIA._merges_en_rango(
                    hoja_origen,
                    int(used_origen.Row),
                    int(used_origen.Row) + int(used_origen.Rows.Count) - 1,
                    int(used_origen.Column),
                    int(used_origen.Column) + int(used_origen.Columns.Count) - 1,
                ),
            }

            return {
                "version_mapa": MapeadorIA.VERSION_MAPA,
                "gst": codigo_gst,
                "archivo_origen": Path(ruta_erp).name,
                "archivo_plantilla": Path(ruta_plantilla).name,
                "regla_principal": (
                    "El objetivo es reconstruir dentro de la zona de captura "
                    "de la hoja ERP la estructura completa del ensayo ERP, "
                    "incluyendo etiquetas, valores y guiones. "
                    "No mapear solamente valores semánticos."
                ),
                "origen": origen,
                "destino": MapeadorIA._serializar_zona_destino(
                    hoja_destino,
                    zona,
                ),
                "dependencias_formulas": dependencias,
            }

        finally:
            if libro_origen is not None:
                try:
                    libro_origen.Close(SaveChanges=False)
                except Exception:
                    pass

            if libro_destino is not None:
                try:
                    libro_destino.Close(SaveChanges=False)
                except Exception:
                    pass

            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass

            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # ============================================================
    # PROMPT
    # ============================================================
    @staticmethod
    def _system_instruction() -> str:
        return """
Eres un analizador de ESTRUCTURA FÍSICA de hojas Excel.

Tu única tarea es devolver un mapa CELDA -> CELDA.

ORIGEN:
El ensayo tal como lo exporta el ERP.

DESTINO:
Exclusivamente la zona segura de captura de la hoja que contiene ERP.
Las referencias de fórmulas muestran qué celdas de esa zona consume el reporte.

REGLAS OBLIGATORIAS:
1. NO selecciones solamente campos semánticamente importantes.
2. Mapea TODAS las celdas no vacías/top-left del origen.
3. Incluye etiquetas, valores, fechas, resultados, nombres, firmas y guiones.
4. "-" es un dato real y jamás debe ignorarse.
5. No compactes filas ni subas valores.
6. Conserva el orden físico del documento.
7. Respeta merges; usa únicamente su top-left.
8. Cada celda no vacía del origen debe aparecer exactamente una vez en operaciones
   o, únicamente si es estructuralmente imposible, en ignorados.
9. Todos los destinos deben estar dentro de la zona proporcionada.
10. No escribas sobre fórmulas.
11. Usa geometría, merges, bordes y referencias de fórmulas como evidencia.
12. Las referencias de fórmulas tienen prioridad alta.
13. No inventes coordenadas.
14. NO GENERES ANCLAS. Python valida la geometría de forma determinista.
15. Confianza >= 0.90 solo si la transformación completa es consistente.
""".strip()

    @staticmethod
    def _es_cuota_diaria_agotada(error: Exception) -> bool:
        texto = f"{type(error).__name__}: {error}".lower()
        marcas = (
            "generaterequestsperdayperprojectpermodel-freetier",
            "generate_content_free_tier_requests",
        )
        return any(m in texto for m in marcas)

    @staticmethod
    def _es_error_temporal(error: Exception) -> bool:
        if MapeadorIA._es_cuota_diaria_agotada(error):
            return False
        texto = f"{type(error).__name__}: {error}".lower()
        marcas = (
            "503", "unavailable", "high demand", "429", "resource_exhausted",
            "too many requests", "500", "502", "504", "timeout", "timed out",
            "connection", "temporarily", "service unavailable", "internal server error",
        )
        return any(m in texto for m in marcas)

    @staticmethod
    def generar_mapa(
        codigo_gst: str, ruta_erp: str, ruta_plantilla: str, cancel_event=None
    ) -> dict:
        api_key = os.getenv("GEMINI_API_KEY", "").strip()
        if not api_key:
            raise RuntimeError("No existe GEMINI_API_KEY en las variables de entorno.")

        try:
            from google import genai
            from google.genai import types
        except ImportError as exc:
            raise RuntimeError("Instala: pip install -U google-genai pydantic") from exc

        verificar_cancelacion()

        _t_total_gemini = time.perf_counter()

        _t = time.perf_counter()
        contexto = MapeadorIA.construir_contexto(codigo_gst, ruta_erp, ruta_plantilla)
        _dt_contexto = time.perf_counter() - _t

        verificar_cancelacion()

        _t = time.perf_counter()
        prompt = (
            "Genera el mapa estructural completo CELDA->CELDA.\n\n"
            + json.dumps(contexto, ensure_ascii=False, separators=(",", ":"), default=str)
        )
        _dt_prompt = time.perf_counter() - _t

        print("\n" + "=" * 78)
        print("GENERANDO MAPA ESTRUCTURAL V2.2 CON GEMINI")
        print("=" * 78)
        print(f"GST: {codigo_gst}")
        print(f"Modelo: {MapeadorIA.MODELO}")
        print(f"Zona segura: {contexto['destino']['zona']['rango']}")
        print(f"Celdas no vacías origen: {len(contexto['origen']['celdas_no_vacias'])}")
        print(f"Dependencias de fórmulas: {len(contexto['dependencias_formulas'])}")
        print(f"Tamaño contexto: {len(prompt):,} caracteres")
        print(f"⏱️ Construir contexto Excel: {_dt_contexto:.3f} s")
        print(f"⏱️ Serializar prompt: {_dt_prompt:.3f} s")
        print("=" * 78)

        cliente = genai.Client(api_key=api_key)
        intento = 0
        espera = 5

        while True:
            verificar_cancelacion()
            intento += 1
            try:
                print(f"🤖 Gemini — intento {intento}...")
                _t_llamada = time.perf_counter()
                respuesta = cliente.models.generate_content(
                    model=MapeadorIA.MODELO,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        system_instruction=MapeadorIA._system_instruction(),
                        response_mime_type="application/json",
                        response_schema=RespuestaMapaIA,
                        temperature=0.0,
                        automatic_function_calling=types.AutomaticFunctionCallingConfig(disable=True),
                    ),
                )
                _dt_llamada = time.perf_counter() - _t_llamada
                print(f"✅ Gemini respondió correctamente en el intento {intento}.")
                print(f"⏱️ Llamada Gemini: {_dt_llamada:.3f} s")
                break
            except Exception as error:
                if MapeadorIA._es_cuota_diaria_agotada(error):
                    raise RuntimeError(
                        "Se agotó la cuota diaria de Gemini para "
                        f"{MapeadorIA.MODELO}. No se harán reintentos automáticos."
                    ) from error
                if not MapeadorIA._es_error_temporal(error):
                    raise
                jitter = random.uniform(0, 2.0)
                pausa = min(5, espera) + jitter
                print(f"⚠️ Error temporal de Gemini: {type(error).__name__}: {error}")
                print(f"⏳ Reintentando automáticamente en {pausa:.0f} segundos...")
                if cancel_event is not None:
                    if cancel_event.wait(timeout=pausa):
                        raise InterruptedError("Procesamiento cancelado por el usuario.")
                else:
                    time.sleep(pausa)
                espera = min(5, espera * 2)

        _t = time.perf_counter()
        if getattr(respuesta, "parsed", None) is not None:
            parsed = respuesta.parsed
            data = parsed.model_dump() if hasattr(parsed, "model_dump") else dict(parsed)
        else:
            texto = (getattr(respuesta, "text", "") or "").strip()
            if not texto:
                raise RuntimeError("Gemini no devolvió ningún mapa.")
            data = json.loads(texto)
        _dt_parseo = time.perf_counter() - _t
        print(f"⏱️ Parsear respuesta Gemini: {_dt_parseo:.3f} s")

        data["version_mapa"] = MapeadorIA.VERSION_MAPA
        data["gst"] = codigo_gst
        data["archivo_plantilla"] = Path(ruta_plantilla).name
        data["hoja_destino"] = contexto["destino"]["hoja"]
        data["zona_captura"] = contexto["destino"]["zona"]
        _t = time.perf_counter()
        data["firma_estructura_origen"] = MapeadorIA.firma_estructura_origen(ruta_erp)
        _dt_firma_final = time.perf_counter() - _t

        print(f"⏱️ Firma estructura ERP final: {_dt_firma_final:.3f} s")
        print(
            f"⏱️ TOTAL generar_mapa (incluye Gemini): "
            f"{time.perf_counter() - _t_total_gemini:.3f} s"
        )
        return data
