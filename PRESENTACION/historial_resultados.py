"""Historial persistente de lotes procesados, agrupado por OT."""

from __future__ import annotations

import os
import re
import sqlite3
import uuid
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


class HistorialResultados:
    def __init__(self, directorio: str | Path):
        self.directorio = Path(directorio)
        self.archivos = self.directorio / "archivos"
        self.archivos.mkdir(parents=True, exist_ok=True)
        self.base = self.directorio / "historial.sqlite3"
        with self._conexion() as db:
            db.execute("""CREATE TABLE IF NOT EXISTS resultados (
                id TEXT PRIMARY KEY,
                ot TEXT NOT NULL,
                fecha TEXT NOT NULL,
                exitosos INTEGER NOT NULL,
                incidencias INTEGER NOT NULL,
                zip_nombre TEXT
            )""")
            db.execute("CREATE INDEX IF NOT EXISTS idx_resultados_fecha ON resultados(fecha DESC)")

    def _conexion(self):
        return sqlite3.connect(self.base, timeout=15)

    @staticmethod
    def _ot(item: dict) -> str:
        ruta = str(item.get("ruta_excel") or "")
        nombre = str(item.get("archivo") or "")
        match = re.search(r"(?i)(?:^|[/\\])OT[-_ ]?(\d+)(?:[/\\]|$)", ruta)
        if not match:
            match = re.search(r"(?<!\d)(\d{3,})_\1-\d+", nombre)
        if not match:
            match = re.search(r"(?i)(?:^|[^A-Z0-9])OT[-_ ]?(\d+)", nombre)
        return f"OT-{match.group(1)}" if match else "OT sin identificar"

    def guardar(self, resultado: dict) -> list[dict]:
        """Crea un registro y un ZIP estable por OT. No almacena lotes cancelados."""
        if str(resultado.get("status", "")).lower() == "cancelado":
            return []
        grupos: dict[str, dict[str, list[dict]]] = defaultdict(lambda: {"exitosos": [], "errores": []})
        for item in resultado.get("archivos_exitosos", []) or []:
            grupos[self._ot(item)]["exitosos"].append(item)
        for item in resultado.get("archivos_con_error", []) or []:
            grupos[self._ot(item)]["errores"].append(item)
        guardados = []
        fecha = datetime.now(ZoneInfo("America/Mexico_City")).isoformat(timespec="seconds")
        for ot, grupo in grupos.items():
            identificador = uuid.uuid4().hex
            nombre_zip = f"{identificador}.zip" if grupo["exitosos"] else None
            exitosos_guardados = 0
            if nombre_zip:
                temporal = self.archivos / f"{identificador}.tmp"
                nombres_usados: set[str] = set()
                try:
                    with zipfile.ZipFile(temporal, "w", zipfile.ZIP_DEFLATED) as z:
                        for item in grupo["exitosos"]:
                            ruta = Path(str(item.get("ruta_excel") or ""))
                            if not ruta.is_file():
                                continue
                            nombre = ruta.name
                            if nombre in nombres_usados:
                                nombre = f"{exitosos_guardados + 1}_{nombre}"
                            nombres_usados.add(nombre)
                            z.write(ruta, arcname=nombre)
                            exitosos_guardados += 1
                    if exitosos_guardados:
                        os.replace(temporal, self.archivos / nombre_zip)
                    else:
                        temporal.unlink(missing_ok=True)
                        nombre_zip = None
                except Exception:
                    temporal.unlink(missing_ok=True)
                    raise
            registro = {
                "id": identificador, "ot": ot, "fecha": fecha,
                "exitosos": exitosos_guardados,
                "incidencias": len(grupo["errores"]), "zip_nombre": nombre_zip,
            }
            with self._conexion() as db:
                db.execute(
                    "INSERT INTO resultados (id, ot, fecha, exitosos, incidencias, zip_nombre) "
                    "VALUES (:id, :ot, :fecha, :exitosos, :incidencias, :zip_nombre)",
                    registro,
                )
            guardados.append(registro)
        return guardados

    def listar(self) -> list[dict]:
        with self._conexion() as db:
            db.row_factory = sqlite3.Row
            rows = db.execute(
                "SELECT id, ot, fecha, exitosos, incidencias, zip_nombre "
                "FROM resultados ORDER BY ot, fecha DESC, id DESC"
            ).fetchall()
        return [dict(row) for row in rows]

    def ruta_descarga(self, identificador: str) -> Path | None:
        with self._conexion() as db:
            row = db.execute(
                "SELECT zip_nombre FROM resultados WHERE id = ?", (identificador,)
            ).fetchone()
        if not row or not row[0]:
            return None
        ruta = self.archivos / row[0]
        return ruta if ruta.is_file() else None
