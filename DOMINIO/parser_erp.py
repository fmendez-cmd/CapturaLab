# DOMINIO/parser_erp.py
import os
import re
from bs4 import BeautifulSoup


class ParserERP:
    @staticmethod
    def extraer_metadatos_archivo(filename: str, ruta_relativa: str = "") -> dict:
        """
        Extrae códigos e identificadores directamente de la nomenclatura del archivo.
        """
        gst_match = re.search(r"GST-(\d+)", filename, re.IGNORECASE)
        req_match = re.search(r"(\d{5})", filename)

        gst_code = f"GST-{gst_match.group(1)}" if gst_match else "Desconocido"
        req_code = req_match.group(1) if req_match else "Sin Requisición"

        return {
            "archivo": filename,
            "ruta": ruta_relativa,
            "req": req_code,
            "gst": gst_code,
            "ensaye": f"Ensaye {gst_code}",
            "formato": f"Plantilla-{gst_code}",
        }

    @staticmethod
    def parsear_html_secuencial(ruta_archivo: str) -> list:
        if not os.path.exists(ruta_archivo):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta_archivo}")

        with open(ruta_archivo, "r", encoding="latin-1", errors="ignore") as f:
            soup = BeautifulSoup(f.read(), "html.parser")

        valores_secuenciales = []
        filas = soup.find_all("tr")

        for fila in filas:
            celdas = fila.find_all(["td", "th"])
            for celda in celdas:
                texto = celda.get_text(strip=True)
                if texto:
                    valores_secuenciales.append(texto)

        return valores_secuenciales
