# DOMINIO/validador_reglas.py


class ValidadorReglas:
    @staticmethod
    def evaluar_registro(datos_ensaye: dict) -> str:
        # Reglas determinísticas ISO/IEC 17025
        if datos_ensaye["gst"] == "Desconocido":
            return "🔴 CRÍTICO: Formato no reconocido en catálogo"
        if datos_ensaye["formato"] == "Sin Asignar":
            return "🟡 REVISAR: Plantilla pendiente de configuración"
        return "🟢 CORRECTO: Listo para generación"
