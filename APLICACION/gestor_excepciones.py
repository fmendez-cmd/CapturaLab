# APLICACION/gestor_excepciones.py
from DOMINIO.validador_reglas import ValidadorReglas


class GestorExcepciones:
    def __init__(self):
        self.errores = []
        self.correctos = []

    def clasificar_lote(self, lista_archivos: list):
        self.errores.clear()
        self.correctos.clear()

        for archivo in lista_archivos:
            estado = ValidadorReglas.evaluar_registro(archivo)
            archivo["estado_validacion"] = estado

            if "🔴" in estado or "🟡" in estado:
                self.errores.append(archivo)
            else:
                self.correctos.append(archivo)

        return {"correctos": self.correctos, "excepciones": self.errores}
