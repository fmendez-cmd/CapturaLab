# INFRAESTRUCTURA/auth_rbac.py

ROLES = {
    "capturista": ["importar_erp", "ver_excepciones"],
    "coordinador": ["importar_erp", "ver_excepciones", "aprobar_lote", "generar_pdf"],
    "jefe_laboratorio": ["ver_excepciones", "firmar_informe", "administrar_plantillas"],
}


class ControlAcceso:
    @staticmethod
    def tiene_permiso(rol: str, accion: str) -> bool:
        permisos = ROLES.get(rol.lower(), [])
        return accion in permisos
