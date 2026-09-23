# DOMINIO/generador_excel.py
#
# Generador Excel basado en mapa estructural v2.
# Solo escribe en la hoja ERP.
# La IA nunca modifica Excel directamente.

from __future__ import annotations

import os
import shutil
import time
from pathlib import Path

from DOMINIO.control_cancelacion import verificar_cancelacion
from DOMINIO.gestor_mapas import GestorMapas
from DOMINIO.validador_mapa import ValidadorMapa
from INFRAESTRUCTURA.drive_connector import DriveConnector
from INFRAESTRUCTURA.config_rutas import TEMP_PLANTILLAS_DIR


class GeneradorExcel:
    PASSWORDS_HOJA = ("686", "20000000")

    XL_CALCULATION_AUTOMATIC = -4105
    XL_CALCULATION_MANUAL = -4135

    @staticmethod
    def _importar_com():
        if os.name != "nt":
            raise RuntimeError(
                "GeneradorExcel requiere Windows + Microsoft Excel."
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
                "Debe existir exactamente una hoja ERP. "
                f"Encontradas: {[str(h.Name) for h in hojas]}"
            )

        return hojas[0]

    @staticmethod
    def _top_left(celda):
        try:
            if bool(celda.MergeCells):
                return celda.MergeArea.Cells(1, 1)
        except Exception:
            pass

        return celda

    @staticmethod
    def _contar_formulas(hoja) -> int:
        used = hoja.UsedRange
        total = 0

        for r in range(1, int(used.Rows.Count) + 1):
            for c in range(1, int(used.Columns.Count) + 1):
                celda = used.Cells(r, c)

                try:
                    if bool(celda.HasFormula):
                        total += 1
                except Exception:
                    pass

        return total

    @staticmethod
    def _estado_proteccion(hoja) -> dict:
        estado = {
            "protegida": False,
            "contents": False,
            "drawing_objects": False,
            "scenarios": False,
        }

        try:
            estado["contents"] = bool(
                hoja.ProtectContents
            )
        except Exception:
            pass

        try:
            estado["drawing_objects"] = bool(
                hoja.ProtectDrawingObjects
            )
        except Exception:
            pass

        try:
            estado["scenarios"] = bool(
                hoja.ProtectScenarios
            )
        except Exception:
            pass

        estado["protegida"] = any(
            (
                estado["contents"],
                estado["drawing_objects"],
                estado["scenarios"],
            )
        )

        return estado

    @staticmethod
    def _desproteger(hoja, estado: dict) -> str | None:
        if not estado["protegida"]:
            print("Hoja ERP sin protección.")
            return None

        print(
            "Desprotegiendo hoja ERP con una "
            "contraseña autorizada..."
        )

        for password in GeneradorExcel.PASSWORDS_HOJA:
            try:
                hoja.Unprotect(Password=password)
            except Exception:
                continue

            try:
                sigue = bool(hoja.ProtectContents)
            except Exception:
                sigue = False

            if not sigue:
                print(
                    "✅ Hoja ERP desprotegida temporalmente."
                )
                # Se devuelve para restaurar exactamente la misma contraseña.
                return password

        raise RuntimeError(
            "No se pudo desproteger la hoja ERP con las "
            "contraseñas autorizadas."
        )

    @staticmethod
    def _reproteger(
        hoja,
        estado: dict,
        password_usado: str | None,
    ):
        if not estado["protegida"]:
            return

        if not password_usado:
            raise RuntimeError(
                "No se conoce la contraseña original con la que "
                "debe restaurarse la protección de la hoja ERP."
            )

        hoja.Protect(
            Password=password_usado,
            DrawingObjects=estado["drawing_objects"],
            Contents=estado["contents"],
            Scenarios=estado["scenarios"],
            UserInterfaceOnly=False,
        )

        print(
            "✅ Protección de la hoja ERP restaurada."
        )

    @staticmethod
    def _aplicar_mapa(
        mapa: dict,
        hoja_origen,
        hoja_destino,
    ) -> int:
        escritos = 0

        for indice, op in enumerate(
            mapa["operaciones"],
            start=1,
        ):
            origen = GeneradorExcel._top_left(
                hoja_origen.Range(
                    str(op["origen"])
                ).Cells(1, 1)
            )

            destino = GeneradorExcel._top_left(
                hoja_destino.Range(
                    str(op["destino"])
                ).Cells(1, 1)
            )

            valor = origen.Value2

            # Una celda mapeada puede estar vacía en un ensayo concreto.
            # La posición NO se compacta ni se desplaza:
            # origen y destino conservan siempre la correspondencia física.
            #
            # Si el origen está vacío, dejamos explícitamente vacío el destino.
            # Un 0 o un "-" NO son vacíos y se copian normalmente.
            try:
                if bool(destino.HasFormula):
                    raise ValueError(
                        f"Operación {indice}: "
                        f"{op['destino']} contiene fórmula."
                    )
            except TypeError:
                pass

            if valor is None or str(valor) == "":
                destino.Value2 = None
                continue

            destino.Value2 = valor
            escritos += 1

        return escritos

    @staticmethod
    def _validar_valores_escritos(
        mapa: dict,
        hoja_origen,
        hoja_destino,
    ):
        errores = []

        for op in mapa["operaciones"]:
            origen = GeneradorExcel._top_left(
                hoja_origen.Range(
                    str(op["origen"])
                ).Cells(1, 1)
            )

            destino = GeneradorExcel._top_left(
                hoja_destino.Range(
                    str(op["destino"])
                ).Cells(1, 1)
            )

            a = origen.Value2
            b = destino.Value2

            # Vacío en origen = vacío en la misma posición destino.
            # Excel COM puede representar una celda vacía como None o "".
            a_vacio = a is None or str(a) == ""
            b_vacio = b is None or str(b) == ""

            if a_vacio and b_vacio:
                continue

            if a_vacio != b_vacio or str(a) != str(b):
                errores.append(
                    (
                        op["origen"],
                        op["destino"],
                        a,
                        b,
                    )
                )

                if len(errores) >= 20:
                    break

        if errores:
            texto = "\n".join(
                f"  {o} -> {d}: origen={a!r}, destino={b!r}"
                for o, d, a, b in errores
            )

            raise ValueError(
                "La validación posterior al pegado encontró "
                "diferencias:\n"
                + texto
            )

    @staticmethod
    def inyectar_datos_dinamicos(
        codigo_gst: str,
        ruta_archivo_erp: str,
        ruta_salida_xlsx: str,
    ) -> str:
        _t_total = time.perf_counter()
        _tiempos = {}
        ruta_plantilla = None

        def _marca(nombre: str, inicio: float):
            _tiempos[nombre] = time.perf_counter() - inicio

        ruta_erp = str(
            Path(ruta_archivo_erp).resolve()
        )

        ruta_salida = str(
            Path(ruta_salida_xlsx).resolve()
        )

        if not Path(ruta_erp).is_file():
            raise FileNotFoundError(
                f"No existe el archivo ERP: {ruta_erp}"
            )

        print()
        print("=" * 78)
        print(
            f"GENERANDO EXCEL CON MAPA ESTRUCTURAL V2 "
            f"PARA {codigo_gst}"
        )
        print("=" * 78)
        print(f"ERP: {ruta_erp}")

        # --------------------------------------------------------
        # 1. Plantilla oficial
        # --------------------------------------------------------
        _t = time.perf_counter()
        ruta_plantilla = str(
            Path(
                DriveConnector.obtener_plantilla_por_gst(
                    codigo_gst
                )
            ).resolve()
        )

        if not Path(ruta_plantilla).is_file():
            raise FileNotFoundError(
                f"No existe la plantilla: {ruta_plantilla}"
            )

        print(f"Plantilla: {ruta_plantilla}")
        _marca("Buscar/descargar plantilla", _t)

        # --------------------------------------------------------
        # 2. Obtener / aprender mapa
        # --------------------------------------------------------
        _t = time.perf_counter()
        mapa = GestorMapas.obtener_o_crear(
            codigo_gst=codigo_gst,
            ruta_erp=ruta_erp,
            ruta_plantilla=ruta_plantilla
        )

        _marca("Obtener mapa V2.2", _t)
        verificar_cancelacion()

        # --------------------------------------------------------
        # 3. Copiar plantilla
        # --------------------------------------------------------
        _t = time.perf_counter()
        Path(ruta_salida).parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        if Path(ruta_salida).exists():
            Path(ruta_salida).unlink()

        shutil.copy2(
            ruta_plantilla,
            ruta_salida,
        )

        _marca("Copiar plantilla", _t)

        pythoncom, win32 = (
            GeneradorExcel._importar_com()
        )

        pythoncom.CoInitialize()

        excel = None
        libro_o = None
        libro_d = None
        guardado = False
        calculo_original = None

        try:
            _t = time.perf_counter()
            excel = win32.DispatchEx(
                "Excel.Application"
            )

            _marca("Iniciar Excel COM", _t)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
            excel.ScreenUpdating = False

            try:
                calculo_original = (
                    excel.Calculation
                )
                excel.Calculation = (
                    GeneradorExcel.XL_CALCULATION_MANUAL
                )
            except Exception:
                calculo_original = None

            _t = time.perf_counter()
            libro_o = excel.Workbooks.Open(
                ruta_erp,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
            )

            libro_d = excel.Workbooks.Open(
                ruta_salida,
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
            )

            _marca("Abrir ERP + plantilla", _t)

            hoja_o = libro_o.Worksheets(1)
            hoja_d = GeneradorExcel._buscar_hoja_erp(
                libro_d
            )

            print(
                f"Hoja ERP: '{hoja_d.Name}'"
            )

            # ----------------------------------------------------
            # 4. Seguridad según origen del mapa
            # ----------------------------------------------------
            mapa_reutilizado = bool(
                mapa.get("_reutilizado_desde_drive", False)
            )

            if mapa_reutilizado:
                print(
                    "⚡ Ruta rápida: mapa V2.2 ya validado "
                    "y reutilizado desde Drive."
                )
                print(
                    "   Se omiten validación estructural completa "
                    "y conteo global de fórmulas."
                )
                formulas_antes = None
            else:
                # Primera creación del mapa: conservar la ruta completa.
                _t = time.perf_counter()
                ValidadorMapa.validar_mapa_abierto(
                    mapa,
                    hoja_o,
                    hoja_d,
                )
                _marca("Validación previa", _t)

                _t = time.perf_counter()
                formulas_antes = (
                    GeneradorExcel._contar_formulas(
                        hoja_d
                    )
                )
                _marca("Contar fórmulas antes", _t)
                print(
                    f"Fórmulas antes: {formulas_antes}"
                )

            # ----------------------------------------------------
            # 5. Protección
            # ----------------------------------------------------
            estado = (
                GeneradorExcel._estado_proteccion(
                    hoja_d
                )
            )

            password_proteccion = GeneradorExcel._desproteger(
                hoja_d,
                estado,
            )

            # ----------------------------------------------------
            # 6. Aplicar mapa conservando posiciones y vacíos
            # ----------------------------------------------------
            _t = time.perf_counter()
            escritos = (
                GeneradorExcel._aplicar_mapa(
                    mapa,
                    hoja_o,
                    hoja_d,
                )
            )

            _marca("Escribir datos", _t)
            print(
                f"Celdas estructurales escritas: "
                f"{escritos}"
            )

            # ----------------------------------------------------
            # 7. Validación valor por valor
            # ----------------------------------------------------
            _t = time.perf_counter()
            GeneradorExcel._validar_valores_escritos(
                mapa,
                hoja_o,
                hoja_d,
            )

            _marca("Validar valores escritos", _t)
            print(
                "✅ Todos los valores mapeados coinciden "
                "con el ensayo ERP."
            )

            # ----------------------------------------------------
            # 8. Fórmulas intactas
            # ----------------------------------------------------
            if mapa_reutilizado:
                # En ruta rápida no recorremos toda la hoja.
                # _aplicar_mapa() ya comprueba HasFormula en CADA
                # celda destino antes de escribir y aborta si alguna
                # operación intentara sobrescribir una fórmula.
                print(
                    "⚡ Conteo global de fórmulas omitido: "
                    "mapa V2.2 reutilizado."
                )
            else:
                # Primera creación del mapa: mantener exactamente
                # la comprobación global histórica.
                _t = time.perf_counter()
                formulas_despues = (
                    GeneradorExcel._contar_formulas(
                        hoja_d
                    )
                )

                _marca("Contar fórmulas después", _t)
                print(
                    f"Fórmulas después: "
                    f"{formulas_despues}"
                )

                if formulas_despues != formulas_antes:
                    raise ValueError(
                        "SEGURIDAD: cambió la cantidad de "
                        "fórmulas. No se guardará."
                    )

            # ----------------------------------------------------
            # 9. Recalcular
            # ----------------------------------------------------
            _t = time.perf_counter()
            try:
                excel.Calculation = (
                    GeneradorExcel.XL_CALCULATION_AUTOMATIC
                )
            except Exception:
                pass

            try:
                excel.CalculateFullRebuild()
            except Exception:
                try:
                    excel.CalculateFull()
                except Exception:
                    pass

            _marca("Recalcular Excel", _t)

            # ----------------------------------------------------
            # 10. Reproteger y guardar
            # ----------------------------------------------------
            _t = time.perf_counter()
            GeneradorExcel._reproteger(
                hoja_d,
                estado,
                password_proteccion,
            )

            libro_d.Save()
            guardado = True
            _marca("Reproteger + guardar", _t)

            _tiempos["TOTAL hasta guardado"] = time.perf_counter() - _t_total

            print()
            print("=" * 78)
            print("⏱️ TIEMPOS DEL ENSAYE")
            print("=" * 78)
            for _nombre, _segundos in _tiempos.items():
                print(f"{_nombre:<32} {_segundos:>10.2f} s")
            print("=" * 78)

            print()
            print("=" * 78)
            print("✅ EXCEL V2 GENERADO CORRECTAMENTE")
            print("=" * 78)
            print(ruta_salida)

            return ruta_salida

        finally:
            if libro_o is not None:
                try:
                    libro_o.Close(
                        SaveChanges=False
                    )
                except Exception:
                    pass

            if libro_d is not None:
                try:
                    libro_d.Close(
                        SaveChanges=bool(guardado)
                    )
                except Exception:
                    pass

            if excel is not None:
                if calculo_original is not None:
                    try:
                        excel.Calculation = (
                            calculo_original
                        )
                    except Exception:
                        pass

                try:
                    excel.Quit()
                except Exception:
                    pass

            pythoncom.CoUninitialize()

            # La plantilla descargada es temporal: Drive es la fuente oficial.
            if ruta_plantilla:
                try:
                    _rp = Path(ruta_plantilla).resolve()
                    _rt = Path(TEMP_PLANTILLAS_DIR).resolve()
                    if _rt == _rp.parent or _rt in _rp.parents:
                        if _rp.exists():
                            _rp.unlink()
                except Exception as error:
                    print(
                        "⚠️ No se pudo eliminar la plantilla temporal: "
                        f"{error}"
                    )

            if (
                not guardado
                and Path(ruta_salida).exists()
            ):
                try:
                    Path(ruta_salida).unlink()
                except Exception:
                    pass
