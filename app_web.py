# -*- coding: utf-8 -*-
"""
app_web.py
==========

Interfaz web (Streamlit) para el motor de descarga de coberturas de
salud. NO reimplementa la logica de Selenium / deteccion CAMPESINO-GENERAL
/ matriz: llama directamente a las funciones que ya existen y funcionan
en `descargar_coberturas.py` (que debe estar en la MISMA carpeta que
este archivo). Asi, cualquier correccion futura al motor se refleja
automaticamente aqui, sin mantener dos copias de la misma logica.

Como correrla:
    pip install -r requirements.txt
    streamlit run app_web.py

Funciona igual en Windows o Linux: crear_driver() en el motor ya
detecta automaticamente si existe Chromium/Chromedriver del sistema
(tipico de un servidor Linux) o si debe usar webdriver-manager (tipico
de una PC de escritorio Windows).
"""

import os
import io
import sys
import queue
import threading
import contextlib
import traceback
from datetime import date

import streamlit as st
import openpyxl

# ------------------------------------------------------------------
# Importa el motor original SIN modificarlo. Debe vivir junto a este
# archivo (mismo directorio).
# ------------------------------------------------------------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import descargar_coberturas as motor
except ImportError as e:
    st.set_page_config(page_title="Coberturas de Salud", page_icon="⚕️")
    st.error(
        "No se encontro `descargar_coberturas.py` en esta misma carpeta. "
        "Este archivo es obligatorio: la app web solo es la interfaz, "
        "toda la logica de descarga vive ahi.\n\n"
        f"Detalle: {e}"
    )
    st.stop()

st.set_page_config(page_title="Coberturas de Salud - MSP", page_icon="⚕️", layout="wide")

# ============================================================
# UTILIDADES COMUNES
# ============================================================

def _asegurar_carpetas():
    os.makedirs(motor.CARPETA_SALIDA, exist_ok=True)
    carpeta_descargas_temp = os.path.join(motor.CARPETA_SALIDA, "_descargas_temp")
    os.makedirs(carpeta_descargas_temp, exist_ok=True)
    carpeta_diagnostico = os.path.join(motor.CARPETA_SALIDA, "_diagnostico")
    return carpeta_descargas_temp, carpeta_diagnostico


def _tabla_desde_filas(filas, columnas):
    """Convierte una lista de listas (como las que arma el motor para
    errores/sin_seguro) en algo que st.dataframe pueda mostrar, sin
    depender de pandas."""
    if not filas:
        return None
    return [dict(zip(columnas, fila)) for fila in filas]


def _reset_estado_auto():
    st.session_state.auto_running = False
    st.session_state.auto_queue = None
    st.session_state.auto_thread = None
    st.session_state.auto_log = []
    st.session_state.auto_progress = (0, 0)
    st.session_state.auto_errores = []
    st.session_state.auto_sin_seguro = []
    st.session_state.auto_done = False
    st.session_state.auto_error_fatal = None


def _init_estado():
    if "auto_running" not in st.session_state:
        _reset_estado_auto()
    if "manual_driver_holder" not in st.session_state:
        st.session_state.manual_driver_holder = None
    if "manual_errores" not in st.session_state:
        st.session_state.manual_errores = []
    if "manual_sin_seguro" not in st.session_state:
        st.session_state.manual_sin_seguro = []
    if "manual_contador" not in st.session_state:
        st.session_state.manual_contador = 0
    if "manual_ultimo_resultado" not in st.session_state:
        st.session_state.manual_ultimo_resultado = None
    if "manual_datos_previos" not in st.session_state:
        st.session_state.manual_datos_previos = None
    if "manual_cedula_buscada" not in st.session_state:
        st.session_state.manual_cedula_buscada = None


_init_estado()

st.title("⚕️ Descarga de Coberturas de Salud")
st.caption(
    "Interfaz web del mismo motor de `descargar_coberturas.py`: consulta "
    "coberturasalud.msp.gob.ec (y app.iess.gob.ec cuando aplica) y organiza "
    "los PDF igual que la version de escritorio."
)

tab_auto, tab_manual = st.tabs(["📂 Modo automático (Excel)", "🧍 Modo manual (un paciente a la vez)"])

# ============================================================
# TAB 1 - MODO AUTOMATICO (equivalente a main())
# ============================================================

with tab_auto:
    st.subheader("Carga masiva desde Excel")
    st.write(
        f"Sube el archivo con las columnas **{motor.COL_NOMBRE}**, **{motor.COL_CEDULA}** y "
        f"**{motor.COL_FECHAS}** (y opcionalmente **{motor.COL_PADRE}** / **{motor.COL_TITULAR}**), "
        "igual que `reporte_inconsistencias.xlsx`."
    )

    archivo_subido = st.file_uploader(
        "Archivo Excel de inconsistencias", type=["xlsx"], key="auto_uploader",
        disabled=st.session_state.auto_running
    )

    col_a, col_b = st.columns(2)
    with col_a:
        escribir_matriz = st.checkbox(
            "Registrar también en la matriz mensual (Instructivo)",
            value=False, disabled=st.session_state.auto_running,
            help="Se agregan filas con Dependencia, fecha de nacimiento y sexo en blanco "
                 "-se completan despues a mano-, igual que el modo por lotes del script original."
        )
    with col_b:
        responsable_matriz = st.text_input(
            "Responsable (para la matriz)", value="", disabled=st.session_state.auto_running or not escribir_matriz
        ).upper()

    registros_preview = None
    if archivo_subido is not None and not st.session_state.auto_running:
        try:
            wb_preview = openpyxl.load_workbook(io.BytesIO(archivo_subido.getvalue()), data_only=True)
            ws_preview = wb_preview[motor.HOJA]
            n_filas = max(ws_preview.max_row - 1, 0)
            st.info(f"El archivo tiene {n_filas} fila(s) de datos (sin contar encabezado).")
        except Exception as e:
            st.warning(f"No se pudo previsualizar el archivo todavia (se validará al iniciar): {e}")

    iniciar = st.button(
        "▶️ Iniciar descarga", type="primary",
        disabled=st.session_state.auto_running or archivo_subido is None
    )

    if iniciar and archivo_subido is not None and not st.session_state.auto_running:
        # Guarda el excel subido exactamente donde el motor espera leerlo.
        with open(motor.ARCHIVO_EXCEL, "wb") as f:
            f.write(archivo_subido.getvalue())

        _reset_estado_auto()
        st.session_state.auto_running = True
        q = queue.Queue()
        st.session_state.auto_queue = q

        def _worker(q, escribir_matriz, responsable_matriz):
            class _QueueWriter:
                def write(self, s):
                    s = s.rstrip("\n")
                    if s.strip():
                        q.put(("log", s))

                def flush(self):
                    pass

            driver_holder = None
            try:
                with contextlib.redirect_stdout(_QueueWriter()):
                    registros = motor.leer_excel(motor.ARCHIVO_EXCEL)
                    if motor.LIMITE_PRUEBA:
                        registros = registros[: motor.LIMITE_PRUEBA]

                    carpeta_descargas_temp, carpeta_diagnostico = _asegurar_carpetas()
                    driver_holder = [motor.crear_driver(carpeta_descargas_temp)]

                    errores, sin_seguro = [], []
                    items = list(enumerate(registros, start=1))

                    def cb_fila(i, estado):
                        q.put(("fila", i, len(items), estado))

                    def cb_progreso(n, total):
                        q.put(("progress", n, total))

                    motor.procesar_lote(
                        items, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                        errores, sin_seguro, callback_fila=cb_fila, callback_progreso=cb_progreso,
                        escribir_matriz=escribir_matriz, responsable_matriz=responsable_matriz
                    )

                    motor._guardar_reportes_finales(errores, sin_seguro)

                q.put(("done", errores, sin_seguro))
            except Exception as e:
                q.put(("error", f"{type(e).__name__}: {e}\n{traceback.format_exc()}"))
            finally:
                if driver_holder is not None:
                    try:
                        driver_holder[0].quit()
                    except Exception:
                        pass

        t = threading.Thread(
            target=_worker, args=(q, escribir_matriz, responsable_matriz), daemon=True
        )
        st.session_state.auto_thread = t
        t.start()
        st.rerun()

    # --- Drena la cola y refresca mientras corre ---
    if st.session_state.auto_running and st.session_state.auto_queue is not None:
        q = st.session_state.auto_queue
        try:
            while True:
                item = q.get_nowait()
                tipo = item[0]
                if tipo == "log":
                    st.session_state.auto_log.append(item[1])
                elif tipo == "progress":
                    st.session_state.auto_progress = (item[1], item[2])
                elif tipo == "fila":
                    pass  # el detalle ya queda en el log
                elif tipo == "done":
                    st.session_state.auto_errores = item[1]
                    st.session_state.auto_sin_seguro = item[2]
                    st.session_state.auto_done = True
                    st.session_state.auto_running = False
                elif tipo == "error":
                    st.session_state.auto_error_fatal = item[1]
                    st.session_state.auto_running = False
        except queue.Empty:
            pass

    n, total = st.session_state.auto_progress
    if total:
        st.progress(min(n / total, 1.0), text=f"{n}/{total} registros procesados")

    if st.session_state.auto_log:
        with st.expander("Registro detallado (log)", expanded=st.session_state.auto_running):
            st.code("\n".join(st.session_state.auto_log[-500:]), language=None)

    if st.session_state.auto_running:
        st.info("Descargando... esta pestaña se actualiza sola.")
        import time as _time
        _time.sleep(1.2)
        st.rerun()

    if st.session_state.auto_error_fatal:
        st.error(f"La descarga se detuvo por un error:\n\n{st.session_state.auto_error_fatal}")

    if st.session_state.auto_done:
        errores = st.session_state.auto_errores
        sin_seguro = st.session_state.auto_sin_seguro
        if not errores and not sin_seguro:
            st.success("Todo se procesó sin errores.")
        else:
            st.warning(f"Terminado: {len(errores)} incidencia(s) de error, {len(sin_seguro)} caso(s) sin seguro.")

        col1, col2 = st.columns(2)
        with col1:
            tabla_err = _tabla_desde_filas(errores, ["Nombre", "Cédula", "Fecha", "Detalle"])
            if tabla_err:
                st.write("**Errores**")
                st.dataframe(tabla_err, use_container_width=True)
                ruta_log = os.path.join(motor.CARPETA_SALIDA, "log_errores.xlsx")
                if os.path.exists(ruta_log):
                    with open(ruta_log, "rb") as f:
                        st.download_button("⬇️ Descargar log_errores.xlsx", f.read(),
                                            file_name="log_errores.xlsx", key="dl_errores")
        with col2:
            tabla_ss = _tabla_desde_filas(
                sin_seguro, ["Nombre Paciente", "Cédula Paciente", "Cédula Consultada", "Rol", "Fecha"]
            )
            if tabla_ss:
                st.write("**Sin seguro**")
                st.dataframe(tabla_ss, use_container_width=True)
                ruta_ss = os.path.join(motor.CARPETA_SALIDA, "sin_seguro.xlsx")
                if os.path.exists(ruta_ss):
                    with open(ruta_ss, "rb") as f:
                        st.download_button("⬇️ Descargar sin_seguro.xlsx", f.read(),
                                            file_name="sin_seguro.xlsx", key="dl_sinseguro")

        if st.button("Empezar otra carga"):
            _reset_estado_auto()
            st.rerun()

# ============================================================
# TAB 2 - MODO MANUAL (equivalente a modo_interactivo())
# ============================================================

with tab_manual:
    st.subheader("Un paciente a la vez")

    # ---- Matriz: dependencias validas (si el Instructivo esta disponible) ----
    dependencias_validas = []
    hay_matriz_disponible = False
    aviso_matriz = None
    try:
        ruta_matriz_actual = motor.localizar_archivo_matriz()
        dependencias_validas = motor.obtener_dependencias_validas(motor.abrir_matriz(ruta_matriz_actual))
        hay_matriz_disponible = True
    except FileNotFoundError as e:
        aviso_matriz = str(e)
    except Exception as e:
        aviso_matriz = f"No se pudo leer la matriz: {e}"

    if aviso_matriz:
        st.warning(
            f"{aviso_matriz}\n\nSe podrán descargar los PDF de cobertura con normalidad, "
            "pero no se llenará la matriz automáticamente."
        )

    responsable = st.text_input("Nombre del responsable (persona que ingresa la información)", key="manual_responsable").upper()

    st.markdown("---")
    st.markdown(f"**Paciente #{st.session_state.manual_contador + 1}**")

    cedula_input = st.text_input("Cédula del paciente", key="manual_cedula", max_chars=10)

    datos_guardados = None
    if cedula_input:
        cedula_normalizada = "".join(ch for ch in cedula_input if ch.isdigit()).zfill(10)[:10]
        if cedula_normalizada != st.session_state.manual_cedula_buscada:
            st.session_state.manual_cedula_buscada = cedula_normalizada
            st.session_state.manual_datos_previos = motor.obtener_datos_guardados_paciente(cedula_normalizada)
        datos_guardados = st.session_state.manual_datos_previos
        if datos_guardados:
            partes = []
            if datos_guardados.get("nombre"):
                partes.append(datos_guardados["nombre"])
            if datos_guardados.get("fecha_nacimiento"):
                partes.append(f"nació {datos_guardados['fecha_nacimiento']}")
            if datos_guardados.get("sexo"):
                partes.append(f"sexo {datos_guardados['sexo']}")
            st.caption(f"Ya se tienen datos guardados de esta cédula: {', '.join(partes)}")

    fecha_nac_defecto = None
    if datos_guardados and datos_guardados.get("fecha_nacimiento"):
        fecha_nac_defecto = motor._parsear_fecha_flexible(datos_guardados["fecha_nacimiento"])
    sexo_defecto = (datos_guardados or {}).get("sexo", "F")

    with st.form("form_manual", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            fecha_atencion = st.date_input("Fecha de atención", value=date.today(), key="manual_fecha_atencion")
            fecha_nacimiento = st.date_input(
                "Fecha de nacimiento del paciente",
                value=fecha_nac_defecto or date(2000, 1, 1),
                min_value=date(1900, 1, 1), max_value=date.today(),
                key="manual_fecha_nac"
            )
            sexo = st.radio("Sexo del paciente", ["F", "M"], index=(0 if sexo_defecto != "M" else 1), horizontal=True)
        with c2:
            if dependencias_validas:
                dependencia = st.selectbox("Dependencia (tipo de consulta)",
                                            options=dependencias_validas + ["OTRA (escribir)"])
                if dependencia == "OTRA (escribir)":
                    dependencia = st.text_input("Escribe la dependencia").upper()
            else:
                dependencia = st.text_input("Dependencia (tipo de consulta)").upper()
            observaciones = st.text_area("Observaciones (opcional)", height=90)

        enviado = st.form_submit_button("🔍 Consultar y descargar cobertura", type="primary")

    if enviado:
        cedula_normalizada = "".join(ch for ch in (cedula_input or "") if ch.isdigit()).zfill(10)[:10]
        if not cedula_input or len(cedula_normalizada) != 10:
            st.error("Ingresa una cédula válida de 10 dígitos.")
        elif not responsable:
            st.error("Ingresa el nombre del responsable.")
        else:
            carpeta_descargas_temp, carpeta_diagnostico = _asegurar_carpetas()
            if st.session_state.manual_driver_holder is None:
                with st.spinner("Abriendo navegador..."):
                    st.session_state.manual_driver_holder = [motor.crear_driver(carpeta_descargas_temp)]
            driver_holder = st.session_state.manual_driver_holder

            st.session_state.manual_contador += 1
            contador = st.session_state.manual_contador

            reg = {
                "cedula": cedula_normalizada,
                "nombre": None,
                "fechas": [fecha_atencion],
                "cedula_padre": None,
                "cedula_titular": None,
            }

            log_buffer = io.StringIO()
            with st.spinner(f"Consultando cobertura de {cedula_normalizada}..."):
                try:
                    with contextlib.redirect_stdout(log_buffer):
                        filas_matriz = motor.procesar_registro(
                            reg, contador, contador, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                            st.session_state.manual_errores, st.session_state.manual_sin_seguro,
                            recolectar_matriz=True
                        )
                        motor.guardar_datos_paciente(
                            cedula_normalizada, nombre=reg.get("nombre"),
                            fecha_nacimiento=fecha_nacimiento, sexo=sexo
                        )
                        filas_escritas = []
                        if hay_matriz_disponible and filas_matriz:
                            try:
                                filas_escritas = motor.agregar_filas_matriz_con_bloqueo(
                                    filas_matriz, dependencia, fecha_nacimiento, sexo, observaciones, responsable
                                )
                            except TimeoutError as e:
                                st.session_state.manual_errores.append(
                                    [reg.get("nombre") or cedula_normalizada, cedula_normalizada,
                                     fecha_atencion.strftime("%d-%m-%Y"), f"(MATRIZ) {e}"]
                                )
                    st.session_state.manual_ultimo_resultado = {
                        "cedula": cedula_normalizada,
                        "nombre": reg.get("nombre"),
                        "log": log_buffer.getvalue(),
                        "filas_matriz": filas_escritas,
                        "error_fatal": None,
                    }
                except Exception as e:
                    st.session_state.manual_ultimo_resultado = {
                        "cedula": cedula_normalizada,
                        "nombre": reg.get("nombre"),
                        "log": log_buffer.getvalue(),
                        "filas_matriz": [],
                        "error_fatal": f"{type(e).__name__}: {e}",
                    }
            st.rerun()

    resultado = st.session_state.manual_ultimo_resultado
    if resultado:
        st.markdown("---")
        st.markdown(f"**Resultado — cédula {resultado['cedula']}**"
                     + (f" ({resultado['nombre']})" if resultado.get("nombre") else ""))
        if resultado["error_fatal"]:
            st.error(resultado["error_fatal"])
        else:
            st.success("Consulta procesada.")
            if resultado["filas_matriz"]:
                for ruta_m, fila in resultado["filas_matriz"]:
                    st.write(f"→ Fila {fila} agregada a la matriz ({os.path.basename(ruta_m)}).")
        if resultado["log"].strip():
            with st.expander("Detalle técnico"):
                st.code(resultado["log"], language=None)

    st.markdown("---")
    col_fin1, col_fin2 = st.columns(2)
    with col_fin1:
        if st.button("🧾 Generar reportes de esta sesión (errores / sin seguro)"):
            motor._guardar_reportes_finales(st.session_state.manual_errores, st.session_state.manual_sin_seguro)
            st.success("Reportes generados en la carpeta PDF_DESCARGADOS del servidor.")
            ruta_log = os.path.join(motor.CARPETA_SALIDA, "log_errores.xlsx")
            ruta_ss = os.path.join(motor.CARPETA_SALIDA, "sin_seguro.xlsx")
            if os.path.exists(ruta_log):
                with open(ruta_log, "rb") as f:
                    st.download_button("⬇️ log_errores.xlsx", f.read(), file_name="log_errores.xlsx", key="dl_man_err")
            if os.path.exists(ruta_ss):
                with open(ruta_ss, "rb") as f:
                    st.download_button("⬇️ sin_seguro.xlsx", f.read(), file_name="sin_seguro.xlsx", key="dl_man_ss")

    with col_fin2:
        if st.button("🛑 Cerrar sesión de navegador"):
            if st.session_state.manual_driver_holder is not None:
                try:
                    st.session_state.manual_driver_holder[0].quit()
                except Exception:
                    pass
                st.session_state.manual_driver_holder = None
            st.success("Navegador cerrado. Puedes seguir ingresando pacientes; se abrirá uno nuevo.")

    if hay_matriz_disponible:
        with st.expander("📋 Generar copia de la matriz para revisar/enviar"):
            modo_fecha = st.selectbox(
                "¿Qué copiar?",
                options=["todo", "hoy", "mes", "rango"],
                format_func=lambda v: {
                    "todo": "Todo", "hoy": "Solo hoy", "mes": "Un mes completo", "rango": "Un rango de fechas"
                }[v]
            )
            fecha_desde_copia = fecha_hasta_copia = None
            if modo_fecha == "mes":
                fecha_desde_copia = st.date_input("Cualquier fecha del mes a copiar", value=date.today(), key="copia_mes")
            elif modo_fecha == "rango":
                fecha_desde_copia = st.date_input("Desde", key="copia_desde")
                fecha_hasta_copia = st.date_input("Hasta", key="copia_hasta")

            if st.button("Generar copia"):
                try:
                    ruta_temporal = os.path.join(motor.CARPETA_SALIDA, "_copias_matriz_web")
                    copia = motor.generar_copia_matriz(
                        carpeta_destino=ruta_temporal, modo_fecha=modo_fecha,
                        fecha_desde=fecha_desde_copia, fecha_hasta=fecha_hasta_copia
                    )
                    with open(copia, "rb") as f:
                        st.download_button(
                            "⬇️ Descargar copia de la matriz", f.read(),
                            file_name=os.path.basename(copia), key="dl_copia_matriz"
                        )
                except Exception as e:
                    st.error(f"No se pudo generar la copia: {e}")
