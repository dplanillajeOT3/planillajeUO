# -*- coding: utf-8 -*-
"""
Interfaz Web para ejecutar descargar_coberturas.py desde un navegador.

COMO USARLO:
    streamlit run app_web.py
"""

import os
import sys
import queue
import threading
from datetime import date
import streamlit as st
import descargar_coberturas as core

# Configuración de la página
st.set_page_config(
    page_title="Descarga de Coberturas - CORESALUD (Web)",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 Sistema de Coberturas Médicas")

# --- Estado de Sesión (Session State) ---
if "log_texto" not in st.session_state:
    st.session_state.log_texto = ""
if "pacientes_cola" not in st.session_state:
    st.session_state.pacientes_cola = []
if "ejecutando" not in st.session_state:
    st.session_state.ejecutando = False

# Sidebar / Menú Lateral
st.sidebar.header("Opciones del Sistema")
modo = st.sidebar.radio("Selecciona el Modo:", ["Lote (Desde Excel)", "Ingreso Manual", "Generar Copia Excel"])


# --- MODO 1: PROCESAMIENTO POR LOTE ---
if modo == "Lote (Desde Excel)":
    st.header("📋 Procesamiento por Lote")
    st.write("Lee los datos directamente desde el archivo `reporte_inconsistencias.xlsx`[cite: 4].")

    quiere_matriz = st.checkbox("¿Agregar también los resultados a la matriz de Excel?", value=True)
    responsable = st.text_input("Nombre del Responsable:").strip().upper()

    if st.button("🚀 Iniciar Descarga por Lotes", disabled=st.session_state.ejecutando):
        if quiere_matriz and not responsable:
            st.warning("⚠️ Debes ingresar el nombre del responsable para continuar.")
        else:
            st.session_state.ejecutando = True
            st.info("Iniciando proceso en segundo plano...")

            def tarea_lote():
                salida_original = sys.stdout
                try:
                    registros = core.leer_excel(core.ARCHIVO_EXCEL)
                    if core.LIMITE_PRUEBA:
                        registros = registros[:core.LIMITE_PRUEBA]

                    items = list(enumerate(registros, start=1))
                    os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                    carpeta_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                    os.makedirs(carpeta_temp, exist_ok=True)
                    carpeta_diag = os.path.join(core.CARPETA_SALIDA, "_diagnostico")

                    driver_holder = [core.crear_driver(carpeta_temp)]
                    errores, sin_seguro = [], []

                    try:
                        core.procesar_lote(
                            items, driver_holder, carpeta_temp, carpeta_diag,
                            errores, sin_seguro,
                            escribir_matriz=quiere_matriz, responsable_matriz=responsable
                        )
                    finally:
                        try:
                            driver_holder[0].quit()
                        except Exception:
                            pass

                    core._guardar_reportes_finales(errores, sin_seguro)
                    st.success("✅ Proceso por lotes finalizado con éxito.")
                except Exception as e:
                    st.error(f"❌ Error durante la ejecución: {e}")
                finally:
                    st.session_state.ejecutando = False

            threading.Thread(target=tarea_lote, daemon=True).start()


# --- MODO 2: INGRESO MANUAL DE PACIENTES ---
elif modo == "Ingreso Manual":
    st.header("👤 Ingreso Manual de Paciente")

    with st.form("form_paciente", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            cedula = st.text_input("Cédula del paciente (10 dígitos):")
            fecha_atencion = st.date_input("Fecha de atención:", date.today())
            dependencia = st.text_input("Dependencia / Tipo de consulta:")
        with col2:
            fecha_nacimiento = st.date_input("Fecha de nacimiento:")
            sexo = st.selectbox("Sexo:", ["Masculino", "Femenino"])
            responsable = st.text_input("Responsable que ingresa:").strip().upper()

        observaciones = st.text_input("Observaciones (opcional):")
        boton_agregar = st.form_submit_button("➕ Agregar a la Cola")

    if boton_agregar:
        cedula_clean = "".join(filter(str.isdigit, cedula)).zfill(10)
        if len(cedula_clean) != 10:
            st.error("Cédula inválida. Debe contener 10 dígitos.")
        elif not responsable:
            st.error("Por favor ingresa el nombre del responsable.")
        else:
            paciente_datos = {
                "cedula": cedula_clean,
                "fecha_atencion": fecha_atencion,
                "dependencia": dependencia,
                "fecha_nacimiento": fecha_nacimiento,
                "sexo_codigo": "M" if sexo == "Masculino" else "F",
                "responsable": responsable,
                "observaciones": observaciones,
                "estado": "En cola"
            }
            st.session_state.pacientes_cola.append(paciente_datos)
            st.success(f"Paciente con Cédula {cedula_clean} agregado a la cola.")

    # Mostrar Tabla de Pacientes en Cola
    if st.session_state.pacientes_cola:
        st.subheader("📋 Queue / Pacientes Ingresados")
        st.dataframe(st.session_state.pacientes_cola)


# --- MODO 3: GENERAR COPIA EN EXCEL ---
elif modo == "Generar Copia Excel":
    st.header("📊 Exportar Copia de Matriz en Excel")

    modo_copia = st.selectbox("¿Qué deseas copiar?", ["todo", "hoy", "mes", "rango"])
    carpeta_destino = st.text_input("Ruta de carpeta destino:", os.path.expanduser("~"))

    fecha_desde = None
    fecha_hasta = None

    if modo_copia == "mes":
        fecha_mes = st.date_input("Selecciona cualquier fecha del mes a exportar:")
        fecha_desde = fecha_mes
    elif modo_copia == "rango":
        col1, col2 = st.columns(2)
        with col1:
            fecha_desde = st.date_input("Desde:")
        with col2:
            fecha_hasta = st.date_input("Hasta:")

    if st.button("📦 Generar Copia Excel"):
        try:
            destino = core.generar_copia_matriz(
                carpeta_destino=carpeta_destino,
                modo_fecha=modo_copia,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta
            )
            st.success(f"✅ Archivo generado exitosamente en:\n`{destino}`")
        except Exception as e:
            st.error(f"❌ Error al generar la copia: {e}")
