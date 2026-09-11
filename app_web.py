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

st.set_page_config(
    page_title="Descarga de Coberturas - CORESALUD (Web)",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 Sistema de Coberturas Médicas")

# --- Estado de Sesión ---
if "pacientes_cola" not in st.session_state:
    st.session_state.pacientes_cola = []
if "ejecutando" not in st.session_state:
    st.session_state.ejecutando = False

# Sidebar
st.sidebar.header("Opciones del Sistema")
modo = st.sidebar.radio("Selecciona el Modo:", ["Ingreso Manual", "Lote (Desde Excel)", "Generar Copia Excel"])


# --- MODO 1: INGRESO MANUAL DE PACIENTES ---
if modo == "Ingreso Manual":
    st.header("👤 Ingreso Manual de Paciente")

    with st.form("form_paciente", clear_on_submit=False):
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
            st.error("❌ Cédula inválida. Debe contener 10 dígitos.")
        elif not dependencia:
            st.error("❌ Por favor ingresa la dependencia.")
        elif not responsable:
            st.error("❌ Por favor ingresa el nombre del responsable.")
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
            st.success(f"✅ Paciente C.I. {cedula_clean} agregado a la cola de espera.")

    # Mostrar Tabla de Pacientes en Cola y Botón de Procesar
    if st.session_state.pacientes_cola:
        st.subheader("📋 Lista de Pacientes por Procesar")
        st.dataframe(st.session_state.pacientes_cola)

        col_proc, col_limp = st.columns([2, 1])
        with col_proc:
            if st.button("🚀 ▶️ PROCESAR PACIENTES AHORA", disabled=st.session_state.ejecutando):
                st.session_state.ejecutando = True
                st.info("⏳ Iniciando la consulta de coberturas y la actualización de la Matriz...")

                errores = []
                sin_seguro = []
                filas_agregadas = 0

                try:
                    os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                    carpeta_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                    os.makedirs(carpeta_temp, exist_ok=True)
                    carpeta_diag = os.path.join(core.CARPETA_SALIDA, "_diagnostico")

                    driver_holder = [core.crear_driver(carpeta_temp)]

                    try:
                        for idx, datos in enumerate(st.session_state.pacientes_cola, start=1):
                            reg = {
                                "cedula": datos["cedula"],
                                "nombre": None,
                                "fechas": [datos["fecha_atencion"]],
                                "cedula_padre": None,
                                "cedula_titular": None,
                            }

                            filas_matriz = core.procesar_registro(
                                reg, idx, idx, driver_holder,
                                carpeta_temp, carpeta_diag,
                                errores, sin_seguro, recolectar_matriz=True
                            )

                            if filas_matriz:
                                esc = core.agregar_filas_matriz_con_bloqueo(
                                    filas_matriz, datos["dependencia"], datos["fecha_nacimiento"],
                                    datos["sexo_codigo"], datos["observaciones"], datos["responsable"]
                                )
                                filas_agregadas += len(esc)
                                datos["estado"] = "OK (Agregado a Matriz)"
                            else:
                                datos["estado"] = "Sin cobertura / Reorganizado"

                    finally:
                        try:
                            driver_holder[0].quit()
                        except Exception:
                            pass

                    core._guardar_reportes_finales(errores, sin_seguro)
                    st.success(f"🎉 ¡Proceso finalizado! Se agregaron {filas_agregadas} fila(s) a la Matriz de Excel.")
                    st.info(f"📁 Los archivos PDF se guardaron en: `{os.path.abspath(core.CARPETA_SALIDA)}`")

                except Exception as e:
                    st.error(f"❌ Ocurrió un error inesperado: {e}")
                finally:
                    st.session_state.ejecutando = False

        with col_limp:
            if st.button("🗑️ Limpiar Lista"):
                st.session_state.pacientes_cola = []
                st.rerun()


# --- MODO 2: PROCESAMIENTO POR LOTE ---
elif modo == "Lote (Desde Excel)":
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

            try:
                registros = core.leer_excel(core.ARCHIVO_EXCEL)[cite: 4]
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
                st.info(f"📁 Revisa la carpeta: `{os.path.abspath(core.CARPETA_SALIDA)}`")
            except Exception as e:
                st.error(f"❌ Error durante la ejecución: {e}")
            finally:
                st.session_state.ejecutando = False


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
            )[cite: 4]
            st.success(f"✅ Archivo generado exitosamente en:\n`{destino}`")
        except Exception as e:
            st.error(f"❌ Error al generar la copia: {e}")
