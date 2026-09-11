# -*- coding: utf-8 -*-
"""
Interfaz Web para ejecutar descargar_coberturas.py desde un navegador.

COMO USARLO:
    streamlit run app_web.py
"""

import os
import sys
import re
from datetime import date, datetime
import streamlit as st
import descargar_coberturas as core

st.set_page_config(
    page_title="Descarga de Coberturas - CORESALUD (Web)",
    page_icon="🏥",
    layout="wide"
)

st.title("🏥 Sistema de Coberturas Médicas")

# --- Estado de Sesión ---
if "ejecutando" not in st.session_state:
    st.session_state.ejecutando = False

# Sidebar
st.sidebar.header("Opciones del Sistema")
modo = st.sidebar.radio("Selecciona el Modo:", ["Ingreso Manual", "Lote (Desde Excel)", "Generar Copia Excel"])


# ============================================================
# MODO 1: INGRESO MANUAL DE PACIENTES
# ============================================================
if modo == "Ingreso Manual":
    st.header("👤 Ingreso Manual de Paciente")

    with st.form("form_paciente", clear_on_submit=True):
        # 1. Responsable de la información al inicio
        responsable = st.text_input("Responsable de la información:").strip().upper()

        col1, col2 = st.columns(2)
        with col1:
            cedula = st.text_input("Cédula del paciente (10 dígitos):")
            fecha_atencion = st.date_input("Fecha de atención:", date.today())
            dependencia = st.text_input("Dependencia / Tipo de consulta:")
            
        with col2:
            # 2. Fecha de nacimiento en formato de texto con autocompletado
            fecha_nac_raw = st.text_input("Fecha de nacimiento (ej. 15081995 o 15-08-1995):", placeholder="DD-MM-YYYY")
            sexo = st.selectbox("Sexo:", ["Masculino", "Femenino"])
            observaciones = st.text_input("Observaciones (opcional):")

        # 3. Botón de acción e ingreso
        boton_agregar = st.form_submit_button("➕ Agregar a la cola y continuar con el siguiente")

    if boton_agregar:
        # Validación y limpieza de la cédula
        cedula_clean = "".join(filter(str.isdigit, cedula)).zfill(10) if cedula else ""
        
        # Formateo automático de la fecha de nacimiento (inserta guiones automáticamente)
        fecha_nac_clean = "".join(filter(str.isdigit, fecha_nac_raw))
        fecha_nac_obj = None

        if len(fecha_nac_clean) == 8:
            # Convierte '15081995' a '15-08-1995'
            fecha_nac_fmt = f"{fecha_nac_clean[:2]}-{fecha_nac_clean[2:4]}-{fecha_nac_clean[4:]}"
            try:
                fecha_nac_obj = datetime.strptime(fecha_nac_fmt, "%d-%m-%Y").date()
            except ValueError:
                fecha_nac_obj = None
        elif len(fecha_nac_raw) == 10 and re.match(r"^\d{2}-\d{2}-\d{4}$", fecha_nac_raw):
            try:
                fecha_nac_obj = datetime.strptime(fecha_nac_raw, "%d-%m-%Y").date()
            except ValueError:
                fecha_nac_obj = None

        # Validación de errores en campos
        if not responsable:
            st.error("❌ Por favor ingresa el nombre del Responsable de la información.")
        elif len(cedula_clean) != 10:
            st.error("❌ Cédula inválida. Debe contener exactamente 10 dígitos.")
        elif not dependencia:
            st.error("❌ Por favor ingresa la Dependencia / Tipo de consulta.")
        elif not fecha_nac_obj:
            st.error("❌ Fecha de nacimiento inválida. Ingresa 8 dígitos (ej. 15081995) o en formato DD-MM-YYYY.")
        else:
            # 4. Procesamiento automático inmediato en el servidor
            st.session_state.ejecutando = True
            st.info(f"⏳ Procesando paciente C.I. {cedula_clean}... Consultando portales y actualizando Matriz.")

            sexo_codigo = "M" if sexo == "Masculino" else "F"
            errores, sin_seguro = [], []

            try:
                os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                carpeta_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                os.makedirs(carpeta_temp, exist_ok=True)
                carpeta_diag = os.path.join(core.CARPETA_SALIDA, "_diagnostico")

                driver_holder = [core.crear_driver(carpeta_temp)]

                try:
                    reg = {
                        "cedula": cedula_clean,
                        "nombre": None,
                        "fechas": [fecha_atencion],
                        "cedula_padre": None,
                        "cedula_titular": None,
                    }

                    filas_matriz = core.procesar_registro(
                        reg, 1, 1, driver_holder,
                        carpeta_temp, carpeta_diag,
                        errores, sin_seguro, recolectar_matriz=True
                    )

                    if filas_matriz:
                        esc = core.agregar_filas_matriz_con_bloqueo(
                            filas_matriz, dependencia, fecha_nac_obj,
                            sexo_codigo, observaciones, responsable
                        )
                        st.toast(f"✅ Paciente C.I. {cedula_clean} procesado correctamente.", icon="🎉")
                        st.success(f"✅ Se registraron {len(esc)} fila(s) en la Matriz Excel del servidor.")

                        # Buscar el PDF generado para permitir la descarga directa al navegador
                        archivos_pdf = []
                        for root, dirs, files in os.walk(core.CARPETA_SALIDA):
                            for f in files:
                                if f.lower().endswith(".pdf") and cedula_clean in f:
                                    archivos_pdf.append(os.path.join(root, f))

                        if archivos_pdf:
                            pdf_path = archivos_pdf[-1]
                            with open(pdf_path, "rb") as pdf_file:
                                st.download_button(
                                    label="📄 Descargar PDF del Paciente a mi PC",
                                    data=pdf_file,
                                    file_name=os.path.basename(pdf_path),
                                    mime="application/pdf"
                                )
                    else:
                        st.warning(f"⚠️ Paciente C.I. {cedula_clean} procesado. No se encontraron coberturas activas.")

                finally:
                    try:
                        driver_holder[0].quit()
                    except Exception:
                        pass

                core._guardar_reportes_finales(errores, sin_seguro)

            except Exception as e:
                st.error(f"❌ Ocurrió un error durante el procesamiento del paciente: {e}")
            finally:
                st.session_state.ejecutando = False


# ============================================================
# MODO 2: PROCESAMIENTO POR LOTE
# ============================================================
elif modo == "Lote (Desde Excel)":
    st.header("📋 Procesamiento por Lote")
    st.write("Lee los datos directamente desde el archivo `reporte_inconsistencias.xlsx`.")

    quiere_matriz = st.checkbox("¿Agregar también los resultados a la matriz de Excel?", value=True)
    responsable = st.text_input("Nombre del Responsable:").strip().upper()

    if st.button("🚀 Iniciar Descarga por Lotes", disabled=st.session_state.ejecutando):
        if quiere_matriz and not responsable:
            st.warning("⚠️ Debes ingresar el nombre del responsable para continuar.")
        else:
            st.session_state.ejecutando = True
            st.info("Iniciando proceso en segundo plano...")

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


# ============================================================
# MODO 3: GENERAR COPIA EN EXCEL
# ============================================================
elif modo == "Generar Copia Excel":
    st.header("📊 Exportar Copia de Matriz en Excel")

    modo_copia = st.selectbox("¿Qué deseas copiar?", ["todo", "hoy", "mes", "rango"])
    
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
            # Guarda temporalmente el Excel copiado en la carpeta de salidas
            destino = core.generar_copia_matriz(
                carpeta_destino=core.CARPETA_SALIDA,
                modo_fecha=modo_copia,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta
            )
            
            st.success("✅ Archivo generado exitosamente en el servidor.")
            
            # Botón para descargar directamente el archivo desde cualquier cliente web
            with open(destino, "rb") as file:
                st.download_button(
                    label="📥 Descargar Matriz Excel a mi PC",
                    data=file,
                    file_name=os.path.basename(destino),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ Error al generar la copia: {e}")
