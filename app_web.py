# -*- coding: utf-8 -*-
"""
Interfaz Web Streamlit adaptada para:
- Autocompletar datos si la cédula ya existe en la base.
- Ejecutar el procesamiento asignando carpetas mensuales automáticamente.
"""

import os
import sys
import io
import datetime
import pandas as pd
import streamlit as st

import descargar_coberturas as core

st.set_page_config(
    page_title="Descarga de Coberturas - CORESALUD",
    page_icon="🏥",
    layout="wide"
)

# Capturador para ver la consola en vivo en la web
class CapturadorLogs(io.TextIOBase):
    def __init__(self, placeholder_log):
        self.placeholder_log = placeholder_log
        self.buffer = ""

    def write(self, texto):
        if texto:
            self.buffer += texto
            self.placeholder_log.code(self.buffer, language="text")
        return len(texto)

    def flush(self):
        pass

st.title("🏥 Descarga de Coberturas Médicas - CORESALUD")

tab_lotes, tab_manual = st.tabs([
    "📂 Procesamiento por Lotes (Excel)",
    "👤 Ingreso Paciente Manual"
])

# ==============================================================================
# TAB 1: BATCH / EXCEL
# ==============================================================================
with tab_lotes:
    st.header("Procesamiento por Lotes")
    archivo_subido = st.file_uploader("Subir archivo Excel de inconsistencias", type=["xlsx"])
    
    if archivo_subido is not None:
        with open(core.ARCHIVO_EXCEL_INCONSISTENCIAS, "wb") as f:
            f.write(archivo_subido.getbuffer())
        st.success("Archivo Excel guardado con éxito.")

    if st.button("🚀 Iniciar Descarga por Lotes", type="primary"):
        log_container = st.empty()
        bar_progreso = st.progress(0)
        
        capturador = CapturadorLogs(log_container)
        salida_original = sys.stdout
        sys.stdout = capturador

        try:
            def callback_progreso(hechos, total):
                bar_progreso.progress(hechos / total if total > 0 else 0)

            core.main(callback_progreso=callback_progreso)
            st.success("🎉 ¡Procesamiento finalizado!")
        except Exception as e:
            st.error(f"Error durante la ejecución: {e}")
        finally:
            sys.stdout = salida_original

# ==============================================================================
# TAB 2: INGRESO MANUAL CON BÚSQUEDA AUTOMÁTICA (PUNTO 4)
# ==============================================================================
with tab_manual:
    st.header("Ingreso Manual de Pacientes")
    st.caption("Si ingresas una cédula registrada anteriormente, el sistema autocompletará sus datos.")

    cedula_ingresada = st.text_input("Cédula del Paciente (10 dígitos)*", max_chars=10)
    
    # Búsqueda en el historial guardado (Punto 4)
    paciente_encontrado = None
    if len(cedula_ingresada) == 10:
        paciente_encontrado = core.buscar_paciente_historial(cedula_ingresada)
        if paciente_encontrado:
            st.info(f"✨ ¡Paciente registrado anteriormente!: {paciente_encontrado.get('apellidos', '')} {paciente_encontrado.get('nombres', '')}")

    with st.form("form_paciente"):
        c1, c2 = st.columns(2)
        with c1:
            responsable = st.text_input("Responsable", value="ADMIN")
            fecha_atencion = st.date_input("Fecha de atención", value=datetime.date.today())
            dependencia = st.selectbox("Dependencia", ["MEDICINA GENERAL (CE)", "EMERGENCIA", "HOSPITALIZACION"])

        with c2:
            nombres = st.text_input("Nombres", value=paciente_encontrado.get("nombres", "") if paciente_encontrado else "")
            apellidos = st.text_input("Apellidos", value=paciente_encontrado.get("apellidos", "") if paciente_encontrado else "")
            sexo = st.selectbox("Sexo", ["Masculino", "Femenino"], index=0 if (not paciente_encontrado or paciente_encontrado.get("sexo")=="M") else 1)

        btn_procesar = st.form_submit_button("⚙️ Procesar Paciente", type="primary")

    if btn_procesar:
        if len(cedula_ingresada) != 10:
            st.error("Ingresa una cédula válida de 10 dígitos.")
        else:
            datos_paciente = {
                "cedula": cedula_ingresada,
                "nombres": nombres,
                "apellidos": apellidos,
                "fecha_atencion": fecha_atencion.strftime("%Y-%m-%d"),
                "dependencia": dependencia,
                "sexo": "M" if sexo == "Masculino" else "F"
            }

            log_box = st.empty()
            capturador = CapturadorLogs(log_box)
            salida_original = sys.stdout
            sys.stdout = capturador

            try:
                core.procesar_paciente(datos_paciente)
                st.success("✅ Paciente procesado con éxito.")
            except Exception as e:
                st.error(f"Error procesando al paciente: {e}")
            finally:
                sys.stdout = salida_original
