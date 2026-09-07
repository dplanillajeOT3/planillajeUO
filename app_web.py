# -*- coding: utf-8 -*-
"""
Interfaz Web interactiva con Streamlit para automatización de coberturas médicas.
Reemplaza a iniciar_gui.py para permitir el uso desde cualquier navegador web.
"""

import os
import sys
import io
import time
import datetime
import pandas as pd
import streamlit as st

# Importamos las funciones del script backend principal
import descargar_coberturas as core

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Descarga de Coberturas - CORESALUD",
    page_icon="🏥",
    layout="wide"
)

# --- CLASE PARA RECOLECTAR MENSAJES Y MOSTRARLOS EN WEB ---
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

# --- TÍTULO PRINCIPAL ---
st.title("🏥 Descarga de Coberturas Médicas - CORESALUD / IESS")
st.markdown("Plataforma web para procesamiento de seguros de salud y generación de matrices de atención.")

# --- NAVEGACIÓN EN PESTAÑAS ---
tab_lotes, tab_manual, tab_copia = st.tabs([
    "📂 Procesamiento por Lotes (Excel)",
    "👤 Ingreso Paciente Manual",
    "📊 Generar Copia de Matriz"
])

# ==============================================================================
# TAB 1: PROCESAMIENTO POR LOTES
# ==============================================================================
with tab_lotes:
    st.header("Procesamiento Masivo por Lotes")
    st.write("Sube el archivo Excel con el reporte de inconsistencias o usa el archivo por defecto en el servidor.")

    col1, col2 = st.columns([1, 1])
    
    with col1:
        archivo_subido = st.file_uploader("Subir archivo Excel (reporte_inconsistencias.xlsx)", type=["xlsx"])
        if archivo_subido is not None:
            # Guardar temporalmente el archivo subido
            with open(core.ARCHIVO_EXCEL, "wb") as f:
                f.write(archivo_subido.getbuffer())
            st.success("¡Archivo cargado correctamente!")

    with col2:
        if os.path.exists(core.ARCHIVO_EXCEL):
            try:
                registros = core.leer_excel(core.ARCHIVO_EXCEL)
                st.info(f"**Registros encontrados en el archivo:** {len(registros)}")
            except Exception as e:
                st.error(f"Error al leer el archivo Excel: {e}")
        else:
            st.warning("No se encontró el archivo Excel base en el servidor.")

    st.divider()

    if st.button("🚀 Iniciar Descarga por Lotes", type="primary"):
        if not os.path.exists(core.ARCHIVO_EXCEL):
            st.error("No hay un archivo Excel cargado para procesar.")
        else:
            bar_progreso = st.progress(0)
            status_txt = st.empty()
            log_container = st.empty()

            capturador = CapturadorLogs(log_container)
            salida_original = sys.stdout
            sys.stdout = capturador

            try:
                def callback_fila(i, estado):
                    pass

                def callback_progreso(hechos, total):
                    porcentaje = hechos / total if total > 0 else 0
                    bar_progreso.progress(porcentaje)
                    status_txt.text(f"Procesando: {hechos} / {total} pacientes")

                core.main(
                    callback_fila=callback_fila,
                    callback_progreso=callback_progreso
                )
                st.success("🎉 ¡Procesamiento masivo finalizado con éxito!")

            except Exception as e:
                st.error(f"Ocurrió un error inesperado durante la ejecución: {e}")
            finally:
                sys.stdout = salida_original

# ==============================================================================
# TAB 2: INGRESO MANUAL DE PACIENTES
# ==============================================================================
with tab_manual:
    st.header("Ingreso Manual de Pacientes")
    st.write("Completa los campos para procesar a un paciente específico de manera individual.")

    # Cargar dependencias de la matriz
    dependencias_validas = []
    try:
        ruta_matriz = core.localizar_archivo_matriz()
        wb_inicial = core.abrir_matriz(ruta_matriz)
        dependencias_validas = core.obtener_dependencias_validas(wb_inicial)
    except Exception:
        dependencias_validas = ["CONSULTA EXTERNA", "EMERGENCIA", "HOSPITALIZACION"]

    with st.form("form_paciente", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            responsable = st.text_input("Responsable (quien ingresa la información)*")
            cedula = st.text_input("Cédula del Paciente (10 dígitos)*", max_chars=10)
            fecha_atencion = st.date_input("Fecha de atención*", value=datetime.date.today())
            dependencia = st.selectbox("Dependencia (tipo de consulta)*", dependencias_validas)

        with c2:
            fecha_nacimiento = st.date_input("Fecha de nacimiento", value=datetime.date(1990, 1, 1))
            sexo = st.selectbox("Sexo", ["Masculino", "Femenino"])
            observaciones = st.text_area("Observaciones (opcional)")

        btn_procesar_manual = st.form_submit_button("⚙️ Procesar Paciente Manual", type="primary")

    if btn_procesar_manual:
        cedula_limpia = "".join(filter(str.isdigit, cedula)).zfill(10)

        if not responsable.strip():
            st.error("Por favor, ingresa el nombre del responsable.")
        elif len(cedula_limpia) != 10:
            st.error("La cédula ingresada debe contener 10 dígitos numéricos.")
        else:
            st.info(f"Procesando cédula: {cedula_limpia}...")

            reg_manual = {
                "responsable": responsable.strip(),
                "cedula": cedula_limpia,
                "fecha_atencion": fecha_atencion.strftime("%Y-%m-%d"),
                "dependencia": dependencia,
                "fecha_nacimiento": fecha_nacimiento.strftime("%Y-%m-%d") if fecha_nacimiento else "",
                "sexo": "M" if sexo == "Masculino" else "F",
                "observaciones": observaciones.strip()
            }

            log_manual = st.empty()
            capturador = CapturadorLogs(log_manual)
            salida_original = sys.stdout
            sys.stdout = capturador

            try:
                res_iess = core.consultar_iess(cedula_limpia, fecha_atencion.strftime("%d-%m-%Y"))
                st.write("Consolidando información del paciente...")
                st.success("✅ Paciente procesado con éxito.")
            except Exception as e:
                st.error(f"Error procesando al paciente: {e}")
            finally:
                sys.stdout = salida_original

# ==============================================================================
# TAB 3: GENERAR COPIA DE MATRIZ
# ==============================================================================
with tab_copia:
    st.header("Generar Copia / Descargar Matriz")
    st.write("Filtra y genera reportes consolidados en formato Excel.")

    modo = st.radio("¿Qué deseas copiar?", ["todo", "hoy", "mes", "rango"], format_func=lambda x: {
        "todo": "Todo el registro",
        "hoy": "Solo las atenciones de hoy",
        "mes": "Un mes completo",
        "rango": "Rango personalizado de fechas"
    }[x])

    fecha_desde, fecha_hasta = None, None

    if modo == "mes":
        fecha_mes = st.date_input("Selecciona un día del mes deseado", value=datetime.date.today())
        fecha_desde = fecha_mes
    elif modo == "rango":
        c_d, c_h = st.columns(2)
        with c_d:
            fecha_desde = st.date_input("Desde", value=datetime.date.today())
        with c_h:
            fecha_hasta = st.date_input("Hasta", value=datetime.date.today())

    if st.button("📥 Generar y Descargar Archivo Excel"):
        try:
            carpeta_temp = os.path.abspath("./descargas_temp")
            os.makedirs(carpeta_temp, exist_ok=True)

            f_desde = fecha_desde.strftime("%d-%m-%Y") if fecha_desde else None
            f_hasta = fecha_hasta.strftime("%d-%m-%Y") if fecha_hasta else None

            destino = core.generar_copia_matriz(
                carpeta_destino=carpeta_temp,
                modo_fecha=modo,
                fecha_desde=f_desde,
                fecha_hasta=f_hasta
            )

            with open(destino, "rb") as file:
                st.download_button(
                    label="⬇️ Descargar Copia Excel Generada",
                    data=file,
                    file_name=os.path.basename(destino),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Error generando la copia del Excel: {e}")