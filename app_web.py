import streamlit as st
import pandas as pd
from datetime import datetime
import os
import zipfile

st.set_page_config(page_title="Descarga de Coberturas - CORESALUD", layout="wide")

# 1. Inicialización de la cola en la sesión si no existe
if "cola_pacientes" not in st.session_state:
    st.session_state.cola_pacientes = []

st.title("📋 Descarga de Coberturas - CORESALUD")

tab_lotes, tab_manual = st.tabs(["📁 Procesamiento por Lotes", "👤 Ingresar Paciente Manual"])

# Lista de dependencias extraída de la imagen
DEPENDENCIAS = [
    "GINECOLOGIA (CE)",
    "MEDICINA FAMILIAR",
    "MEDICINA GENERAL (CE)",
    "OBSTETRICIA (CE)",
    "PEDIATRIA (CE)",
    "PSICOLOGIA (CE)",
    "NUTRICION Y DIETETICA"
]

# ==========================================
# 1. PESTAÑA: POR LOTES
# ==========================================
with tab_lotes:
    col_btn1, col_btn2, col_métrica = st.columns([2, 2, 3])
    with col_btn1:
        btn_iniciar_lote = st.button("🚀 Iniciar descarga (por lotes)", use_container_width=True)
    with col_btn2:
        btn_reintentar = st.button("🔄 Reintentar fallidos", use_container_width=True)

    archivo_subido = st.file_uploader("Seleccionar matriz Excel (.xlsx)", type=["xlsx"])

    if archivo_subido:
        df_lote = pd.read_excel(archivo_subido)
        st.dataframe(df_lote, use_container_width=True, height=250)
        st.success(f"Cargados {len(df_lote)} registros desde el archivo.")

    st.markdown("**Detalle:**")
    st.code("Sistema listo para recibir lotes...", language="bash")


# ==========================================
# 2. PESTAÑA: MANUAL / INTERACTIVO
# ==========================================
with tab_manual:
    st.subheader("Ingresar pacientes manualmente")
    
    # Formulario interactivo
    with st.form(key="form_paciente", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            responsable = st.text_input("Responsable (persona que ingresa la información):")
            cedula = st.text_input("Cédula del paciente (*):")
            
            # Fecha de atención: por defecto HOY, formato DIA-MES-AÑO
            fecha_atencion = st.date_input("Fecha de atención (*):", value=datetime.now(), format="DD/MM/YYYY")
            dependencia = st.selectbox("Dependencia (tipo de consulta):", DEPENDENCIAS)
            
        with col2:
            # Fecha de nacimiento: en blanco (value=None), formato DIA-MES-AÑO
            fecha_nacimiento = st.date_input("Fecha de nacimiento (*):", value=None, format="DD/MM/YYYY")
            sexo = st.radio("Sexo (*):", ["M", "F"], horizontal=True)
            observaciones = st.text_input("Observaciones (opcional):")

        # Botón para agregar a la cola
        btn_agregar = st.form_submit_button("➕ Agregar a la cola y seguir con el siguiente", type="primary")

    # Lógica al dar clic en Agregar
    if btn_agregar:
        if not cedula or fecha_nacimiento is None:
            st.error("Por favor completa los campos obligatorios: Cédula y Fecha de nacimiento.")
        else:
            nuevo_paciente = {
                "#": len(st.session_state.cola_pacientes) + 1,
                "Responsable": responsable,
                "Cedula": cedula,
                "Fecha atención": fecha_atencion.strftime("%d-%m-%Y"),
                "Dependencia": dependencia,
                "Fecha nacimiento": fecha_nacimiento.strftime("%d-%m-%Y"),
                "Sexo": sexo,
                "Observaciones": observaciones,
                "Estado": "Pendiente"
            }
            st.session_state.cola_pacientes.append(nuevo_paciente)
            st.success(f"Paciente con Cédula {cedula} agregado a la cola.")

    st.divider()

    # Métrica de la cola
    st.write(f"**{len(st.session_state.cola_pacientes)} en cola**")

    # Tabla en tiempo real con los registros de la sesión
    df_cola = pd.DataFrame(st.session_state.cola_pacientes)
    st.dataframe(df_cola, use_container_width=True, height=200)

    # Descargar paquete ZIP (Matriz + PDFs)
    if len(st.session_state.cola_pacientes) > 0:
        if st.button("📦 Descargar Matriz y PDFs (.ZIP)"):
            # Generar Excel dinámico
            df_cola.to_excel("matriz_manual.xlsx", index=False)
            
            # Crear archivo ZIP en el servidor
            zip_filename = "resultado_coberturas.zip"
            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                if os.path.exists("matriz_manual.xlsx"):
                    zipf.write("matriz_manual.xlsx")
                # Incluye los archivos del directorio de resultados
                if os.path.exists("PDF_DESCARGADOS"):
                    for root, dirs, files in os.walk("PDF_DESCARGADOS"):
                        for file in files:
                            zipf.write(os.path.join(root, file))

            with open(zip_filename, "rb") as f:
                st.download_button(
                    label="⬇️ Confirmar Descarga de ZIP",
                    data=f,
                    file_name=zip_filename,
                    mime="application/zip"
                )

    st.markdown("**Detalle:**")
    st.code(f"Registros temporales almacenados: {len(st.session_state.cola_pacientes)}", language="bash")
