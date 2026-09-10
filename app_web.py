import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Descarga de Coberturas - CORESALUD", layout="wide")

st.title("📋 Descarga de Coberturas - CORESALUD")

# Definir las pestañas de izquierda a derecha tal como en tus imágenes
tab_lotes, tab_manual = st.tabs(["📁 Procesamiento por Lotes", "👤 Ingresar Paciente Manual"])

# ==========================================
# 1. PESTAÑA: POR LOTES (Pantalla Principal)
# ==========================================
with tab_lotes:
    # Barra de botones superiores
    col_btn1, col_btn2, col_btn3, col_btn4, col_métrica = st.columns([2, 2, 2, 2, 3])
    
    with col_btn1:
        btn_iniciar = st.button("🚀 Iniciar descarga (por lotes)", use_container_width=True)
    with col_btn2:
        btn_reintentar = st.button("🔄 Reintentar fallidos", use_container_width=True)
    with col_btn3:
        btn_excel = st.button("📊 Generar copia en Excel", use_container_width=True)
    with col_métrica:
        st.subheader("0 / 10 pacientes")

    # Carga de archivo
    archivo_subido = st.file_uploader("Seleccionar matriz Excel", type=["xlsx"])

    # Tabla principal de pacientes (Dataframe interactivo)
    st.markdown("**Pacientes en cola:**")
    
    # Ejemplo de estructura de datos para la tabla
    datos_ejemplo = [
        {"#": 1, "Paciente": "ARAUZ ARAUZ MARIA SUSANA", "Cedula": "1710224963", "Fecha(s)": "02-07-2026", "Estado": "Pendiente"},
        {"#": 2, "Paciente": "BENALCAZAR TORRES MARIA REBECA", "Cedula": "1704995313", "Fecha(s)": "08-07-2026", "Estado": "Pendiente"},
        {"#": 3, "Paciente": "COLLAGUAZO TORRES ANALI MARIANA", "Cedula": "1755260781", "Fecha(s)": "02-07-2026", "Estado": "Pendiente"},
    ]
    st.dataframe(pd.DataFrame(datos_ejemplo), use_container_width=True, height=250)

    # Console / Terminal de logs abajo (simula el cuadro negro 'Detalle')
    st.markdown("**Detalle:**")
    st.code("Matriz cargada correctamente. Listo para iniciar...", language="bash")


# ==========================================
# 2. PESTAÑA: MANUAL / INTERACTIVO
# ==========================================
with tab_manual:
    st.subheader("Ingresar pacientes manualmente")
    
    # Formulario
    col1, col2 = st.columns(2)
    
    with col1:
        responsable = st.text_input("Responsable (persona que ingresa la información):")
        cedula = st.text_input("Cédula del paciente:")
        fecha_atencion = st.date_input("Fecha de atención:", value=datetime.now())
        dependencia = st.selectbox("Dependencia (tipo de consulta):", ["Consulta Externa", "Emergencia", "Hospitalización"])
        
    with col2:
        fecha_nacimiento = st.date_input("Fecha de nacimiento:")
        sexo = st.radio("Sexo:", ["M", "F"], horizontal=True)
        observaciones = st.text_input("Observaciones (opcional):")

    btn_agregar = st.button("➕ Agregar a la cola y seguir con el siguiente", type="primary")

    st.divider()

    # Métricas y acciones secundarias
    col_stat, col_act1, col_act2 = st.columns([3, 2, 2])
    with col_stat:
        st.write("**0 en cola, 0 completados**")
    with col_act1:
        st.button("Reintentar fallidos ", key="m_reintentar")
    with col_act2:
        st.button("Generar copia en Excel", key="m_excel")

    # Tabla manual y consola de salida
    st.dataframe(pd.DataFrame(columns=["#", "Cedula", "Fecha atencion", "Dependencia", "Estado"]), use_container_width=True)
    
    st.markdown("**Detalle:**")
    st.code("Matriz cargada: INSTRUCTIVO5 SEPTIEMBRE 2026.xlsx", language="bash")
