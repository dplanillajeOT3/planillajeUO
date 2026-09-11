import os
import shutil
import zipfile
import time
from datetime import datetime
import pandas as pd
import streamlit as st

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# -----------------------------------------------------------------------------
# 1. CONFIGURACIÓN DEL DRIVER SELENIUM EN NUBE (HEADLESS)
# -----------------------------------------------------------------------------
def crear_driver(carpeta_descargas):
    os.makedirs(carpeta_descargas, exist_ok=True)
    
    RUTAS_CHROMIUM = ("/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome")
    RUTAS_CHROMEDRIVER = ("/usr/bin/chromedriver", "/usr/lib/chromium/chromedriver", "/usr/lib/chromium-browser/chromedriver")
    
    binario_chromium = next((r for r in RUTAS_CHROMIUM if os.path.exists(r)), None)
    ruta_driver_sistema = next((r for r in RUTAS_CHROMEDRIVER if os.path.exists(r)), None)

    opciones = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.abspath(carpeta_descargas),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    opciones.add_experimental_option("prefs", prefs)

    opciones.add_argument("--headless=new")
    opciones.add_argument("--no-sandbox")
    opciones.add_argument("--disable-dev-shm-usage")
    opciones.add_argument("--disable-gpu")
    opciones.add_argument("--window-size=1920,1080")

    if binario_chromium:
        opciones.binary_location = binario_chromium

    if ruta_driver_sistema:
        servicio = Service(ruta_driver_sistema)
    else:
        servicio = Service(ChromeDriverManager().install())

    driver = webdriver.Chrome(service=servicio, options=opciones)
    
    try:
        driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": os.path.abspath(carpeta_descargas)
        })
    except Exception:
        pass

    return driver


# -----------------------------------------------------------------------------
# 2. LÓGICA DE PROCESAMIENTO / SCRAPING CON SELENIUM
# -----------------------------------------------------------------------------
def procesar_cola_con_selenium(df_pacientes, carpeta_pdf, area_log):
    driver = None
    try:
        area_log.code("Iniciando navegador Chrome Headless...", language="bash")
        driver = crear_driver(carpeta_pdf)
        
        total = len(df_pacientes)
        for idx, fila in df_pacientes.iterrows():
            cedula = fila.get("Cedula") or fila.get("Cédula")
            area_log.code(f"[{idx+1}/{total}] Procesando Cédula: {cedula}...", language="bash")
            
            # ------------------------------------------------------------------
            # AQUÍ COPIAS TU LÓGICA DE NAVEGACIÓN Y DESCARGA EXISTENTE
            # Ejemplo:
            # driver.get("https://sitio-web-coberturas.com")
            # driver.find_element(...).send_keys(cedula)
            # ------------------------------------------------------------------
            time.sleep(1) # Simulación de descarga
            
            # Actualizar estado
            df_pacientes.at[idx, "Estado"] = "Completado"

        area_log.code("Proceso de Scraping finalizado correctamente.", language="bash")
    except Exception as e:
        area_log.code(f"Error durante el procesamiento: {str(e)}", language="bash")
    finally:
        if driver:
            driver.quit()
            
    return df_pacientes


# -----------------------------------------------------------------------------
# 3. INTERFAZ EN STREAMLIT
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Descarga de Coberturas - CORESALUD", layout="wide")

if "cola_pacientes" not in st.session_state:
    st.session_state.cola_pacientes = []

st.title("📋 Descarga de Coberturas - CORESALUD")

tab_lotes, tab_manual = st.tabs(["📁 Procesamiento por Lotes", "👤 Ingresar Paciente Manual"])

DEPENDENCIAS = [
    "GINECOLOGIA (CE)",
    "MEDICINA FAMILIAR",
    "MEDICINA GENERAL (CE)",
    "OBSTETRICIA (CE)",
    "PEDIATRIA (CE)",
    "PSICOLOGIA (CE)",
    "NUTRICION Y DIETETICA"
]

DIR_DESCARGAS = "PDF_DESCARGADOS"

# ==========================================
# 1. PESTAÑA: POR LOTES
# ==========================================
with tab_lotes:
    col_btn1, col_btn2 = st.columns([2, 2])
    
    archivo_subido = st.file_uploader("Seleccionar matriz Excel (.xlsx)", type=["xlsx"])
    log_lotes = st.empty()

    if archivo_subido:
        df_lote = pd.read_excel(archivo_subido)
        st.dataframe(df_lote, use_container_width=True, height=200)

        with col_btn1:
            if st.button("🚀 Iniciar descarga (por lotes)", type="primary"):
                log_lotes.code("Iniciando lote desde Excel...", language="bash")
                df_resultado = procesar_cola_con_selenium(df_lote, DIR_DESCARGAS, log_lotes)
                
                # Guardar Excel de Salida
                df_resultado.to_excel("Reporte_Lote_Procesado.xlsx", index=False)
                
                # Crear ZIP
                zip_filename = "Resultado_Lote.zip"
                with zipfile.ZipFile(zip_filename, 'w') as zipf:
                    zipf.write("Reporte_Lote_Procesado.xlsx")
                    if os.path.exists(DIR_DESCARGAS):
                        for root, _, files in os.walk(DIR_DESCARGAS):
                            for f in files:
                                zipf.write(os.path.join(root, f))
                
                with open(zip_filename, "rb") as f:
                    st.download_button("📦 Descargar Matriz y PDFs (.ZIP)", data=f, file_name=zip_filename, mime="application/zip")


# ==========================================
# 2. PESTAÑA: MANUAL / INTERACTIVO (CORREGIDO)
# ==========================================
with tab_manual:
    st.subheader("Ingresar pacientes manualmente")
    
    # IMPORTANTE: Desactivamos el limpiado automático 'clear_on_submit=False'
    with st.form(key="form_paciente", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            responsable = st.text_input("Responsable (persona que ingresa la información):")
            cedula = st.text_input("Cédula del paciente (*):")
            fecha_atencion = st.date_input("Fecha de atención (*):", value=datetime.now(), format="DD/MM/YYYY")
            dependencia = st.selectbox("Dependencia (tipo de consulta):", DEPENDENCIAS)
            
        with col2:
            fecha_nacimiento = st.date_input("Fecha de nacimiento (*):", value=None, format="DD/MM/YYYY")
            sexo = st.radio("Sexo (*):", ["M", "F"], horizontal=True)
            observaciones = st.text_input("Observaciones (opcional):")

        # Botón para enviar el formulario
        btn_agregar = st.form_submit_button("➕ Agregar a la cola y seguir con el siguiente", type="primary")

    # Validación y guardado en la cola de la sesión
    if btn_agregar:
        if not cedula.strip() or fecha_nacimiento is None:
            st.error("Por favor completa los campos obligatorios: Cédula y Fecha de nacimiento.")
        else:
            nuevo_paciente = {
                "#": len(st.session_state.cola_pacientes) + 1,
                "Responsable": responsable,
                "Cedula": cedula,
                "Fecha atencion": fecha_atencion.strftime("%d-%m-%Y"),
                "Dependencia": dependencia,
                "Fecha nacimiento": fecha_nacimiento.strftime("%d-%m-%Y"),
                "Sexo": sexo,
                "Observaciones": observaciones,
                "Estado": "Pendiente"
            }
            st.session_state.cola_pacientes.append(nuevo_paciente)
            st.success(f"¡Paciente con Cédula {cedula} agregado correctamente a la cola!")
            st.rerun()  # Actualiza la interfaz inmediatamente
    st.divider()
    st.write(f"**{len(st.session_state.cola_pacientes)} en cola**")

    df_cola = pd.DataFrame(st.session_state.cola_pacientes)
    st.dataframe(df_cola, use_container_width=True, height=180)

    log_manual = st.empty()

    if len(st.session_state.cola_pacientes) > 0:
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            if st.button("🚀 Procesar Cola Manual con Selenium"):
                df_res = procesar_cola_con_selenium(df_cola, DIR_DESCARGAS, log_manual)
                
                # Exportar Matriz
                df_res.to_excel("Matriz_Manual_Procesada.xlsx", index=False)
                
                # Crear ZIP
                zip_filename = "Resultado_Manual.zip"
                with zipfile.ZipFile(zip_filename, 'w') as zipf:
                    zipf.write("Matriz_Manual_Procesada.xlsx")
                    if os.path.exists(DIR_DESCARGAS):
                        for root, _, files in os.walk(DIR_DESCARGAS):
                            for f in files:
                                zipf.write(os.path.join(root, f))
                
                with open(zip_filename, "rb") as f:
                    st.download_button("📦 Descargar Matriz y PDFs (.ZIP)", data=f, file_name=zip_filename, mime="application/zip")
