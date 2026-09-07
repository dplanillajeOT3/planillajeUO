import os
import re
import json
from datetime import datetime, date

import streamlit as st
import openpyxl
from openpyxl.formula.translate import Translator

# Importar el motor original de descargas
import descargar_coberturas as motor

# Configuración inicial de la página
st.set_page_config(
    page_title="Gestión de Coberturas de Salud",
    page_icon="📋",
    layout="wide"
)

# Estilos visuales sencillos
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; }
    </style>
""", unsafe_allow_html=True)

st.title("📋 Sistema de Gestión y Descarga de Coberturas")

# ============================================================
# FUNCIONES AUXILIARES PARA MANEJO DE MATRIZ Y DATOS
# ============================================================

def cargar_datos_pacientes():
    if os.path.exists(motor.RUTA_DATOS_PACIENTES):
        try:
            with open(motor.RUTA_DATOS_PACIENTES, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def guardar_datos_pacientes(datos):
    try:
        with open(motor.RUTA_DATOS_PACIENTES, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Error guardando base de pacientes: {e}")

def buscar_siguiente_fila_vacia(ws):
    fila = 2
    while True:
        val = ws.cell(row=fila, column=motor.COL_MATRIZ_CEDULA).value
        if val is None or str(val).strip() == "":
            return fila
        fila += 1

def copiar_formulas_si_es_necesario(ws, fila_destino):
    fila_origen = fila_destino - 1
    if fila_origen < 2:
        return
    for col in motor._COLUMNAS_FORMULA_MATRIZ:
        celda_origen = ws.cell(row=fila_origen, column=col)
        celda_destino = ws.cell(row=fila_destino, column=col)
        if celda_origen.value and str(celda_origen.value).startswith("="):
            formula_orig = str(celda_origen.value)
            try:
                formula_trad = Translator(formula_orig, origin=celda_origen.coordinate).translate_formula(celda_destino.coordinate)
                celda_destino.value = formula_trad
            except Exception:
                celda_destino.value = formula_orig

def insertar_en_matriz_excel(info_fila, f_nacimiento, sexo_letra, observacion=""):
    os.makedirs(motor.CARPETA_RESPALDOS_MATRIZ, exist_ok=True)
    if os.path.exists(motor.ARCHIVO_MATRIZ):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = os.path.join(motor.CARPETA_RESPALDOS_MATRIZ, f"INSTRUCTIVO_{stamp}.xlsx")
        try:
            import shutil
            shutil.copy2(motor.ARCHIVO_MATRIZ, backup)
        except Exception:
            pass

    wb = openpyxl.load_workbook(motor.ARCHIVO_MATRIZ)
    ws = wb[motor.HOJA_MATRIZ]

    fila = buscar_siguiente_fila_vacia(ws)
    copiar_formulas_si_es_necesario(ws, fila)

    fecha_atencion = info_fila.get("fecha_atencion")
    cedula_paciente = info_fila.get("cedula_paciente")
    nombre_paciente = info_fila.get("nombre_paciente")
    institucion = info_fila.get("institucion", "IESS")
    tipo_seguro = info_fila.get("tipo_seguro_beneficiario", motor.TIPO_SEGURO_GENERAL_TITULAR)
    parentesco = info_fila.get("parentesco", motor.PARENTESCO_TITULAR)
    cedula_afiliado = info_fila.get("cedula_afiliado", cedula_paciente)
    nombre_afiliado = info_fila.get("nombre_afiliado", nombre_paciente)

    ws.cell(row=fila, column=1, value=fecha_atencion.strftime("%Y-%m-%d") if isinstance(fecha_atencion, (date, datetime)) else fecha_atencion)
    ws.cell(row=fila, column=2, value=institucion)
    ws.cell(row=fila, column=5, value="10")
    ws.cell(row=fila, column=6, value=tipo_seguro)
    ws.cell(row=fila, column=8, value=cedula_paciente)
    ws.cell(row=fila, column=9, value=nombre_paciente)
    ws.cell(row=fila, column=10, value=f_nacimiento.strftime("%Y-%m-%d") if isinstance(f_nacimiento, (date, datetime)) else f_nacimiento)
    ws.cell(row=fila, column=11, value="FEMENINO" if sexo_letra == "F" else "MASCULINO")
    ws.cell(row=fila, column=13, value=cedula_afiliado)
    ws.cell(row=fila, column=14, value=nombre_afiliado)
    ws.cell(row=fila, column=17, value=parentesco)
    ws.cell(row=fila, column=19, value=observacion)

    wb.save(motor.ARCHIVO_MATRIZ)
    return fila

# ============================================================
# INTERFAZ EN STREAMLIT (COLUMNAS LOTES Y MANUAL)
# ============================================================

col_lotes, col_manual = st.columns(2)

# --- PANEL IZQUIERDO: DESCARGA POR LOTES ---
with col_lotes:
    st.subheader("📂 MODO POR LOTES (Excel)")
    
    archivo_excel = st.file_uploader("Cargar archivo de reporte (.xlsx)", type=["xlsx"])

    if archivo_excel:
        with open(motor.ARCHIVO_EXCEL, "wb") as f:
            f.write(archivo_excel.getbuffer())
        st.success("Archivo subido y listo para ser procesado.")

    if st.button("🚀 Iniciar Descarga por Lotes", key="btn_lotes"):
        if not os.path.exists(motor.ARCHIVO_EXCEL):
            st.error("Por favor sube un archivo Excel primero.")
        else:
            registros = motor.leer_excel(motor.ARCHIVO_EXCEL)
            barra_progreso = st.progress(0)
            texto_estado = st.empty()

            os.makedirs(motor.CARPETA_SALIDA, exist_ok=True)
            carpeta_temp = os.path.join(motor.CARPETA_SALIDA, "_descargas_temp")
            os.makedirs(carpeta_temp, exist_ok=True)
            carpeta_diag = os.path.join(motor.CARPETA_SALIDA, "_diagnostico")

            driver_holder = [motor.crear_driver(carpeta_temp)]
            errores, sin_seguro = [], []

            try:
                total = len(registros)
                for idx, reg in enumerate(registros, start=1):
                    texto_estado.info(f"Procesando {idx}/{total}: {reg['nombre']} (C.I: {reg['cedula']})")
                    motor.procesar_registro(
                        reg, idx, total, driver_holder, carpeta_temp, carpeta_diag,
                        errores, sin_seguro
                    )
                    barra_progreso.progress(idx / total)

                motor._guardar_reportes_finales(errores, sin_seguro)
                st.success("¡Proceso por lotes finalizado exitosamente!")
            except Exception as e:
                st.error(f"Error procesando el lote: {e}")
            finally:
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass

# --- PANEL DERECHO: MODO MANUAL E INSERCIÓN ---
with col_manual:
    st.subheader("👤 MODO MANUAL / INTERACTIVO")

    with st.form("form_registro_manual"):
        cedula_in = st.text_input("Cédula Paciente (*)")
        f_atencion_in = st.date_input("Fecha de Atención (*)", value=date.today())
        f_nacimiento_in = st.date_input("Fecha de Nacimiento (*)")
        sexo_in = st.radio("Sexo (*)", ["M", "F"], horizontal=True)
        obs_in = st.text_input("Observación")

        btn_enviar_manual = st.form_submit_button("⚡ Descargar e Insertar en Matriz")

    if btn_enviar_manual:
        cedula_clean = re.sub(r"\D", "", cedula_in).zfill(10)
        if len(cedula_clean) != 10:
            st.error("La cédula ingresada debe contener 10 dígitos.")
        else:
            st.info(f"Procesando paciente C.I: {cedula_clean}...")
            
            # Guardar/Actualizar base local de pacientes
            pacientes = cargar_datos_pacientes()
            pacientes[cedula_clean] = {
                "fecha_nacimiento": f_nacimiento_in.strftime("%d-%m-%Y"),
                "sexo": sexo_in
            }
            guardar_datos_pacientes(pacientes)

            reg = {
                "cedula": cedula_clean,
                "nombre": None,
                "fechas": [f_atencion_in],
                "cedula_padre": None,
                "cedula_titular": None
            }

            os.makedirs(motor.CARPETA_SALIDA, exist_ok=True)
            carpeta_temp = os.path.join(motor.CARPETA_SALIDA, "_descargas_temp")
            os.makedirs(carpeta_temp, exist_ok=True)
            carpeta_diag = os.path.join(motor.CARPETA_SALIDA, "_diagnostico")

            driver_holder = [motor.crear_driver(carpeta_temp)]
            errores, sin_seguro = [], []

            try:
                filas_matriz = motor.procesar_registro(
                    reg, 1, 1, driver_holder, carpeta_temp, carpeta_diag,
                    errores, sin_seguro, recolectar_matriz=True
                )

                if sin_seguro:
                    st.warning("El paciente no registra seguro activo.")
                elif errores or not filas_matriz:
                    st.error("Error al consultar o procesar la cobertura.")
                else:
                    info_fila = filas_matriz[0]
                    fila_excel = insertar_en_matriz_excel(info_fila, f_nacimiento_in, sexo_in, obs_in)
                    st.success(f"¡Éxito! Registro ingresado correctamente en la Matriz Excel (Fila {fila_excel}).")

            except Exception as e:
                st.error(f"Error procesando registro: {e}")
            finally:
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass
