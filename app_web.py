# -*- coding: utf-8 -*-
"""
app_web.py - Version web (Streamlit) de descargar_coberturas.py

A DIFERENCIA de la version anterior, este archivo NO reimplementa la
logica de scraping: la reutiliza directamente desde descargar_coberturas.py
(aqui importado como "core"), que ya trae:

  - crear_driver()  -> ya detecta Chromium/Chromedriver instalados via
    apt (packages.txt) y ya corre en modo headless (MODO_HEADLESS=True),
    asi que sirve tal cual tanto en escritorio como en la nube.
  - procesar_registro() / procesar_lote() -> la MISMA logica de deteccion
    de GENERAL/CAMPESINO/MENOR/PADRE/MADRE/TITULAR, IESS, etc.
  - localizar_archivo_matriz() / agregar_filas_matriz_con_bloqueo() ->
    el MISMO llenado del Excel INSTRUCTIVO con formulas intactas
    (nada de exportar un DataFrame plano con pandas).

Este archivo SOLO se encarga de la interfaz (Streamlit), de mostrar el
log en vivo, y de empaquetar el resultado (PDFs + matriz) en un .zip
descargable, ya que en un servidor en la nube el disco es temporal y
nunca hay que depender de que el archivo se quede ahi.

REQUISITOS EN EL REPOSITORIO (junto a este archivo y a
descargar_coberturas.py):

  packages.txt:
      chromium
      chromium-driver

  requirements.txt (agrega streamlit a lo que ya tenia el script de
  escritorio):
      streamlit
      selenium
      webdriver-manager
      openpyxl
      pypdf

  Un archivo "INSTRUCTIVO..._....xlsx" (la plantilla real, con las
  formulas de MAESTRO) subido UNA VEZ junto a descargar_coberturas.py.
  Si no esta, esta misma app te deja subirlo desde el navegador.
"""

import os
import io
import zipfile
import contextlib
from datetime import datetime

import streamlit as st

import descargar_coberturas as core

st.set_page_config(page_title="Descarga de Coberturas - CORESALUD", layout="wide")

DEPENDENCIAS_RESPALDO = [
    "GINECOLOGIA (CE)", "MEDICINA FAMILIAR", "MEDICINA GENERAL (CE)",
    "OBSTETRICIA (CE)", "PEDIATRIA (CE)", "PSICOLOGIA (CE)", "NUTRICION Y DIETETICA",
]

# ============================================================
# LOG EN VIVO: redirige los print() reales del motor a la pantalla
# ============================================================


class _EscritorLog(io.TextIOBase):
    """Reemplaza a stdout mientras corre el motor: cada print() de
    descargar_coberturas.py llega aqui en vez de a una consola que
    nadie ve, y se va mostrando en el recuadro de la app."""

    def __init__(self, placeholder, buffer_key):
        self.placeholder = placeholder
        self.buffer_key = buffer_key

    def write(self, texto):
        if texto:
            st.session_state[self.buffer_key] += texto
            # se muestra solo la cola del log para no saturar el navegador
            self.placeholder.code(st.session_state[self.buffer_key][-4000:], language="bash")
        return len(texto)

    def flush(self):
        pass


@contextlib.contextmanager
def log_en_vivo(placeholder, buffer_key):
    st.session_state[buffer_key] = ""
    with contextlib.redirect_stdout(_EscritorLog(placeholder, buffer_key)):
        yield


# ============================================================
# PLANTILLA DEL INSTRUCTIVO (una sola vez, queda junto al motor)
# ============================================================


def _asegurar_plantilla_matriz():
    try:
        core.localizar_archivo_matriz()
        return True
    except FileNotFoundError:
        st.warning(
            "Todavia no hay ninguna plantilla 'INSTRUCTIVO...xlsx' en el servidor. "
            "Subela una sola vez; de ahi en adelante cada mes se genera solo a partir de ella."
        )
        plantilla = st.file_uploader(
            "Subir plantilla INSTRUCTIVO (.xlsx)", type=["xlsx"], key="plantilla_instructivo"
        )
        if plantilla is not None:
            destino = os.path.join(core.BASE_DIR, plantilla.name)
            with open(destino, "wb") as f:
                f.write(plantilla.getbuffer())
            st.success(f"Plantilla guardada como '{plantilla.name}'.")
            st.rerun()
        return False


# ============================================================
# EMPAQUETADO FINAL EN ZIP (PDFs + matriz del mes + reportes)
# ============================================================


def _zip_resultado():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        # PDF_DESCARGADOS/ completo (incluye GENERAL/, CAMPESINO/,
        # log_errores.xlsx y sin_seguro.xlsx, que el motor ya guarda ahi)
        if os.path.isdir(core.CARPETA_SALIDA):
            for raiz, _, archivos in os.walk(core.CARPETA_SALIDA):
                if os.path.basename(raiz) in ("_descargas_temp", "_diagnostico"):
                    continue
                for nombre_archivo in archivos:
                    ruta_completa = os.path.join(raiz, nombre_archivo)
                    ruta_relativa = os.path.relpath(ruta_completa, core.CARPETA_SALIDA)
                    zf.write(ruta_completa, os.path.join("PDF_DESCARGADOS", ruta_relativa))
        # matriz del mes actual (con formulas intactas)
        try:
            ruta_matriz = core.localizar_archivo_matriz()
            zf.write(ruta_matriz, os.path.basename(ruta_matriz))
        except FileNotFoundError:
            pass
    buffer.seek(0)
    return buffer


# ============================================================
# INTERFAZ
# ============================================================

st.title("📋 Descarga de Coberturas - CORESALUD")

tab_lotes, tab_manual = st.tabs(["📁 Procesamiento por lotes", "👤 Ingresar paciente manual"])

# ------------------------------------------------------------
# PESTAÑA 1: POR LOTES (equivalente a main() del script de escritorio)
# ------------------------------------------------------------
with tab_lotes:
    hay_matriz = _asegurar_plantilla_matriz()

    responsable_lote = st.text_input(
        "Responsable (persona que ingresa la informacion):", key="resp_lote"
    )
    archivo_subido = st.file_uploader(
        "Excel de inconsistencias (mismo formato de siempre: Nombre Excel, "
        "Cédula Excel, Fechas Excel, etc.)",
        type=["xlsx"], key="excel_lote",
    )

    log_lotes = st.empty()
    barra_lotes = st.empty()
    zip_lotes = st.empty()

    puede_iniciar = archivo_subido is not None and hay_matriz
    if st.button("🚀 Iniciar descarga por lotes", type="primary", disabled=not puede_iniciar):
        ruta_temp_excel = os.path.join(core.BASE_DIR, "_subida_temp_lote.xlsx")
        with open(ruta_temp_excel, "wb") as f:
            f.write(archivo_subido.getbuffer())

        with log_en_vivo(log_lotes, "log_buffer_lote"):
            registros = core.leer_excel(ruta_temp_excel)
            print(f"Se leyeron {len(registros)} paciente(s) del Excel.\n")

            os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
            carpeta_descargas_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
            os.makedirs(carpeta_descargas_temp, exist_ok=True)
            carpeta_diagnostico = os.path.join(core.CARPETA_SALIDA, "_diagnostico")

            print("Iniciando Chrome (headless)...")
            driver_holder = [core.crear_driver(carpeta_descargas_temp)]

            errores, sin_seguro = [], []
            items = list(enumerate(registros, start=1))
            total = len(items)

            def _progreso(n, total_):
                barra_lotes.progress(n / total_, text=f"{n}/{total_} paciente(s) procesado(s)")

            try:
                core.procesar_lote(
                    items, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                    errores, sin_seguro, callback_progreso=_progreso,
                    escribir_matriz=True, responsable_matriz=responsable_lote,
                )
            finally:
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass

            core._guardar_reportes_finales(errores, sin_seguro)

        st.success("Lote terminado.")
        zip_lotes.download_button(
            "📦 Descargar PDFs + Matriz (.zip)",
            data=_zip_resultado(),
            file_name=f"coberturas_lote_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip",
        )

# ------------------------------------------------------------
# PESTAÑA 2: MANUAL (equivalente a modo_interactivo(), un Chrome
# compartido para todos los pacientes que se vayan agregando)
# ------------------------------------------------------------
with tab_manual:
    hay_matriz = _asegurar_plantilla_matriz()

    if "historial_manual" not in st.session_state:
        st.session_state.historial_manual = []
    if "driver_manual" not in st.session_state:
        st.session_state.driver_manual = None

    with st.form(key="form_paciente", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            responsable = st.text_input("Responsable:")
            cedula = st.text_input("Cédula del paciente (*):")
            fecha_atencion = st.date_input(
                "Fecha de atención (*):", value=datetime.now(), format="DD/MM/YYYY"
            )
            try:
                dependencias_validas = core.obtener_dependencias_validas(
                    core.abrir_matriz(core.localizar_archivo_matriz())
                ) if hay_matriz else DEPENDENCIAS_RESPALDO
            except Exception:
                dependencias_validas = DEPENDENCIAS_RESPALDO
            dependencia = st.selectbox("Dependencia:", dependencias_validas)
        with col2:
            fecha_nacimiento = st.date_input(
                "Fecha de nacimiento (*):", value=None, format="DD/MM/YYYY"
            )
            sexo = st.radio("Sexo (*):", ["M", "F"], horizontal=True)
            observaciones = st.text_input("Observaciones (opcional):")

        enviar = st.form_submit_button(
            "➕ Procesar y agregar a la matriz", type="primary", disabled=not hay_matriz
        )

    log_manual = st.empty()

    if enviar:
        if not cedula or fecha_nacimiento is None:
            st.error("Completa Cédula y Fecha de nacimiento.")
        else:
            reg = {
                "cedula": cedula.strip().zfill(10),
                "nombre": None,
                "fechas": [fecha_atencion],
                "cedula_padre": None,
                "cedula_titular": None,
            }

            with log_en_vivo(log_manual, "log_buffer_manual"):
                os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                carpeta_descargas_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                os.makedirs(carpeta_descargas_temp, exist_ok=True)
                carpeta_diagnostico = os.path.join(core.CARPETA_SALIDA, "_diagnostico")

                if st.session_state.driver_manual is None:
                    print("Iniciando Chrome (headless), se mantiene abierto para los "
                          "siguientes pacientes que agregues...")
                    st.session_state.driver_manual = [core.crear_driver(carpeta_descargas_temp)]

                errores, sin_seguro = [], []
                filas_matriz = core.procesar_registro(
                    reg, 1, 1, st.session_state.driver_manual,
                    carpeta_descargas_temp, carpeta_diagnostico,
                    errores, sin_seguro, recolectar_matriz=True,
                )

                estado = "Pendiente"
                if filas_matriz:
                    try:
                        filas_escritas = core.agregar_filas_matriz_con_bloqueo(
                            filas_matriz, dependencia, fecha_nacimiento, sexo,
                            observaciones, responsable,
                        )
                        estado = f"OK ({len(filas_escritas)} fila(s) en la matriz)"
                    except TimeoutError as e:
                        estado = f"ERROR matriz: {e}"
                elif errores:
                    estado = f"ERROR: {errores[-1][-1]}"
                elif sin_seguro:
                    estado = "SIN SEGURO"

            st.session_state.historial_manual.append({
                "Cédula": reg["cedula"], "Fecha": fecha_atencion.strftime("%d-%m-%Y"),
                "Dependencia": dependencia, "Estado": estado,
            })

    if st.session_state.historial_manual:
        st.divider()
        st.write(f"**{len(st.session_state.historial_manual)} paciente(s) procesado(s) en esta sesión**")
        st.dataframe(st.session_state.historial_manual, use_container_width=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.download_button(
                "📦 Descargar PDFs + Matriz (.zip)",
                data=_zip_resultado(),
                file_name=f"coberturas_manual_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip",
            )
        with col_b:
            if st.button("🔴 Cerrar Chrome y terminar sesión"):
                if st.session_state.driver_manual is not None:
                    try:
                        st.session_state.driver_manual[0].quit()
                    except Exception:
                        pass
                    st.session_state.driver_manual = None
                st.success("Chrome cerrado. Puedes seguir descargando el ZIP arriba cuando quieras.")
