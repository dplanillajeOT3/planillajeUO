# -*- coding: utf-8 -*-
"""
app_web.py - Version web (Streamlit) de descargar_coberturas.py

Reutiliza el motor real (descargar_coberturas.py, aqui "core") en vez de
reimplementar la logica de scraping/matriz:

  - core.crear_driver()            -> Chrome headless (cloud-aware)
  - core.procesar_registro/lote()  -> deteccion GENERAL/CAMPESINO/etc.
  - core.localizar_archivo_matriz / core.agregar_filas_matriz_con_bloqueo
        -> llenado real del INSTRUCTIVO con formulas intactas

Este archivo se encarga solo de la interfaz, el log en vivo, el
reintento de fallidos y el empaquetado final en .zip.

PENDIENTE (fuera de este archivo, ver conversacion): integracion con
Google Drive por Unidad/Año/Mes y clasificacion IESS/CAMPESINO/ISSFA/
ISSPOL. Mientras tanto, la app guarda en el disco temporal del
contenedor y todo se descarga como .zip al terminar.
"""

import os
import io
import zipfile
import contextlib
from datetime import datetime

import streamlit as st

import descargar_coberturas as core

st.set_page_config(page_title="Coberturas CORESALUD", layout="wide", page_icon="📋")

DEPENDENCIAS_RESPALDO = [
    "GINECOLOGIA (CE)", "MEDICINA FAMILIAR", "MEDICINA GENERAL (CE)",
    "OBSTETRICIA (CE)", "PEDIATRIA (CE)", "PSICOLOGIA (CE)", "NUTRICION Y DIETETICA",
]

# ------------------------------------------------------------
# ESTETICA: interfaz mas compacta (punto 6)
# ------------------------------------------------------------
st.markdown("""
<style>
div.block-container{padding-top:1.2rem; padding-bottom:1rem; max-width:1150px;}
h1{font-size:1.6rem !important; margin-bottom:0.6rem !important;}
h2, h3{font-size:1.05rem !important;}
label, .stMarkdown p{font-size:0.85rem !important;}
div[data-testid="stTextInput"] input, div[data-testid="stDateInput"] input,
div[data-testid="stSelectbox"] div, div[data-testid="stNumberInput"] input{
    font-size:0.85rem !important; padding:0.35rem 0.5rem !important;
}
button{font-size:0.82rem !important; padding:0.3rem 0.9rem !important;}
div[data-testid="stDataFrame"]{font-size:0.8rem !important;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# LOG EN VIVO: redirige los print() reales del motor a la pantalla
# ============================================================


class _EscritorLog(io.TextIOBase):
    def __init__(self, placeholder, buffer_key):
        self.placeholder = placeholder
        self.buffer_key = buffer_key

    def write(self, texto):
        if texto:
            st.session_state[self.buffer_key] += texto
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
# PLANTILLA DEL INSTRUCTIVO
# ============================================================


def _asegurar_plantilla_matriz():
    try:
        core.localizar_archivo_matriz()
        return True
    except FileNotFoundError:
        st.warning(
            "No hay ninguna plantilla 'INSTRUCTIVO...xlsx' en el servidor. "
            "Subela una sola vez; los demas meses se generan solos a partir de ella."
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
# EMPAQUETADO FINAL EN ZIP
# ============================================================


def _zip_resultado():
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.isdir(core.CARPETA_SALIDA):
            for raiz, _, archivos in os.walk(core.CARPETA_SALIDA):
                if os.path.basename(raiz) in ("_descargas_temp", "_diagnostico"):
                    continue
                for nombre_archivo in archivos:
                    ruta_completa = os.path.join(raiz, nombre_archivo)
                    ruta_relativa = os.path.relpath(ruta_completa, core.CARPETA_SALIDA)
                    zf.write(ruta_completa, os.path.join("PDF_DESCARGADOS", ruta_relativa))
        try:
            ruta_matriz = core.localizar_archivo_matriz()
            zf.write(ruta_matriz, os.path.basename(ruta_matriz))
        except FileNotFoundError:
            pass
    buffer.seek(0)
    return buffer


def _obtener_driver(carpeta_descargas_temp):
    """Un solo Chrome por sesion de navegador (se reutiliza entre envios
    del formulario manual y entre reintentos), para no acumular varios
    procesos de Chrome en memoria (punto 8)."""
    if st.session_state.get("driver_compartido") is None:
        print("Iniciando Chrome (headless)...")
        st.session_state.driver_compartido = [core.crear_driver(carpeta_descargas_temp)]
    return st.session_state.driver_compartido


def _cerrar_driver():
    driver_holder = st.session_state.get("driver_compartido")
    if driver_holder is not None:
        try:
            driver_holder[0].quit()
        except Exception:
            pass
        st.session_state.driver_compartido = None


def _procesar_un_registro(driver_holder, carpeta_descargas_temp, carpeta_diagnostico, reg, datos_matriz):
    """Procesa un paciente (se usa tanto al ingresarlo por primera vez
    como al reintentarlo despues) y devuelve el texto de estado."""
    errores, sin_seguro = [], []
    filas_matriz = core.procesar_registro(
        reg, 1, 1, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
        errores, sin_seguro, recolectar_matriz=True,
    )
    if filas_matriz:
        try:
            filas = core.agregar_filas_matriz_con_bloqueo(
                filas_matriz, datos_matriz["dependencia"], datos_matriz["fecha_nacimiento"],
                datos_matriz["sexo"], datos_matriz["observaciones"], datos_matriz["responsable"],
            )
            return f"OK ({len(filas)} fila(s))"
        except TimeoutError as e:
            return f"ERROR matriz: {e}"
    elif errores:
        return f"ERROR: {errores[-1][-1]}"
    elif sin_seguro:
        return "SIN SEGURO"
    return "ERROR: sin resultado"


# ============================================================
# INTERFAZ
# ============================================================

st.title("📋 Descarga de Coberturas - CORESALUD")

for clave, valor in {
    "historial_manual": [], "fallidos_lote": [], "errores_lote_acum": [],
    "sin_seguro_lote_acum": [], "driver_compartido": None,
}.items():
    if clave not in st.session_state:
        st.session_state[clave] = valor

tab_lotes, tab_manual = st.tabs(["📁 Por lotes", "👤 Paciente manual"])

# ------------------------------------------------------------
# PESTAÑA 1: POR LOTES
# ------------------------------------------------------------
with tab_lotes:
    hay_matriz = _asegurar_plantilla_matriz()

    c1, c2 = st.columns([1, 2])
    with c1:
        responsable_lote = st.text_input("Responsable:", key="resp_lote")
    with c2:
        archivo_subido = st.file_uploader(
            "Excel de inconsistencias (.xlsx)", type=["xlsx"], key="excel_lote"
        )

    log_lotes = st.empty()
    barra_lotes = st.empty()

    def _correr_lote(items, titulo_boton_zip):
        os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
        carpeta_descargas_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
        os.makedirs(carpeta_descargas_temp, exist_ok=True)
        carpeta_diagnostico = os.path.join(core.CARPETA_SALIDA, "_diagnostico")

        with log_en_vivo(log_lotes, "log_buffer_lote"):
            driver_holder = _obtener_driver(carpeta_descargas_temp)
            errores, sin_seguro = [], []
            total = len(items)

            def _progreso(n, total_):
                barra_lotes.progress(n / total_, text=f"{n}/{total_} procesado(s)")

            core.procesar_lote(
                items, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                errores, sin_seguro, callback_progreso=_progreso,
                escribir_matriz=True, responsable_matriz=responsable_lote,
            )
            core._guardar_reportes_finales(errores, sin_seguro)

        # cedulas que fallaron -> para reintentar SOLO esas, con el mismo
        # registro (punto 4), sin volver a leerlas como si fueran nuevas
        cedulas_con_error = {fila[1] for fila in errores}
        st.session_state.fallidos_lote = [
            (indice, reg) for indice, reg in items if reg["cedula"] in cedulas_con_error
        ]
        st.session_state.errores_lote_acum += errores
        st.session_state.sin_seguro_lote_acum += sin_seguro

        if st.session_state.fallidos_lote:
            st.warning(f"{len(st.session_state.fallidos_lote)} paciente(s) quedaron con error.")
        else:
            st.success("Lote terminado sin errores pendientes.")

        st.download_button(
            titulo_boton_zip, data=_zip_resultado(),
            file_name=f"coberturas_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip", key=f"zip_{datetime.now().timestamp()}",
        )

    puede_iniciar = archivo_subido is not None and hay_matriz
    if st.button("🚀 Iniciar descarga por lotes", type="primary", disabled=not puede_iniciar):
        ruta_temp_excel = os.path.join(core.BASE_DIR, "_subida_temp_lote.xlsx")
        with open(ruta_temp_excel, "wb") as f:
            f.write(archivo_subido.getbuffer())
        registros = core.leer_excel(ruta_temp_excel)
        items = list(enumerate(registros, start=1))
        st.session_state.errores_lote_acum = []
        st.session_state.sin_seguro_lote_acum = []
        _correr_lote(items, "📦 Descargar PDFs + Matriz (.zip)")

    # --- Reintento masivo de fallidos (punto 4) ---
    if st.session_state.fallidos_lote:
        st.divider()
        st.write(f"⚠️ **{len(st.session_state.fallidos_lote)} paciente(s) pendiente(s) de reintentar** "
                 f"(mismos registros, no se leen de nuevo del Excel):")
        st.dataframe(
            [{"Cédula": r["cedula"], "Nombre": r["nombre"]} for _, r in st.session_state.fallidos_lote],
            use_container_width=True, height=150,
        )
        if st.button("🔁 Reintentar todos los fallidos", type="primary"):
            _correr_lote(st.session_state.fallidos_lote, "📦 Descargar resultado del reintento (.zip)")

# ------------------------------------------------------------
# PESTAÑA 2: MANUAL
# ------------------------------------------------------------
with tab_manual:
    hay_matriz = _asegurar_plantilla_matriz()

    try:
        dependencias_validas = core.obtener_dependencias_validas(
            core.abrir_matriz(core.localizar_archivo_matriz())
        ) if hay_matriz else DEPENDENCIAS_RESPALDO
    except Exception:
        dependencias_validas = DEPENDENCIAS_RESPALDO

    col1, col2, col3 = st.columns(3)
    with col1:
        responsable = st.text_input("Responsable:", key="campo_responsable")
        cedula = st.text_input("Cédula (*):", key="campo_cedula")
    with col2:
        fecha_atencion = st.date_input(
            "Fecha de atención (*):", value=datetime.now(), format="DD/MM/YYYY", key="campo_fecha_atencion"
        )
        fecha_nacimiento = st.date_input(
            "Fecha de nacimiento (*):", value=None, format="DD/MM/YYYY", key="campo_fecha_nac"
        )
    with col3:
        dependencia = st.selectbox("Dependencia:", dependencias_validas, key="campo_dependencia")
        sexo = st.radio("Sexo (*):", ["M", "F"], horizontal=True, key="campo_sexo")

    observaciones = st.text_input("Observaciones:", key="campo_observaciones")

    log_manual = st.empty()

    if st.button("➕ Procesar y agregar a la matriz", type="primary", disabled=not hay_matriz):
        if not cedula or fecha_nacimiento is None:
            st.error("Completa Cédula y Fecha de nacimiento.")
        else:
            reg = {
                "cedula": cedula.strip().zfill(10), "nombre": None,
                "fechas": [fecha_atencion], "cedula_padre": None, "cedula_titular": None,
            }
            datos_matriz = {
                "dependencia": dependencia, "fecha_nacimiento": fecha_nacimiento,
                "sexo": sexo, "observaciones": observaciones, "responsable": responsable,
            }

            with log_en_vivo(log_manual, "log_buffer_manual"):
                os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                carpeta_descargas_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                os.makedirs(carpeta_descargas_temp, exist_ok=True)
                carpeta_diagnostico = os.path.join(core.CARPETA_SALIDA, "_diagnostico")
                driver_holder = _obtener_driver(carpeta_descargas_temp)

                estado = _procesar_un_registro(
                    driver_holder, carpeta_descargas_temp, carpeta_diagnostico, reg, datos_matriz
                )

            fila_historial = {
                "Cédula": reg["cedula"], "Fecha": fecha_atencion.strftime("%d/%m/%Y"), "Estado": estado,
            }
            if estado.startswith("ERROR") or estado == "SIN SEGURO":
                fila_historial["_reg"] = reg
                fila_historial["_datos"] = datos_matriz
            st.session_state.historial_manual.append(fila_historial)

            # limpia SOLO los campos del paciente; responsable y fecha de
            # atencion (que vuelve a quedar en "hoy") se mantienen (punto 3)
            for k in ("campo_cedula", "campo_fecha_nac", "campo_sexo", "campo_observaciones", "campo_dependencia"):
                st.session_state.pop(k, None)
            st.rerun()

    if st.session_state.historial_manual:
        st.divider()
        fallidos = [h for h in st.session_state.historial_manual if "_reg" in h]
        st.write(f"**{len(st.session_state.historial_manual)} procesado(s) en esta sesión** "
                 f"({len(fallidos)} con error)")
        st.dataframe(
            [{"Cédula": h["Cédula"], "Fecha": h["Fecha"], "Estado": h["Estado"]}
             for h in st.session_state.historial_manual],
            use_container_width=True, height=200,
        )

        cA, cB, cC = st.columns(3)
        with cA:
            st.download_button(
                "📦 Descargar PDFs + Matriz (.zip)", data=_zip_resultado(),
                file_name=f"coberturas_manual_{datetime.now():%Y%m%d_%H%M}.zip", mime="application/zip",
            )
        with cB:
            if fallidos and st.button(f"🔁 Reintentar los {len(fallidos)} fallidos"):
                os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                carpeta_descargas_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                os.makedirs(carpeta_descargas_temp, exist_ok=True)
                carpeta_diagnostico = os.path.join(core.CARPETA_SALIDA, "_diagnostico")
                driver_holder = _obtener_driver(carpeta_descargas_temp)

                with log_en_vivo(log_manual, "log_buffer_manual"):
                    for h in fallidos:
                        estado = _procesar_un_registro(
                            driver_holder, carpeta_descargas_temp, carpeta_diagnostico, h["_reg"], h["_datos"]
                        )
                        h["Estado"] = estado
                        if not (estado.startswith("ERROR") or estado == "SIN SEGURO"):
                            h.pop("_reg", None)
                            h.pop("_datos", None)
                st.rerun()
        with cC:
            if st.button("🔴 Cerrar Chrome"):
                _cerrar_driver()
                st.success("Chrome cerrado.")
