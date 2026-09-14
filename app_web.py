# -*- coding: utf-8 -*-
"""
app_web.py - Version web (Streamlit) de descargar_coberturas.py, con
login por unidad y sincronizacion con Google Drive.

Reutiliza el motor real (descargar_coberturas.py, aqui "core"):
  - core.crear_driver() / procesar_registro() / procesar_lote()
  - core.localizar_archivo_matriz() / agregar_filas_matriz_con_bloqueo()

Y drive_utils.py para:
  - Login por unidad (st.secrets["usuarios"])
  - Bajar el INSTRUCTIVO del mes actual desde Drive al iniciar sesion
  - Subir la matriz + clasificar y subir cada PDF a IESS/CAMPESINO/
    ISSFA/ISSPOL dentro de PLANILLAJE COBERTURAS/<unidad>/<año>/<mes>

REQUISITOS (junto a este archivo, descargar_coberturas.py y drive_utils.py):

  packages.txt:
      chromium
      chromium-driver

  requirements.txt:
      streamlit>=1.31
      selenium
      webdriver-manager
      openpyxl
      pypdf
      google-api-python-client
      google-auth

  Secrets de la app (Settings -> Secrets en Streamlit Cloud):
      [usuarios]  -> ya lo tienes (clave = nombre de unidad, password, nombre)
      [gcp_service_account] -> la cuenta de servicio (rotada tras el
      incidente de esta conversacion, la version anterior quedo expuesta)
"""

import os
import io
import zipfile
import contextlib
from datetime import datetime, date

import streamlit as st

import descargar_coberturas as core
import drive_utils as drv

st.set_page_config(page_title="Coberturas CORESALUD", layout="wide", page_icon="📋")

DEPENDENCIAS_RESPALDO = [
    "GINECOLOGIA (CE)", "MEDICINA FAMILIAR", "MEDICINA GENERAL (CE)",
    "OBSTETRICIA (CE)", "PEDIATRIA (CE)", "PSICOLOGIA (CE)", "NUTRICION Y DIETETICA",
]

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
# LOGIN POR UNIDAD
# ============================================================


def _pantalla_login():
    st.title("📋 Descarga de Coberturas - CORESALUD")
    st.subheader("Iniciar sesión")
    usuarios = dict(st.secrets.get("usuarios", {}))
    if not usuarios:
        st.error("No hay unidades configuradas en los Secrets de la app (sección [usuarios]).")
        return

    claves = list(usuarios.keys())
    clave = st.selectbox("Unidad:", claves, format_func=lambda k: usuarios[k]["nombre"])
    contrasena = st.text_input("Contraseña:", type="password", key="campo_password_login")

    if st.button("Ingresar", type="primary"):
        if contrasena == usuarios[clave]["password"]:
            st.session_state.unidad = {"clave": clave, "nombre": usuarios[clave]["nombre"]}
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")


# ============================================================
# LOG EN VIVO
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
# GOOGLE DRIVE: conexion, matriz del mes, sincronizacion
# ============================================================


def _servicio_drive():
    if "gcp_service_account" not in st.secrets:
        return None
    try:
        return drv.obtener_servicio_drive(st.secrets["gcp_service_account"])
    except Exception as e:
        st.error(f"No se pudo conectar con Google Drive: {e}")
        return None


def _carpeta_mes_actual(servicio, fecha=None):
    drive_id = drv.id_unidad_compartida(servicio)
    nombre_unidad = st.session_state.unidad["nombre"]
    id_carpeta_mes = drv.carpeta_de_unidad_anio_mes(servicio, drive_id, nombre_unidad, fecha or date.today())
    ids_tipo = drv.asegurar_subcarpetas_tipo_seguro(servicio, drive_id, id_carpeta_mes)
    return drive_id, id_carpeta_mes, ids_tipo


def _asegurar_matriz_desde_drive():
    """Antes de procesar cualquier paciente, se asegura de tener local
    (en core.BASE_DIR) el INSTRUCTIVO del mes actual de ESTA unidad,
    bajandolo de Drive si hace falta. Se hace una sola vez por sesion."""
    if st.session_state.get("matriz_lista"):
        return True

    servicio = _servicio_drive()
    if servicio is None:
        return False

    try:
        drive_id, id_carpeta_mes, _ = _carpeta_mes_actual(servicio)
        archivo = drv.buscar_archivo_por_patron(servicio, drive_id, id_carpeta_mes, r"^INSTRUCTIVO.*\.xlsx$")
        if archivo is not None:
            destino = os.path.join(core.BASE_DIR, archivo["name"])
            drv.descargar_archivo(servicio, archivo["id"], destino)
            st.session_state.matriz_lista = True
            return True
    except Exception as e:
        st.warning(f"No se pudo revisar la matriz en Drive: {e}")

    # Todavia no hay ningun INSTRUCTIVO para esta unidad/mes en Drive:
    # se deja subir uno una sola vez (queda local y se sube a Drive de una vez)
    st.warning(
        f"No hay ninguna plantilla 'INSTRUCTIVO...xlsx' en Drive para "
        f"{st.session_state.unidad['nombre']} este mes. Subela una sola vez."
    )
    plantilla = st.file_uploader("Subir plantilla INSTRUCTIVO (.xlsx)", type=["xlsx"], key="plantilla_instructivo")
    if plantilla is not None:
        destino = os.path.join(core.BASE_DIR, plantilla.name)
        with open(destino, "wb") as f:
            f.write(plantilla.getbuffer())
        try:
            drive_id, id_carpeta_mes, _ = _carpeta_mes_actual(servicio)
            drv.subir_o_reemplazar_archivo(servicio, destino, plantilla.name, id_carpeta_mes, drv.MIME_XLSX)
            st.success("Plantilla guardada y subida a Drive.")
        except Exception as e:
            st.warning(f"Se guardó localmente pero no se pudo subir a Drive todavía: {e}")
        st.session_state.matriz_lista = True
        st.rerun()
    return False


def _sincronizar_con_drive():
    servicio = _servicio_drive()
    if servicio is None:
        st.error("No hay credenciales de Google Drive configuradas (gcp_service_account).")
        return

    with st.spinner("Sincronizando con Drive..."):
        try:
            drive_id, id_carpeta_mes, ids_tipo = _carpeta_mes_actual(servicio)
        except Exception as e:
            st.error(f"No se pudo ubicar la carpeta en Drive: {e}")
            return

        try:
            ruta_matriz = core.localizar_archivo_matriz()
            drv.subir_o_reemplazar_archivo(
                servicio, ruta_matriz, os.path.basename(ruta_matriz), id_carpeta_mes, drv.MIME_XLSX
            )
        except FileNotFoundError:
            pass

        subidos, sin_clasificar = 0, []
        if os.path.isdir(core.CARPETA_SALIDA):
            for raiz, _, archivos in os.walk(core.CARPETA_SALIDA):
                if os.path.basename(raiz) in ("_descargas_temp", "_diagnostico"):
                    continue
                for nombre_archivo in archivos:
                    if not nombre_archivo.lower().endswith(".pdf"):
                        continue
                    ruta_completa = os.path.join(raiz, nombre_archivo)
                    with open(ruta_completa, "rb") as f:
                        contenido = f.read()
                    texto = core.texto_pdf(contenido)
                    esquema = drv.determinar_esquema_cobertura(texto, core.es_campesino)
                    if esquema is None:
                        sin_clasificar.append(nombre_archivo)
                        continue
                    ruta_relativa = os.path.relpath(ruta_completa, core.CARPETA_SALIDA)
                    nombre_drive = ruta_relativa.replace(os.sep, " - ")
                    drv.subir_o_reemplazar_archivo(
                        servicio, ruta_completa, nombre_drive, ids_tipo[esquema], "application/pdf"
                    )
                    subidos += 1

    mensaje = f"✅ Sincronizado: {subidos} PDF(s) subidos a IESS/CAMPESINO/ISSFA/ISSPOL y matriz actualizada en Drive."
    st.success(mensaje)
    if sin_clasificar:
        st.warning(
            f"{len(sin_clasificar)} PDF(s) no se pudieron clasificar (ninguna fila de la tabla "
            f"decía 'SI REGISTRA COBERTURA'; revísalos a mano): " + ", ".join(sin_clasificar)
        )


# ============================================================
# UTILIDADES DE PROCESAMIENTO (compartidas por lotes y manual)
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

if "unidad" not in st.session_state:
    _pantalla_login()
    st.stop()

for clave, valor in {
    "historial_manual": [], "fallidos_lote": [], "errores_lote_acum": [],
    "sin_seguro_lote_acum": [], "driver_compartido": None, "matriz_lista": False,
}.items():
    if clave not in st.session_state:
        st.session_state[clave] = valor

col_titulo, col_sesion = st.columns([3, 1])
with col_titulo:
    st.title("📋 Descarga de Coberturas - CORESALUD")
with col_sesion:
    st.write("")
    st.caption(f"Sesión: **{st.session_state.unidad['nombre']}**")
    if st.button("🔴 Cerrar sesión"):
        _cerrar_driver()
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

hay_matriz = _asegurar_matriz_desde_drive()
if not hay_matriz:
    st.stop()

tab_lotes, tab_manual = st.tabs(["📁 Por lotes", "👤 Paciente manual"])

# ------------------------------------------------------------
# PESTAÑA 1: POR LOTES
# ------------------------------------------------------------
with tab_lotes:
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

        colz1, colz2 = st.columns(2)
        with colz1:
            st.download_button(
                titulo_boton_zip, data=_zip_resultado(),
                file_name=f"coberturas_{datetime.now():%Y%m%d_%H%M}.zip",
                mime="application/zip", key=f"zip_{datetime.now().timestamp()}",
            )
        with colz2:
            if st.button("☁️ Sincronizar con Drive", key=f"sync_{datetime.now().timestamp()}"):
                _sincronizar_con_drive()

    puede_iniciar = archivo_subido is not None
    if st.button("🚀 Iniciar descarga por lotes", type="primary", disabled=not puede_iniciar):
        ruta_temp_excel = os.path.join(core.BASE_DIR, "_subida_temp_lote.xlsx")
        with open(ruta_temp_excel, "wb") as f:
            f.write(archivo_subido.getbuffer())
        registros = core.leer_excel(ruta_temp_excel)
        items = list(enumerate(registros, start=1))
        st.session_state.errores_lote_acum = []
        st.session_state.sin_seguro_lote_acum = []
        _correr_lote(items, "📦 Descargar PDFs + Matriz (.zip)")

    if st.session_state.fallidos_lote:
        st.divider()
        st.write(f"⚠️ **{len(st.session_state.fallidos_lote)} paciente(s) pendiente(s) de reintentar**:")
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
    try:
        dependencias_validas = core.obtener_dependencias_validas(
            core.abrir_matriz(core.localizar_archivo_matriz())
        )
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

    if st.button("➕ Procesar y agregar a la matriz", type="primary"):
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

        cA, cB, cC, cD = st.columns(4)
        with cA:
            st.download_button(
                "📦 Descargar .zip", data=_zip_resultado(),
                file_name=f"coberturas_manual_{datetime.now():%Y%m%d_%H%M}.zip", mime="application/zip",
            )
        with cB:
            if st.button("☁️ Sincronizar con Drive"):
                _sincronizar_con_drive()
        with cC:
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
        with cD:
            if st.button("🔴 Cerrar Chrome"):
                _cerrar_driver()
                st.success("Chrome cerrado.")
