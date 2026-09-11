# -*- coding: utf-8 -*-
"""
app_web.py
==========

Interfaz web (Streamlit) para el motor de descarga de coberturas de
salud. NO reimplementa la logica de Selenium / deteccion CAMPESINO-GENERAL
/ matriz: llama directamente a las funciones que ya existen y funcionan
en `descargar_coberturas.py` (que debe estar en la MISMA carpeta que
este archivo). Asi, cualquier correccion futura al motor se refleja
automaticamente aqui, sin mantener dos copias de la misma logica.

Como correrla:
    pip install -r requirements.txt
    streamlit run app_web.py

Funciona igual en Windows o Linux: crear_driver() en el motor ya
detecta automaticamente si existe Chromium/Chromedriver del sistema
(tipico de un servidor Linux) o si debe usar webdriver-manager (tipico
de una PC de escritorio Windows).
"""

import os
import io
import sys
import time
import uuid
import queue
import zipfile
import threading
import contextlib
import traceback
from datetime import date

import streamlit as st
import openpyxl

# ------------------------------------------------------------------
# Importa el motor original SIN modificarlo. Debe vivir junto a este
# archivo (mismo directorio).
# ------------------------------------------------------------------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import descargar_coberturas as motor
except ImportError as e:
    st.set_page_config(page_title="Coberturas de Salud", page_icon="⚕️")
    st.error(
        "No se encontro `descargar_coberturas.py` en esta misma carpeta. "
        "Este archivo es obligatorio: la app web solo es la interfaz, "
        "toda la logica de descarga vive ahi.\n\n"
        f"Detalle: {e}"
    )
    st.stop()

st.set_page_config(page_title="Coberturas de Salud - MSP", page_icon="⚕️", layout="wide")

# ============================================================
# LOGIN POR UNIDAD (aisla los archivos de cada unidad operativa)
# ============================================================
#
# Los usuarios/contraseñas NUNCA van en este archivo ni en GitHub: se
# definen en los "Secrets" de la app en Streamlit Cloud (Settings >
# Secrets), con este formato:
#
#   [usuarios]
#   pumamaqui = { password = "clave-segura-1", nombre = "Casa de Acogida Pumamaqui" }
#   san_juan  = { password = "clave-segura-2", nombre = "UO San Juan" }
#   ...  (una linea por cada una de las 21 unidades)
#
# El texto antes de "=" (ej. "pumamaqui") es lo que la unidad escribe
# como usuario. Cada unidad, al iniciar sesion, solo ve y modifica sus
# PROPIOS archivos (matriz, PDFs, datos de pacientes guardados) -nunca
# los de otra unidad-, porque a partir de aqui se redirige todo el
# motor a una subcarpeta propia de esa unidad.

def _cargar_usuarios():
    try:
        return {k: dict(v) for k, v in st.secrets["usuarios"].items()}
    except Exception:
        return {}


if "auth_unidad" not in st.session_state:
    st.session_state.auth_unidad = None

if st.session_state.auth_unidad is None:
    st.title("⚕️ Coberturas de Salud — Acceso")
    usuarios = _cargar_usuarios()

    if not usuarios:
        st.error(
            "Todavía no hay usuarios configurados. En Streamlit Cloud, entra a "
            "**Settings → Secrets** de esta app y agrega la sección `[usuarios]` "
            "(ver el comentario al inicio de `app_web.py` para el formato exacto)."
        )
        st.stop()

    with st.form("form_login"):
        usuario_in = st.text_input("Usuario (unidad operativa)")
        clave_in = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar", type="primary")

    if entrar:
        datos_usuario = usuarios.get(usuario_in.strip())
        if datos_usuario and clave_in == datos_usuario.get("password"):
            st.session_state.auth_unidad = usuario_in.strip()
            st.session_state.auth_unidad_nombre = datos_usuario.get("nombre", usuario_in.strip())
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")

    st.stop()

# ------------------------------------------------------------------
# A partir de aqui, ya hay una unidad autenticada. Se redirige TODO el
# motor (matriz, PDFs, datos_pacientes.json, respaldos) a una carpeta
# exclusiva de esta unidad, para que nunca se mezcle con las de las
# otras 20. Cada unidad debe subir su propio INSTRUCTIVO...xlsx una
# sola vez (mas abajo se le pide si todavia no lo tiene).
# ------------------------------------------------------------------
_UNIDAD = st.session_state.auth_unidad
_CARPETA_BASE_ORIGINAL = motor.BASE_DIR
_carpeta_unidad = os.path.join(_CARPETA_BASE_ORIGINAL, "DATOS_UNIDADES", _UNIDAD)
os.makedirs(_carpeta_unidad, exist_ok=True)

motor.BASE_DIR = _carpeta_unidad
motor.ARCHIVO_EXCEL = os.path.join(_carpeta_unidad, "reporte_inconsistencias.xlsx")
motor.CARPETA_SALIDA = os.path.join(_carpeta_unidad, "PDF_DESCARGADOS")
motor.RUTA_DATOS_PACIENTES = os.path.join(_carpeta_unidad, "datos_pacientes.json")
motor.CARPETA_RESPALDOS_MATRIZ = os.path.join(_carpeta_unidad, "_respaldos_matriz")

# ============================================================
# SINCRONIZACION CON GOOGLE DRIVE (persistencia entre redeploys)
# ============================================================
#
# Requiere que en los Secrets de Streamlit Cloud exista la seccion
# [gcp_service_account] con el contenido del .json de la cuenta de
# servicio (ver README.md). Si no esta configurada, la app sigue
# funcionando igual mas no persiste entre redeploys.

_DRIVE_CARPETA_RAIZ_ID_DEFECTO = "1OQ9dkzDOtbQT90mljvoPvzToJcPOSeH6"


def _drive_disponible():
    try:
        return "gcp_service_account" in st.secrets
    except Exception:
        return False


@st.cache_resource
def _drive_cliente():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    info = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _drive_carpeta_raiz_id():
    try:
        return st.secrets["drive"]["carpeta_raiz_id"]
    except Exception:
        return _DRIVE_CARPETA_RAIZ_ID_DEFECTO


def _drive_obtener_o_crear_carpeta(nombre, carpeta_padre_id):
    servicio = _drive_cliente()
    nombre_escapado = nombre.replace("'", "\\'")
    query = (
        f"'{carpeta_padre_id}' in parents and name = '{nombre_escapado}' "
        "and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    resultado = servicio.files().list(q=query, fields="files(id, name)", pageSize=1).execute()
    archivos = resultado.get("files", [])
    if archivos:
        return archivos[0]["id"]
    metadata = {"name": nombre, "mimeType": "application/vnd.google-apps.folder", "parents": [carpeta_padre_id]}
    carpeta = servicio.files().create(body=metadata, fields="id").execute()
    return carpeta["id"]


def _drive_listar(carpeta_id):
    servicio = _drive_cliente()
    archivos, token = [], None
    while True:
        resultado = servicio.files().list(
            q=f"'{carpeta_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=token, pageSize=200
        ).execute()
        archivos.extend(resultado.get("files", []))
        token = resultado.get("nextPageToken")
        if not token:
            break
    return archivos


def _drive_subir_o_actualizar(ruta_local, nombre_remoto, carpeta_id):
    from googleapiclient.http import MediaFileUpload
    servicio = _drive_cliente()
    existentes = [
        f for f in _drive_listar(carpeta_id)
        if f["name"] == nombre_remoto and f["mimeType"] != "application/vnd.google-apps.folder"
    ]
    media = MediaFileUpload(ruta_local, resumable=False)
    if existentes:
        servicio.files().update(fileId=existentes[0]["id"], media_body=media).execute()
    else:
        servicio.files().create(body={"name": nombre_remoto, "parents": [carpeta_id]}, media_body=media).execute()


def _drive_descargar_archivo(archivo_id, ruta_destino):
    from googleapiclient.http import MediaIoBaseDownload
    servicio = _drive_cliente()
    solicitud = servicio.files().get_media(fileId=archivo_id)
    os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)
    with open(ruta_destino, "wb") as f:
        descargador = MediaIoBaseDownload(f, solicitud)
        listo = False
        while not listo:
            _, listo = descargador.next_chunk()


def _drive_subir_carpeta(carpeta_local, carpeta_id):
    """Sube/actualiza TODO el contenido de carpeta_local dentro de
    carpeta_id en Drive, recursivamente (espejo local -> nube)."""
    if not carpeta_id or not os.path.isdir(carpeta_local):
        return
    for nombre in os.listdir(carpeta_local):
        if nombre.startswith("."):
            continue
        ruta = os.path.join(carpeta_local, nombre)
        if os.path.isdir(ruta):
            sub_id = _drive_obtener_o_crear_carpeta(nombre, carpeta_id)
            _drive_subir_carpeta(ruta, sub_id)
        else:
            _drive_subir_o_actualizar(ruta, nombre, carpeta_id)


def _drive_descargar_carpeta(carpeta_id, carpeta_local):
    """Descarga TODO el contenido de carpeta_id de Drive a carpeta_local,
    recursivamente (nube -> local; se usa para restaurar al iniciar sesion)."""
    if not carpeta_id:
        return
    os.makedirs(carpeta_local, exist_ok=True)
    for item in _drive_listar(carpeta_id):
        ruta_local = os.path.join(carpeta_local, item["name"])
        if item["mimeType"] == "application/vnd.google-apps.folder":
            _drive_descargar_carpeta(item["id"], ruta_local)
        else:
            _drive_descargar_archivo(item["id"], ruta_local)


def _drive_sincronizar_ahora(mensaje_spinner="Guardando en la nube..."):
    if not st.session_state.get("drive_carpeta_unidad_id"):
        return
    with st.spinner(mensaje_spinner):
        try:
            _drive_subir_carpeta(_carpeta_unidad, st.session_state.drive_carpeta_unidad_id)
        except Exception as e:
            st.warning(f"No se pudo guardar en la nube: {e}")


# Al entrar, restaura de Drive lo que ya existiera de esta unidad
# (asi el disco temporal de Streamlit Cloud "recupera" lo que se
# habia perdido en el ultimo redeploy).
if "drive_carpeta_unidad_id" not in st.session_state:
    if _drive_disponible():
        try:
            with st.spinner("Conectando con Google Drive y restaurando los datos de tu unidad..."):
                _raiz_id = _drive_carpeta_raiz_id()
                _id_unidad = _drive_obtener_o_crear_carpeta(_UNIDAD, _raiz_id)
                st.session_state.drive_carpeta_unidad_id = _id_unidad
                _drive_descargar_carpeta(_id_unidad, _carpeta_unidad)
        except Exception as e:
            st.session_state.drive_carpeta_unidad_id = None
            st.warning(
                f"No se pudo conectar con Google Drive todavía ({e}). "
                "Se sigue trabajando con el disco temporal; usa el botón "
                "'📦 Descargar todos los PDF (.zip)' como respaldo mientras tanto."
            )
    else:
        st.session_state.drive_carpeta_unidad_id = None

# ============================================================
# UTILIDADES COMUNES
# ============================================================

def _asegurar_carpetas():
    os.makedirs(motor.CARPETA_SALIDA, exist_ok=True)
    carpeta_descargas_temp = os.path.join(motor.CARPETA_SALIDA, "_descargas_temp")
    os.makedirs(carpeta_descargas_temp, exist_ok=True)
    carpeta_diagnostico = os.path.join(motor.CARPETA_SALIDA, "_diagnostico")
    return carpeta_descargas_temp, carpeta_diagnostico


def _tabla_desde_filas(filas, columnas):
    """Convierte una lista de listas (como las que arma el motor para
    errores/sin_seguro) en algo que st.dataframe pueda mostrar, sin
    depender de pandas."""
    if not filas:
        return None
    return [dict(zip(columnas, fila)) for fila in filas]


def _boton_descargar_zip(etiqueta, key):
    """Comprime toda la carpeta PDF_DESCARGADOS (PDFs + reportes) en un
    .zip y ofrece descargarlo al navegador del usuario. Sirve sin
    importar en que maquina este corriendo la app (local o en la nube),
    porque la descarga viaja por el propio navegador, no depende de
    que el usuario tenga acceso al disco del servidor."""
    if not os.path.isdir(motor.CARPETA_SALIDA):
        st.caption("Todavía no hay nada descargado en esta sesión.")
        return
    if st.button(etiqueta, key=key):
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for raiz, _dirs, archivos in os.walk(motor.CARPETA_SALIDA):
                for nombre_archivo in archivos:
                    ruta_completa = os.path.join(raiz, nombre_archivo)
                    ruta_relativa = os.path.relpath(ruta_completa, motor.CARPETA_SALIDA)
                    zf.write(ruta_completa, ruta_relativa)
        buffer.seek(0)
        st.download_button(
            "⬇️ Descargar PDF_DESCARGADOS.zip", buffer.getvalue(),
            file_name="PDF_DESCARGADOS.zip", mime="application/zip", key=key + "_dl"
        )


def _reset_estado_auto():
    st.session_state.auto_running = False
    st.session_state.auto_queue = None
    st.session_state.auto_thread = None
    st.session_state.auto_log = []
    st.session_state.auto_progress = (0, 0)
    st.session_state.auto_errores = []
    st.session_state.auto_sin_seguro = []
    st.session_state.auto_done = False
    st.session_state.auto_error_fatal = None
    st.session_state.auto_indices_error = set()
    st.session_state.auto_escribir_matriz = False
    st.session_state.auto_responsable_matriz = ""


def _init_estado():
    if "auto_running" not in st.session_state:
        _reset_estado_auto()
    if "manual_errores" not in st.session_state:
        st.session_state.manual_errores = []
    if "manual_sin_seguro" not in st.session_state:
        st.session_state.manual_sin_seguro = []
    if "manual_contador" not in st.session_state:
        st.session_state.manual_contador = 0
    if "manual_datos_previos" not in st.session_state:
        st.session_state.manual_datos_previos = None
    if "manual_cedula_buscada" not in st.session_state:
        st.session_state.manual_cedula_buscada = None
    if "manual_queue_in" not in st.session_state:
        st.session_state.manual_queue_in = queue.Queue()
    if "manual_queue_out" not in st.session_state:
        st.session_state.manual_queue_out = queue.Queue()
    if "manual_worker_thread" not in st.session_state:
        st.session_state.manual_worker_thread = None
    if "manual_items" not in st.session_state:
        st.session_state.manual_items = []  # cola visible: [{id, cedula, fecha, estado, ...}]


def _worker_manual(q_in, q_out, carpeta_descargas_temp, carpeta_diagnostico):
    """Corre en un hilo aparte, uno por sesion de usuario. Mantiene UN
    solo Chrome abierto (igual que modo_interactivo en consola) y va
    tomando pacientes de la cola en el orden en que llegan, sin que la
    interfaz tenga que esperar a que termine cada uno."""
    driver_holder = [None]
    while True:
        item = q_in.get()
        if item is None:  # senal para terminar el hilo
            if driver_holder[0] is not None:
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass
            return
        if item.get("accion") == "cerrar_driver":
            if driver_holder[0] is not None:
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass
                driver_holder[0] = None
            continue

        item_id = item["id"]
        q_out.put(("procesando", item_id, None))
        log_buffer = io.StringIO()
        try:
            with contextlib.redirect_stdout(log_buffer):
                if driver_holder[0] is None:
                    driver_holder[0] = motor.crear_driver(carpeta_descargas_temp)

                reg = {
                    "cedula": item["cedula"],
                    "nombre": None,
                    "fechas": [item["fecha_atencion"]],
                    "cedula_padre": None,
                    "cedula_titular": None,
                }
                errores_local, sin_seguro_local = [], []
                filas_matriz = motor.procesar_registro(
                    reg, item["contador"], item["contador"], driver_holder,
                    carpeta_descargas_temp, carpeta_diagnostico,
                    errores_local, sin_seguro_local, recolectar_matriz=True
                )
                motor.guardar_datos_paciente(
                    item["cedula"], nombre=reg.get("nombre"),
                    fecha_nacimiento=item["fecha_nacimiento"], sexo=item["sexo"]
                )
                filas_escritas = []
                if item.get("hay_matriz") and filas_matriz:
                    try:
                        filas_escritas = motor.agregar_filas_matriz_con_bloqueo(
                            filas_matriz, item["dependencia"], item["fecha_nacimiento"],
                            item["sexo"], item["observaciones"], item["responsable"]
                        )
                    except TimeoutError as e:
                        errores_local.append([
                            reg.get("nombre") or item["cedula"], item["cedula"],
                            item["fecha_atencion"].strftime("%d-%m-%Y"), f"(MATRIZ) {e}"
                        ])

            q_out.put(("listo", item_id, {
                "nombre": reg.get("nombre"),
                "log": log_buffer.getvalue(),
                "filas_matriz": filas_escritas,
                "errores": errores_local,
                "sin_seguro": sin_seguro_local,
            }))
        except Exception as e:
            try:
                driver_holder[0].quit()
            except Exception:
                pass
            driver_holder[0] = None
            q_out.put(("error", item_id, {
                "detalle": f"{type(e).__name__}: {e}",
                "log": log_buffer.getvalue(),
            }))


def _lanzar_auto(items, escribir_matriz, responsable_matriz, es_reintento=False):
    """Arranca el hilo de descarga para 'items' (lista de (indice_original,
    registro)). Si es_reintento=True, sigue acumulando en las mismas
    listas de errores/sin_seguro que ya se tenian (no se pierde el
    historial de la corrida anterior); si es False, arranca de cero."""
    if es_reintento:
        errores_base = list(st.session_state.auto_errores)
        sin_seguro_base = list(st.session_state.auto_sin_seguro)
        st.session_state.auto_log.append(f"--- Reintentando {len(items)} registro(s) fallido(s) ---")
    else:
        errores_base, sin_seguro_base = [], []
        st.session_state.auto_log = []

    st.session_state.auto_running = True
    st.session_state.auto_done = False
    st.session_state.auto_error_fatal = None
    st.session_state.auto_progress = (0, len(items))
    st.session_state.auto_indices_error = set()
    st.session_state.auto_escribir_matriz = escribir_matriz
    st.session_state.auto_responsable_matriz = responsable_matriz

    q = queue.Queue()
    st.session_state.auto_queue = q

    def _worker(q, items, escribir_matriz, responsable_matriz, errores, sin_seguro):
        class _QueueWriter:
            def write(self, s):
                s = s.rstrip("\n")
                if s.strip():
                    q.put(("log", s))

            def flush(self):
                pass

        driver_holder = None
        try:
            with contextlib.redirect_stdout(_QueueWriter()):
                carpeta_descargas_temp, carpeta_diagnostico = _asegurar_carpetas()
                driver_holder = [motor.crear_driver(carpeta_descargas_temp)]

                def cb_fila(i, estado):
                    q.put(("fila", i, len(items), estado))

                def cb_progreso(n, total):
                    q.put(("progress", n, total))

                motor.procesar_lote(
                    items, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                    errores, sin_seguro, callback_fila=cb_fila, callback_progreso=cb_progreso,
                    escribir_matriz=escribir_matriz, responsable_matriz=responsable_matriz
                )
                motor._guardar_reportes_finales(errores, sin_seguro)

            q.put(("done", errores, sin_seguro))
        except Exception as e:
            q.put(("error", f"{type(e).__name__}: {e}\n{traceback.format_exc()}"))
        finally:
            if driver_holder is not None:
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass

    t = threading.Thread(
        target=_worker, args=(q, items, escribir_matriz, responsable_matriz, errores_base, sin_seguro_base),
        daemon=True
    )
    st.session_state.auto_thread = t
    t.start()


_init_estado()

col_titulo, col_sesion = st.columns([4, 1])
with col_titulo:
    st.title("⚕️ Descarga de Coberturas de Salud")
    st.caption(
        f"Unidad: **{st.session_state.get('auth_unidad_nombre', _UNIDAD)}** — "
        "interfaz web del mismo motor de `descargar_coberturas.py`."
    )
with col_sesion:
    st.write("")
    if st.button("🚪 Cerrar sesión"):
        _drive_sincronizar_ahora("Guardando todo en la nube antes de salir...")
        st.session_state.auth_unidad = None
        st.session_state.auth_unidad_nombre = None
        st.session_state.drive_carpeta_unidad_id = None
        st.rerun()

if st.session_state.get("drive_carpeta_unidad_id"):
    st.caption("☁️ Conectado a Google Drive — tus archivos se pueden respaldar en la nube.")
else:
    st.caption("⚠️ Sin conexión a Google Drive — los archivos solo viven en este servidor temporal.")

tab_auto, tab_manual = st.tabs(["📂 Modo automático (Excel)", "🧍 Modo manual (un paciente a la vez)"])

# ============================================================
# TAB 1 - MODO AUTOMATICO (equivalente a main())
# ============================================================

with tab_auto:
    st.subheader("Carga masiva desde Excel")
    st.write(
        f"Sube el archivo con las columnas **{motor.COL_NOMBRE}**, **{motor.COL_CEDULA}** y "
        f"**{motor.COL_FECHAS}** (y opcionalmente **{motor.COL_PADRE}** / **{motor.COL_TITULAR}**), "
        "igual que `reporte_inconsistencias.xlsx`."
    )

    archivo_subido = st.file_uploader(
        "Archivo Excel de inconsistencias", type=["xlsx"], key="auto_uploader",
        disabled=st.session_state.auto_running
    )

    col_a, col_b = st.columns(2)
    with col_a:
        escribir_matriz = st.checkbox(
            "Registrar también en la matriz mensual (Instructivo)",
            value=False, disabled=st.session_state.auto_running,
            help="Se agregan filas con Dependencia, fecha de nacimiento y sexo en blanco "
                 "-se completan despues a mano-, igual que el modo por lotes del script original."
        )
    with col_b:
        responsable_matriz = st.text_input(
            "Responsable (para la matriz)", value="", disabled=st.session_state.auto_running or not escribir_matriz
        ).upper()

    registros_preview = None
    if archivo_subido is not None and not st.session_state.auto_running:
        try:
            wb_preview = openpyxl.load_workbook(io.BytesIO(archivo_subido.getvalue()), data_only=True)
            ws_preview = wb_preview[motor.HOJA]
            n_filas = max(ws_preview.max_row - 1, 0)
            st.info(f"El archivo tiene {n_filas} fila(s) de datos (sin contar encabezado).")
        except Exception as e:
            st.warning(f"No se pudo previsualizar el archivo todavia (se validará al iniciar): {e}")

    iniciar = st.button(
        "▶️ Iniciar descarga", type="primary",
        disabled=st.session_state.auto_running or archivo_subido is None
    )

    if iniciar and archivo_subido is not None and not st.session_state.auto_running:
        # Guarda el excel subido exactamente donde el motor espera leerlo.
        with open(motor.ARCHIVO_EXCEL, "wb") as f:
            f.write(archivo_subido.getvalue())

        try:
            registros = motor.leer_excel(motor.ARCHIVO_EXCEL)
            if motor.LIMITE_PRUEBA:
                registros = registros[: motor.LIMITE_PRUEBA]
        except Exception as e:
            st.error(f"No se pudo leer el Excel: {e}")
        else:
            _reset_estado_auto()
            items = list(enumerate(registros, start=1))
            _lanzar_auto(items, escribir_matriz, responsable_matriz, es_reintento=False)
            st.rerun()

    # --- Drena la cola y refresca mientras corre ---
    if st.session_state.auto_running and st.session_state.auto_queue is not None:
        q = st.session_state.auto_queue
        try:
            while True:
                item = q.get_nowait()
                tipo = item[0]
                if tipo == "log":
                    st.session_state.auto_log.append(item[1])
                elif tipo == "progress":
                    st.session_state.auto_progress = (item[1], item[2])
                elif tipo == "fila":
                    _, i, _total_items, estado = item
                    if estado == "Con errores":
                        st.session_state.auto_indices_error.add(i)
                elif tipo == "done":
                    st.session_state.auto_errores = item[1]
                    st.session_state.auto_sin_seguro = item[2]
                    st.session_state.auto_done = True
                    st.session_state.auto_running = False
                    _drive_sincronizar_ahora("Tanda terminada, guardando en la nube...")
                elif tipo == "error":
                    st.session_state.auto_error_fatal = item[1]
                    st.session_state.auto_running = False
        except queue.Empty:
            pass

    n, total = st.session_state.auto_progress
    if total:
        st.progress(min(n / total, 1.0), text=f"{n}/{total} registros procesados")

    if st.session_state.auto_log:
        with st.expander("Registro detallado (log)", expanded=st.session_state.auto_running):
            st.code("\n".join(st.session_state.auto_log[-500:]), language=None)

    if st.session_state.auto_running:
        st.info("Descargando... esta pestaña se actualiza sola.")
        time.sleep(1.2)
        st.rerun()

    if st.session_state.auto_error_fatal:
        st.error(f"La descarga se detuvo por un error:\n\n{st.session_state.auto_error_fatal}")

    if st.session_state.auto_done:
        errores = st.session_state.auto_errores
        sin_seguro = st.session_state.auto_sin_seguro
        if not errores and not sin_seguro:
            st.success("Todo se procesó sin errores.")
        else:
            st.warning(f"Terminado: {len(errores)} incidencia(s) de error, {len(sin_seguro)} caso(s) sin seguro.")

        col1, col2 = st.columns(2)
        with col1:
            tabla_err = _tabla_desde_filas(errores, ["Nombre", "Cédula", "Fecha", "Detalle"])
            if tabla_err:
                st.write("**Errores**")
                st.dataframe(tabla_err, use_container_width=True)
                ruta_log = os.path.join(motor.CARPETA_SALIDA, "log_errores.xlsx")
                if os.path.exists(ruta_log):
                    with open(ruta_log, "rb") as f:
                        st.download_button("⬇️ Descargar log_errores.xlsx", f.read(),
                                            file_name="log_errores.xlsx", key="dl_errores")
        with col2:
            tabla_ss = _tabla_desde_filas(
                sin_seguro, ["Nombre Paciente", "Cédula Paciente", "Cédula Consultada", "Rol", "Fecha"]
            )
            if tabla_ss:
                st.write("**Sin seguro**")
                st.dataframe(tabla_ss, use_container_width=True)
                ruta_ss = os.path.join(motor.CARPETA_SALIDA, "sin_seguro.xlsx")
                if os.path.exists(ruta_ss):
                    with open(ruta_ss, "rb") as f:
                        st.download_button("⬇️ Descargar sin_seguro.xlsx", f.read(),
                                            file_name="sin_seguro.xlsx", key="dl_sinseguro")

        col_zip, col_nube = st.columns(2)
        with col_zip:
            _boton_descargar_zip("📦 Descargar todos los PDF (.zip)", key="zip_auto")
        with col_nube:
            if st.session_state.get("drive_carpeta_unidad_id") and st.button("☁️ Guardar en la nube ahora", key="sync_auto"):
                _drive_sincronizar_ahora()
                st.success("Guardado en Google Drive.")

        if st.session_state.auto_indices_error:
            st.markdown("---")
            n_fallidos = len(st.session_state.auto_indices_error)
            st.write(f"⚠️ {n_fallidos} registro(s) del Excel original terminaron con error.")
            if st.button(f"🔁 Reintentar los {n_fallidos} fallido(s)", type="primary"):
                try:
                    registros_todos = motor.leer_excel(motor.ARCHIVO_EXCEL)
                except Exception as e:
                    st.error(f"No se pudo releer el Excel original para reintentar: {e}")
                else:
                    indices = sorted(st.session_state.auto_indices_error)
                    items_retry = [
                        (i, registros_todos[i - 1]) for i in indices if 0 < i <= len(registros_todos)
                    ]
                    if items_retry:
                        _lanzar_auto(
                            items_retry, st.session_state.auto_escribir_matriz,
                            st.session_state.auto_responsable_matriz, es_reintento=True
                        )
                        st.rerun()
                    else:
                        st.warning("No se encontraron esas filas en el Excel actual (¿se reemplazó el archivo?).")

        if st.button("Empezar otra carga"):
            _reset_estado_auto()
            st.rerun()

# ============================================================
# TAB 2 - MODO MANUAL (equivalente a modo_interactivo())
# ============================================================

with tab_manual:
    st.subheader("Un paciente a la vez")

    # ---- Matriz: dependencias validas (si el Instructivo esta disponible) ----
    dependencias_validas = []
    hay_matriz_disponible = False
    aviso_matriz = None
    try:
        ruta_matriz_actual = motor.localizar_archivo_matriz()
        dependencias_validas = motor.obtener_dependencias_validas(motor.abrir_matriz(ruta_matriz_actual))
        hay_matriz_disponible = True
    except FileNotFoundError as e:
        aviso_matriz = str(e)
    except Exception as e:
        aviso_matriz = f"No se pudo leer la matriz: {e}"

    if aviso_matriz:
        st.warning(
            "Esta unidad todavía no tiene su archivo INSTRUCTIVO (matriz mensual) configurado. "
            "Se podrán descargar los PDF de cobertura con normalidad, pero no se llenará la "
            "matriz automáticamente hasta que subas uno."
        )
        instructivo_subido = st.file_uploader(
            "Sube el INSTRUCTIVO...xlsx de esta unidad (solo hace falta una vez)",
            type=["xlsx"], key="uploader_instructivo"
        )
        if instructivo_subido is not None:
            os.makedirs(motor.BASE_DIR, exist_ok=True)
            ruta_guardado = os.path.join(motor.BASE_DIR, instructivo_subido.name)
            with open(ruta_guardado, "wb") as f:
                f.write(instructivo_subido.getvalue())
            st.success(f"Guardado como {instructivo_subido.name}. Recargando...")
            st.rerun()

    responsable = st.text_input("Nombre del responsable (persona que ingresa la información)", key="manual_responsable").upper()

    # Arranca el hilo de trabajo de esta sesion (una sola vez).
    carpeta_descargas_temp, carpeta_diagnostico = _asegurar_carpetas()
    hilo = st.session_state.manual_worker_thread
    if hilo is None or not hilo.is_alive():
        hilo = threading.Thread(
            target=_worker_manual,
            args=(st.session_state.manual_queue_in, st.session_state.manual_queue_out,
                  carpeta_descargas_temp, carpeta_diagnostico),
            daemon=True
        )
        st.session_state.manual_worker_thread = hilo
        hilo.start()

    st.markdown("---")
    st.markdown(f"**Paciente #{st.session_state.manual_contador + 1}**")

    cedula_input = st.text_input("Cédula del paciente", key="manual_cedula", max_chars=10)

    datos_guardados = None
    if cedula_input:
        cedula_normalizada = "".join(ch for ch in cedula_input if ch.isdigit()).zfill(10)[:10]
        if cedula_normalizada != st.session_state.manual_cedula_buscada:
            st.session_state.manual_cedula_buscada = cedula_normalizada
            st.session_state.manual_datos_previos = motor.obtener_datos_guardados_paciente(cedula_normalizada)
        datos_guardados = st.session_state.manual_datos_previos
        if datos_guardados:
            partes = []
            if datos_guardados.get("nombre"):
                partes.append(datos_guardados["nombre"])
            if datos_guardados.get("fecha_nacimiento"):
                partes.append(f"nació {datos_guardados['fecha_nacimiento']}")
            if datos_guardados.get("sexo"):
                partes.append(f"sexo {datos_guardados['sexo']}")
            st.caption(f"Ya se tienen datos guardados de esta cédula: {', '.join(partes)}")

    fecha_nac_defecto = None
    if datos_guardados and datos_guardados.get("fecha_nacimiento"):
        fecha_nac_defecto = motor._parsear_fecha_flexible(datos_guardados["fecha_nacimiento"])
    sexo_defecto = (datos_guardados or {}).get("sexo", "F")

    PLACEHOLDER_DEP = "-- Selecciona --"
    with st.form("form_manual", clear_on_submit=False):
        c1, c2 = st.columns(2)
        with c1:
            fecha_atencion = st.date_input("Fecha de atención", value=date.today(), key="manual_fecha_atencion")
            fecha_nacimiento = st.date_input(
                "Fecha de nacimiento del paciente",
                value=fecha_nac_defecto or date(2000, 1, 1),
                min_value=date(1900, 1, 1), max_value=date.today(),
                key="manual_fecha_nac"
            )
            sexo = st.radio("Sexo del paciente", ["F", "M"], index=(0 if sexo_defecto != "M" else 1),
                             horizontal=True, key="manual_sexo")
        with c2:
            if dependencias_validas:
                dependencia = st.selectbox(
                    "Dependencia (tipo de consulta)",
                    options=[PLACEHOLDER_DEP] + dependencias_validas + ["OTRA (escribir)"],
                    key="manual_dependencia"
                )
                if dependencia == "OTRA (escribir)":
                    dependencia = st.text_input("Escribe la dependencia", key="manual_dependencia_otra").upper()
            else:
                dependencia = st.text_input("Dependencia (tipo de consulta)", key="manual_dependencia_otra").upper()
            observaciones = st.text_area("Observaciones (opcional)", height=90, key="manual_observaciones")

        enviado = st.form_submit_button("➕ Agregar a la cola y continuar con el siguiente", type="primary")

    if enviado:
        cedula_normalizada = "".join(ch for ch in (cedula_input or "") if ch.isdigit()).zfill(10)[:10]
        if not cedula_input or len(cedula_normalizada) != 10:
            st.error("Ingresa una cédula válida de 10 dígitos.")
        elif not responsable:
            st.error("Ingresa el nombre del responsable.")
        elif dependencias_validas and dependencia == PLACEHOLDER_DEP:
            st.error("Selecciona una dependencia.")
        elif not dependencia:
            st.error("Ingresa la dependencia.")
        else:
            st.session_state.manual_contador += 1
            item_id = str(uuid.uuid4())
            item = {
                "id": item_id,
                "contador": st.session_state.manual_contador,
                "cedula": cedula_normalizada,
                "fecha_atencion": fecha_atencion,
                "dependencia": dependencia,
                "fecha_nacimiento": fecha_nacimiento,
                "sexo": sexo,
                "observaciones": observaciones,
                "responsable": responsable,
                "hay_matriz": hay_matriz_disponible,
            }
            st.session_state.manual_items.insert(0, {
                "id": item_id, "cedula": cedula_normalizada,
                "fecha": fecha_atencion.strftime("%d-%m-%Y"),
                "estado": "En cola", "nombre": None, "detalle": "", "filas_matriz": [],
                "original": item,
            })
            st.session_state.manual_queue_in.put(item)

            # Limpia el formulario para el siguiente paciente. Se conserva
            # SOLO el responsable (key "manual_responsable", que queda
            # intacto porque no se toca aqui); todo lo demas, incluida la
            # dependencia, vuelve a su valor por defecto (la fecha de
            # atencion vuelve a ser la de hoy).
            for k in ("manual_cedula", "manual_fecha_atencion", "manual_fecha_nac", "manual_sexo",
                      "manual_dependencia", "manual_dependencia_otra", "manual_observaciones",
                      "manual_cedula_buscada", "manual_datos_previos"):
                st.session_state.pop(k, None)
            st.rerun()

    # --- Drena resultados que el hilo de trabajo ya vaya terminando ---
    try:
        while True:
            tipo, item_id, datos = st.session_state.manual_queue_out.get_nowait()
            for it in st.session_state.manual_items:
                if it["id"] != item_id:
                    continue
                if tipo == "procesando":
                    it["estado"] = "Procesando..."
                elif tipo == "listo":
                    it["estado"] = "Listo"
                    it["nombre"] = datos.get("nombre")
                    it["detalle"] = datos.get("log", "")
                    it["filas_matriz"] = datos.get("filas_matriz", [])
                    st.session_state.manual_errores.extend(datos.get("errores", []))
                    st.session_state.manual_sin_seguro.extend(datos.get("sin_seguro", []))
                elif tipo == "error":
                    it["estado"] = "Error"
                    it["detalle"] = datos.get("detalle", "")
                break
    except queue.Empty:
        pass

    st.markdown("---")
    st.markdown("**Cola de pacientes de esta sesión**")
    if not st.session_state.manual_items:
        st.caption("Todavía no has agregado ningún paciente.")
    else:
        iconos = {"En cola": "🕓", "Procesando...": "⏳", "Listo": "✅", "Error": "❌"}
        for it in st.session_state.manual_items:
            etiqueta = f"{iconos.get(it['estado'], '')} {it['cedula']} — {it['fecha']} — {it['estado']}"
            if it.get("nombre"):
                etiqueta += f" ({it['nombre']})"
            with st.expander(etiqueta, expanded=(it["estado"] == "Error")):
                if it["filas_matriz"]:
                    for ruta_m, fila in it["filas_matriz"]:
                        st.write(f"→ Fila {fila} agregada a la matriz ({os.path.basename(ruta_m)}).")
                if it["detalle"].strip():
                    st.code(it["detalle"], language=None)
                if it["estado"] == "Error" and st.button("🔁 Reintentar este paciente", key=f"retry_{it['id']}"):
                    nuevo_id = str(uuid.uuid4())
                    nuevo_item = dict(it["original"])
                    nuevo_item["id"] = nuevo_id
                    st.session_state.manual_items.insert(0, {
                        "id": nuevo_id, "cedula": it["cedula"], "fecha": it["fecha"],
                        "estado": "En cola", "nombre": None, "detalle": "", "filas_matriz": [],
                        "original": nuevo_item,
                    })
                    st.session_state.manual_queue_in.put(nuevo_item)
                    st.rerun()

    if any(it["estado"] in ("En cola", "Procesando...") for it in st.session_state.manual_items):
        time.sleep(1.2)
        st.rerun()

    st.markdown("---")
    col_fin1, col_fin2, col_fin3, col_fin4 = st.columns(4)
    with col_fin1:
        if st.button("🧾 Generar reportes de esta sesión (errores / sin seguro)"):
            motor._guardar_reportes_finales(st.session_state.manual_errores, st.session_state.manual_sin_seguro)
            st.success("Reportes generados.")
            _drive_sincronizar_ahora()
            ruta_log = os.path.join(motor.CARPETA_SALIDA, "log_errores.xlsx")
            ruta_ss = os.path.join(motor.CARPETA_SALIDA, "sin_seguro.xlsx")
            if os.path.exists(ruta_log):
                with open(ruta_log, "rb") as f:
                    st.download_button("⬇️ log_errores.xlsx", f.read(), file_name="log_errores.xlsx", key="dl_man_err")
            if os.path.exists(ruta_ss):
                with open(ruta_ss, "rb") as f:
                    st.download_button("⬇️ sin_seguro.xlsx", f.read(), file_name="sin_seguro.xlsx", key="dl_man_ss")
    with col_fin2:
        if st.button("🛑 Cerrar sesión de navegador"):
            st.session_state.manual_queue_in.put({"accion": "cerrar_driver"})
            st.success("Se cerrará el navegador en cuanto termine el paciente actual (si hay alguno en curso).")
    with col_fin3:
        _boton_descargar_zip("📦 Descargar todos los PDF (.zip)", key="zip_manual")
    with col_fin4:
        if st.session_state.get("drive_carpeta_unidad_id") and st.button("☁️ Guardar en la nube ahora", key="sync_manual"):
            _drive_sincronizar_ahora()
            st.success("Guardado en Google Drive.")

    if hay_matriz_disponible:
        with st.expander("📋 Generar copia de la matriz para revisar/enviar"):
            modo_fecha = st.selectbox(
                "¿Qué copiar?",
                options=["todo", "hoy", "mes", "rango"],
                format_func=lambda v: {
                    "todo": "Todo", "hoy": "Solo hoy", "mes": "Un mes completo", "rango": "Un rango de fechas"
                }[v]
            )
            fecha_desde_copia = fecha_hasta_copia = None
            if modo_fecha == "mes":
                fecha_desde_copia = st.date_input("Cualquier fecha del mes a copiar", value=date.today(), key="copia_mes")
            elif modo_fecha == "rango":
                fecha_desde_copia = st.date_input("Desde", key="copia_desde")
                fecha_hasta_copia = st.date_input("Hasta", key="copia_hasta")

            if st.button("Generar copia"):
                try:
                    ruta_temporal = os.path.join(motor.CARPETA_SALIDA, "_copias_matriz_web")
                    copia = motor.generar_copia_matriz(
                        carpeta_destino=ruta_temporal, modo_fecha=modo_fecha,
                        fecha_desde=fecha_desde_copia, fecha_hasta=fecha_hasta_copia
                    )
                    with open(copia, "rb") as f:
                        st.download_button(
                            "⬇️ Descargar copia de la matriz", f.read(),
                            file_name=os.path.basename(copia), key="dl_copia_matriz"
                        )
                except Exception as e:
                    st.error(f"No se pudo generar la copia: {e}")
