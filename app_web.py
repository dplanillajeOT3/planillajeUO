# -*- coding: utf-8 -*-
"""
drive_utils.py - Conexion con Google Drive para PLANILLAJE COBERTURAS.

Estructura esperada (Unidad Compartida "PLANILLAJE COBERTURAS"):

    PLANILLAJE COBERTURAS/  (Unidad Compartida)
      <NOMBRE DE LA UNIDAD>/        p.ej. "CENTRO DE SALUD COLINAS DEL NORTE"
        2026/
          9 SEPTIEMBRE/
            INSTRUCTIVO5 SEPTIEMBRE 2026.xlsx
            IESS/
            CAMPESINO/
            ISSFA/
            ISSPOL/

No se toca nada de descargar_coberturas.py: este modulo solo mueve
archivos hacia/desde Drive. La logica real de descarga y de llenado de
la matriz sigue siendo la de "core" (descargar_coberturas.py), operando
sobre una copia local mientras corre la sesion.

CREDENCIALES: requiere que en st.secrets exista la seccion
[gcp_service_account] (la cuenta de servicio ya creada, con la Unidad
Compartida "PLANILLAJE COBERTURAS" compartida con su correo como
"Administrador de contenido" o superior).

IMPORTANTE: no se pudo probar contra la API real de Drive desde este
entorno (sin acceso a internet aqui), asi que antes de usarlo en un
lote real conviene probar primero con "probar_conexion()" (mas abajo)
o el boton de diagnostico que se agrega en app_web.py.
"""

import io
import os
import re
import unicodedata

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from google.oauth2 import service_account

NOMBRE_UNIDAD_COMPARTIDA = "PLANILLAJE COBERTURAS"
CARPETAS_TIPO_SEGURO = ["IESS", "CAMPESINO", "ISSFA", "ISSPOL"]

MIME_CARPETA = "application/vnd.google-apps.folder"
MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]


def _quitar_tildes(txt):
    return "".join(
        c for c in unicodedata.normalize("NFD", txt or "") if unicodedata.category(c) != "Mn"
    )


# ============================================================
# AUTENTICACION
# ============================================================

_SERVICIO = None


def obtener_servicio_drive(secrets_gcp):
    """secrets_gcp: st.secrets['gcp_service_account'] (dict-like)."""
    global _SERVICIO
    if _SERVICIO is not None:
        return _SERVICIO
    credenciales = service_account.Credentials.from_service_account_info(
        dict(secrets_gcp), scopes=["https://www.googleapis.com/auth/drive"]
    )
    _SERVICIO = build("drive", "v3", credentials=credenciales)
    return _SERVICIO


# ============================================================
# NAVEGACION DE CARPETAS
# ============================================================


def listar_unidades_compartidas_visibles(servicio):
    """Todas las Unidades Compartidas (Shared Drives) que esta cuenta de
    servicio puede ver ahora mismo, sin filtrar por nombre. Util para
    diagnosticar: si esta lista sale vacia, el problema es de permisos
    (no se comparti\u00f3 la Unidad Compartida en si con el correo de la
    cuenta de servicio, o solo se comparti\u00f3 una carpeta de adentro)."""
    vistas = []
    token = None
    while True:
        resultado = servicio.drives().list(
            pageSize=100, pageToken=token, fields="nextPageToken, drives(id, name)"
        ).execute()
        vistas += resultado.get("drives", [])
        token = resultado.get("nextPageToken")
        if not token:
            break
    return vistas


def id_unidad_compartida(servicio, nombre=NOMBRE_UNIDAD_COMPARTIDA):
    visibles = listar_unidades_compartidas_visibles(servicio)

    objetivo = nombre.strip().casefold()
    for unidad in visibles:
        if unidad["name"].strip().casefold() == objetivo:
            return unidad["id"]

    if not visibles:
        raise FileNotFoundError(
            f"Esta cuenta de servicio no ve NINGUNA Unidad Compartida (0 resultados). "
            f"Seguramente se compartio una carpeta de ADENTRO de '{nombre}' con su correo, "
            "en vez de compartir la Unidad Compartida completa. Hay que compartir la Unidad "
            "Compartida en si (clic derecho sobre su nombre en la barra lateral de Drive -> "
            "'Administrar miembros' / 'Compartir') como Administrador de contenido."
        )

    nombres_vistos = ", ".join(f"'{u['name']}'" for u in visibles)
    raise FileNotFoundError(
        f"Esta cuenta de servicio SI ve Unidades Compartidas, pero ninguna se llama "
        f"exactamente '{nombre}'. Las que ve son: {nombres_vistos}. "
        "Revisa mayusculas/espacios, o ajusta NOMBRE_UNIDAD_COMPARTIDA en drive_utils.py."
    )


def buscar_subcarpeta(servicio, drive_id, id_padre, nombre_exacto):
    nombre_escapado = nombre_exacto.replace("'", "\\'")
    q = (
        f"'{id_padre}' in parents and name = '{nombre_escapado}' "
        f"and mimeType = '{MIME_CARPETA}' and trashed = false"
    )
    resultado = servicio.files().list(
        q=q, corpora="drive", driveId=drive_id, includeItemsFromAllDrives=True,
        supportsAllDrives=True, fields="files(id, name)",
    ).execute()
    archivos = resultado.get("files", [])
    return archivos[0]["id"] if archivos else None


def crear_subcarpeta(servicio, drive_id, id_padre, nombre):
    metadata = {"name": nombre, "mimeType": MIME_CARPETA, "parents": [id_padre]}
    carpeta = servicio.files().create(
        body=metadata, supportsAllDrives=True, fields="id"
    ).execute()
    return carpeta["id"]


def buscar_o_crear_subcarpeta(servicio, drive_id, id_padre, nombre):
    id_existente = buscar_subcarpeta(servicio, drive_id, id_padre, nombre)
    return id_existente or crear_subcarpeta(servicio, drive_id, id_padre, nombre)


def buscar_archivo_por_patron(servicio, drive_id, id_padre, patron_regex):
    """Busca, entre los archivos DENTRO de id_padre (no subcarpetas), el
    primero cuyo nombre haga match con patron_regex (case-insensitive).
    Se usa para encontrar 'INSTRUCTIVO...xlsx' sin saber el nombre exacto."""
    resultado = servicio.files().list(
        q=f"'{id_padre}' in parents and trashed = false",
        corpora="drive", driveId=drive_id, includeItemsFromAllDrives=True,
        supportsAllDrives=True, fields="files(id, name, mimeType)",
    ).execute()
    expresion = re.compile(patron_regex, re.IGNORECASE)
    for archivo in resultado.get("files", []):
        if archivo["mimeType"] != MIME_CARPETA and expresion.search(archivo["name"]):
            return archivo
    return None


def carpeta_de_unidad_anio_mes(servicio, drive_id, nombre_unidad, fecha, crear_si_falta=True):
    """Navega Unidad Compartida -> <nombre_unidad> -> <año> -> '<mes_num> <MES>'.
    Devuelve el id de la carpeta del mes. Si crear_si_falta=True, crea la
    cadena de carpetas que falte (no debería hacer falta para el año/mes,
    pero sí sirve para robustez si todavia no existe ese mes).

    La carpeta de cada unidad vive directamente en la raíz de la Unidad
    Compartida, así que su "padre" es el propio drive_id (en la API de
    Drive, el id de la raíz de una Unidad Compartida es igual a su
    driveId)."""
    id_carpeta_unidad = buscar_subcarpeta(servicio, drive_id, drive_id, nombre_unidad)
    if id_carpeta_unidad is None:
        if not crear_si_falta:
            raise FileNotFoundError(f"No existe la carpeta de la unidad '{nombre_unidad}' en Drive.")
        id_carpeta_unidad = crear_subcarpeta(servicio, drive_id, id_unidad, nombre_unidad)

    nombre_anio = str(fecha.year)
    id_carpeta_anio = buscar_subcarpeta(servicio, drive_id, id_carpeta_unidad, nombre_anio)
    if id_carpeta_anio is None:
        if not crear_si_falta:
            raise FileNotFoundError(f"No existe la carpeta '{nombre_anio}' dentro de '{nombre_unidad}'.")
        id_carpeta_anio = crear_subcarpeta(servicio, drive_id, id_carpeta_unidad, nombre_anio)

    nombre_mes = f"{fecha.month} {MESES_ES[fecha.month - 1]}"
    id_carpeta_mes = buscar_subcarpeta(servicio, drive_id, id_carpeta_anio, nombre_mes)
    if id_carpeta_mes is None:
        # por si el mes ya existe pero con otro formato de nombre (p.ej.
        # solo "SEPTIEMBRE" o con guiones), se busca de forma mas flexible
        # antes de crear uno nuevo y terminar con dos carpetas del mismo mes
        resultado = servicio.files().list(
            q=f"'{id_carpeta_anio}' in parents and mimeType = '{MIME_CARPETA}' and trashed = false",
            corpora="drive", driveId=drive_id, includeItemsFromAllDrives=True,
            supportsAllDrives=True, fields="files(id, name)",
        ).execute()
        patron = re.compile(re.escape(MESES_ES[fecha.month - 1]), re.IGNORECASE)
        for carpeta in resultado.get("files", []):
            if patron.search(_quitar_tildes(carpeta["name"]).upper()):
                id_carpeta_mes = carpeta["id"]
                break

    if id_carpeta_mes is None:
        if not crear_si_falta:
            raise FileNotFoundError(f"No existe la carpeta del mes '{nombre_mes}'.")
        id_carpeta_mes = crear_subcarpeta(servicio, drive_id, id_carpeta_anio, nombre_mes)

    return id_carpeta_mes


def asegurar_subcarpetas_tipo_seguro(servicio, drive_id, id_carpeta_mes):
    """Devuelve {'IESS': id, 'CAMPESINO': id, 'ISSFA': id, 'ISSPOL': id},
    creando las que falten."""
    ids = {}
    for nombre in CARPETAS_TIPO_SEGURO:
        ids[nombre] = buscar_o_crear_subcarpeta(servicio, drive_id, id_carpeta_mes, nombre)
    return ids


# ============================================================
# DESCARGA / SUBIDA DE ARCHIVOS
# ============================================================


def descargar_archivo(servicio, file_id, ruta_destino):
    request = servicio.files().get_media(fileId=file_id, supportsAllDrives=True)
    with io.FileIO(ruta_destino, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        listo = False
        while not listo:
            _, listo = downloader.next_chunk()
    return ruta_destino


def subir_o_reemplazar_archivo(servicio, ruta_local, nombre_archivo, id_carpeta, mime_type):
    """Si ya existe un archivo con ese nombre en la carpeta, lo actualiza
    (mantiene el mismo file_id / enlace); si no, lo crea."""
    resultado = servicio.files().list(
        q=f"'{id_carpeta}' in parents and name = '{nombre_archivo}' and trashed = false",
        supportsAllDrives=True, includeItemsFromAllDrives=True, fields="files(id)",
    ).execute()
    media = MediaFileUpload(ruta_local, mimetype=mime_type, resumable=True)
    existentes = resultado.get("files", [])
    if existentes:
        return servicio.files().update(
            fileId=existentes[0]["id"], media_body=media, supportsAllDrives=True
        ).execute()
    metadata = {"name": nombre_archivo, "parents": [id_carpeta]}
    return servicio.files().create(
        body=metadata, media_body=media, supportsAllDrives=True, fields="id"
    ).execute()


# ============================================================
# CLASIFICACION DEL PDF: IESS / CAMPESINO / ISSFA / ISSPOL
# ============================================================
#
# El PDF de coberturasalud.msp.gob.ec trae SIEMPRE una tabla con 3 filas
# (IESS, ISSFA, ISSPOL) y, en la columna "Registro de Cobertura de
# Atención de Salud", cada fila dice "SI REGISTRA COBERTURA" o "NO
# REGISTRA COBERTURA". El motor actual (descargar_coberturas.py) solo
# mira la fila IESS (estado_cobertura_iess); aqui se agrega el mismo
# analisis para ISSFA e ISSPOL, sin tocar el archivo original.


def _estado_fila(texto_sin_tildes, etiqueta, limite):
    patron = rf"\b{re.escape(etiqueta)}\b(?:(?!\b{re.escape(limite)}\b).)*?(SI REGISTRA COBERTURA|NO REGISTRA COBERTURA)"
    m = re.search(patron, texto_sin_tildes, re.DOTALL)
    if not m:
        return None
    return m.group(1) == "SI REGISTRA COBERTURA"


def determinar_esquema_cobertura(texto_pdf, es_campesino_fn):
    """Devuelve 'ISSFA', 'ISSPOL', 'CAMPESINO' o 'IESS' segun cual fila de
    la tabla del PDF confirma cobertura ('SI REGISTRA COBERTURA'). Si
    ninguna fila la confirma, devuelve None (sin seguro / indeterminado;
    ese caso ya se reporta aparte como sin_seguro/errores y no debe
    subirse a ninguna carpeta).

    es_campesino_fn: pasar core.es_campesino (se inyecta desde afuera
    para no duplicar esa deteccion aqui)."""
    t = _quitar_tildes(texto_pdf or "").upper()

    if _estado_fila(t, "ISSFA", "ISSPOL") is True:
        return "ISSFA"
    if _estado_fila(t, "ISSPOL", "RED PRIVADA") is True:
        return "ISSPOL"
    if _estado_fila(t, "IESS", "ISSFA") is True:
        return "CAMPESINO" if es_campesino_fn(texto_pdf) else "IESS"
    return None


# ============================================================
# DIAGNOSTICO (para probar la conexion antes de un lote real)
# ============================================================


def probar_conexion(secrets_gcp, nombre_unidad):
    """Devuelve una lista de mensajes de diagnostico (en vez de lanzar
    excepciones sueltas), pensada para mostrarse directo en la app con
    st.write(...) antes de confiar en la integracion para un lote real."""
    mensajes = []
    try:
        servicio = obtener_servicio_drive(secrets_gcp)
        mensajes.append("✅ Autenticación con la cuenta de servicio: OK")
    except Exception as e:
        mensajes.append(f"❌ No se pudo autenticar: {e}")
        return mensajes

    try:
        drive_id = id_unidad_compartida(servicio)
        mensajes.append(f"✅ Unidad Compartida '{NOMBRE_UNIDAD_COMPARTIDA}' encontrada.")
    except Exception as e:
        mensajes.append(f"❌ {e}")
        return mensajes

    id_carpeta_unidad = buscar_subcarpeta(servicio, drive_id, drive_id, nombre_unidad)
    if id_carpeta_unidad:
        mensajes.append(f"✅ Carpeta de la unidad '{nombre_unidad}' encontrada.")
    else:
        mensajes.append(f"⚠️ No existe todavía una carpeta '{nombre_unidad}' dentro de "
                         f"'{NOMBRE_UNIDAD_COMPARTIDA}' (se creará sola en el primer uso).")

    return mensajes
