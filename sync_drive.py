# -*- coding: utf-8 -*-
"""
sync_drive.py - Sube a Google Drive lo que descarga la version de escritorio.

La version de escritorio sigue funcionando EXACTAMENTE igual (descarga los
PDF y llena la matriz en las carpetas locales). Este modulo solo agrega, al
terminar cada paciente (modo manual) o cada lote, un paso que:

  1. Sube cada PDF nuevo a  PLANILLAJE COBERTURAS/<unidad>/<año>/<mes>/IESS
     o /CAMPESINO  (segun la carpeta GENERAL/CAMPESINO donde el motor lo guardo).
  2. Sube la matriz INSTRUCTIVO del mes a la carpeta de ese mes.

Y al abrir el programa, si todavia no hay matriz del mes actual en el PC,
la baja de Drive.

Si Drive no esta configurado (o falla), el programa sigue funcionando como
siempre: los errores de Drive solo se muestran en el registro, nunca
detienen una descarga.

CONFIGURACION (una sola vez por PC, la pide el programa al abrirse y se
guarda en config_drive.json junto a este archivo):
  - la unidad (carpeta de la unidad dentro de PLANILLAJE COBERTURAS)
  - el archivo .json de la cuenta de servicio de Google

Requiere:  pip install google-api-python-client google-auth
"""

import os
import re
import json
import threading
from datetime import date

import descargar_coberturas as core

try:
    import drive_utils as drv
    _ERROR_IMPORTACION = None
except Exception as e:  # faltan las librerias de Google
    drv = None
    _ERROR_IMPORTACION = e

# carpeta en Drive (clave) -> nombre completo de la unidad
UNIDADES = {
    "COTOCOLLAO": "CENTRO DE SALUD COTOCOLLAO",
    "JAIME_ROLDOS": "CENTRO DE SALUD JAIME ROLDOS AGUILERA",
    "CARCELEN_BAJO": "CENTRO DE SALUD CARCELEN BAJO",
    "COLINAS_DEL_NORTE": "CENTRO DE SALUD COLINAS DEL NORTE",
    "CORAZON_DE_JESUS": "CENTRO DE SALUD CORAZON DE JESUS",
    "CALACALI": "CENTRO DE SALUD CALACALI",
    "EL_CONDADO": "CENTRO DE SALUD EL CONDADO",
    "PISULI": "CENTRO DE SALUD PISULI",
    "POMASQUI": "CENTRO DE SALUD POMASQUI",
    "SAN_ANTONIO": "CENTRO DE SALUD SAN ANTONIO",
    "COMITE_DEL_PUEBLO": "CENTRO DE SALUD COMITÉ DEL PUEBLO",
    "1E": "CENTRO DE SALUD 1E",
    "LA_BOTA": "CENTRO DE SALUD LA BOTA",
    "COTOCOLLAO_ALTO": "CENTRO DE SALUD COTOCOLLAO ALTO",
    "MENA_DEL_HIERRO": "CENTRO DE SALUD MENA DEL HIERRO",
    "NONO": "CENTRO DE SALUD NONO",
    "ATAHUALPA": "CENTRO DE SALUD ATAHUALPA",
    "CHAVEZPAMBA": "CENTRO DE SALUD CHAVEZPAMBA",
    "PERUCHO": "CENTRO DE SALUD PERUCHO",
    "PUELLARO": "CENTRO DE SALUD PUELLARO",
    "SAN_JOSE_DE_MINAS": "CENTRO DE SALUD SAN JOSE DE MINAS",
}

CONFIG_RUTA = os.path.join(core.BASE_DIR, "config_drive.json")
REGISTRO_RUTA = os.path.join(core.BASE_DIR, "drive_subidos.json")

_LOCK = threading.Lock()
_CACHE_CARPETAS = {}
_AVISO_CUOTA_DADO = False


# ============================================================
# CONFIGURACION
# ============================================================


def cargar_config():
    """Devuelve {'unidad': <carpeta>, 'credenciales': <ruta .json>} o None."""
    try:
        with open(CONFIG_RUTA, encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, ValueError):
        return None
    unidad = cfg.get("unidad")
    cred = cfg.get("credenciales")
    if not unidad or not cred:
        return None
    if not os.path.isabs(cred):
        cred = os.path.join(core.BASE_DIR, cred)
    if not os.path.isfile(cred):
        return None
    return {"unidad": unidad, "credenciales": cred}


def activo():
    return drv is not None and cargar_config() is not None


def configurar_primera_vez(root):
    """Si Drive no esta configurado en este PC, pregunta la unidad y el
    archivo de credenciales. Si se cancela, el programa sigue sin Drive."""
    if drv is None or cargar_config() is not None:
        return
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    if not messagebox.askyesno(
        "Google Drive",
        "Este PC todavia no esta conectado a Google Drive.\n\n"
        "¿Configurarlo ahora? (lo que se descargue se subira solo a Drive)\n\n"
        "Si dices que no, el programa funciona igual pero solo guarda en este PC.",
        parent=root,
    ):
        return

    nombres = sorted(UNIDADES.values())
    por_nombre = {v: k for k, v in UNIDADES.items()}
    resultado = {}

    dlg = tk.Toplevel(root)
    dlg.title("Configurar Google Drive")
    dlg.geometry("520x240")
    dlg.transient(root)
    dlg.grab_set()

    ttk.Label(dlg, text="Unidad de este PC:").pack(anchor="w", padx=14, pady=(14, 2))
    var_unidad = tk.StringVar()
    ttk.Combobox(dlg, textvariable=var_unidad, values=nombres, state="readonly", width=60).pack(padx=14)

    ttk.Label(dlg, text="Archivo de credenciales de Google (.json):").pack(anchor="w", padx=14, pady=(14, 2))
    var_cred = tk.StringVar()
    fila = ttk.Frame(dlg)
    fila.pack(fill="x", padx=14)
    ttk.Entry(fila, textvariable=var_cred, width=52).pack(side="left")
    ttk.Button(
        fila, text="Buscar...",
        command=lambda: var_cred.set(filedialog.askopenfilename(
            parent=dlg, filetypes=[("JSON", "*.json")]) or var_cred.get()),
    ).pack(side="left", padx=6)

    def guardar():
        if not var_unidad.get() or not os.path.isfile(var_cred.get()):
            messagebox.showwarning("Falta informacion", "Elige la unidad y el archivo .json.", parent=dlg)
            return
        resultado["unidad"] = por_nombre[var_unidad.get()]
        resultado["credenciales"] = var_cred.get()
        dlg.destroy()

    ttk.Button(dlg, text="Guardar", command=guardar).pack(pady=18)
    root.wait_window(dlg)

    if resultado:
        with open(CONFIG_RUTA, "w", encoding="utf-8") as f:
            json.dump(resultado, f, ensure_ascii=False, indent=2)


# ============================================================
# AUXILIARES
# ============================================================


def _servicio():
    cfg = cargar_config()
    with open(cfg["credenciales"], encoding="utf-8") as f:
        info = json.load(f)
    return drv.obtener_servicio_drive(info)


def _carpetas_de_mes(servicio, anio, mes):
    clave = (anio, mes)
    if clave not in _CACHE_CARPETAS:
        id_raiz = drv.id_carpeta_raiz(servicio)
        _CACHE_CARPETAS[clave] = drv.carpetas_de_mes(
            servicio, id_raiz, cargar_config()["unidad"], anio, mes
        )
    return _CACHE_CARPETAS[clave]


def _es_matriz(nombre):
    return bool(re.match(r"^INSTRUCTIVO.*\.xlsx$", nombre, re.IGNORECASE)) and not nombre.startswith("~$")


def _firma(ruta):
    return [os.path.getsize(ruta), int(os.path.getmtime(ruta))]


def _leer_registro():
    try:
        with open(REGISTRO_RUTA, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _guardar_registro(registro):
    try:
        with open(REGISTRO_RUTA, "w", encoding="utf-8") as f:
            json.dump(registro, f, ensure_ascii=False)
    except OSError:
        pass


def _explicar_error(e, log):
    """Mensaje claro para el error de cuota de las cuentas de servicio."""
    global _AVISO_CUOTA_DADO
    texto = str(e)
    if ("storageQuotaExceeded" in texto or "do not have storage quota" in texto) and not _AVISO_CUOTA_DADO:
        _AVISO_CUOTA_DADO = True
        log("   [Drive] Google no deja a la cuenta de servicio CREAR archivos nuevos en una carpeta "
            "de 'Mi unidad'. Hay que cambiar a autorizacion con la cuenta de Google duena de la "
            "carpeta (avisale a quien te ayuda con el programa).")


# ============================================================
# TRAER LA MATRIZ DEL MES (al abrir el programa)
# ============================================================


def traer_matriz_si_falta(log=print):
    """Si en este PC no hay matriz del mes actual y en Drive si, la baja."""
    if not activo():
        return
    try:
        hoy = date.today()
        for nombre in os.listdir(core.BASE_DIR):
            if _es_matriz(nombre) and drv.fecha_de_texto(nombre) == (hoy.year, hoy.month):
                return
        servicio = _servicio()
        id_mes, _ = _carpetas_de_mes(servicio, hoy.year, hoy.month)
        archivo = drv.buscar_archivo_por_patron(servicio, id_mes, r"^INSTRUCTIVO.*\.xlsx$")
        if archivo:
            drv.descargar_archivo(servicio, archivo["id"], os.path.join(core.BASE_DIR, archivo["name"]))
            log(f"[Drive] Matriz del mes descargada: {archivo['name']}")
    except Exception as e:
        log(f"[Drive] No se pudo traer la matriz del mes: {e}")


# ============================================================
# SUBIR LO DESCARGADO
# ============================================================


def sincronizar(log=print):
    """Sube a Drive los PDF y matrices nuevos o modificados. Nunca lanza
    excepciones: cualquier problema se informa por 'log'."""
    if drv is None:
        log(f"   [Drive] No disponible (falta instalar las librerias de Google: {_ERROR_IMPORTACION}).")
        return
    if not activo():
        return
    with _LOCK:
        try:
            _sincronizar(log)
        except Exception as e:
            log(f"   [Drive] ERROR: {e}")
            _explicar_error(e, log)


def _sincronizar(log):
    servicio = _servicio()
    registro = _leer_registro()
    hoy = date.today()
    subidos, fallos = 0, 0

    # 1) PDF
    if os.path.isdir(core.CARPETA_SALIDA):
        for raiz, carpetas, archivos in os.walk(core.CARPETA_SALIDA):
            carpetas[:] = [c for c in carpetas if c not in ("_descargas_temp", "_diagnostico")]
            for nombre in archivos:
                if not nombre.lower().endswith(".pdf"):
                    continue
                ruta = os.path.join(raiz, nombre)
                rel = os.path.relpath(ruta, core.CARPETA_SALIDA)
                clave = "pdf|" + rel
                firma = _firma(ruta)
                if registro.get(clave) == firma:
                    continue
                try:
                    anio_mes, carpeta_tipo, nombre_drive = drv.destino_de_pdf_local(rel)
                    if carpeta_tipo is None:
                        log(f"   [Drive] Omitido (ruta inesperada): {rel}")
                        registro[clave] = firma
                        continue
                    anio, mes = anio_mes or (hoy.year, hoy.month)
                    _, ids_tipo = _carpetas_de_mes(servicio, anio, mes)
                    drv.subir_o_reemplazar_archivo(servicio, ruta, nombre_drive, ids_tipo[carpeta_tipo], "application/pdf")
                    registro[clave] = firma
                    subidos += 1
                except Exception as e:
                    fallos += 1
                    log(f"   [Drive] No se pudo subir {rel}: {e}")
                    _explicar_error(e, log)

    # 2) matrices INSTRUCTIVO
    for nombre in os.listdir(core.BASE_DIR):
        if not _es_matriz(nombre):
            continue
        ruta = os.path.join(core.BASE_DIR, nombre)
        clave = "xlsx|" + nombre
        firma = _firma(ruta)
        if registro.get(clave) == firma:
            continue
        anio_mes = drv.fecha_de_texto(nombre)
        if anio_mes is None:
            continue
        try:
            id_mes, _ = _carpetas_de_mes(servicio, *anio_mes)
            drv.subir_o_reemplazar_archivo(servicio, ruta, nombre, id_mes, drv.MIME_XLSX)
            registro[clave] = firma
            subidos += 1
        except Exception as e:
            fallos += 1
            log(f"   [Drive] No se pudo subir la matriz {nombre}: {e}")
            _explicar_error(e, log)

    _guardar_registro(registro)
    if subidos or fallos:
        log(f"   [Drive] {subidos} archivo(s) subido(s) a Drive" + (f", {fallos} con error." if fallos else "."))
