# -*- coding: utf-8 -*-

"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Apunta directamente al ejecutable local sin llamar a ChromeDriverManager()
service = Service(executable_path='./chromedriver.exe')
driver = webdriver.Chrome(service=service)
# Apunta directamente al ejecutable local sin llamar a ChromeDriverManager()
service = Service(executable_path='./chromedriver.exe')
driver = webdriver.Chrome(service=service)
DESCARGA AUTOMATICA DE PDFs DE COBERTURA DE SALUD DESDE CORESALUD

Lee cedulas y fechas desde un archivo Excel y consulta:
https://coberturasalud.msp.gob.ec

Guarda los PDFs en carpetas por paciente, separadas en dos carpetas
raiz segun el tipo de afiliacion detectado en el propio PDF:

    PDF_DESCARGADOS/GENERAL/<paciente>/...
    PDF_DESCARGADOS/CAMPESINO/<paciente>/...

FUNCIONES:
- PACIENTE
- MENOR
- PADRE
- MADRE
- TITULAR
- COBERTURA IESS (solo CAMPESINO, desde app.iess.gob.ec)
- COBERTURA JEFE DE FAMILIA (solo CAMPESINO)
- No sobrescribe PDFs existentes
- No elimina PDFs anteriores
- No vuelve a descargar una fecha que ya existe
- Si aparece una fecha nueva, la descarga
- Detecta MADRE/PADRE desde el PDF (flujo GENERAL)
- Detecta CAMPESINO desde el PDF y enruta a la carpeta CAMPESINO
- Detecta si el CAMPESINO es titular (Jefe de Familia del SSC) o dependiente
- Si es dependiente, descarga tambien COBERTURA IESS (calificacion IESS) y,
  a partir del acreditador que ahi aparece, COBERTURA JEFE DE FAMILIA
- Usa el nombre como respaldo
- Reintentos de Chrome
- Diagnostico de errores
- log_errores.xlsx
- sin_seguro.xlsx
- Creación diferida de carpetas: No crea la carpeta del paciente si no hay PDFs descargados.

IMPORTANTE:
Este archivo mantiene:
    main(callback_fila=None, callback_progreso=None)
para que pueda seguir siendo utilizado por la interfaz u otro archivo Python.

NOTA SOBRE EL FLUJO CAMPESINO (app.iess.gob.ec):
Ese sitio no se pudo inspeccionar directamente durante el desarrollo (esta
bloqueado para herramientas automaticas de lectura web), por lo que los
selectores de Selenium se escribieron en base a capturas de pantalla y al
texto visible del formulario ("Identificación del Asegurado", "Fecha de
Consulta", "Contingencia", "Aceptar"), con varias estrategias de respaldo.
Si falla, revisa las capturas en PDF_DESCARGADOS/_diagnostico
(los archivos de la calificación IESS llevan "_iess_intento" en el nombre).
"""

# ============================================================
# IMPORTACIONES
# ============================================================

import os
import sys
import re
import time
import json
import base64
import shutil
import unicodedata
import contextlib
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.formula.translate import Translator

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
# LECTOR PDF
# ============================================================

try:
    from pypdf import PdfReader
except ImportError:
    from PyPDF2 import PdfReader

# ============================================================
# CONFIGURACION DE RUTA Y EXCEL
# ============================================================

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARCHIVO_EXCEL = os.path.join(BASE_DIR, "reporte_inconsistencias.xlsx")
HOJA = "Sheet1"

# ============================================================
# COLUMNAS DEL EXCEL
# ============================================================

COL_NOMBRE = "Nombre Excel"
COL_CEDULA = "Cédula Excel"
COL_PADRE = "COBERTURA PADRE"
COL_TITULAR = "COBERTURA TITULAR"
COL_FECHAS = "Fechas Excel"

# ============================================================
# ARCHIVO DE LA MATRIZ (INSTRUCTIVO) - MODO MANUAL / INTERACTIVO
# ============================================================
#
# Este es el archivo Excel "INSTRUCTIVO5_SEPTIEMBRE_2026.xlsx" que ya
# trae, en la hoja MES_2026, las formulas precargadas (codigo de
# dependencia, numero de expediente, edad por fecha de nacimiento,
# codigo de tipo de seguro, codigo de parentesco, etc.). El modo manual
# (modo_interactivo) solo escribe en las columnas que son de llenado
# manual; las columnas con formula se dejan intactas para que Excel las
# siga calculando solas.
ARCHIVO_MATRIZ = os.path.join(BASE_DIR, "INSTRUCTIVO5_SEPTIEMBRE_2026.xlsx")
HOJA_MATRIZ = "MES_2026"
HOJA_MAESTRO = "MAESTRO"

# Carpeta donde se guarda una copia de respaldo del archivo de la matriz
# cada vez que el modo manual va a escribir en el, por seguridad (nunca
# se sobrescribe el archivo sin dejar una copia previa).
CARPETA_RESPALDOS_MATRIZ = os.path.join(BASE_DIR, "_respaldos_matriz")

# Archivo donde se recuerdan datos de pacientes ya ingresados antes
# (fecha de nacimiento, sexo, nombre), para no tener que volver a
# escribirlos cada vez que se ingresa la misma cedula.
RUTA_DATOS_PACIENTES = os.path.join(BASE_DIR, "datos_pacientes.json")

# Columna (1 = A) de la matriz que se usa para detectar cual es la
# siguiente fila vacia: "C.I PACIENTE".
COL_MATRIZ_CEDULA = 8

# ------------------------------------------------------------------
# Textos de "TIPO SEGURO (BENEFICIARIO)" (columna F de la matriz),
# copiados EXACTAMENTE como aparecen en MAESTRO!G4:G14, para que el
# VLOOKUP de la columna G (codigo) los reconozca. Se asignan solos,
# segun lo que el propio motor de descarga ya determino (campesino o
# no, titular o dependiente, menor o adulto). Si alguna vez el maestro
# cambia esos textos, hay que actualizarlos aqui tambien.
# ------------------------------------------------------------------
TIPO_SEGURO_CAMPESINO_TITULAR = "SEGURO SOCIAL CAMPESINO"
TIPO_SEGURO_CAMPESINO_JUBILADO = "SEGURO SOCIAL CAMPESINO JUBILADO"
TIPO_SEGURO_CAMPESINO_DEPENDIENTE = "SEGURO SOCIAL CAMPESINO"
TIPO_SEGURO_GENERAL_TITULAR = "ACTIVO (SEGURO GENERAL)"
TIPO_SEGURO_GENERAL_CONYUGE = "CONYUGUE"
TIPO_SEGURO_GENERAL_MENOR = "MENORES DEPENDIENTES"

# Textos de "PARENTESCO" (columna Q de la matriz), copiados EXACTAMENTE
# como aparecen en MAESTRO!P5:P8.
PARENTESCO_CONYUGE = "CONYUGE"
PARENTESCO_HIJO = "HIJO/HIJA"
PARENTESCO_PARIENTE = "PARIENTE"
PARENTESCO_TITULAR = "TITULAR"

# ============================================================
# CARPETAS Y URL
# ============================================================

CARPETA_SALIDA = os.path.join(BASE_DIR, "PDF_DESCARGADOS")
URL_CORESALUD = "https://coberturasalud.msp.gob.ec"

# Sitio de "Calificación Atención Médica" del IESS, usado UNICAMENTE
# para pacientes del seguro CAMPESINO que resultan ser dependientes
# (no titulares). De aqui sale el PDF "COBERTURA IESS".
URL_IESS_CALIFICACION = "https://app.iess.gob.ec/gestion-calificacion-derecho-web/public/formulariosContacto.jsf"
CONTINGENCIA_IESS = "Enfermedad"  # Siempre es "Enfermedad" segun lo indicado.

# Nombres de las subcarpetas raiz segun el tipo de afiliacion detectado.
SUBCARPETA_GENERAL = "GENERAL"
SUBCARPETA_CAMPESINO = "CAMPESINO"

MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]

def nombre_mes_es(fecha):
    return MESES_ES[fecha.month - 1]

def nombre_carpeta_mes(fecha):
    """Nombre de la carpeta del mes para una fecha dada, ej: 'SEPTIEMBRE
    2026'. Se usa tanto para organizar los PDF descargados como para
    nombrar el archivo de la matriz de cada mes."""
    return f"{nombre_mes_es(fecha)} {fecha.year}"

# ============================================================
# CONFIGURACION DE EJECUCION
# ============================================================

LIMITE_PRUEBA = None  # None = todos, o un entero ej: 3
ESPERA_MAX = 25
MODO_HEADLESS = True
PAUSA_ENTRE_CONSULTAS = 2
MAX_INTENTOS_POR_FECHA = 3

# Cache global para evitar releer archivos PDF varias veces en disco
_CACHE_TEXTO_PDF = {}

# Mapeo de meses (Español e Inglés) a número de mes
MESES_MAPA = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
    "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12,
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
    "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12
}

MESES_INGLES = [
    "", "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]
MESES_ESPANOL = [
    "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

# ============================================================
# QUITAR TILDES
# ============================================================

def quitar_tildes(txt):
    if txt is None:
        return ""
    txt = str(txt)
    txt = unicodedata.normalize("NFKD", txt)
    return "".join(c for c in txt if not unicodedata.combining(c)).upper()

# ============================================================
# PARSEAR FECHAS DEL EXCEL
# ============================================================

def parsear_fechas(texto_celda):
    fechas = []
    if texto_celda is None:
        return fechas

    texto_celda = str(texto_celda)

    for y, m, d in re.findall(r"(\d{4}),\s*(\d{1,2}),\s*(\d{1,2})", texto_celda):
        try:
            fechas.append(date(int(y), int(m), int(d)))
        except ValueError:
            pass

    for y, m, d in re.findall(r"(?<!\d)(\d{4})-(\d{1,2})-(\d{1,2})(?!\d)", texto_celda):
        try:
            fechas.append(date(int(y), int(m), int(d)))
        except ValueError:
            pass

    return sorted(set(fechas))

# ============================================================
# LEER EXCEL
# ============================================================

def leer_excel(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    ws = wb[HOJA]

    encabezados = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]

    idx_nombre = encabezados.index(COL_NOMBRE)
    idx_cedula = encabezados.index(COL_CEDULA)
    idx_fechas = encabezados.index(COL_FECHAS)

    idx_padre = encabezados.index(COL_PADRE) if COL_PADRE in encabezados else None
    idx_titular = encabezados.index(COL_TITULAR) if COL_TITULAR in encabezados else None

    def leer_cedula_opcional(row, idx, cedula_paciente):
        if idx is None:
            return None
        raw = row[idx]
        if raw in (None, ""):
            return None
        try:
            val = str(int(raw)).zfill(10)
        except Exception:
            val = str(raw).strip().zfill(10)

        if val == cedula_paciente:
            return None
        return val

    registros = []

    for row in ws.iter_rows(min_row=2, values_only=True):
        nombre = row[idx_nombre]
        cedula_raw = row[idx_cedula]
        fechas_raw = row[idx_fechas]

        if not nombre or cedula_raw is None or not fechas_raw:
            continue

        try:
            cedula = str(int(cedula_raw)).zfill(10)
        except Exception:
            cedula = str(cedula_raw).strip().zfill(10)

        fechas = parsear_fechas(str(fechas_raw))
        cedula_padre = leer_cedula_opcional(row, idx_padre, cedula)
        cedula_titular = leer_cedula_opcional(row, idx_titular, cedula)

        if fechas:
            registros.append({
                "cedula": cedula,
                "nombre": str(nombre).strip(),
                "fechas": fechas,
                "cedula_padre": cedula_padre,
                "cedula_titular": cedula_titular
            })

    return registros

# ============================================================
# CREAR DRIVER
# ============================================================

def _limpiar_lock_wdm():
    """webdriver-manager usa un archivo de bloqueo (.wdm-lock-...) para
    evitar que dos procesos descarguen el chromedriver al mismo tiempo.
    Si un proceso anterior se cerro de golpe (se cerro la ventana, se
    mato el proceso, se quedo sin internet a mitad de la descarga, etc.)
    ese archivo se queda huerfano y todas las ejecuciones siguientes
    truenan con "Timed out waiting for webdriver-manager lock", aunque
    en realidad no haya nada corriendo. Aqui se borran esos archivos de
    bloqueo antes de intentar instalar/usar el driver, tanto en la
    carpeta por defecto de webdriver-manager (~/.wdm) como en la que
    indique la variable de entorno WDM_LOCAL si esta definida."""
    carpetas = [os.path.join(os.path.expanduser("~"), ".wdm")]
    wdm_local = os.environ.get("WDM_LOCAL")
    if wdm_local:
        carpetas.append(os.path.join(wdm_local, ".wdm"))

    for carpeta in carpetas:
        if not os.path.isdir(carpeta):
            continue
        try:
            for nombre in os.listdir(carpeta):
                if ".wdm-lock" in nombre:
                    ruta = os.path.join(carpeta, nombre)
                    try:
                        os.remove(ruta)
                        print(f"   (Se limpio un archivo de bloqueo antiguo: {ruta})")
                    except Exception:
                        pass
        except Exception:
            pass


def _instalar_chromedriver(reintentos=3):
    """Llama a ChromeDriverManager().install() con reintentos. Si truena
    por el bloqueo (.wdm-lock) o por timeout de red, limpia el lock y
    reintenta antes de rendirse."""
    ultimo_error = None
    for intento in range(1, reintentos + 1):
        try:
            return ChromeDriverManager().install()
        except Exception as e:
            ultimo_error = e
            print(f"   Aviso: fallo instalando/verificando chromedriver "
                  f"(intento {intento}/{reintentos}): {e}")
            _limpiar_lock_wdm()
            time.sleep(2)
    raise ultimo_error


def crear_driver(carpeta_descargas):
    # Rutas tipicas del navegador y su driver cuando se instalan por apt
    # en un contenedor Linux (Streamlit Community Cloud, Render, etc,
    # via un archivo "packages.txt" con "chromium" y "chromium-driver").
    # Si existen, se usan directamente -mas rapido y confiable que dejar
    # que webdriver_manager intente detectar la version del navegador el
    # mismo-. Si no existen (por ejemplo en la PC de escritorio con
    # Windows), se sigue exactamente igual que siempre.
    _RUTAS_CHROMIUM = ("/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome")
    _RUTAS_CHROMEDRIVER = (
        "/usr/bin/chromedriver", "/usr/lib/chromium/chromedriver", "/usr/lib/chromium-browser/chromedriver"
    )
    binario_chromium = next((r for r in _RUTAS_CHROMIUM if os.path.exists(r)), None)
    ruta_chromedriver_sistema = next((r for r in _RUTAS_CHROMEDRIVER if os.path.exists(r)), None)

    opciones = webdriver.ChromeOptions()
    prefs = {
        "download.default_directory": os.path.abspath(carpeta_descargas),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    opciones.add_experimental_option("prefs", prefs)
    opciones.add_argument("--disable-popup-blocking")

    if binario_chromium:
        opciones.binary_location = binario_chromium

    # Fuerza una escala de renderizado fija (1.0), sin importar la escala
    # de pantalla/DPI configurada en Windows en cada PC (Windows 11 suele
    # traer 125%/150% por defecto en pantallas nuevas, Windows 10 a veces
    # 100%). Sin esto, Page.printToPDF y las capturas de pantalla pueden
    # renderizar el texto con distinto ajuste de columnas segun la PC,
    # aunque el sitio responda exactamente igual.
    opciones.add_argument("--force-device-scale-factor=1")
    opciones.add_argument("--high-dpi-support=1")

    # Reduce diferencias de compatibilidad entre instalaciones de Windows
    # (permisos, perfiles nuevos, antivirus/politicas locales), y es
    # obligatorio ademas para correr como root dentro de un contenedor
    # Linux (Streamlit Cloud, Render, etc).
    opciones.add_argument("--no-sandbox")
    opciones.add_argument("--disable-dev-shm-usage")

    if MODO_HEADLESS:
        opciones.add_argument("--headless=new")
        opciones.add_argument("--window-size=1400,1000")
        opciones.add_argument("--disable-gpu")

    if ruta_chromedriver_sistema:
        servicio = Service(ruta_chromedriver_sistema)
    else:
        _limpiar_lock_wdm()
        servicio = Service(_instalar_chromedriver())

    driver = webdriver.Chrome(service=servicio, options=opciones)
    driver.set_script_timeout(20)

    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": os.path.abspath(carpeta_descargas)
            }
        )
    except Exception:
        pass

    return driver

# ============================================================
# AUXILIARES CHROME / DESCARGAS
# ============================================================

def _archivos_pdf_listos(carpeta):
    try:
        nombres = os.listdir(carpeta)
    except FileNotFoundError:
        return set()

    en_progreso = any(n.endswith(".crdownload") or n.endswith(".tmp") for n in nombres)
    if en_progreso:
        return set()

    return {n for n in nombres if n.lower().endswith(".pdf")}

def _mensaje_error_visible(driver):
    selectores = [
        "[role='alert']", ".alert", ".swal2-popup", ".swal2-html-container",
        ".toast", ".Toastify__toast", ".error", ".mensaje-error",
        ".p-toast-message", ".ant-message"
    ]
    for selector in selectores:
        try:
            elementos = driver.find_elements(By.CSS_SELECTOR, selector)
            for el in elementos:
                if el.is_displayed():
                    texto = (el.text or "").strip()
                    if texto:
                        return texto
        except Exception:
            continue
    return None

def esperar_pdf_o_blob(driver, carpeta_descargas_temp, archivos_previos, timeout):
    fin = time.time() + timeout
    while time.time() < fin:
        actuales = _archivos_pdf_listos(carpeta_descargas_temp)
        nuevos = actuales - archivos_previos
        if nuevos:
            nombre = sorted(nuevos)[0]
            return ("archivo", os.path.join(carpeta_descargas_temp, nombre))

        blob_url = None
        for tag in ("iframe", "embed", "object"):
            try:
                elementos = driver.find_elements(By.TAG_NAME, tag)
            except Exception:
                elementos = []
            for el in elementos:
                try:
                    src = el.get_attribute("src") or el.get_attribute("data") or ""
                except Exception:
                    src = ""
                if "blob:" in src:
                    blob_url = src
                    break
            if blob_url:
                break

        if not blob_url and driver.current_url.startswith("blob:"):
            blob_url = driver.current_url

        if blob_url:
            return ("blob", blob_url)

        mensaje = _mensaje_error_visible(driver)
        if mensaje:
            return ("mensaje", mensaje)

        time.sleep(0.5)

    return (None, None)

def consultar_y_obtener_pdf_bytes(driver, cedula, fecha_ddmmyyyy, carpeta_descargas_temp):
    driver.get(URL_CORESALUD)

    campo_cedula = WebDriverWait(driver, ESPERA_MAX).until(
        EC.presence_of_element_located((By.ID, "cedula"))
    )
    campo_cedula.clear()
    campo_cedula.send_keys(cedula)

    from selenium.webdriver.common.keys import Keys
    campo_fecha = WebDriverWait(driver, ESPERA_MAX).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="DD-MM-YYYY"]'))
    )
    campo_fecha.click()
    campo_fecha.send_keys(Keys.CONTROL, "a")
    campo_fecha.send_keys(Keys.DELETE)
    campo_fecha.send_keys(fecha_ddmmyyyy)
    campo_fecha.send_keys(Keys.TAB)

    archivos_previos = _archivos_pdf_listos(carpeta_descargas_temp)

    boton = WebDriverWait(driver, ESPERA_MAX).until(
        EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Consultar']"))
    )
    boton.click()

    tipo, valor = esperar_pdf_o_blob(driver, carpeta_descargas_temp, archivos_previos, ESPERA_MAX)

    if tipo == "archivo":
        time.sleep(0.3)
        with open(valor, "rb") as f:
            pdf_bytes = f.read()
        try:
            os.remove(valor)
        except OSError:
            pass
        return pdf_bytes

    if tipo == "blob":
        resultado_b64 = driver.execute_async_script(
            """
            var uri = arguments[0];
            var callback = arguments[1];
            fetch(uri)
                .then(r => r.blob())
                .then(blob => {
                    var reader = new FileReader();
                    reader.onloadend = function() {
                        callback(reader.result.split(',')[1]);
                    };
                    reader.readAsDataURL(blob);
                })
                .catch(err => callback('ERROR:' + err));
            """,
            valor
        )
        if resultado_b64.startswith("ERROR:"):
            raise RuntimeError("No se pudo leer el blob del PDF: " + resultado_b64)
        return base64.b64decode(resultado_b64)

    if tipo == "mensaje":
        raise RuntimeError("El sitio respondio: " + valor)

    raise RuntimeError("Tiempo de espera agotado: no aparecio PDF ni mensaje en pantalla")

# ============================================================
# SITIO IESS - CALIFICACION ATENCION MEDICA (solo CAMPESINO)
# ============================================================
#
# A diferencia de coberturasalud.msp.gob.ec, este sitio no entrega un PDF
# descargable: muestra el resultado como una pagina web normal y hay que
# "Guardar como PDF" (equivalente a Page.printToPDF de Chrome DevTools).
#
# El formulario tiene 3 campos visibles:
#   "Identificación del Asegurado: *"  -> texto (cedula)
#   "Fecha de Consulta:"               -> campo de fecha (con icono calendario)
#   "Contingencia: *"                  -> lista desplegable, aqui siempre "Enfermedad"
# y un boton "Aceptar".
#
# No fue posible inspeccionar el HTML real del sitio (bloqueado para
# herramientas automaticas), asi que se usan varias estrategias de
# localizacion de elementos por texto visible, con reintentos y capturas
# de diagnostico si algo falla.

def _ubicar_campo_por_etiqueta(driver, texto_etiqueta):
    """Busca un input/textarea cercano a una etiqueta con el texto dado."""
    candidatos = [
        f"//label[contains(normalize-space(.), '{texto_etiqueta}')]/following::input[1]",
        f"//*[self::label or self::span or self::div][contains(normalize-space(.), '{texto_etiqueta}')]/following::input[1]",
        f"//*[contains(normalize-space(text()), '{texto_etiqueta}')]/following::input[1]",
    ]
    for xpath in candidatos:
        try:
            el = driver.find_element(By.XPATH, xpath)
            if el.is_displayed():
                return el
        except Exception:
            continue
    return None

def _click_robusto(driver, elemento):
    """Intenta un clic normal de Selenium; si algo lo intercepta (overlay,
    elemento fuera de vista, etc.) cae a un clic disparado por JS."""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elemento)
    except Exception:
        pass
    try:
        elemento.click()
        return True
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", elemento)
            return True
        except Exception:
            return False

def _click_boton_aceptar(driver, timeout=None):
    """
    Localiza y hace clic en el boton "Aceptar" (id real:
    formConsulta:btnAceptar), con un mini-reintento propio. Se necesita
    porque en este sitio el boton arranca deshabilitado y se habilita via
    una actualizacion AJAX de PrimeFaces al elegir Contingencia: justo
    despues de habilitarse, el elemento puede quedar "stale" por otra
    actualizacion, o el panel desplegable puede tapar el punto del clic
    un instante (ElementClickInterceptedException). Reintentar localizando
    de nuevo cada vez (en vez de reusar la misma referencia) resuelve
    ambos casos.
    """
    if timeout is None:
        timeout = ESPERA_MAX

    xpath_boton = (
        "//button[@id='formConsulta:btnAceptar'] | "
        "//button[normalize-space()='Aceptar'] | "
        "//input[@type='submit' and contains(@value,'Aceptar')] | "
        "//*[@role='button' and contains(normalize-space(.),'Aceptar')]"
    )

    fin = time.time() + timeout
    ultimo_error = None
    while time.time() < fin:
        try:
            boton = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xpath_boton))
            )
            if _click_robusto(driver, boton):
                return True
        except Exception as e:
            ultimo_error = e
        time.sleep(0.5)

    if ultimo_error:
        raise ultimo_error
    raise RuntimeError("No se pudo hacer clic en 'Aceptar'")

def _seleccionar_contingencia(driver, valor="Enfermedad"):
    """
    Selecciona la opcion de Contingencia. Estructura REAL confirmada con
    el HTML del sitio (no una suposicion):

        <div id="formConsulta:contingencia_select" class="ui-selectonemenu ..."
             role="combobox" aria-haspopup="listbox" aria-expanded="false">
            <select id="formConsulta:contingencia_select_input" ...>  (oculto, accesibilidad)
                <option value="">Seleccione...</option>
                <option value="2">Maternidad</option>
                <option value="14">Enfermedad</option>
                <option value="15">Emergencia</option>
            </select>
            <label id="formConsulta:contingencia_select_label">Seleccione...</label>
            <div class="ui-selectonemenu-trigger">...</div>
        </div>
        <div id="formConsulta:contingencia_select_panel" class="... ui-helper-hidden ...">
            <ul id="formConsulta:contingencia_select_items">
                <li id="formConsulta:contingencia_select_2" data-label="Enfermedad">Enfermedad</li>
                ...
            </ul>
        </div>

    OJO: el boton "Aceptar" (formConsulta:btnAceptar) arranca deshabilitado
    (disabled="disabled") y SOLO se habilita cuando el <select> oculto
    dispara su evento "change" real (el sitio tiene un onchange inline que
    hace una llamada PrimeFaces.ab(...) para habilitarlo). Por eso aqui se
    hace clic de verdad sobre el <li> de la lista (como lo haria una
    persona) en vez de solo cambiar el <select> por JavaScript: es la
    forma mas confiable de que esa llamada se dispare igual que en un uso
    normal del sitio.
    """
    id_widget = "formConsulta:contingencia_select"

    try:
        widget = driver.find_element(By.ID, id_widget)
    except Exception:
        # Respaldo por si el id cambiara en el futuro.
        candidatos = driver.find_elements(
            By.XPATH, "//*[contains(@id,'contingencia') and contains(@class,'ui-selectonemenu')]"
        )
        widget = candidatos[0] if candidatos else None

    if widget is None:
        return False

    xpath_opcion = (
        f"//li[@data-label='{valor}' and contains(@id,'contingencia_select')]"
        f" | //li[contains(@class,'ui-selectonemenu-item') and normalize-space(text())='{valor}']"
    )

    for intento_abrir in range(2):  # a veces el primer clic no alcanza a abrir el panel
        if not _click_robusto(driver, widget):
            continue
        try:
            opcion_el = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.XPATH, xpath_opcion))
            )
        except Exception:
            continue

        if _click_robusto(driver, opcion_el):
            # Confirmamos que la etiqueta visible realmente cambio (esto
            # demuestra que PrimeFaces proceso el clic, no solo que el
            # <li> existia en el DOM).
            try:
                WebDriverWait(driver, 5).until(
                    lambda d: quitar_tildes(
                        d.find_element(By.ID, f"{id_widget}_label").text
                    ) == quitar_tildes(valor)
                )
            except Exception:
                continue  # ni la etiqueta cambio: reintentar el ciclo de apertura

            # El panel de opciones a veces se queda abierto (visualmente)
            # encima del boton "Aceptar" y luego intercepta el clic. Nos
            # aseguramos de que se cierre antes de seguir.
            try:
                WebDriverWait(driver, 5).until(
                    EC.invisibility_of_element_located((By.ID, f"{id_widget}_panel"))
                )
            except Exception:
                # Si no se cerro solo, lo cerramos nosotros haciendo clic
                # de nuevo en el widget (togglea abrir/cerrar) o con Escape.
                try:
                    from selenium.webdriver.common.keys import Keys
                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                except Exception:
                    pass
                time.sleep(0.3)

            return True

    return False

def _establecer_fecha_iess(driver, fecha_ddmmyyyy):
    """
    Escribe la fecha en el campo "Fecha de Consulta". Es un p:calendar de
    PrimeFaces cuyo input real tiene id "formConsulta:fec_calendar_input"
    (confirmado con el HTML del sitio) y esta marcado como readonly a
    proposito, para forzar el uso del calendario visual. En vez de
    navegar ese calendario (fragil), se escribe el valor directamente
    por JavaScript en el input y se disparan los eventos que PrimeFaces
    escucha, que es lo que finalmente se envia en el formulario.
    """
    candidatos_id = [
        "formConsulta:fec_calendar_input",
        "formConsulta:fecha_calendar_input",
    ]
    campo = None
    for id_campo in candidatos_id:
        try:
            campo = driver.find_element(By.ID, id_campo)
            break
        except Exception:
            continue

    if campo is None:
        try:
            campo = driver.find_element(By.CSS_SELECTOR, "input.hasDatepicker")
        except Exception:
            campo = None

    if campo is None:
        campo = _ubicar_campo_por_etiqueta(driver, "Fecha de Consulta")

    if campo is None:
        return False

    valor = fecha_ddmmyyyy.replace("-", "/")  # dd/mm/yyyy, igual al formato que devuelve el sitio

    try:
        driver.execute_script(
            """
            var input = arguments[0], valor = arguments[1];
            input.removeAttribute('readonly');
            input.value = valor;
            input.dispatchEvent(new Event('input', {bubbles: true}));
            input.dispatchEvent(new Event('change', {bubbles: true}));
            input.dispatchEvent(new Event('blur', {bubbles: true}));
            """,
            campo, valor
        )
        return True
    except Exception:
        return False

def _guardar_diagnostico_iess(driver, carpeta_diagnostico, nombre_base):
    """Guarda una captura de pantalla Y el HTML de la pagina en el momento
    del fallo. El HTML es mucho mas util que la imagen para ajustar los
    selectores (permite ver ids/clases reales sin depender de otra
    captura de pantalla)."""
    try:
        os.makedirs(carpeta_diagnostico, exist_ok=True)
    except Exception:
        return
    try:
        driver.save_screenshot(os.path.join(carpeta_diagnostico, f"{nombre_base}.png"))
    except Exception:
        pass
    try:
        with open(os.path.join(carpeta_diagnostico, f"{nombre_base}.html"), "w", encoding="utf-8") as f:
            f.write(driver.page_source)
    except Exception:
        pass

def _describir_error(e):
    """Descripcion legible de una excepcion de Selenium. Muchas veces
    str(e) trae la primera linea vacia (por ejemplo "Message: " sin nada
    mas en TimeoutException sin mensaje explicito), asi que se busca la
    primera linea NO vacia, y si no hay ninguna se usa el nombre de la
    clase de la excepcion."""
    texto = str(e).strip()
    lineas = [l.strip() for l in texto.splitlines() if l.strip()]
    if lineas:
        return f"{type(e).__name__}: {lineas[0]}"
    return type(e).__name__

def consultar_iess_calificacion_pdf(driver, cedula, fecha_ddmmyyyy, carpeta_diagnostico):
    """
    Consulta https://app.iess.gob.ec/.../formulariosContacto.jsf para una
    cedula y fecha, y devuelve los bytes del PDF generado con
    Page.printToPDF (equivalente a "Guardar como PDF").
    """
    from selenium.webdriver.common.keys import Keys

    driver.get(URL_IESS_CALIFICACION)

    try:
        campo_cedula = WebDriverWait(driver, ESPERA_MAX).until(
            EC.presence_of_element_located((By.ID, "formConsulta:cedula_text"))
        )
    except Exception:
        campo_cedula = WebDriverWait(driver, ESPERA_MAX).until(
            lambda d: _ubicar_campo_por_etiqueta(d, "Identificación del Asegurado")
            or _ubicar_campo_por_etiqueta(d, "Identificacion del Asegurado")
        )
    campo_cedula.click()
    campo_cedula.send_keys(Keys.CONTROL, "a")
    campo_cedula.send_keys(Keys.DELETE)
    campo_cedula.send_keys(cedula)

    _establecer_fecha_iess(driver, fecha_ddmmyyyy)

    if not _seleccionar_contingencia(driver, CONTINGENCIA_IESS):
        raise RuntimeError("No se pudo seleccionar la Contingencia 'Enfermedad'")

    # El boton arranca deshabilitado y solo se habilita tras la llamada
    # que dispara la seleccion de Contingencia; puede tardar un poco.
    # Ademas, justo despues de habilitarse puede quedar "stale" por otra
    # actualizacion AJAX, o el panel de Contingencia puede tapar el clic
    # un instante: por eso se localiza y hace clic en un mini-reintento,
    # en vez de una sola vez.
    _click_boton_aceptar(driver)

    def _resultado_listo(d):
        # OJO: "CALIFICACION ATENCION MEDICA" aparece TANTO en el
        # formulario como en el resultado, asi que no sirve para
        # distinguir si ya se cargo el resultado (eso causaba que se
        # tomara la pagina "Verificando... Protegido por ALTCHA" como si
        # ya fuera el resultado final). Se usan marcadores que SOLO
        # aparecen en la pagina de resultado.
        fuente_norm = quitar_tildes(d.page_source)

        if "VERIFICANDO" in fuente_norm and "ALTCHA" in fuente_norm:
            return False  # el captcha automatico del sitio aun esta verificando

        if "APELLIDOS Y NOMBRES" in fuente_norm:
            return True  # pagina de resultado, con datos del asegurado

        if "CON COBERTURA IESS" in fuente_norm or "SIN COBERTURA IESS" in fuente_norm:
            return True

        if _mensaje_error_visible(d):
            return True

        return False

    try:
        WebDriverWait(driver, ESPERA_MAX).until(_resultado_listo)
    except Exception:
        _guardar_diagnostico_iess(driver, carpeta_diagnostico, f"{cedula}_{fecha_ddmmyyyy}_timeout")
        raise RuntimeError("Tiempo de espera agotado esperando el resultado de la calificacion IESS")

    mensaje = _mensaje_error_visible(driver)
    if mensaje:
        raise RuntimeError("El sitio IESS respondio: " + mensaje)

    resultado = driver.execute_cdp_cmd("Page.printToPDF", {
        "printBackground": True,
        "preferCSSPageSize": True
    })
    return base64.b64decode(resultado["data"])

def consultar_iess_con_reintentos(driver_holder, cedula, fecha_str, carpeta_descargas_temp, carpeta_diagnostico):
    ultimo_error = None
    for intento in range(1, MAX_INTENTOS_POR_FECHA + 1):
        try:
            if not sesion_viva(driver_holder[0]):
                print("   Sesion de Chrome caida, reiniciando...")
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass
                driver_holder[0] = crear_driver(carpeta_descargas_temp)

            pdf = consultar_iess_calificacion_pdf(
                driver_holder[0], cedula, fecha_str, carpeta_diagnostico
            )
            return (pdf, None)

        except Exception as e:
            ultimo_error = _describir_error(e)
            print(f"      (IESS) intento {intento}/{MAX_INTENTOS_POR_FECHA} fallo: {ultimo_error}")

            _guardar_diagnostico_iess(
                driver_holder[0], carpeta_diagnostico, f"{cedula}_{fecha_str}_iess_intento{intento}"
            )

            try:
                driver_holder[0].quit()
            except Exception:
                pass

            driver_holder[0] = crear_driver(carpeta_descargas_temp)
            time.sleep(2)

    return (None, ultimo_error)

# ============================================================
# EXTRAER TEXTO DE PDF
# ============================================================

def texto_pdf(pdf_bytes):
    import io
    lector = PdfReader(io.BytesIO(pdf_bytes))
    textos = []
    for pagina in lector.pages:
        try:
            textos.append(pagina.extract_text() or "")
        except Exception:
            textos.append("")
    return "\n".join(textos)

# ============================================================
# FECHA YA DESCARGADA
# ============================================================

def fecha_ya_descargada(carpeta_paciente, fecha_objetivo, tipo_cobertura=None):
    if not os.path.exists(carpeta_paciente):
        return False

    if not isinstance(fecha_objetivo, date):
        return False

    try:
        archivos = os.listdir(carpeta_paciente)
    except Exception:
        return False

    f_dd_mm_yyyy = fecha_objetivo.strftime("%d-%m-%Y")
    f_dd_mm_yyyy_slash = fecha_objetivo.strftime("%d/%m/%Y")
    f_yyyy_mm_dd = fecha_objetivo.strftime("%Y-%m-%d")

    dia_str = f"{fecha_objetivo.day:02d}"
    mes_num = fecha_objetivo.month
    ano_str = str(fecha_objetivo.year)

    mes_en = MESES_INGLES[mes_num]
    mes_es = MESES_ESPANOL[mes_num]

    f_texto_en = f"{dia_str} {mes_en} {ano_str}"
    f_texto_es = f"{dia_str} {mes_es} {ano_str}"

    tipo_norm = quitar_tildes(tipo_cobertura) if tipo_cobertura else None

    for archivo in archivos:
        if not archivo.lower().endswith(".pdf"):
            continue

        if tipo_norm:
            nombre_archivo_norm = quitar_tildes(archivo)
            if f"COBERTURA {tipo_norm}" not in nombre_archivo_norm:
                continue

        ruta_pdf = os.path.join(carpeta_paciente, archivo)

        try:
            if ruta_pdf not in _CACHE_TEXTO_PDF:
                with open(ruta_pdf, "rb") as f:
                    contenido = f.read()
                _CACHE_TEXTO_PDF[ruta_pdf] = texto_pdf(contenido)

            txt = _CACHE_TEXTO_PDF[ruta_pdf]

            if (f_dd_mm_yyyy in txt or 
                f_dd_mm_yyyy_slash in txt or 
                f_yyyy_mm_dd in txt or
                f_texto_en.lower() in txt.lower() or
                f_texto_es.lower() in txt.lower()):
                return True

            coincidencias_texto = re.findall(r"\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b", txt)
            for d, m_txt, y in coincidencias_texto:
                m_upper = m_txt.upper()
                if m_upper in MESES_MAPA:
                    try:
                        if date(int(y), MESES_MAPA[m_upper], int(d)) == fecha_objetivo:
                            return True
                    except ValueError:
                        pass

            coincidencias_num1 = re.findall(r"\b(\d{1,2})[-/](\d{1,2})[-/](\d{4})\b", txt)
            for d, m, y in coincidencias_num1:
                try:
                    if date(int(y), int(m), int(d)) == fecha_objetivo:
                        return True
                except ValueError:
                    pass

            coincidencias_num2 = re.findall(r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b", txt)
            for y, m, d in coincidencias_num2:
                try:
                    if date(int(y), int(m), int(d)) == fecha_objetivo:
                        return True
                except ValueError:
                    pass

        except Exception:
            continue

    return False

def _texto_contiene_fecha(txt, fecha_objetivo):
    """Misma logica de coincidencia de fechas que fecha_ya_descargada,
    pero aplicada directamente a un texto ya extraido (para poder
    reutilizar el texto de un PDF que ya existe en disco sin duplicar
    la funcion fecha_ya_descargada original)."""
    f_dd_mm_yyyy = fecha_objetivo.strftime("%d-%m-%Y")
    f_dd_mm_yyyy_slash = fecha_objetivo.strftime("%d/%m/%Y")
    f_yyyy_mm_dd = fecha_objetivo.strftime("%Y-%m-%d")

    dia_str = f"{fecha_objetivo.day:02d}"
    mes_num = fecha_objetivo.month
    ano_str = str(fecha_objetivo.year)

    mes_en = MESES_INGLES[mes_num]
    mes_es = MESES_ESPANOL[mes_num]

    f_texto_en = f"{dia_str} {mes_en} {ano_str}"
    f_texto_es = f"{dia_str} {mes_es} {ano_str}"

    if (f_dd_mm_yyyy in txt or f_dd_mm_yyyy_slash in txt or f_yyyy_mm_dd in txt or
            f_texto_en.lower() in txt.lower() or f_texto_es.lower() in txt.lower()):
        return True

    for d, m_txt, y in re.findall(r"\b(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\b", txt):
        m_upper = m_txt.upper()
        if m_upper in MESES_MAPA:
            try:
                if date(int(y), MESES_MAPA[m_upper], int(d)) == fecha_objetivo:
                    return True
            except ValueError:
                pass

    for d, m, y in re.findall(r"\b(\d{1,2})[-/](\d{1,2})[-/](\d{4})\b", txt):
        try:
            if date(int(y), int(m), int(d)) == fecha_objetivo:
                return True
        except ValueError:
            pass

    for y, m, d in re.findall(r"\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b", txt):
        try:
            if date(int(y), int(m), int(d)) == fecha_objetivo:
                return True
        except ValueError:
            pass

    return False

def _texto_pdf_existente_para_fecha(carpeta_paciente, fecha, tipo_cobertura):
    """Si ya existe (de una corrida anterior) un PDF 'COBERTURA
    {tipo_cobertura}' que corresponde exactamente a esa fecha, devuelve su
    texto. Se usa para poder clasificar GENERAL/CAMPESINO y titular/
    dependiente sin tener que volver a descargar un PDF que ya esta en
    disco (reanudacion de una corrida anterior)."""
    if not carpeta_paciente or not os.path.exists(carpeta_paciente):
        return None

    tipo_norm = quitar_tildes(tipo_cobertura)
    try:
        archivos = os.listdir(carpeta_paciente)
    except Exception:
        return None

    for archivo in archivos:
        if not archivo.lower().endswith(".pdf"):
            continue
        if f"COBERTURA {tipo_norm}" not in quitar_tildes(archivo):
            continue

        ruta_pdf = os.path.join(carpeta_paciente, archivo)
        try:
            if ruta_pdf not in _CACHE_TEXTO_PDF:
                with open(ruta_pdf, "rb") as f:
                    contenido = f.read()
                _CACHE_TEXTO_PDF[ruta_pdf] = texto_pdf(contenido)
            txt = _CACHE_TEXTO_PDF[ruta_pdf]
        except Exception:
            continue

        if _texto_contiene_fecha(txt, fecha):
            return txt

    return None

# ============================================================
# PARSER DE CONTENIDO DE PDF
# ============================================================

def es_campesino(texto):
    """
    True si el PDF corresponde al seguro CAMPESINO. La palabra
    "CAMPESINO" no siempre aparece (por ejemplo, un jubilado del SSC solo
    dice "jubilado sistema de pensiones / jubilado del ssc", sin la
    palabra "campesino" en ningun lado), asi que tambien se reconoce por
    la sigla "SSC" (Seguro Social Campesino), que es exclusiva de este
    tipo de seguro.
    """
    t = quitar_tildes(texto)
    return "CAMPESINO" in t or "SSC" in t

def es_menor_de_edad(texto):
    t = quitar_tildes(texto)
    return "MENOR DE EDAD" in t or "MENOR DE 18" in t or "MENOR DE DIECIOCHO" in t

def es_titular_campesino(texto):
    """
    True si el PDF de PACIENTE/CAMPESINO indica que la persona consultada
    es, ella misma, la titular del seguro campesino (no un dependiente) y
    por lo tanto debe ir SOLA, sin el paquete de COBERTURA IESS + JEFE DE
    FAMILIA. Esto pasa en dos casos:

      - Es el "Jefe de Familia del SSC" (cabeza de familia que acredita a
        otros).
      - Es "Jubilado del SSC" (jubilado/pensionista del seguro campesino,
        que se hace atender el mismo, igual que el jubilado del seguro
        general).

    Si no aparece ninguno de los dos, es un dependiente del titular y si
    necesita el paquete completo.
    """
    t = quitar_tildes(texto)
    return "JEFE DE FAMILIA" in t or "JUBILADO" in t

def es_jubilado_campesino(texto):
    """True si, dentro de un PDF de CAMPESINO titular, la persona es
    especificamente 'Jubilado del SSC' (en vez de 'Jefe de Familia del
    SSC'). Se usa solo para elegir el texto correcto de TIPO SEGURO
    (BENEFICIARIO) en la matriz (CA vs JC)."""
    return "JUBILADO" in quitar_tildes(texto)

def extraer_acreditador(texto_iess):
    """
    Respaldo para extraer el acreditador desde el texto del PDF.
    Se mantiene por compatibilidad, pero cuando Selenium sigue en la pagina
    de resultados se prefiere extraer_acreditador_dom(), porque el DOM no
    depende de como Chrome acomoda las columnas al imprimir el PDF.
    """
    if not texto_iess:
        return None, None

    texto_norm = quitar_tildes(texto_iess)

    # Busca filas que contengan una cedula de 10 digitos despues de un numero
    # de fila. Se permiten saltos de linea entre cedula, nombre y afiliacion.
    patrones = [
        r"\b\d{1,3}\s+(\d{10})\s+(.{4,100}?)\s+"
        r"(JEFE\s+DE\s+FAMILIA|TITULAR|CONYUGE)",
        r"\b(\d{10})\s+(.{4,100}?)\s+"
        r"(JEFE\s+DE\s+FAMILIA|TITULAR|CONYUGE)",
    ]

    for patron in patrones:
        m = re.search(patron, texto_norm, re.IGNORECASE | re.DOTALL)
        if m:
            return m.group(1).strip(), re.sub(r"\s+", " ", m.group(2)).strip()

    # Respaldo: despues de la seccion ACREDITADOR, tomar la primera cedula
    # que no sea obviamente la del paciente.
    idx = texto_norm.rfind("ACREDITADOR")
    if idx != -1:
        resto = texto_norm[idx:]
        m = re.search(r"\b(\d{10})\b", resto)
        if m:
            cedula = m.group(1)
            return cedula, None

    return None, None


def extraer_acreditador_dom(driver, cedula_paciente=None):
    """
    Extrae el acreditador directamente de la pagina de resultados del IESS.

    Esta es la estrategia PRINCIPAL. No depende del PDF ni de la escala DPI,
    zoom o version de Chrome de la computadora. Busca la tabla que contiene
    el encabezado ACREDITADOR y, dentro de esa tabla, una cedula de 10
    digitos distinta de la cedula consultada.

    Devuelve (cedula, nombre). El nombre es opcional y se usa solo como
    referencia para los mensajes del programa.
    """
    try:
        tablas = driver.find_elements(By.XPATH, "//table")
    except Exception:
        return None, None

    cedula_paciente = str(cedula_paciente).strip() if cedula_paciente else None

    # 1) Prioridad: tabla cuyo texto menciona ACREDITADOR.
    tablas_acreditador = []
    for tabla in tablas:
        try:
            txt = quitar_tildes(tabla.text or "")
        except Exception:
            continue
        if "ACREDITADOR" in txt:
            tablas_acreditador.append(tabla)

    candidatas = tablas_acreditador if tablas_acreditador else tablas

    for tabla in candidatas:
        try:
            filas = tabla.find_elements(By.XPATH, ".//tr")
        except Exception:
            filas = []

        for fila in filas:
            try:
                celdas = fila.find_elements(By.XPATH, "./th|./td")
                textos = [(c.text or "").strip() for c in celdas]
            except Exception:
                continue

            cedulas = []
            for txt in textos:
                # Una celda normalmente contiene exactamente la cedula.
                m = re.fullmatch(r"\s*(\d{10})\s*", txt)
                if m:
                    cedulas.append(m.group(1))
                else:
                    # Respaldo por si la celda trae texto adicional.
                    for m2 in re.finditer(r"(?<!\d)(\d{10})(?!\d)", txt):
                        cedulas.append(m2.group(1))

            for cedula in cedulas:
                if cedula_paciente and cedula == cedula_paciente:
                    continue

                nombre = None
                # El nombre normalmente esta en una celda vecina.
                for txt in textos:
                    limpio = re.sub(r"\s+", " ", txt).strip()
                    if (len(limpio) >= 5 and
                            not re.fullmatch(r"\d{10}", limpio) and
                            "ACREDITADOR" not in quitar_tildes(limpio)):
                        nombre = limpio
                        break

                return cedula, nombre

    # 2) Ultimo respaldo: todas las celdas de todas las tablas.
    try:
        celdas = driver.find_elements(By.XPATH, "//table//td | //table//th")
    except Exception:
        celdas = []

    for celda in celdas:
        try:
            txt = (celda.text or "").strip()
        except Exception:
            continue

        m = re.fullmatch(r"\d{10}", txt)
        if m:
            cedula = m.group(1)
            if cedula_paciente and cedula == cedula_paciente:
                continue
            return cedula, None

    return None, None


def extraer_cedulas_acreditadores_dom(driver, cedula_paciente=None):
    """
    Extrae cedulas de acreditadores desde las tablas del DOM.

    Se priorizan las tablas que contienen el encabezado ACREDITADOR para
    evitar tomar numeros de otras tablas de la pagina.
    """
    cedulas = []
    cedula_paciente = str(cedula_paciente).strip() if cedula_paciente else None

    try:
        tablas = driver.find_elements(By.XPATH, "//table")
    except Exception:
        return cedulas

    tablas_acreditador = []
    for tabla in tablas:
        try:
            txt = quitar_tildes(tabla.text or "")
        except Exception:
            continue
        if "ACREDITADOR" in txt:
            tablas_acreditador.append(tabla)

    candidatas = tablas_acreditador if tablas_acreditador else tablas

    for tabla in candidatas:
        try:
            celdas = tabla.find_elements(By.XPATH, ".//td | .//th")
        except Exception:
            continue

        for celda in celdas:
            try:
                texto = (celda.text or "").strip()
            except Exception:
                continue

            for m in re.finditer(r"(?<!\d)(\d{10})(?!\d)", texto):
                cedula_f = m.group(1)
                if cedula_paciente and cedula_f == cedula_paciente:
                    continue
                if cedula_f not in cedulas:
                    cedulas.append(cedula_f)

    return cedulas
def extraer_cedulas_acreditadores(texto_iess, cedula_paciente=None):
    """
    Devuelve las cedulas (10 digitos) de los acreditadores que aparecen en
    la tabla "Nro. C.I. Acreditador / Nombres del Acreditador / Tipo de
    Afiliacion Acreditador" del PDF de Calificacion del IESS. Se usa para
    el flujo GENERAL: un menor de edad puede tener acreditados tanto al
    PADRE como a la MADRE al mismo tiempo, y un adulto puede tener
    acreditado a su CONYUGE.

    IMPORTANTE: esta funcion solo extrae la CEDULA, no el nombre ni el
    "Tipo de Afiliacion" del acreditador. Se comprobo con un caso real
    que, para el GENERAL, ese "Tipo de Afiliacion" NO dice "Padre" ni
    "Madre": dice el propio tipo de afiliacion del acreditador al IESS
    (por ejemplo "Afiliado Seguro General Tiempo Completo"), asi que no
    sirve para reconocer el parentesco. Ademas, el nombre y ese texto
    suelen venir partidos en varias lineas dentro del PDF (por el ancho
    de columna de la tabla), lo que hace muy poco confiable tratar de
    extraerlos con precision. Por eso solo se saca la cedula: con ella se
    consulta despues a coberturasalud.msp.gob.ec (igual que con las
    columnas del Excel), que si trae el nombre real, y desde ahi se
    determina PADRE/MADRE por genero como siempre.

    Devuelve una lista de cedulas (strings de 10 digitos), sin repetir y
    sin incluir la del propio paciente, en el orden en que aparecen.
    Puede devolver una lista vacia si no se reconoce ninguna fila (en ese
    caso el llamador debe usar la columna del Excel como respaldo, si la
    tiene).
    """
    if not texto_iess:
        return []

    idx_ultimo = texto_iess.upper().rfind("ACREDITADOR")
    if idx_ultimo == -1:
        return []

    resto = texto_iess[idx_ultimo:]

    cedulas = []
    for m in re.finditer(r"\b\d{1,3}\s+(\d{10})\b", resto):
        cedula_f = m.group(1)
        if cedula_paciente and cedula_f == cedula_paciente:
            continue
        if cedula_f in cedulas:
            continue
        cedulas.append(cedula_f)

    return cedulas

# ============================================================
# NOMBRES COMUNES (para distinguir PADRE/MADRE por el primer nombre,
# sin depender de ninguna libreria externa que podria no estar
# instalada en la computadora donde corre el script)
# ============================================================

NOMBRES_FEMENINOS_COMUNES = frozenset({
    "MARIA", "ANA", "ROSA", "CARMEN", "LUZ", "GLORIA", "PATRICIA", "SANDRA", "MONICA",
    "VERONICA", "SILVIA", "DIANA", "CLAUDIA", "ADRIANA", "PAOLA", "FANNY", "GABRIELA",
    "VALERIA", "CAMILA", "DANIELA", "ANDREA", "KAREN", "JESSICA", "JENNIFER", "STEFANY",
    "STEPHANY", "ESTEFANIA", "FERNANDA", "ALEXANDRA", "XIMENA", "TATIANA", "VIVIANA",
    "SUSANA", "ELIZABETH", "MARGARITA", "TERESA", "DOLORES", "MERCEDES", "CONSUELO",
    "PIEDAD", "ESPERANZA", "AMPARO", "BEATRIZ", "ISABEL", "SOLEDAD", "PILAR", "LOURDES",
    "CECILIA", "VICTORIA", "ALICIA", "LORENA", "PAULINA", "JOHANNA", "JOHANA", "YOLANDA",
    "NANCY", "NELLY", "NORMA", "GLADYS", "MYRIAM", "MIRIAM", "IRENE", "INES", "LUCIA",
    "MARLENE", "MARITZA", "JANETH", "JANET", "KATTY", "KATHY", "KARLA", "CARLA", "DAYANA",
    "DAYSI", "DEYSI", "ERIKA", "EDITH", "ELSA", "ELENA", "EVELYN", "FABIOLA", "FLOR",
    "GEOVANNA", "GIOCONDA", "GUADALUPE", "HAYDEE", "HERLINDA", "HORTENSIA", "JACQUELINE",
    "JAQUELINE", "JIMENA", "JOSEFINA", "JULIA", "JULISSA", "KATHERINE", "KATHERIN", "LAURA",
    "LEONOR", "LETICIA", "LIGIA", "LILIANA", "LINDA", "LUISA", "MABEL", "MAGALY", "MALENA",
    "MANUELA", "MARCIA", "MARIANA", "MARIBEL", "MARINA", "MARISOL", "MARTHA", "MAYRA",
    "MELANIE", "MELIDA", "MIRELLA", "NADIA", "NATALY", "NATALIA", "NAYELI", "NAOMI",
    "NOEMI", "OLGA", "PAMELA", "PRISCILA", "RAQUEL", "REBECA", "ROCIO", "ROSARIO", "RUTH",
    "SAMANTA", "SAMANTHA", "SARA", "SASKYA", "SCARLETH", "SHIRLEY", "SOFIA", "SONIA",
    "TANIA", "TAMARA", "VALENTINA", "VANESA", "VANESSA", "WENDY", "YADIRA", "YESENIA",
    "ZOILA", "ALEXA", "NICOLE", "AMELIA", "EMILIA", "EMMA", "ANTONELLA", "RENATA",
    "ALLISON", "ASHLEY", "BRITTANY", "KIMBERLY", "MADELEINE", "MELISSA", "MICHELLE",
    "MISHELL", "PIERINA", "GABRIELLA", "ANGELICA", "ANGELA", "AURORA", "BERTHA", "BLANCA",
    "CATALINA", "CINDY", "CRISTINA", "DALIA", "DENISSE", "DORIS", "ELVIRA", "ESTELA",
    "GEORGINA", "GRACIELA", "GUILLERMINA", "HILDA", "IVONNE", "JACKELINE", "JEANNETTE",
    "JOAQUINA", "JUANA", "KATIUSKA", "LADY", "LIZBETH", "LUPE", "MAGDALENA", "MARISELA",
    "MARIUXI", "MAYERLY", "MERY", "MICAELA", "MILENA", "NAYARA", "NUBIA", "PERLA",
    "PIEDAD", "RENE", "ROSSANA", "ROXANA", "SELENA", "SORAYA", "SUJEY", "TATIANA",
    "VALESKA", "YAJAIRA", "ZULEMA",
})

NOMBRES_MASCULINOS_COMUNES = frozenset({
    "JOSE", "JUAN", "LUIS", "CARLOS", "JORGE", "MIGUEL", "PEDRO", "PABLO", "FRANCISCO",
    "JAVIER", "ANDRES", "DIEGO", "DANIEL", "DAVID", "EDUARDO", "FERNANDO", "RICARDO",
    "ROBERTO", "ALFREDO", "ALFONSO", "ANTONIO", "ANGEL", "MANUEL", "MARIO", "MARCO",
    "MARCOS", "MAURICIO", "NELSON", "OSCAR", "PATRICIO", "RAFAEL", "RAMIRO", "RENE",
    "RODRIGO", "RUBEN", "SANTIAGO", "SEGUNDO", "SERGIO", "VICTOR", "WILSON", "WASHINGTON",
    "WILLIAM", "XAVIER", "ALEJANDRO", "ALEXIS", "AMADO", "ARMANDO", "ARTURO", "AUGUSTO",
    "BOLIVAR", "BYRON", "CESAR", "CRISTIAN", "CRISTOBAL", "DARIO", "EDGAR", "EDISON",
    "EDWIN", "ELVIS", "EMILIO", "ENRIQUE", "ERICK", "ERNESTO", "ESTEBAN", "EUGENIO",
    "EZEQUIEL", "FABIAN", "FABRICIO", "FAUSTO", "FELIX", "FIDEL", "FRANKLIN", "GABRIEL",
    "GALO", "GEOVANNY", "GERARDO", "GERMAN", "GILBERTO", "GONZALO", "GUIDO", "GUILLERMO",
    "GUSTAVO", "HECTOR", "HENRY", "HERNAN", "HUGO", "HUMBERTO", "IGNACIO", "IVAN",
    "JACINTO", "JAIME", "JEFFERSON", "JEISON", "JHON", "JHONNY", "JOEL", "JONATHAN",
    "JONATAN", "JOFFRE", "JULIO", "KEVIN", "KLEBER", "LENIN", "LEONARDO", "LEONIDAS",
    "LUCAS", "MANOLO", "MARLON", "MATEO", "MAXIMO", "MELVIN", "MILTON", "MOISES",
    "NESTOR", "NICOLAS", "NORBERTO", "OMAR", "ORLANDO", "OSWALDO", "PAUL", "PIERO",
    "PIETRO", "PLUTARCO", "RAUL", "REINALDO", "REMIGIO", "RENAN", "RICHARD", "ROLANDO",
    "ROMEL", "RONALD", "SAMUEL", "SAUL", "SIXTO", "TEODORO", "TOBIAS", "TOMAS", "ULISES",
    "VINICIO", "VLADIMIR", "WALTER", "WELLINGTON", "WILFRIDO", "WILMER", "WILDER",
    "YOFRE", "ALVARO", "AMILCAR", "ANIBAL", "BENIGNO", "BENJAMIN", "BRAYAN", "BRYAN",
    "CAMILO", "CLAUDIO", "CONSTANTINO", "CORNELIO", "DARWIN", "DEMETRIO", "DENNIS",
    "DIONISIO", "DOMINGO", "EFRAIN", "ELIAS", "EMERSON", "ERASMO", "FAUSTINO", "FEDERICO",
    "FRANK", "GEOVANY", "GIOVANNY", "HAROLD", "HIPOLITO", "HOLGER", "ISAAC", "ISIDRO",
    "JACOB", "JEISSON", "JEREMY", "JHAIR", "JOAQUIN", "JONAS", "JOSUE", "JUSTINO",
    "KLEVER", "LEIVER", "LEODAN", "LEOPOLDO", "LUCIANO", "MAICOL", "MARCELO", "MATIAS",
    "MAXIMILIANO", "MELCHOR", "MODESTO", "NAPOLEON", "NAUM", "NEPTALI", "NILO", "NOE",
    "OCTAVIO", "OLMEDO", "OSVALDO", "OTTO", "PACIFICO", "PASTOR", "PLACIDO", "POLIVIO",
    "PORFIRIO", "QUINTIN", "RAMON", "REMBERTO", "RIGOBERTO", "RODOLFO", "ROGELIO",
    "ROMAN", "RUPERTO", "SALOMON", "SALVADOR", "SANTOS", "SEBASTIAN", "SERAFIN", "SILVIO",
    "SIMON", "SOCRATES", "TARQUINO", "TEOFILO", "TIMOLEON", "URBANO", "VALENTIN",
    "VICENTE", "VIRGILIO", "WALDO", "WILLAN", "YOVANNY", "ZOILO",
})

# Nombres masculinos comunes que terminan en "A" (excepcion a la regla de
# "termina en A => femenino").
_EXCEPCIONES_MASCULINOS_TERMINAN_A = frozenset({"LUCA", "NICOLA", "JOSUA", "ELISHA"})

def _adivinar_genero_por_terminacion(nombre):
    """Heuristica de respaldo (solo si el nombre no esta en las listas de
    arriba): en espanol, un nombre de pila que termina en 'A' es casi
    siempre femenino."""
    if not nombre:
        return None
    if nombre in _EXCEPCIONES_MASCULINOS_TERMINAN_A:
        return "PADRE"
    if nombre.endswith("A") and len(nombre) >= 3:
        return "MADRE"
    return None

def extraer_nombre_persona(texto):
    m = re.search(
        r"COBERTURA DE SALUD\s*\n(.+?)\n\s*N[uú]mero de documento",
        texto,
        re.IGNORECASE | re.DOTALL
    )
    if m:
        return m.group(1).strip()
    return None

def detectar_sexo_desde_pdf(texto):
    t = quitar_tildes(texto)
    patrones_femenino = [r"SEXO\s*[:\-]?\s*FEMENINO", r"SEXO\s*[:\-]?\s*F", r"\bFEMENINO\b", r"\bMUJER\b"]
    for patron in patrones_femenino:
        if re.search(patron, t):
            return "MADRE"

    patrones_masculino = [r"SEXO\s*[:\-]?\s*MASCULINO", r"SEXO\s*[:\-]?\s*M", r"\bMASCULINO\b", r"\bHOMBRE\b"]
    for patron in patrones_masculino:
        if re.search(patron, t):
            return "PADRE"
    return None

def detectar_genero_por_nombre(nombre_completo):
    """
    Adivina PADRE/MADRE a partir del primer nombre de pila de la persona.
    El formato tipico en estos PDF es "APELLIDO APELLIDO NOMBRE NOMBRE",
    asi que se descartan las primeras 2 palabras (apellidos) cuando hay
    mas de 2, igual que antes.

    Orden de intentos:
      1. Lista de nombres femeninos/masculinos comunes incluida en este
         mismo archivo (no depende de ninguna libreria externa, asi que
         funciona aunque el equipo donde corre el script no tenga nada
         instalado aparte de lo basico).
      2. Heuristica de terminacion ("termina en A" => femenino).
      3. Si esta instalada, la libreria gender_guesser como ultimo
         respaldo (capa extra, ya no es la unica fuente).
    """
    if not nombre_completo:
        return None

    palabras = [p for p in nombre_completo.split() if p.isalpha()]
    hay_apellidos_claros = len(palabras) > 2
    candidatos = palabras[2:] if hay_apellidos_claros else palabras
    candidatos_norm = [quitar_tildes(p) for p in candidatos]

    # 1) Lista de nombres comunes (segura de usar siempre: solo dispara
    # si la palabra es un nombre de pila conocido, apellido o no).
    for palabra in candidatos_norm:
        if palabra in NOMBRES_FEMENINOS_COMUNES:
            return "MADRE"
        if palabra in NOMBRES_MASCULINOS_COMUNES:
            return "PADRE"

    # 2) Heuristica de terminacion ("termina en A" => femenino). Muchos
    # apellidos ecuatorianos (de origen kichwa, por ejemplo) tambien
    # terminan en "A", asi que esta heuristica solo se aplica cuando
    # estamos razonablemente seguros de que "candidatos" son en verdad
    # nombres de pila (4 palabras o mas: 2 apellidos + al menos 1
    # nombre), no cuando solo hay 1 o 2 palabras en total.
    if hay_apellidos_claros:
        for palabra in candidatos_norm:
            rol = _adivinar_genero_por_terminacion(palabra)
            if rol:
                return rol

    # 3) Respaldo opcional: libreria externa, solo si esta instalada
    try:
        import gender_guesser.detector as genderdet
    except ImportError:
        return None

    d = genderdet.Detector(case_sensitive=False)
    for palabra in candidatos:
        genero = d.get_gender(palabra.capitalize())
        if genero in ("male", "mostly_male"):
            return "PADRE"
        if genero in ("female", "mostly_female"):
            return "MADRE"

    return None

def determinar_rol_familiar(texto_pdf):
    rol = detectar_sexo_desde_pdf(texto_pdf)
    if rol:
        return rol
    nombre = extraer_nombre_persona(texto_pdf)
    rol = detectar_genero_por_nombre(nombre)
    if rol:
        return rol
    return "PADRE"

def estado_cobertura_iess(texto):
    """Busca en el texto del PDF si la fila del IESS (no ISSFA ni ISSPOL)
    dice 'SI REGISTRA COBERTURA' o 'NO REGISTRA COBERTURA'. Devuelve
    True, False, o None si no se encontro NINGUNO de los dos marcadores
    -- esto casi siempre pasa cuando la pagina de coberturasalud fallo o
    devolvio un PDF/pagina que no es el resultado real (caida del sitio,
    glitch, pagina a medio cargar), y NO significa que el paciente no
    tenga seguro."""
    t = quitar_tildes(texto)
    m = re.search(
        r"\bIESS\b(?:(?!\bISSFA\b).)*?(SI REGISTRA COBERTURA|NO REGISTRA COBERTURA)",
        t,
        re.DOTALL
    )
    if not m:
        return None
    return m.group(1) == "SI REGISTRA COBERTURA"

def es_resultado_indeterminado(texto):
    """True si el PDF no trae ninguno de los marcadores esperados de
    resultado: normalmente significa que coberturasalud.msp.gob.ec fallo
    o devolvio contenido inesperado, y este resultado NUNCA debe
    reportarse como "sin seguro" sin antes reintentar/confirmar."""
    return estado_cobertura_iess(texto) is None

def iess_tiene_cobertura(texto):
    return estado_cobertura_iess(texto) is True

def es_sin_seguro(texto):
    return estado_cobertura_iess(texto) is False

def siguiente_numero_archivo(carpeta_paciente, tipo_cobertura):
    patron = re.compile(
        rf"^(\d+)\s+COBERTURA\s+{re.escape(tipo_cobertura)}\.pdf$",
        re.IGNORECASE
    )
    numeros = []
    try:
        archivos = os.listdir(carpeta_paciente)
    except FileNotFoundError:
        return 1

    for archivo in archivos:
        coincidencia = patron.match(archivo)
        if coincidencia:
            try:
                numeros.append(int(coincidencia.group(1)))
            except ValueError:
                pass

    if not numeros:
        return 1

    return max(numeros) + 1

# ============================================================
# CARPETA DEL PACIENTE (GENERAL / CAMPESINO)
# ============================================================

def ruta_carpeta_general(nombre, fecha):
    return os.path.join(CARPETA_SALIDA, nombre_carpeta_mes(fecha), SUBCARPETA_GENERAL, nombre)

def ruta_carpeta_campesino(nombre, fecha):
    return os.path.join(CARPETA_SALIDA, nombre_carpeta_mes(fecha), SUBCARPETA_CAMPESINO, nombre)

def localizar_carpeta_paciente(nombre, fecha):
    """
    Devuelve la carpeta del paciente PARA EL MES DE 'fecha' si ya existe
    en GENERAL o en CAMPESINO (de una corrida anterior). Si todavia no
    existe en ninguna de las dos, devuelve None: el tipo se determina
    recien al descargar y leer el primer PDF de ese paciente.
    """
    ruta_campesino = ruta_carpeta_campesino(nombre, fecha)
    if os.path.exists(ruta_campesino):
        return ruta_campesino

    ruta_general = ruta_carpeta_general(nombre, fecha)
    if os.path.exists(ruta_general):
        return ruta_general

    return None

# ============================================================
# REINTENTOS Y MANEJO DE SESION CHROME
# ============================================================

def sesion_viva(driver):
    try:
        _ = driver.title
        return True
    except Exception:
        return False

def consultar_con_reintentos(driver_holder, cedula, fecha_str, carpeta_descargas_temp, carpeta_diagnostico):
    """Devuelve (pdf_bytes, error, sin_seguro_confirmado):

      - Resultado normal y claro (con o sin seguro) desde coberturasalud:
        (pdf, None, False).

      - Si coberturasalud falla repetidas veces (caido, timeout) o
        devuelve contenido sin el resultado esperado (pagina a medio
        cargar / glitch del sitio -- ANTES esto se confundia con "sin
        seguro"), se hace una verificacion de respaldo en el sitio del
        IESS, SOLO para decidir como reportarlo (nunca para reemplazar
        el PDF oficial de coberturasalud):
          * El IESS confirma que SI tiene seguro -> sigue siendo un
            ERROR, no "sin seguro" (pdf=None, error=mensaje,
            sin_seguro_confirmado=False), para que quede en la lista de
            errores y se pueda reintentar con el boton "Reintentar" mas
            tarde y obtener el PDF real.
          * El IESS confirma que NO tiene seguro -> se puede reportar
            sin seguro con confianza aunque no se haya podido bajar el
            PDF de coberturasalud (pdf=None, error=None,
            sin_seguro_confirmado=True).
          * El IESS tampoco responde con claridad -> error comun
            (pdf=None, error=mensaje, sin_seguro_confirmado=False).
    """
    ultimo_error = None
    for intento in range(1, MAX_INTENTOS_POR_FECHA + 1):
        try:
            if not sesion_viva(driver_holder[0]):
                print("   Sesion de Chrome caida, reiniciando...")
                try:
                    driver_holder[0].quit()
                except Exception:
                    pass
                driver_holder[0] = crear_driver(carpeta_descargas_temp)

            pdf = consultar_y_obtener_pdf_bytes(
                driver_holder[0], cedula, fecha_str, carpeta_descargas_temp
            )

            if es_resultado_indeterminado(texto_pdf(pdf)):
                raise RuntimeError(
                    "coberturasalud.msp.gob.ec respondio sin el resultado esperado "
                    "(posible caida o glitch del sitio)"
                )

            return (pdf, None, False)

        except Exception as e:
            ultimo_error = str(e).splitlines()[0]
            print(f"      intento {intento}/{MAX_INTENTOS_POR_FECHA} fallo: {ultimo_error}")

            try:
                os.makedirs(carpeta_diagnostico, exist_ok=True)
                nombre_shot = f"{cedula}_{fecha_str}_intento{intento}.png"
                driver_holder[0].save_screenshot(os.path.join(carpeta_diagnostico, nombre_shot))
            except Exception:
                pass

            try:
                driver_holder[0].quit()
            except Exception:
                pass

            driver_holder[0] = crear_driver(carpeta_descargas_temp)
            time.sleep(2)

    # coberturasalud no dio un resultado claro tras todos los intentos:
    # se verifica en el IESS SOLO para decidir la clasificacion, sin
    # arriesgarse a marcar como "sin seguro" a alguien que si tiene
    # mientras el sitio principal esta caido.
    print("      coberturasalud sigue fallando; verificando en el IESS como respaldo...")
    try:
        if not sesion_viva(driver_holder[0]):
            driver_holder[0] = crear_driver(carpeta_descargas_temp)
        pdf_iess = consultar_iess_calificacion_pdf(
            driver_holder[0], cedula, fecha_str, carpeta_diagnostico
        )
        texto_iess_norm = quitar_tildes(texto_pdf(pdf_iess)).upper()
        if "SIN COBERTURA IESS" in texto_iess_norm:
            print("      (el IESS confirma que NO tiene seguro)")
            return (None, None, True)
        if "CON COBERTURA IESS" in texto_iess_norm:
            print("      (el IESS confirma que SI tiene seguro; hay que reintentar "
                  "mas tarde para obtener el PDF oficial de coberturasalud)")
            ultimo_error = (
                (ultimo_error or "coberturasalud no disponible")
                + " -- el IESS confirma que SI tiene seguro; reintentar mas tarde"
            )
    except Exception as e:
        print(f"      (la verificacion de respaldo en el IESS tambien fallo: {_describir_error(e)})")

    return (None, ultimo_error, False)

# ============================================================
# MOTOR DE DESCARGA POR REGISTRO (reutilizado por el modo Excel y por
# el modo manual/interactivo)
# ============================================================

def procesar_registro(reg, i, total, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                       errores, sin_seguro, callback_fila=None, recolectar_matriz=False):
    """
    Procesa UN registro (una cedula con una o mas fechas): descarga su(s)
    PDF(s) de cobertura y, si corresponde, los de PADRE/MADRE/TITULAR o
    del paquete CAMPESINO (IESS + JEFE DE FAMILIA). Esta es exactamente
    la misma logica que antes vivia adentro del bucle de main(); se
    saco de ahi para poder llamarla tambien una sola vez por vez desde
    el modo manual (modo_interactivo), con un reg armado a partir de lo
    que el usuario va escribiendo por consola en vez de leerlo del
    Excel de inconsistencias.

    reg["nombre"] puede venir en None (modo manual: todavia no se sabe
    el nombre real del paciente). En ese caso se descarga siempre en
    linea el PDF propio del paciente (no se puede revisar si "ya existe"
    sin saber en que carpeta buscar) y, apenas se lee ese PDF, se toma
    el nombre real desde ahi mismo (igual que hace el resto del script
    para reconocer MENOR/CAMPESINO/etc.), y se guarda de vuelta en
    reg["nombre"] por si el llamador lo quiere usar despues.

    Si recolectar_matriz=True, devuelve una lista con un diccionario de
    datos por cada fecha en la que se confirmo que el paciente SI tiene
    seguro (para llenar despues una fila de la matriz MES_2026). Si el
    paciente no tiene seguro o hubo un error consultando su propio PDF,
    esa fecha no genera fila de matriz (ya queda registrada aparte en
    sin_seguro/errores). Si recolectar_matriz=False (modo Excel de
    siempre) se devuelve None y el comportamiento es identico al de
    antes.
    """
    cedula = reg["cedula"]
    nombre = reg["nombre"]
    fechas = reg["fechas"]
    cedula_padre = reg["cedula_padre"]
    cedula_titular = reg["cedula_titular"]

    filas_matriz = [] if recolectar_matriz else None

    def guardar_pdf_en_carpeta(carpeta_paciente, contenido_bytes, tipo_cobertura):
        """Crea la carpeta solo en el momento que se vaya a guardar el PDF."""
        os.makedirs(carpeta_paciente, exist_ok=True)

        numero_nuevo = siguiente_numero_archivo(carpeta_paciente, tipo_cobertura)
        nombre_pdf = f"{numero_nuevo} COBERTURA {tipo_cobertura}.pdf"
        ruta_final = os.path.join(carpeta_paciente, nombre_pdf)

        while os.path.exists(ruta_final):
            numero_nuevo += 1
            nombre_pdf = f"{numero_nuevo} COBERTURA {tipo_cobertura}.pdf"
            ruta_final = os.path.join(carpeta_paciente, nombre_pdf)

        with open(ruta_final, "wb") as f:
            f.write(contenido_bytes)

        return nombre_pdf, ruta_final

    def descargar_familiar(cedula_fam, rol_fijo, fecha, fecha_str, carpeta_paciente, nombre_paciente,
                            permitir_ambos_padres=False, info_matriz=None):
        if rol_fijo == "PADRE":
            if not permitir_ambos_padres:
                ya_existe_padre = fecha_ya_descargada(carpeta_paciente, fecha, "PADRE")
                ya_existe_madre = fecha_ya_descargada(carpeta_paciente, fecha, "MADRE")

                if ya_existe_padre or ya_existe_madre:
                    rol_existente = "PADRE" if ya_existe_padre else "MADRE"
                    print(f"   [{fecha_str}] OMITIDO: ya existe COBERTURA {rol_existente} para esta fecha.")
                    return
            # Si permitir_ambos_padres=True (viene de la busqueda automatica
            # en el IESS y encontro tanto PADRE como MADRE), no se hace este
            # chequeo temprano: puede que ya se haya guardado uno de los dos
            # y ahora se este agregando el otro. El chequeo de mas abajo por
            # rol_final ya evita guardar un duplicado exacto.
        else:
            if fecha_ya_descargada(carpeta_paciente, fecha, rol_fijo):
                print(f"   [{fecha_str}] OMITIDO: ya existe COBERTURA {rol_fijo} para esta fecha.")
                return

        pdf_fam, error_fam, sin_seguro_confirmado = consultar_con_reintentos(
            driver_holder, cedula_fam, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
        )

        if sin_seguro_confirmado:
            print(f"   [{fecha_str}] AVISO: {rol_fijo} ({cedula_fam}) NO TIENE SEGURO (confirmado via IESS).")
            sin_seguro.append([nombre_paciente, reg["cedula"], cedula_fam, rol_fijo, fecha_str])
            return

        if error_fam:
            print(f"   [{fecha_str}] ERROR {rol_fijo}: {error_fam}")
            errores.append([nombre_paciente, cedula_fam, fecha_str, f"({rol_fijo}) {error_fam}"])
            return

        texto_fam = texto_pdf(pdf_fam)

        if es_sin_seguro(texto_fam):
            print(f"   [{fecha_str}] AVISO: {rol_fijo} ({cedula_fam}) NO TIENE SEGURO.")
            sin_seguro.append([nombre_paciente, reg["cedula"], cedula_fam, rol_fijo, fecha_str])
            return

        rol_final = determinar_rol_familiar(texto_fam) if rol_fijo == "PADRE" else rol_fijo

        if fecha_ya_descargada(carpeta_paciente, fecha, rol_final):
            print(f"   [{fecha_str}] OMITIDO: la fecha ya existe como COBERTURA {rol_final}.")
            return

        nombre_pdf_fam, ruta_fam = guardar_pdf_en_carpeta(carpeta_paciente, pdf_fam, rol_final)
        _CACHE_TEXTO_PDF[ruta_fam] = texto_fam

        if info_matriz is not None and "parentesco" not in info_matriz:
            nombre_fam = extraer_nombre_persona(texto_fam)
            info_matriz["cedula_afiliado"] = cedula_fam
            info_matriz["nombre_afiliado"] = (nombre_fam or "").strip() or cedula_fam
            if rol_fijo == "PADRE":
                info_matriz["parentesco"] = PARENTESCO_HIJO
                info_matriz["tipo_seguro_beneficiario"] = TIPO_SEGURO_GENERAL_MENOR
            else:
                info_matriz["parentesco"] = PARENTESCO_CONYUGE
                info_matriz["tipo_seguro_beneficiario"] = TIPO_SEGURO_GENERAL_CONYUGE

        print(f"   [{fecha_str}] OK ({rol_final}) -> {nombre_pdf_fam}")
        time.sleep(PAUSA_ENTRE_CONSULTAS)

    def descargar_jefe_de_familia_campesino(cedula_jefe, nombre_jefe_ref, fecha, fecha_str,
                                             carpeta_paciente, nombre_paciente, cedula_paciente,
                                             info_matriz=None):
        """Descarga la COBERTURA JEFE DE FAMILIA (desde coberturasalud) para
        el acreditador de un paciente CAMPESINO dependiente."""
        if fecha_ya_descargada(carpeta_paciente, fecha, "JEFE DE FAMILIA"):
            print(f"   [{fecha_str}] OMITIDO: ya existe COBERTURA JEFE DE FAMILIA para esta fecha.")
            return

        pdf_jefe, error_jefe, sin_seguro_confirmado = consultar_con_reintentos(
            driver_holder, cedula_jefe, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
        )

        if sin_seguro_confirmado:
            print(f"   [{fecha_str}] AVISO: JEFE DE FAMILIA ({cedula_jefe}) NO TIENE SEGURO (confirmado via IESS).")
            sin_seguro.append([nombre_paciente, cedula_paciente, cedula_jefe, "JEFE DE FAMILIA", fecha_str])
            return

        if error_jefe:
            print(f"   [{fecha_str}] ERROR JEFE DE FAMILIA ({nombre_jefe_ref or cedula_jefe}): {error_jefe}")
            errores.append([nombre_paciente, cedula_jefe, fecha_str, f"(JEFE DE FAMILIA) {error_jefe}"])
            return

        texto_jefe = texto_pdf(pdf_jefe)

        if es_sin_seguro(texto_jefe):
            print(f"   [{fecha_str}] AVISO: JEFE DE FAMILIA ({cedula_jefe}) NO TIENE SEGURO.")
            sin_seguro.append([nombre_paciente, cedula_paciente, cedula_jefe, "JEFE DE FAMILIA", fecha_str])
            return

        nombre_pdf_jefe, ruta_jefe = guardar_pdf_en_carpeta(carpeta_paciente, pdf_jefe, "JEFE DE FAMILIA")
        _CACHE_TEXTO_PDF[ruta_jefe] = texto_jefe

        if info_matriz is not None and "parentesco" not in info_matriz:
            nombre_jefe_real = extraer_nombre_persona(texto_jefe) or nombre_jefe_ref
            info_matriz["cedula_afiliado"] = cedula_jefe
            info_matriz["nombre_afiliado"] = (nombre_jefe_real or "").strip() or cedula_jefe
            info_matriz["parentesco"] = PARENTESCO_PARIENTE
            info_matriz["tipo_seguro_beneficiario"] = TIPO_SEGURO_CAMPESINO_DEPENDIENTE

        print(f"   [{fecha_str}] OK (JEFE DE FAMILIA: {nombre_jefe_ref or cedula_jefe}) -> {nombre_pdf_jefe}")
        time.sleep(PAUSA_ENTRE_CONSULTAS)

    def descargar_paquete_campesino(cedula_paciente, fecha, fecha_str, carpeta_paciente, nombre_paciente,
                                     info_matriz=None):
        """Para un paciente CAMPESINO que es dependiente (no titular):
        descarga COBERTURA IESS desde app.iess.gob.ec, extrae de ahi la
        cedula del acreditador (jefe de familia), y descarga su COBERTURA
        JEFE DE FAMILIA desde coberturasalud. Si el paciente es el propio
        titular (Jefe de Familia del SSC), esta funcion no debe llamarse:
        ese paciente va solo, sin este paquete."""
        if fecha_ya_descargada(carpeta_paciente, fecha, "IESS"):
            print(f"   [{fecha_str}] IESS ya existe; se reutiliza para buscar al Jefe de Familia.")
            texto_iess = _texto_pdf_existente_para_fecha(carpeta_paciente, fecha, "IESS")
            if not texto_iess:
                pdf_iess, error_iess = consultar_iess_con_reintentos(
                    driver_holder, cedula_paciente, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
                )
                if error_iess:
                    print(f"   [{fecha_str}] ERROR IESS: {error_iess}")
                    errores.append([nombre_paciente, cedula_paciente, fecha_str, f"(IESS) {error_iess}"])
                    return
                texto_iess = texto_pdf(pdf_iess)
        else:
            pdf_iess, error_iess = consultar_iess_con_reintentos(
                driver_holder, cedula_paciente, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
            )

            if error_iess:
                print(f"   [{fecha_str}] ERROR IESS: {error_iess}")
                errores.append([nombre_paciente, cedula_paciente, fecha_str, f"(IESS) {error_iess}"])
                return

            texto_iess = texto_pdf(pdf_iess)
            nombre_pdf_iess, ruta_iess = guardar_pdf_en_carpeta(carpeta_paciente, pdf_iess, "IESS")
            _CACHE_TEXTO_PDF[ruta_iess] = texto_iess

            print(f"   [{fecha_str}] OK (IESS) -> {nombre_pdf_iess}")

        cedula_acreditador, nombre_acreditador = extraer_acreditador_dom(
            driver_holder[0], cedula_paciente
        )

        if not cedula_acreditador:
            cedula_acreditador, nombre_acreditador = extraer_acreditador(texto_iess)

        if not cedula_acreditador or cedula_acreditador == cedula_paciente:
            _guardar_diagnostico_sin_acreditador(
                cedula_paciente, fecha_str,
                locals().get("pdf_iess", b""),
                texto_iess,
                "jefe_familia"
            )
            print(f"   [{fecha_str}] AVISO: no se pudo identificar al Jefe de Familia "
                  f"(DOM ni PDF). Se guardo diagnostico en _diagnostico.")
            return

        time.sleep(PAUSA_ENTRE_CONSULTAS)
        descargar_jefe_de_familia_campesino(
            cedula_acreditador, nombre_acreditador, fecha, fecha_str,
            carpeta_paciente, nombre_paciente, cedula_paciente,
            info_matriz=info_matriz
        )

    def _guardar_diagnostico_sin_acreditador(cedula_paciente, fecha_str, pdf_iess, texto_iess, etiqueta):
        try:
            os.makedirs(carpeta_diagnostico, exist_ok=True)
            base = f"{cedula_paciente}_{fecha_str}_iess_sin_{etiqueta}"
            with open(os.path.join(carpeta_diagnostico, base + ".pdf"), "wb") as f:
                f.write(pdf_iess)
            with open(os.path.join(carpeta_diagnostico, base + ".txt"), "w", encoding="utf-8") as f:
                f.write(texto_iess)
        except Exception:
            pass

    def buscar_padres_via_iess(cedula_paciente, fecha_str):
        pdf_iess, error_iess = consultar_iess_con_reintentos(
            driver_holder, cedula_paciente, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
        )
        if error_iess:
            print(f"   [{fecha_str}] AVISO: no se pudo consultar el IESS para ubicar al padre/madre "
                  f"automaticamente ({error_iess}); se intenta con el Excel si trae esa columna.")
            return []

        texto_iess = texto_pdf(pdf_iess)

        cedulas = extraer_cedulas_acreditadores_dom(driver_holder[0], cedula_paciente)[:2]
        if not cedulas:
            cedulas = extraer_cedulas_acreditadores(texto_iess, cedula_paciente)[:2]

        if not cedulas:
            print(f"   [{fecha_str}] AVISO: la consulta al IESS no encontro ningun acreditador "
                  f"(padre/madre) reconocible (se guarda copia de diagnostico); se intenta con el "
                  f"Excel si trae esa columna.")
            _guardar_diagnostico_sin_acreditador(cedula_paciente, fecha_str, pdf_iess, texto_iess, "padre_madre")

        return cedulas

    def buscar_conyuge_via_iess(cedula_paciente, fecha_str):
        pdf_iess, error_iess = consultar_iess_con_reintentos(
            driver_holder, cedula_paciente, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
        )
        if error_iess:
            print(f"   [{fecha_str}] AVISO: no se pudo consultar el IESS para ubicar al conyuge/titular "
                  f"automaticamente ({error_iess}); se intenta con el Excel si trae esa columna.")
            return None

        texto_iess = texto_pdf(pdf_iess)

        cedulas = extraer_cedulas_acreditadores_dom(
            driver_holder[0], cedula_paciente
        )

        if not cedulas:
            cedulas = extraer_cedulas_acreditadores(texto_iess, cedula_paciente)

        if cedulas:
            return cedulas[0]

        print(f"   [{fecha_str}] AVISO: la consulta al IESS no encontro ningun acreditador "
              f"(conyuge/titular) reconocible (se guarda copia de diagnostico); se intenta con el "
              f"Excel si trae esa columna.")
        _guardar_diagnostico_sin_acreditador(cedula_paciente, fecha_str, pdf_iess, texto_iess, "conyuge")

        return None

    carpeta_paciente_actual = localizar_carpeta_paciente(nombre, fechas[0]) if nombre else None
    if carpeta_paciente_actual is None and (cedula_padre or cedula_titular):
        carpeta_paciente_actual = ruta_carpeta_general(nombre, fechas[0])

    errores_antes = len(errores)

    if callback_fila:
        callback_fila(i, "Procesando...")

    extra = []
    if cedula_padre:
        extra.append(f"padre/madre: {cedula_padre}")
    if cedula_titular:
        extra.append(f"titular: {cedula_titular}")

    print(f"\n[{i}/{total}] {nombre or cedula} - {cedula} - {len(fechas)} fecha(s)" +
          (" - " + ", ".join(extra) if extra else ""))

    for fecha in fechas:
        fecha_str = fecha.strftime("%d-%m-%Y")

        if nombre is not None:
            carpeta_chequeo = localizar_carpeta_paciente(nombre, fecha) or ruta_carpeta_general(nombre, fecha)
            existe_paciente = fecha_ya_descargada(carpeta_chequeo, fecha, "PACIENTE")
            existe_menor = fecha_ya_descargada(carpeta_chequeo, fecha, "MENOR")
        else:
            # Paciente nuevo del modo manual: todavia no se conoce su
            # nombre real (se obtiene recien del propio PDF), asi que no
            # hay ninguna carpeta que revisar todavia. Se consulta
            # siempre en linea para este primer PDF; una vez descargado
            # y guardado no se vuelve a duplicar, porque el chequeo de
            # "ya existe" de mas abajo (carpeta_paciente_final) ya se
            # hace con el nombre real.
            carpeta_chequeo = None
            existe_paciente = False
            existe_menor = False

        texto = None  # texto del PDF de PACIENTE/MENOR de esta fecha, si se conoce

        if existe_paciente or existe_menor:
            tipo_existente = "MENOR" if existe_menor else "PACIENTE"
            print(f"   [{fecha_str}] OMITIDO paciente: ya existe COBERTURA {tipo_existente} para esta fecha.")
            carpeta_paciente_actual = carpeta_chequeo
            texto = _texto_pdf_existente_para_fecha(carpeta_chequeo, fecha, tipo_existente)
        else:
            pdf_bytes, error, sin_seguro_confirmado = consultar_con_reintentos(
                driver_holder, cedula, fecha_str, carpeta_descargas_temp, carpeta_diagnostico
            )

            if sin_seguro_confirmado:
                print(f"   [{fecha_str}] AVISO: paciente NO TIENE SEGURO (confirmado via IESS).")
                sin_seguro.append([nombre or cedula, cedula, cedula, "PACIENTE", fecha_str])
            elif error:
                print(f"   [{fecha_str}] ERROR paciente: {error}")
                errores.append([nombre or cedula, cedula, fecha_str, error])
            else:
                texto_descargado = texto_pdf(pdf_bytes)

                if nombre is None:
                    nombre_detectado = extraer_nombre_persona(texto_descargado)
                    nombre = (nombre_detectado or "").strip() or cedula
                    reg["nombre"] = nombre
                    print(f"   [{fecha_str}] Paciente identificado: {nombre}")

                if es_sin_seguro(texto_descargado):
                    print(f"   [{fecha_str}] AVISO: paciente NO TIENE SEGURO.")
                    sin_seguro.append([nombre, cedula, cedula, "PACIENTE", fecha_str])
                else:
                    menor = es_menor_de_edad(texto_descargado)
                    tipo_cobertura = "MENOR" if menor else "PACIENTE"

                    campesino = es_campesino(texto_descargado)
                    tipo_seguro = SUBCARPETA_CAMPESINO if campesino else SUBCARPETA_GENERAL
                    carpeta_paciente_final = os.path.join(
                        CARPETA_SALIDA, nombre_carpeta_mes(fecha), tipo_seguro, nombre
                    )

                    if fecha_ya_descargada(carpeta_paciente_final, fecha, tipo_cobertura):
                        print(f"   [{fecha_str}] OMITIDO: la fecha ya existe.")
                    else:
                        nombre_pdf, ruta_final = guardar_pdf_en_carpeta(
                            carpeta_paciente_final, pdf_bytes, tipo_cobertura
                        )
                        _CACHE_TEXTO_PDF[ruta_final] = texto_descargado

                        etiqueta = tipo_seguro
                        if menor:
                            etiqueta += " / MENOR"

                        print(f"   [{fecha_str}] OK -> {nombre_pdf} ({etiqueta})")

                    carpeta_paciente_actual = carpeta_paciente_final
                    texto = texto_descargado

        time.sleep(PAUSA_ENTRE_CONSULTAS)

        info_matriz_fecha = None
        if recolectar_matriz and texto:
            info_matriz_fecha = {
                "cedula_paciente": cedula,
                "nombre_paciente": nombre,
                "fecha_atencion": fecha,
            }

        if texto and carpeta_paciente_actual:
            if es_campesino(texto):
                if not es_titular_campesino(texto):
                    descargar_paquete_campesino(
                        cedula, fecha, fecha_str, carpeta_paciente_actual, nombre,
                        info_matriz=info_matriz_fecha
                    )
                else:
                    if info_matriz_fecha is not None:
                        info_matriz_fecha["parentesco"] = PARENTESCO_TITULAR
                        info_matriz_fecha["cedula_afiliado"] = cedula
                        info_matriz_fecha["nombre_afiliado"] = nombre
                        info_matriz_fecha["tipo_seguro_beneficiario"] = (
                            TIPO_SEGURO_CAMPESINO_JUBILADO if es_jubilado_campesino(texto)
                            else TIPO_SEGURO_CAMPESINO_TITULAR
                        )
                if info_matriz_fecha is not None:
                    info_matriz_fecha["institucion"] = "SEGURO CAMPESINO"
                    if "parentesco" not in info_matriz_fecha:
                        # No se pudo identificar al Jefe de Familia (fallo de
                        # red, o no se reconocio en el DOM/PDF): se deja el
                        # parentesco mas probable (PARIENTE) para que la fila
                        # no quede sin ese dato, aunque CI/nombre del
                        # afiliado queden en blanco para completar a mano.
                        info_matriz_fecha["parentesco"] = PARENTESCO_PARIENTE
            else:
                menor_general = es_menor_de_edad(texto)

                if menor_general:
                    cedulas_padres_iess = buscar_padres_via_iess(cedula, fecha_str)
                    if cedulas_padres_iess:
                        permitir_ambos = len(cedulas_padres_iess) > 1
                        for cedula_f in cedulas_padres_iess:
                            descargar_familiar(
                                cedula_f, "PADRE", fecha, fecha_str, carpeta_paciente_actual, nombre,
                                permitir_ambos_padres=permitir_ambos, info_matriz=info_matriz_fecha
                            )
                    elif cedula_padre:
                        descargar_familiar(cedula_padre, "PADRE", fecha, fecha_str, carpeta_paciente_actual, nombre,
                                            info_matriz=info_matriz_fecha)

                    if cedula_titular:
                        descargar_familiar(cedula_titular, "TITULAR", fecha, fecha_str, carpeta_paciente_actual, nombre,
                                            info_matriz=info_matriz_fecha)
                else:
                    cedula_conyuge_iess = buscar_conyuge_via_iess(cedula, fecha_str)
                    if cedula_conyuge_iess:
                        descargar_familiar(cedula_conyuge_iess, "TITULAR", fecha, fecha_str, carpeta_paciente_actual, nombre,
                                            info_matriz=info_matriz_fecha)
                    elif cedula_titular:
                        descargar_familiar(cedula_titular, "TITULAR", fecha, fecha_str, carpeta_paciente_actual, nombre,
                                            info_matriz=info_matriz_fecha)

                    if cedula_padre:
                        descargar_familiar(cedula_padre, "PADRE", fecha, fecha_str, carpeta_paciente_actual, nombre,
                                            info_matriz=info_matriz_fecha)

                if info_matriz_fecha is not None:
                    info_matriz_fecha["institucion"] = "IESS"
                    if "parentesco" not in info_matriz_fecha:
                        # No se encontro ningun acreditador (ni por IESS ni por
                        # el Excel): se asume que el propio paciente es el
                        # titular de su seguro general.
                        info_matriz_fecha["parentesco"] = (
                            PARENTESCO_HIJO if menor_general else PARENTESCO_TITULAR
                        )
                        if info_matriz_fecha["parentesco"] == PARENTESCO_TITULAR:
                            info_matriz_fecha["cedula_afiliado"] = cedula
                            info_matriz_fecha["nombre_afiliado"] = nombre
                            info_matriz_fecha["tipo_seguro_beneficiario"] = TIPO_SEGURO_GENERAL_TITULAR
                        else:
                            info_matriz_fecha["tipo_seguro_beneficiario"] = TIPO_SEGURO_GENERAL_MENOR
        else:
            if carpeta_paciente_actual and (cedula_padre or cedula_titular):
                if cedula_padre:
                    descargar_familiar(cedula_padre, "PADRE", fecha, fecha_str, carpeta_paciente_actual, nombre)
                if cedula_titular:
                    descargar_familiar(cedula_titular, "TITULAR", fecha, fecha_str, carpeta_paciente_actual, nombre)

        if info_matriz_fecha is not None:
            filas_matriz.append(info_matriz_fecha)

    # Si la carpeta fue creada previamente pero quedó totalmente vacía, eliminarla
    if carpeta_paciente_actual and os.path.exists(carpeta_paciente_actual):
        try:
            if not os.listdir(carpeta_paciente_actual):
                os.rmdir(carpeta_paciente_actual)
        except Exception:
            pass

    hubo_error = len(errores) > errores_antes

    if callback_fila:
        callback_fila(i, "Con errores" if hubo_error else "OK")

    return filas_matriz

# ============================================================
# REPORTES FINALES (log_errores.xlsx / sin_seguro.xlsx)
# ============================================================

def _guardar_reportes_finales(errores, sin_seguro):
    if errores:
        wb_log = openpyxl.Workbook()
        ws_log = wb_log.active
        ws_log.title = "Errores"
        ws_log.append(["Nombre", "Cedula", "Fecha", "Detalle"])

        for fila in errores:
            ws_log.append(fila)

        ruta_log = os.path.join(CARPETA_SALIDA, "log_errores.xlsx")
        wb_log.save(ruta_log)
        print(f"\n{len(errores)} incidencia(s) de error.\nRevisa: {ruta_log}")

    if sin_seguro:
        wb_ss = openpyxl.Workbook()
        ws_ss = wb_ss.active
        ws_ss.title = "SinSeguro"
        ws_ss.append(["Nombre Paciente", "Cedula Paciente", "Cedula Consultada", "Rol", "Fecha"])

        for fila in sin_seguro:
            ws_ss.append(fila)

        ruta_ss = os.path.join(CARPETA_SALIDA, "sin_seguro.xlsx")
        wb_ss.save(ruta_ss)
        print(f"{len(sin_seguro)} caso(s) sin seguro.\nRevisa: {ruta_ss}")

    if not errores and not sin_seguro:
        print("\nTodo se proceso sin errores.")

# ============================================================
# MAIN (MODO EXCEL, IGUAL QUE ANTES)
# ============================================================

def procesar_lote(items, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                   errores, sin_seguro, callback_fila=None, callback_progreso=None,
                   escribir_matriz=False, responsable_matriz=None):
    """Procesa una lista de (indice_original, registro): la primera vez
    es el lote completo leido del Excel, y tambien se usa para
    reintentar solo un subconjunto (las filas que fallaron) desde la
    interfaz, sin tener que repetir todo el lote. 'errores' y
    'sin_seguro' reciben los resultados (se espera que estas listas ya
    traigan lo acumulado de corridas anteriores, si aplica, para que el
    reporte final no pierda nada).

    Si 'escribir_matriz' es True, cada paciente procesado que SI tiene
    seguro tambien se agrega a la matriz de Excel DEL MES QUE LE
    CORRESPONDA segun su fecha de atencion (nunca se mezclan fechas de
    meses distintos en el mismo archivo), con Dependencia, fecha de
    nacimiento y sexo EN BLANCO (el modo por lotes no trae esos datos;
    se completan despues a mano) y 'responsable_matriz' como Responsable
    en todas las filas. Si 'escribir_matriz' es False (por defecto), el
    comportamiento es igual que antes: solo se descargan los PDF, sin
    tocar la matriz."""
    total = len(items)
    for n, (indice, reg) in enumerate(items, start=1):
        filas_matriz = procesar_registro(
            reg, indice, total, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
            errores, sin_seguro, callback_fila=callback_fila, recolectar_matriz=escribir_matriz
        )
        if escribir_matriz and filas_matriz:
            try:
                agregar_filas_matriz_con_bloqueo(
                    filas_matriz, "", None, "", "", responsable_matriz or ""
                )
            except TimeoutError as e:
                print(f"   ERROR guardando en la matriz: {e}")
                errores.append([reg.get("nombre") or reg["cedula"], reg["cedula"], "", f"(MATRIZ) {e}"])
        if callback_progreso:
            callback_progreso(n, total)


def main(callback_fila=None, callback_progreso=None):
    registros = leer_excel(ARCHIVO_EXCEL)

    if LIMITE_PRUEBA:
        registros = registros[:LIMITE_PRUEBA]

    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    carpeta_descargas_temp = os.path.join(CARPETA_SALIDA, "_descargas_temp")
    os.makedirs(carpeta_descargas_temp, exist_ok=True)
    carpeta_diagnostico = os.path.join(CARPETA_SALIDA, "_diagnostico")

    driver_holder = [crear_driver(carpeta_descargas_temp)]

    errores = []
    sin_seguro = []
    items = list(enumerate(registros, start=1))

    try:
        procesar_lote(
            items, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
            errores, sin_seguro, callback_fila=callback_fila, callback_progreso=callback_progreso
        )
    finally:
        try:
            driver_holder[0].quit()
        except Exception:
            pass

    _guardar_reportes_finales(errores, sin_seguro)

# ============================================================
# ENTRADA INTERACTIVA (MODO MANUAL, UN PACIENTE A LA VEZ)
# ============================================================

def _preguntar(mensaje, permitir_vacio=False):
    while True:
        valor = input(mensaje).strip()
        if valor or permitir_vacio:
            return valor
        print("   Este dato es obligatorio, intenta de nuevo.")

def pedir_cedula():
    while True:
        texto = _preguntar("Cedula del paciente: ")
        solo_digitos = re.sub(r"\D", "", texto)
        if not solo_digitos:
            print("   Ingresa solo numeros.")
            continue
        try:
            cedula = str(int(solo_digitos)).zfill(10)
        except ValueError:
            print("   Cedula invalida, intenta de nuevo.")
            continue
        if len(cedula) != 10:
            print("   La cedula debe tener 10 digitos.")
            continue
        return cedula

def _parsear_fecha_flexible(texto):
    texto = texto.strip()
    for formato in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    return None

def pedir_fecha(mensaje, permitir_vacio_hoy=False, no_futura=False, valor_por_defecto=None):
    while True:
        texto = input(mensaje).strip()
        if not texto:
            if valor_por_defecto:
                return valor_por_defecto
            if permitir_vacio_hoy:
                return date.today()
        fecha = _parsear_fecha_flexible(texto)
        if not fecha:
            print("   Fecha invalida. Usa el formato DD-MM-YYYY (ej: 05-09-2026).")
            continue
        if no_futura and fecha > date.today():
            print("   Esa fecha todavia no ha pasado, revisa el dato.")
            continue
        return fecha

def pedir_sexo(valor_por_defecto=None):
    while True:
        sugerencia = f" [Enter = {valor_por_defecto}]" if valor_por_defecto else ""
        texto = _preguntar(f"Sexo del paciente (M/F){sugerencia}: ").strip().upper()
        if not texto and valor_por_defecto:
            return valor_por_defecto
        if texto in ("M", "F"):
            return texto
        print("   Ingresa 'M' o 'F'.")

def pedir_confirmacion(mensaje):
    while True:
        texto = input(mensaje).strip().upper()
        if texto in ("S", "SI"):
            return True
        if texto in ("N", "NO"):
            return False
        print("   Responde S (si) o N (no).")

def pedir_texto_opcional(mensaje):
    return input(mensaje).strip()

# ============================================================
# LECTURA / ESCRITURA DE LA MATRIZ (INSTRUCTIVO / hoja MES_2026)
# ============================================================

# Columnas con formula que la propia plantilla ya trae precargadas y que
# NO se deben pisar (C, D, G, L, O, P, R). Si alguna vez se necesita
# agregar una fila mas alla de donde llegan esas formulas, se "arrastran"
# hacia abajo traduciendo las referencias relativas.
_COLUMNAS_FORMULA_MATRIZ = [3, 4, 7, 12, 15, 16, 18]

def _nombre_matriz_para_mes(nombre_archivo_base, fecha):
    """A partir del nombre de un archivo de matriz que ya existe (el mas
    reciente), arma el nombre que deberia tener el archivo del mes de
    'fecha': si el nombre base ya trae un mes y anio reconocibles (ej.
    'INSTRUCTIVO5 AGOSTO 2026.xlsx'), se reemplazan por el mes/anio
    nuevos, conservando el separador que se venia usando (espacio o
    guion bajo). Si no se reconoce ningun mes en el nombre, simplemente
    se le agrega el mes y anio nuevos al final."""
    base_sin_ext, ext = os.path.splitext(nombre_archivo_base)
    mes_nuevo = nombre_mes_es(fecha)
    anio_nuevo = str(fecha.year)

    patron = re.compile(r"(" + "|".join(MESES_ES) + r")([ _-]*)(\d{4})", re.IGNORECASE)
    m = patron.search(base_sin_ext.upper())
    if m:
        separador = m.group(2) or " "
        inicio, fin = m.span()
        base_nueva = base_sin_ext[:inicio] + f"{mes_nuevo}{separador}{anio_nuevo}" + base_sin_ext[fin:]
    else:
        base_nueva = f"{base_sin_ext} {mes_nuevo} {anio_nuevo}"

    return base_nueva + ext

def _clonar_matriz_para_mes_actual(ruta_base, fecha_actual):
    """Genera el archivo de la matriz del mes de 'fecha_actual' clonando
    TODA la estructura (encabezados, formulas, formato, y la hoja
    MAESTRO con sus dependencias) del archivo de matriz mas reciente que
    ya existe, pero vaciando las filas de pacientes de HOJA_MATRIZ para
    que el mes nuevo arranque limpio. Asi no hace falta subir el Excel a
    mano cada mes: se genera solo, y como siempre se clona del archivo
    mas reciente, si ese archivo tiene cambios (una dependencia nueva en
    MAESTRO, una columna nueva, etc.) el mes nuevo los hereda
    automaticamente."""
    nombre_nuevo = _nombre_matriz_para_mes(os.path.basename(ruta_base), fecha_actual)
    ruta_nueva = os.path.join(os.path.dirname(ruta_base), nombre_nuevo)

    if os.path.exists(ruta_nueva):
        return ruta_nueva

    wb = openpyxl.load_workbook(ruta_base)
    ws = wb[HOJA_MATRIZ]

    # Se borran solo las filas de datos (fila 2 en adelante); encabezados,
    # formulas y formato quedan tal como estaban en el archivo base.
    if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)

    wb.save(ruta_nueva)
    print(f"Se genero automaticamente la matriz del mes: {os.path.basename(ruta_nueva)}\n"
          f"   (a partir de: {os.path.basename(ruta_base)})")
    return ruta_nueva

_SUFIJO_LOCK_DATOS_PACIENTES = ".lock"

@contextlib.contextmanager
def _bloqueo_datos_pacientes(timeout=15):
    """Mismo mecanismo de candado que se usa para la matriz (ver
    bloqueo_matriz), aplicado a datos_pacientes.json: evita que dos PC
    (o dos ventanas abiertas al mismo tiempo) que comparten este archivo
    por red se pisen entre si al guardar. ANTES, si dos PC guardaban
    casi al mismo tiempo, la ultima en guardar podia sobrescribir el
    archivo entero solo con lo que ella tenia en memoria, borrando sin
    querer los pacientes que la otra PC acababa de agregar."""
    ruta_lock = RUTA_DATOS_PACIENTES + _SUFIJO_LOCK_DATOS_PACIENTES
    fin = time.time() + timeout
    adquirido = False
    while time.time() < fin:
        try:
            fd = os.open(ruta_lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            adquirido = True
            break
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(ruta_lock) > 30:
                    os.remove(ruta_lock)
                    continue
            except OSError:
                pass
            time.sleep(0.2)
    try:
        yield
    finally:
        if adquirido:
            try:
                os.remove(ruta_lock)
            except OSError:
                pass

def _cargar_datos_pacientes():
    """Lee datos_pacientes.json. Si el archivo no existe todavia (nunca
    se ha guardado nadie), devuelve un diccionario vacio -eso es normal.
    Si el archivo SI existe pero no se pudo leer o esta corrupto, lanza
    la excepcion en vez de devolver un diccionario vacio en silencio:
    devolver vacio ahi seria peligroso, porque un guardado inmediatamente
    despues sobrescribiria el archivo real con uno casi vacio, borrando
    todo lo que ya se tenia guardado."""
    if not os.path.exists(RUTA_DATOS_PACIENTES):
        return {}
    with open(RUTA_DATOS_PACIENTES, "r", encoding="utf-8") as f:
        return json.load(f)

def _guardar_datos_pacientes(datos):
    with open(RUTA_DATOS_PACIENTES, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False, indent=2)

def obtener_datos_guardados_paciente(cedula):
    """Devuelve un dict con los datos que se recuerdan de esta cedula de
    una vez anterior (nombre, fecha_nacimiento como texto DD-MM-YYYY,
    sexo como 'M'/'F'), o None si es la primera vez que se ve, o si el
    archivo no se pudo leer en este momento (por ejemplo, otra PC lo
    esta guardando justo ahora)."""
    try:
        with _bloqueo_datos_pacientes():
            return _cargar_datos_pacientes().get(cedula)
    except Exception as e:
        print(f"   AVISO: no se pudieron leer los datos guardados de pacientes ({RUTA_DATOS_PACIENTES}): {e}")
        return None

def guardar_datos_paciente(cedula, nombre=None, fecha_nacimiento=None, sexo=None):
    """Recuerda estos datos de la cedula para la proxima vez que se
    ingrese, para no tener que volver a escribirlos. Solo actualiza los
    campos que se pasen (los que sean None/vacios no se tocan, no borran
    lo que ya estaba guardado). Toma el candado del archivo mientras lee
    y vuelve a guardar, para que dos PC no se pisen entre si; si por
    algun motivo no se puede leer el archivo existente con seguridad,
    NO se guarda nada esta vez (mejor perder este ultimo dato que
    arriesgarse a borrar todo lo demas que ya estaba guardado)."""
    try:
        with _bloqueo_datos_pacientes():
            datos = _cargar_datos_pacientes()
            actual = datos.get(cedula, {})
            if nombre:
                actual["nombre"] = nombre
            if fecha_nacimiento:
                actual["fecha_nacimiento"] = (
                    fecha_nacimiento.strftime("%d-%m-%Y") if hasattr(fecha_nacimiento, "strftime")
                    else fecha_nacimiento
                )
            if sexo:
                actual["sexo"] = sexo
            datos[cedula] = actual
            _guardar_datos_pacientes(datos)
    except Exception as e:
        print(f"   AVISO: no se pudieron guardar los datos de la cedula {cedula} para la proxima vez: {e}")

def localizar_archivo_matriz(fecha=None):
    """Devuelve la ruta del archivo de la matriz del MES DE 'fecha' (por
    defecto, el mes actual si no se indica). Ya NO hace falta subirlo a
    mano cada mes, NI mezclar en el mismo archivo pacientes de meses
    distintos: si no existe todavia un archivo para ese mes puntual, se
    genera solo clonando la estructura del archivo INSTRUCTIVO...xlsx
    mas reciente que ya exista en la carpeta del script (ver
    _clonar_matriz_para_mes_actual). Solo hay que subir el Excel una
    vez; de ahi en adelante cada mes que haga falta (incluido uno
    anterior, si se ingresa una fecha de atencion atrasada) se genera
    solo a partir del mas reciente que ya exista, heredando cualquier
    cambio que se le haya hecho."""
    fecha = fecha or date.today()
    patron_mes_pedido = re.compile(
        re.escape(nombre_mes_es(fecha)) + r"[ _-]*" + re.escape(str(fecha.year)), re.IGNORECASE
    )

    if os.path.exists(ARCHIVO_MATRIZ) and patron_mes_pedido.search(os.path.basename(ARCHIVO_MATRIZ).upper()):
        return ARCHIVO_MATRIZ

    candidatos = []
    try:
        for nombre_archivo in os.listdir(BASE_DIR):
            nombre_norm = nombre_archivo.upper()
            if (nombre_norm.startswith("INSTRUCTIVO") and nombre_norm.endswith(".XLSX")
                    and not nombre_archivo.startswith("~$")):
                candidatos.append(os.path.join(BASE_DIR, nombre_archivo))
    except Exception:
        pass

    if not candidatos:
        raise FileNotFoundError(
            "No se encontro ningun archivo 'INSTRUCTIVO...xlsx' en la carpeta del script:\n"
            f"{BASE_DIR}\nSube uno (una sola vez) para que sirva de base; los demas meses "
            "se generaran solos a partir de el."
        )

    candidatos.sort(key=os.path.getmtime, reverse=True)

    # ¿Ya existe, por nombre, un archivo del mes pedido entre los
    # candidatos? (ej. si alguien lo subio a mano igual, o ya se genero
    # antes en una corrida anterior).
    for ruta in candidatos:
        if patron_mes_pedido.search(os.path.basename(ruta).upper()):
            return ruta

    if len(candidatos) > 1:
        print("AVISO: se encontro mas de un archivo INSTRUCTIVO...xlsx en la carpeta del "
              "script; se usara el modificado mas recientemente como base para generar "
              f"la matriz de {nombre_carpeta_mes(fecha)}: " + os.path.basename(candidatos[0]))

    return _clonar_matriz_para_mes_actual(candidatos[0], fecha)

def _hacer_respaldo_matriz(ruta):
    os.makedirs(CARPETA_RESPALDOS_MATRIZ, exist_ok=True)
    marca = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.splitext(os.path.basename(ruta))[0]
    destino = os.path.join(CARPETA_RESPALDOS_MATRIZ, f"{base}_{marca}.xlsx")
    shutil.copy2(ruta, destino)
    return destino

def _carpeta_descargas_windows():
    """Devuelve la carpeta 'Descargas' del usuario de Windows (la que
    Windows use en ese momento, aunque el usuario la haya movido a otra
    unidad/ruta desde las propiedades de la carpeta). Si algo falla
    (no es Windows, no se pudo consultar), cae de respaldo a
    <carpeta del usuario>\\Downloads, que es lo correcto en el 99% de
    los casos."""
    respaldo = os.path.join(os.path.expanduser("~"), "Downloads")
    if sys.platform != "win32":
        return respaldo
    try:
        import ctypes
        from ctypes import wintypes
        import uuid

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", wintypes.DWORD),
                ("Data2", wintypes.WORD),
                ("Data3", wintypes.WORD),
                ("Data4", ctypes.c_byte * 8),
            ]

            def __init__(self, uuid_):
                ctypes.Structure.__init__(self)
                self.Data1, self.Data2, self.Data3, self.Data4[0], self.Data4[1], resto = uuid_.fields
                for i in range(2, 8):
                    self.Data4[i] = (resto >> (8 - i - 1) * 8) & 0xFF

        FOLDERID_DOWNLOADS = uuid.UUID("{374DE290-123F-4565-9164-39C4925E467B}")
        guid = GUID(FOLDERID_DOWNLOADS)
        buf = ctypes.c_wchar_p()
        resultado = ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, 0, ctypes.byref(buf))
        if resultado == 0 and buf.value:
            return buf.value
        return respaldo
    except Exception:
        return respaldo


def _fecha_en_rango(valor_celda, fecha_desde, fecha_hasta):
    """Compara el valor de la celda 'Fecha atencion' (puede venir como
    date, datetime o texto DD/MM/YYYY, segun como lo haya dejado Excel)
    contra el rango [fecha_desde, fecha_hasta], ambos inclusive."""
    if valor_celda in (None, ""):
        return False
    if isinstance(valor_celda, datetime):
        fecha_celda = valor_celda.date()
    elif isinstance(valor_celda, date):
        fecha_celda = valor_celda
    else:
        fecha_celda = _parsear_fecha_flexible(str(valor_celda))
        if not fecha_celda:
            return False
    return fecha_desde <= fecha_celda <= fecha_hasta


def generar_copia_matriz(carpeta_destino=None, modo_fecha="todo",
                          fecha_desde=None, fecha_hasta=None):
    """Genera una copia del archivo de la matriz para revisar o enviar,
    sin arriesgar el archivo original que se sigue usando. No modifica
    nada del Instructivo real.

    carpeta_destino: donde se guarda la copia. Si no se indica, se usa
        la carpeta de Descargas de Windows del usuario actual.

    modo_fecha: "todo" (copia todas las filas, tal como antes),
        "hoy" (solo filas cuya 'Fecha atencion' es la fecha de hoy),
        "mes" (todas las filas del mes/anio de 'fecha_desde'; si
        'fecha_desde' no se indica, se usa el mes/anio actual), o
        "rango" (filas cuya 'Fecha atencion' esta entre 'fecha_desde' y
        'fecha_hasta', ambas inclusive; ambas son obligatorias en este
        modo).

    Devuelve la ruta del archivo generado."""
    ruta = localizar_archivo_matriz()
    marca = datetime.now().strftime("%Y-%m-%d_%H-%M")
    base = os.path.splitext(os.path.basename(ruta))[0]

    carpeta_destino = carpeta_destino or _carpeta_descargas_windows()
    os.makedirs(carpeta_destino, exist_ok=True)

    modo_fecha = (modo_fecha or "todo").lower()

    if modo_fecha == "todo":
        nombre_copia = f"{base} - copia {marca}.xlsx"
        destino = os.path.join(carpeta_destino, nombre_copia)
        shutil.copy2(ruta, destino)
        return destino

    if modo_fecha == "hoy":
        fecha_desde = fecha_hasta = date.today()
        sufijo = f"dia {fecha_desde.strftime('%d-%m-%Y')}"
    elif modo_fecha == "mes":
        base_mes = fecha_desde or date.today()
        fecha_desde = base_mes.replace(day=1)
        if base_mes.month == 12:
            fecha_hasta = base_mes.replace(day=31)
        else:
            fecha_hasta = base_mes.replace(month=base_mes.month + 1, day=1) - timedelta(days=1)
        sufijo = f"mes {fecha_desde.strftime('%m-%Y')}"
    elif modo_fecha == "rango":
        if not fecha_desde or not fecha_hasta:
            raise ValueError("Para modo_fecha='rango' hay que indicar fecha_desde y fecha_hasta.")
        sufijo = f"{fecha_desde.strftime('%d-%m-%Y')} a {fecha_hasta.strftime('%d-%m-%Y')}"
    else:
        raise ValueError("modo_fecha debe ser 'todo', 'hoy', 'mes' o 'rango'.")

    wb_origen = openpyxl.load_workbook(ruta)
    ws_origen = wb_origen[HOJA_MATRIZ]

    filas_a_borrar = []
    for fila in range(2, ws_origen.max_row + 1):
        valor_fecha = ws_origen.cell(row=fila, column=5).value  # E = Fecha atencion
        if not _fecha_en_rango(valor_fecha, fecha_desde, fecha_hasta):
            filas_a_borrar.append(fila)

    # Se borra de abajo hacia arriba para no desfasar los indices de fila
    # segun se van eliminando.
    for fila in reversed(filas_a_borrar):
        ws_origen.delete_rows(fila, 1)

    nombre_copia = f"{base} - copia {sufijo} {marca}.xlsx"
    destino = os.path.join(carpeta_destino, nombre_copia)
    wb_origen.save(destino)
    return destino

def abrir_matriz(ruta):
    """Abre el archivo de la matriz (haciendo antes una copia de
    respaldo, por seguridad: nunca se toca el archivo real sin dejar
    una copia previa en _respaldos_matriz/). 'ruta' debe venir de
    localizar_archivo_matriz()."""
    if not os.path.exists(ruta):
        raise FileNotFoundError(
            f"No se encontro el archivo de la matriz: {ruta}\n"
            "Debe estar guardado en la misma carpeta que este script."
        )
    ruta_respaldo = _hacer_respaldo_matriz(ruta)
    print(f"(Respaldo de la matriz guardado en: {ruta_respaldo})")
    return openpyxl.load_workbook(ruta)

def obtener_dependencias_validas(wb_matriz):
    """Lee la lista de DEPENDENCIA (tipo de consulta) directamente de
    MAESTRO!D4:D.., para no mantener una lista separada en este script:
    si el maestro cambia, esto se actualiza solo."""
    ws = wb_matriz[HOJA_MAESTRO]
    valores = []
    fila = 4
    while True:
        valor = ws.cell(row=fila, column=4).value  # columna D
        if valor in (None, ""):
            break
        valores.append(str(valor).strip())
        fila += 1
    return valores

def pedir_dependencia(dependencias_validas):
    if dependencias_validas:
        print("   Dependencias disponibles:")
        for idx, valor in enumerate(dependencias_validas, start=1):
            print(f"     {idx}) {valor}")
        print("     0) Otra (escribir manualmente)")
        while True:
            eleccion = _preguntar("   Elige un numero (o escribe el nombre): ")
            if eleccion.isdigit():
                num = int(eleccion)
                if num == 0:
                    return _preguntar("   Escribe la dependencia: ").upper()
                if 1 <= num <= len(dependencias_validas):
                    return dependencias_validas[num - 1]
                print("   Numero fuera de rango.")
                continue
            return eleccion.upper()
    return _preguntar("Dependencia (tipo de consulta): ").upper()

def siguiente_fila_matriz(ws):
    fila = 2
    while ws.cell(row=fila, column=COL_MATRIZ_CEDULA).value not in (None, ""):
        fila += 1
    return fila

def _ultima_fila_con_formula(ws, columna):
    fila = 2
    ultima = 1
    while ws.cell(row=fila, column=columna).value not in (None, ""):
        ultima = fila
        fila += 1
        if fila > 5000:
            break
    return ultima

def asegurar_formulas_fila(ws, fila_objetivo):
    """Si fila_objetivo esta mas alla de donde llegan las formulas
    precargadas de la plantilla, las copia hacia esa fila traduciendo
    las referencias relativas (lo mismo que "arrastrar" la formula hacia
    abajo a mano en Excel)."""
    for columna in _COLUMNAS_FORMULA_MATRIZ:
        origen = _ultima_fila_con_formula(ws, columna)
        formula_origen = ws.cell(row=origen, column=columna).value
        if not formula_origen or not str(formula_origen).startswith("="):
            continue
        if fila_objetivo <= origen:
            continue
        celda_origen = ws.cell(row=origen, column=columna).coordinate
        celda_destino = ws.cell(row=fila_objetivo, column=columna).coordinate
        nueva_formula = Translator(formula_origen, origin=celda_origen).translate_formula(celda_destino)
        ws.cell(row=fila_objetivo, column=columna).value = nueva_formula

def _normalizar_fecha_celda(valor):
    """Convierte lo que venga en una celda de fecha (datetime, date, o
    texto DD/MM/YYYY) a un date de Python, para poder comparar fechas
    sin importar como haya quedado guardado el valor."""
    if valor in (None, ""):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return _parsear_fecha_flexible(str(valor))

def existe_fila_matriz(ws, cedula, fecha_atencion, dependencia):
    """True si ya hay una fila en la matriz con esta MISMA cedula de
    paciente (columna H), esta MISMA fecha de atencion (columna E) Y
    esta MISMA dependencia (columna B). Se usa para no duplicar al
    paciente si se lo vuelve a ingresar para la misma fecha Y la misma
    dependencia; si cambia la fecha de atencion O la dependencia (por
    ejemplo, la misma fecha pero otra consulta distinta), SI debe
    quedar como una fila nueva."""
    cedula = str(cedula).strip()
    fecha_atencion = _normalizar_fecha_celda(fecha_atencion)
    dependencia = str(dependencia or "").strip().upper()
    fila = 2
    while ws.cell(row=fila, column=COL_MATRIZ_CEDULA).value not in (None, ""):
        cedula_fila = str(ws.cell(row=fila, column=COL_MATRIZ_CEDULA).value or "").strip()
        if cedula_fila == cedula:
            fecha_fila = _normalizar_fecha_celda(ws.cell(row=fila, column=5).value)
            dependencia_fila = str(ws.cell(row=fila, column=2).value or "").strip().upper()
            if fecha_fila == fecha_atencion and dependencia_fila == dependencia:
                return True
        fila += 1
    return False

def agregar_fila_matriz(ws, info, dependencia, fecha_nacimiento, sexo, observaciones, responsable):
    """Escribe una fila nueva en MES_2026 con lo que ya determino el
    motor de descarga (institucion, cedula/nombre del paciente, fecha de
    atencion, parentesco, cedula/nombre del afiliado, tipo de seguro
    beneficiario) mas lo que se pidio en el formulario (dependencia,
    fecha de nacimiento, sexo, observaciones, responsable). El CIE10 y
    "primera/subsecuente" y "diagnostico presuntivo/definitivo" se dejan
    en blanco a proposito (se completan en otro proceso).

    Si esta MISMA cedula de paciente YA tiene una fila con esta MISMA
    fecha de atencion Y esta MISMA dependencia, no se agrega una fila
    duplicada -se devuelve None-; si cambia la fecha de atencion o la
    dependencia (otra visita, u otra consulta el mismo dia), SI se
    agrega normalmente.

    Las columnas con formula de la plantilla (codigo de dependencia,
    numero de expediente, codigo de tipo de seguro, edad, duracion,
    codigo de parentesco) no se tocan: se siguen calculando solas en
    Excel apenas se abra el archivo. CI AFILIADO / APELLIDOS Y NOMBRES
    (afiliado) SI se escriben como valor fijo, porque la formula
    original de la plantilla solo sabe poner al propio paciente cuando
    es titular; cuando no lo es, hay que poner los datos reales del
    titular/acreditador que el motor de descarga ya identifico.
    """
    if existe_fila_matriz(ws, info["cedula_paciente"], info["fecha_atencion"], dependencia):
        return None

    fila = siguiente_fila_matriz(ws)
    asegurar_formulas_fila(ws, fila)

    ws.cell(row=fila, column=1, value=info.get("institucion", ""))                 # A
    ws.cell(row=fila, column=2, value=dependencia)                                 # B
    c_fecha = ws.cell(row=fila, column=5, value=info["fecha_atencion"])            # E
    c_fecha.number_format = "DD/MM/YYYY"
    ws.cell(row=fila, column=6, value=info.get("tipo_seguro_beneficiario", ""))    # F
    c_cedula = ws.cell(row=fila, column=8, value=info["cedula_paciente"])          # H
    c_cedula.number_format = "@"
    ws.cell(row=fila, column=9, value=str(info["nombre_paciente"]).upper())        # I
    ws.cell(row=fila, column=10, value=sexo)                                       # J
    c_nac = ws.cell(row=fila, column=11, value=fecha_nacimiento)                   # K
    c_nac.number_format = "DD/MM/YYYY"
    ws.cell(row=fila, column=17, value=info.get("parentesco", ""))                 # Q
    c_afiliado = ws.cell(row=fila, column=19, value=info.get("cedula_afiliado", ""))  # S
    c_afiliado.number_format = "@"
    ws.cell(row=fila, column=20, value=str(info.get("nombre_afiliado", "")).upper())  # T
    ws.cell(row=fila, column=22, value=observaciones)                              # V
    ws.cell(row=fila, column=23, value=responsable)                                # W

    return fila

_SUFIJO_LOCK_MATRIZ = ".lock"

@contextlib.contextmanager
def bloqueo_matriz(ruta_matriz, timeout=30):
    """Candado simple basado en un archivo, para que dos (o mas) PC que
    comparten la matriz por red no se pisen entre si al guardarla al
    mismo tiempo. ANTES cada PC mantenia su propia copia del archivo
    abierta en memoria toda la sesion y la sobreescribia completa cada
    vez que guardaba un paciente; si dos PC lo hacian a la vez, la
    ultima en guardar borraba los cambios que la otra acababa de hacer.
    Con este candado, cada PC espera su turno para leer la version MAS
    RECIENTE del archivo, agregar su fila, y guardar -- sin perder lo
    que la otra PC ya escribio.

    Se implementa creando '<matriz>.lock' de forma EXCLUSIVA (falla si
    ya existe); si ya esta tomado, se espera y reintenta hasta
    'timeout' segundos. Si el candado quedo huerfano (una PC se cerro
    de golpe con el candado puesto) y tiene mas de 2 minutos de
    antiguedad, se considera vencido y se libera solo."""
    ruta_lock = ruta_matriz + _SUFIJO_LOCK_MATRIZ
    fin = time.time() + timeout
    adquirido = False
    while time.time() < fin:
        try:
            fd = os.open(ruta_lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            adquirido = True
            break
        except FileExistsError:
            try:
                antiguedad = time.time() - os.path.getmtime(ruta_lock)
                if antiguedad > 120:
                    os.remove(ruta_lock)
                    continue
            except OSError:
                pass
            time.sleep(0.3)

    if not adquirido:
        raise TimeoutError(
            "No se pudo obtener acceso exclusivo a la matriz: otra computadora la esta "
            "guardando en este momento. Intenta de nuevo en unos segundos."
        )

    try:
        yield
    finally:
        try:
            os.remove(ruta_lock)
        except OSError:
            pass

def agregar_filas_matriz_con_bloqueo(filas_info, dependencia, fecha_nacimiento,
                                      sexo, observaciones, responsable):
    """Version segura-para-red de agregar_fila_matriz: agrega cada fila
    a la matriz DEL MES QUE LE CORRESPONDE segun su propia fecha de
    atencion -nunca fuerza todo al mes actual-, para que fechas de
    meses distintos jamas se mezclen en el mismo archivo; si hace falta,
    esa matriz mensual se genera sola (ver localizar_archivo_matriz).
    Agrupa las filas por mes para abrir/guardar cada archivo una sola
    vez, y toma el candado de cada uno mientras escribe -- asi dos PC
    usando la matriz al mismo tiempo nunca se borran los cambios entre
    si. Si alguna fila es un duplicado exacto (misma cedula, misma
    fecha de atencion Y misma dependencia) de una que ya esta en la
    matriz de ese mes, se salta -no se agrega de nuevo-; si cambia la
    fecha o la dependencia, SI se agrega como fila nueva. Devuelve la
    lista de (ruta_matriz, numero_de_fila) de las filas que quedaron
    escritas (sin contar los duplicados saltados)."""
    grupos = {}
    orden_rutas = []
    for info in filas_info:
        ruta_matriz = localizar_archivo_matriz(info["fecha_atencion"])
        if ruta_matriz not in grupos:
            grupos[ruta_matriz] = []
            orden_rutas.append(ruta_matriz)
        grupos[ruta_matriz].append(info)

    resultados = []
    for ruta_matriz in orden_rutas:
        filas_escritas_aqui = []
        with bloqueo_matriz(ruta_matriz):
            wb = abrir_matriz(ruta_matriz)
            ws = wb[HOJA_MATRIZ]
            for info in grupos[ruta_matriz]:
                fila = agregar_fila_matriz(
                    ws, info, dependencia, fecha_nacimiento, sexo, observaciones, responsable
                )
                if fila is None:
                    print(f"   (Ya existe una fila para la cedula {info['cedula_paciente']} con fecha "
                          f"{info['fecha_atencion'].strftime('%d-%m-%Y')} y esta misma dependencia; "
                          "no se duplica.)")
                    continue
                filas_escritas_aqui.append(fila)
                resultados.append((ruta_matriz, fila))
            if filas_escritas_aqui:
                wb.save(ruta_matriz)
    return resultados

# ============================================================
# MODO MANUAL / INTERACTIVO (un paciente a la vez)
# ============================================================

def modo_interactivo():
    print("\n=== MODO MANUAL (un paciente a la vez) ===\n")

    try:
        ruta_matriz = localizar_archivo_matriz()
    except FileNotFoundError as e:
        print(f"AVISO: {e}")
        print("Se descargaran los PDF de cobertura con normalidad, pero NO se podra "
              "llenar la matriz automaticamente.\n")
        ruta_matriz = None
        dependencias_validas = []
    else:
        # Esta apertura es solo para leer la lista de dependencias
        # validas (MAESTRO); NO se guarda en memoria para escribir mas
        # tarde, porque si esta computadora se queda con esta copia
        # abierta durante toda la sesion mientras otra PC tambien esta
        # agregando pacientes a la misma matriz por red, la que guarde
        # de ultimo terminaria borrando lo que la otra ya escribio. Cada
        # fila se agrega releyendo el archivo tal como este en ese
        # momento, y en la matriz del MES QUE LE CORRESPONDA segun su
        # propia fecha de atencion (ver agregar_filas_matriz_con_bloqueo
        # mas abajo), no siempre esta del mes actual.
        dependencias_validas = obtener_dependencias_validas(abrir_matriz(ruta_matriz))

    hay_matriz_disponible = ruta_matriz is not None

    responsable = _preguntar("Nombre del responsable (persona que ingresa la informacion): ").upper()

    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    carpeta_descargas_temp = os.path.join(CARPETA_SALIDA, "_descargas_temp")
    os.makedirs(carpeta_descargas_temp, exist_ok=True)
    carpeta_diagnostico = os.path.join(CARPETA_SALIDA, "_diagnostico")

    driver_holder = [crear_driver(carpeta_descargas_temp)]
    errores = []
    sin_seguro = []
    contador = 0
    filas_agregadas = 0

    try:
        while True:
            contador += 1
            print(f"\n--- Paciente #{contador} ---")

            cedula = pedir_cedula()

            datos_guardados = obtener_datos_guardados_paciente(cedula)
            if datos_guardados:
                partes = []
                if datos_guardados.get("nombre"):
                    partes.append(datos_guardados["nombre"])
                if datos_guardados.get("fecha_nacimiento"):
                    partes.append(f"nacio {datos_guardados['fecha_nacimiento']}")
                if datos_guardados.get("sexo"):
                    partes.append(f"sexo {datos_guardados['sexo']}")
                print(f"   (Ya se tienen datos guardados de esta cedula: {', '.join(partes)})")

            fecha_atencion = pedir_fecha(
                f"Fecha de atencion (DD-MM-YYYY) [Enter = hoy {date.today().strftime('%d-%m-%Y')}]: ",
                permitir_vacio_hoy=True
            )
            dependencia = pedir_dependencia(dependencias_validas)

            fecha_nacimiento_guardada = None
            if datos_guardados and datos_guardados.get("fecha_nacimiento"):
                fecha_nacimiento_guardada = _parsear_fecha_flexible(datos_guardados["fecha_nacimiento"])
            sugerencia_fecha_nac = (
                f" [Enter = {fecha_nacimiento_guardada.strftime('%d-%m-%Y')}]" if fecha_nacimiento_guardada else ""
            )
            fecha_nacimiento = pedir_fecha(
                f"Fecha de nacimiento del paciente (DD-MM-YYYY){sugerencia_fecha_nac}: ", no_futura=True,
                valor_por_defecto=fecha_nacimiento_guardada
            )

            sexo = pedir_sexo(valor_por_defecto=(datos_guardados or {}).get("sexo"))
            observaciones = pedir_texto_opcional("Observaciones (Enter para dejar en blanco): ")

            reg = {
                "cedula": cedula,
                "nombre": None,
                "fechas": [fecha_atencion],
                "cedula_padre": None,
                "cedula_titular": None,
            }

            filas_matriz = procesar_registro(
                reg, contador, contador, driver_holder, carpeta_descargas_temp, carpeta_diagnostico,
                errores, sin_seguro, recolectar_matriz=True
            )

            guardar_datos_paciente(
                cedula, nombre=reg.get("nombre"), fecha_nacimiento=fecha_nacimiento, sexo=sexo
            )

            if hay_matriz_disponible and filas_matriz:
                try:
                    filas_escritas = agregar_filas_matriz_con_bloqueo(
                        filas_matriz, dependencia, fecha_nacimiento,
                        sexo, observaciones, responsable
                    )
                except TimeoutError as e:
                    print(f"   ERROR guardando en la matriz: {e}")
                    errores.append([reg.get("nombre") or cedula, cedula, fecha_atencion.strftime("%d-%m-%Y"),
                                     f"(MATRIZ) {e}"])
                else:
                    filas_agregadas += len(filas_escritas)
                    for ruta_matriz_fila, fila in filas_escritas:
                        print(f"   -> Fila {fila} agregada a la matriz ({os.path.basename(ruta_matriz_fila)}).")

            if not pedir_confirmacion("\n¿Deseas ingresar otro paciente? (S/N): "):
                break

    finally:
        try:
            driver_holder[0].quit()
        except Exception:
            pass

    _guardar_reportes_finales(errores, sin_seguro)

    if hay_matriz_disponible and filas_agregadas > 0:
        print(f"\nSe guardaron {filas_agregadas} fila(s) en la(s) matriz(ces) correspondiente(s) "
              "(cada una en el archivo del mes de su fecha de atencion).")

        if pedir_confirmacion("¿Deseas generar una copia del archivo de la matriz en este momento? (S/N): "):
            print("   ¿Que quieres copiar?")
            print("     1) Todo")
            print("     2) Solo hoy")
            print("     3) Un mes completo")
            print("     4) Un rango de fechas")
            opcion = _preguntar("   Elige un numero [1]: ").strip() or "1"

            modo_fecha = "todo"
            fecha_desde = fecha_hasta = None
            if opcion == "2":
                modo_fecha = "hoy"
            elif opcion == "3":
                modo_fecha = "mes"
                fecha_desde = pedir_fecha(
                    "   Cualquier fecha del mes que quieres copiar (DD-MM-YYYY) "
                    f"[Enter = mes actual, {date.today().strftime('%m-%Y')}]: ",
                    permitir_vacio_hoy=True
                )
            elif opcion == "4":
                modo_fecha = "rango"
                fecha_desde = pedir_fecha("   Desde (DD-MM-YYYY): ")
                fecha_hasta = pedir_fecha("   Hasta (DD-MM-YYYY): ")

            carpeta_defecto = _carpeta_descargas_windows()
            carpeta_destino = _preguntar(
                f"   Carpeta de destino [Enter = Descargas de Windows: {carpeta_defecto}]: "
            ).strip() or carpeta_defecto

            copia = generar_copia_matriz(
                carpeta_destino=carpeta_destino,
                modo_fecha=modo_fecha,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta
            )
            print(f"Copia generada exitosamente en:\n{copia}")

if __name__ == "__main__":
    main()
