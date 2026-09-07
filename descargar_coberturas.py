# -*- coding: utf-8 -*-
"""
Backend principal para descarga de coberturas e integración de matriz.
Incluye:
- Organización de PDFs y Excel por mes (ej. "SEPTIEMBRE 2026").
- Impresión en consola con Nombres y Apellidos del paciente.
- Base de datos local/cache en JSON para autocompletar datos repetidos.
"""

import os
import re
import json
import time
import datetime
import pandas as pd
from datetime import datetime as dt

# --- RUTAS Y ARCHIVOS BASE ---
CARPETA_BASE = os.path.abspath("./RESULTADOS_COBERTURAS")
ARCHIVO_HISTORIAL = os.path.abspath("./historial_pacientes.json")
ARCHIVO_MATRIZ_BASE = os.path.abspath("./MATRIZ_BASE_MODELO.xlsx")
ARCHIVO_EXCEL_INCONSISTENCIAS = os.path.abspath("./reporte_inconsistencias.xlsx")

# --- NOMBRES DE MESES EN ESPAÑOL ---
MESES_ES = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

# ==============================================================================
# GESTIÓN DEL HISTORIAL DE PACIENTES (PUNTO 4)
# ==============================================================================
def cargar_historial():
    """Carga la base de datos de pacientes previamente consultados."""
    if os.path.exists(ARCHIVO_HISTORIAL):
        try:
            with open(ARCHIVO_HISTORIAL, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def guardar_paciente_historial(cedula, datos_paciente):
    """Guarda o actualiza a un paciente en la base de datos local."""
    historial = cargar_historial()
    historial[str(cedula)] = datos_paciente
    try:
        with open(ARCHIVO_HISTORIAL, "w", encoding="utf-8") as f:
            json.dump(historial, f, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"⚠️ No se pudo guardar el historial local: {e}")

def buscar_paciente_historial(cedula):
    """Busca si el paciente ya fue procesado con anterioridad."""
    historial = cargar_historial()
    return historial.get(str(cedula), None)

# ==============================================================================
# ORGANIZACIÓN POR CARPETAS DE MESES (PUNTO 1 Y 2)
# ==============================================================================
def obtener_carpeta_mes(fecha_str=None):
    """
    Retorna la ruta de la carpeta del mes según la fecha proporcionada (YYYY-MM-DD o DD-MM-YYYY).
    Ejemplo de carpeta: "./RESULTADOS_COBERTURAS/SEPTIEMBRE 2026"
    """
    if fecha_str:
        try:
            if "-" in fecha_str and len(fecha_str.split("-")[0]) == 4:
                fecha_obj = dt.strptime(fecha_str, "%Y-%m-%d")
            else:
                fecha_obj = dt.strptime(fecha_str, "%d-%m-%Y")
        except Exception:
            fecha_obj = dt.now()
    else:
        fecha_obj = dt.now()

    nombre_mes = f"{MESES_ES[fecha_obj.month]} {fecha_obj.year}"
    ruta_mes = os.path.join(CARPETA_BASE, nombre_mes)
    os.makedirs(ruta_mes, exist_ok=True)
    return ruta_mes, nombre_mes

def obtener_o_crear_matriz_mes(fecha_str=None):
    """Genera la matriz del mes basándose en el modelo inicial si no existe."""
    ruta_carpeta, nombre_mes = obtener_carpeta_mes(fecha_str)
    ruta_matriz_mes = os.path.join(ruta_carpeta, f"MATRIZ_COBERTURAS_{nombre_mes.replace(' ', '_')}.xlsx")

    # Si la matriz de ese mes no existe, copiamos el modelo base
    if not os.path.exists(ruta_matriz_mes):
        if os.path.exists(ARCHIVO_MATRIZ_BASE):
            import shutil
            shutil.copy(ARCHIVO_MATRIZ_BASE, ruta_matriz_mes)
            print(f"📂 Nueva matriz creada para el mes: {nombre_mes}")
        else:
            # Crear un Excel básico si no existe el modelo
            df_vacio = pd.DataFrame(columns=["CEDULA", "NOMBRES_APELLIDOS", "FECHA_ATENCION", "DEPENDENCIA", "ESTADO"])
            df_vacio.to_excel(ruta_matriz_mes, index=False)
            print(f"📄 Matriz generada automáticamente en: {ruta_matriz_mes}")

    return ruta_matriz_mes

# ==============================================================================
# PROCESAMIENTO GENERAL
# ==============================================================================
def procesar_paciente(paciente_data):
    """
    Procesa a un único paciente, revisa si ya existe en el historial (Punto 4),
    imprime Nombres y Apellidos en consola (Punto 3) y guarda PDF en carpeta mensual (Punto 1).
    """
    cedula = str(paciente_data.get("cedula", "")).zfill(10)
    fecha_atencion = paciente_data.get("fecha_atencion", dt.now().strftime("%Y-%m-%d"))
    
    # 1. Obtener carpetas correspondientes al mes
    ruta_carpeta_mes, nombre_mes = obtener_carpeta_mes(fecha_atencion)
    ruta_matriz = obtener_o_crear_matriz_mes(fecha_atencion)

    # 2. Revisar si el paciente ya fue ingresado previamente (Punto 4)
    paciente_guardado = buscar_paciente_historial(cedula)
    
    if paciente_guardado:
        nombres = paciente_guardado.get("nombres", "NO REGISTRADO")
        apellidos = paciente_guardado.get("apellidos", "")
        nombre_completo = f"{apellidos} {nombres}".strip()
        print(f"\n⚡ [HISTORIAL ENCONTRADO] Paciente: {nombre_completo} (Cédula: {cedula})")
        print(f"📁 Usando datos guardados previamente. Guardando en: {nombre_mes}")
        
        # Actualizamos la fecha de última atención
        paciente_guardado["ultima_atencion"] = fecha_atencion
        guardar_paciente_historial(cedula, paciente_guardado)
        return paciente_guardado

    # 3. Si no existe en el historial, procesar consulta (Simulación/Consulta Web)
    print(f"\n🔍 Consultando paciente nuevo Cédula: {cedula} en portales del CORE...")
    
    # --- AQUÍ VA LA LÓGICA DE SELENIUM / CONSULTA WEB ---
    # Ejemplo de nombres obtenidos de la consulta:
    nombres_obtenidos = paciente_data.get("nombres", "JUAN CARLOS")
    apellidos_obtenidos = paciente_data.get("apellidos", "PEREZ LOPEZ")
    nombre_completo = f"{apellidos_obtenidos} {nombres_obtenidos}".strip()

    # PUNTO 3: Imprimir nombres y apellidos en la consola al momento de crear archivo/carpeta
    print(f"👤 PACIENTE DETECTADO: {nombre_completo}")
    print(f"📁 Guardando comprobantes PDF en: {ruta_carpeta_mes}")

    # Guardar paciente procesado en historial local (Punto 4)
    registro_nuevo = {
        "cedula": cedula,
        "nombres": nombres_obtenidos,
        "apellidos": apellidos_obtenidos,
        "sexo": paciente_data.get("sexo", "M"),
        "fecha_nacimiento": paciente_data.get("fecha_nacimiento", ""),
        "ultima_atencion": fecha_atencion,
        "dependencia": paciente_data.get("dependencia", "")
    }
    guardar_paciente_historial(cedula, registro_nuevo)

    return registro_nuevo

def procesar_lista_pacientes(lista_pacientes):
    """Procesa una lista de pacientes uno por uno."""
    for pac in lista_pacientes:
        procesar_paciente(pac)

def main(callback_fila=None, callback_progreso=None):
    """Ejecución por lotes."""
    if os.path.exists(ARCHIVO_EXCEL_INCONSISTENCIAS):
        df = pd.read_excel(ARCHIVO_EXCEL_INCONSISTENCIAS)
        total = len(df)
        for idx, row in df.iterrows():
            paciente = {
                "cedula": str(row.get("CEDULA", "")),
                "fecha_atencion": str(row.get("FECHA_ATENCION", dt.now().strftime("%Y-%m-%d"))),
                "dependencia": str(row.get("DEPENDENCIA", ""))
            }
            procesar_paciente(paciente)
            if callback_progreso:
                callback_progreso(idx + 1, total)
