# -*- coding: utf-8 -*-
"""
probar_drive.py - Prueba la conexion con Google Drive ANTES de usarla.

Ejecutalo con:   python probar_drive.py

Comprueba, paso a paso:
  1. que la cuenta de servicio se autentica,
  2. que ve la carpeta "PLANILLAJE COBERTURAS",
  3. que encuentra (o crea) la carpeta de tu unidad / año / mes,
  4. que PUEDE SUBIR un archivo nuevo (esta es la prueba importante: si
     falla aqui, ninguna subida de PDF va a funcionar) y lo borra.

Usa la misma configuracion que el programa (config_drive.json).
Si todavia no existe, te pregunta la unidad y el archivo .json.
"""

import json
import os
import sys
from datetime import date

import sync_drive
import drive_utils as drv


def _pedir_config():
    print("Todavia no hay config_drive.json. Unidades disponibles:")
    claves = sorted(sync_drive.UNIDADES)
    for i, k in enumerate(claves, 1):
        print(f"  {i:2}. {sync_drive.UNIDADES[k]}")
    n = int(input("Numero de tu unidad: "))
    ruta = input("Ruta del archivo .json de credenciales: ").strip().strip('"')
    cfg = {"unidad": claves[n - 1], "credenciales": ruta}
    with open(sync_drive.CONFIG_RUTA, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    return sync_drive.cargar_config()


def main():
    if sync_drive.drv is None:
        print("❌ Faltan las librerias de Google. Instalalas con:\n"
              "   pip install google-api-python-client google-auth")
        return

    cfg = sync_drive.cargar_config() or _pedir_config()
    if not cfg:
        print("❌ Configuracion invalida (revisa la ruta del .json).")
        return
    print(f"Unidad: {cfg['unidad']}")

    try:
        with open(cfg["credenciales"], encoding="utf-8") as f:
            info = json.load(f)
        servicio = drv.obtener_servicio_drive(info)
        print(f"✅ 1. Autenticacion OK ({info.get('client_email')})")
    except Exception as e:
        print(f"❌ 1. No se pudo autenticar: {e}")
        return

    try:
        id_raiz = drv.id_carpeta_raiz(servicio)
        print(f"✅ 2. Carpeta '{drv.NOMBRE_CARPETA_RAIZ}' encontrada")
    except Exception as e:
        print(f"❌ 2. {e}")
        return

    hoy = date.today()
    try:
        id_mes, ids_tipo = drv.carpetas_de_mes(servicio, id_raiz, cfg["unidad"], hoy.year, hoy.month)
        print(f"✅ 3. Carpeta {cfg['unidad']} / {hoy.year} / {hoy.month} lista (con IESS y CAMPESINO)")
    except Exception as e:
        print(f"❌ 3. No se pudo ubicar/crear la carpeta del mes: {e}")
        return

    archivo = drv.buscar_archivo_por_patron(servicio, id_mes, r"^INSTRUCTIVO.*\.xlsx$")
    if archivo:
        print(f"✅    Matriz encontrada en Drive: {archivo['name']}")
    else:
        print("⚠️    No hay matriz INSTRUCTIVO en la carpeta de este mes todavia.")

    ruta_prueba = os.path.join(sync_drive.core.BASE_DIR, "_prueba_drive.txt")
    with open(ruta_prueba, "w", encoding="utf-8") as f:
        f.write("prueba de conexion - se puede borrar\n")
    try:
        res = drv.subir_o_reemplazar_archivo(
            servicio, ruta_prueba, "_prueba_conexion.txt", ids_tipo["IESS"], "text/plain"
        )
        print("✅ 4. Subida de un archivo nuevo: OK")
        try:
            servicio.files().delete(fileId=res["id"]).execute()
            print("✅    Archivo de prueba borrado.")
        except Exception:
            print("⚠️    No se pudo borrar el archivo de prueba (borralo a mano de la carpeta IESS).")
    except Exception as e:
        texto = str(e)
        print(f"❌ 4. No se pudo subir un archivo nuevo: {texto[:300]}")
        if "storageQuotaExceeded" in texto or "do not have storage quota" in texto:
            print("\n   → Google no deja a una cuenta de servicio CREAR archivos en una carpeta de\n"
                  "     'Mi unidad'. Hay que cambiar a autorizacion con la cuenta de Google duena de\n"
                  "     la carpeta. Pasale este mensaje a quien te ayuda con el programa.")
    finally:
        try:
            os.remove(ruta_prueba)
        except OSError:
            pass


if __name__ == "__main__":
    main()
    if sys.platform.startswith("win"):
        input("\nPresiona Enter para cerrar...")
