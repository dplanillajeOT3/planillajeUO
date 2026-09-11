# App web de Coberturas de Salud

## Qué es esto

`app_web.py` es una interfaz web (Streamlit) para el mismo motor de
`descargar_coberturas.py`. No reescribe la lógica de Selenium, la
detección GENERAL/CAMPESINO ni el manejo de la matriz: **llama
directamente** a las funciones que ya existen en ese archivo. Por eso
`descargar_coberturas.py` debe estar en la **misma carpeta** que
`app_web.py` — la app web no funciona sin él.

## Archivos

- `app_web.py` — la interfaz web (nueva).
- `descargar_coberturas.py` — el motor (tu archivo original, sin tocar).
- `requirements.txt` — dependencias de Python.
- `packages.txt` — solo se usa si despliegas en un servidor Linux tipo
  Streamlit Community Cloud o Render (instala Chromium vía `apt`). En
  Windows no hace nada, se puede ignorar.

## Cómo correrla

```bash
pip install -r requirements.txt
streamlit run app_web.py
```

Se abre en el navegador en `http://localhost:8501`. Funciona igual en
Windows o Linux: `crear_driver()` (dentro de `descargar_coberturas.py`)
ya detecta solo si hay Chromium/Chromedriver del sistema (Linux) o si
debe usar `webdriver-manager` (Windows), sin que tengas que cambiar
nada.

Para que el modo manual pueda llenar la matriz automáticamente, el
archivo `INSTRUCTIVO..._2026.xlsx` debe estar en la misma carpeta que
`descargar_coberturas.py`, igual que en la versión de consola.

## Los dos modos

**Modo automático (pestaña "📂 Modo automático")**
Sube el Excel de inconsistencias, opcionalmente marca "Registrar en la
matriz" e indica el responsable, y presiona "Iniciar descarga". Corre
en segundo plano (un hilo aparte) y la pantalla se refresca sola
mostrando el progreso y el log, igual que la consola. Al terminar
puedes descargar `log_errores.xlsx` y `sin_seguro.xlsx`.

**Modo manual (pestaña "🧍 Modo manual")**
Reemplaza los `input()` de `modo_interactivo()` por un formulario:
cédula, fecha de atención, dependencia (tomada de la hoja MAESTRO si
hay matriz disponible), fecha de nacimiento, sexo y observaciones.
Mantiene un solo navegador abierto durante toda la sesión (igual que
la consola), así que la segunda consulta en adelante es más rápida.
Puedes generar reportes o cerrar el navegador en cualquier momento con
los botones de abajo, y generar una copia de la matriz (todo / hoy /
un mes / un rango) igual que al final del modo interactivo de consola.

## Notas importantes

- **Un navegador Chrome por sesión de usuario.** Si varias personas
  usan la app web al mismo tiempo en modo manual, cada una abre su
  propio Chrome en el servidor. Para un equipo pequeño (una oficina)
  está bien; si va a ser un servidor compartido por muchas personas a
  la vez, conviene limitar cuántas pestañas/sesiones simultáneas se
  permiten.
- **El progreso del modo automático se pierde si recargas la página**
  del navegador a la fuerza (F5) mientras está corriendo — el proceso
  en el servidor sigue, pero Streamlit pierde la referencia a esa
  sesión. Evita recargar manualmente mientras la barra de progreso
  esté activa; la app ya se refresca sola.
- Los PDF descargados, los backups de la matriz y los reportes se
  guardan en el servidor donde corre `streamlit run`, dentro de
  `PDF_DESCARGADOS/`, exactamente igual que con la versión de consola.
