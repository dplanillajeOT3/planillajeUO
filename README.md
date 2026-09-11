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

## Login por unidad (21 unidades operativas)

La app ahora pide usuario y contraseña antes de mostrar nada. Cada
usuario corresponde a una unidad operativa y queda con sus propios
archivos (matriz, PDFs, `datos_pacientes.json`) completamente
separados de las demás — nunca se mezclan.

**Las contraseñas NUNCA van en `app_web.py` ni en GitHub.** Se
configuran en Streamlit Community Cloud: entra a tu app en
[share.streamlit.io](https://share.streamlit.io), luego
**⋮ (menú) → Settings → Secrets**, y pega algo como esto (una línea
por cada una de las 21 unidades):

```toml
[usuarios]
pumamaqui = { password = "clave-segura-1", nombre = "Casa de Acogida Pumamaqui" }
san_juan  = { password = "clave-segura-2", nombre = "UO San Juan" }
calderon  = { password = "clave-segura-3", nombre = "UO Calderón" }
# ...continúa con las 21
```

- La palabra antes del `=` (ej. `pumamaqui`) es lo que la unidad
  escribe como **usuario** al entrar. Usa algo simple, sin espacios ni
  tildes (minúsculas, guion bajo si hace falta).
- `password` es la contraseña que tú defines para esa unidad.
- `nombre` es solo el texto bonito que se muestra arriba una vez
  adentro (puede tener espacios y tildes).

Después de guardar los Secrets, la app se reinicia sola y ya puedes
repartir usuario/contraseña a cada unidad.

**La primera vez que una unidad entra**, en la pestaña de modo manual
le va a pedir subir su propio `INSTRUCTIVO...xlsx` (una sola vez);
de ahí en adelante esa unidad ya tiene su matriz mensual propia,
separada de las otras 20.

### ⚠️ Pendiente: esto todavía NO resuelve la pérdida de datos en redeploys

~~Esta separación por unidad ya es un avance...~~ **Actualización: ya está resuelto.** La app ahora se conecta a Google Drive (carpeta raíz `1OQ9dkzDOtbQT90mljvoPvzToJcPOSeH6`, en la cuenta `planillajeot3@gmail.com`) para respaldar y restaurar los archivos de cada unidad. Falta un solo paso de configuración de tu parte:

## Conectar con Google Drive (persistencia real)

1. Sigue los pasos para crear un **proyecto de Google Cloud + cuenta de servicio + llave JSON**, y comparte la carpeta de Drive con esa cuenta de servicio (esto ya se explicó en el chat; resumen: Google Cloud Console → nuevo proyecto → habilitar "Google Drive API" → Credenciales → Crear cuenta de servicio → pestaña Claves → Crear clave JSON → descarga el archivo → comparte tu carpeta de Drive con el correo `...@...iam.gserviceaccount.com` de esa cuenta, como Editor).
2. Abre el archivo `.json` que se descargó con un editor de texto (Bloc de notas sirve). Copia su contenido completo.
3. En Streamlit Cloud, entra a tu app → **⋮ → Settings → Secrets**, y agrega (junto a la sección `[usuarios]` que ya tenías) algo como esto, **pegando tus propios valores** del JSON (los nombres de la izquierda deben quedar EXACTAMENTE así; los valores de la derecha son los que cambian):

```toml
[gcp_service_account]
type = "service_account"
project_id = "tu-project-id"
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "app-planillaje@tu-project-id.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

   Todos esos campos vienen tal cual en el `.json` descargado (solo cambia el formato de `{ }` a líneas `campo = "valor"`). El campo `private_key` es el más delicado: debe quedar en una sola línea, entre comillas, con los `\n` literales tal como aparecen en el JSON — cópialo y pégalo tal cual, sin editarlo a mano.

4. (Opcional) Si en algún momento quieres usar OTRA carpeta raíz de Drive en vez de la que ya viene por defecto, agrega también:

```toml
[drive]
carpeta_raiz_id = "el-id-de-tu-carpeta"
```

5. Guarda los Secrets. La app se reinicia sola. Al volver a entrar con cualquier usuario, arriba debe aparecer "☁️ Conectado a Google Drive" en vez de la advertencia amarilla.

**Cómo queda funcionando:** al iniciar sesión, la app crea (si no existe) una subcarpeta con el nombre exacto del usuario dentro de tu carpeta de Drive, y descarga ahí lo que ya hubiera de esa unidad. Mientras trabajas, hay un botón **"☁️ Guardar en la nube ahora"** en ambas pestañas, y además se guarda solo: al terminar una tanda del modo automático, al generar reportes en el modo manual, y al cerrar sesión. Así, aunque Streamlit Cloud borre el disco temporal en el próximo redeploy, la próxima vez que esa unidad entre, todo se restaura desde Drive.


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
