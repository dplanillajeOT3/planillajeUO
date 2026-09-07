# -*- coding: utf-8 -*-
"""
Ventana grafica para correr descargar_coberturas.py sin ver la consola negra.

COMO USARLO:
    python iniciar_gui.py

  (o hacer doble clic si tienes asociado .py a python en Windows)

Requiere los mismos paquetes que descargar_coberturas.py, mas nada extra:
tkinter viene incluido con Python en Windows.

Tiene dos formas de trabajar:
  - "Iniciar descarga": el modo de siempre, por lotes, leyendo
    reporte_inconsistencias.xlsx.
  - "Ingresar paciente manual": abre un formulario para escribir
    pacientes uno tras otro SIN esperar a que el anterior termine de
    procesarse: cada uno que agregas entra a una cola y se van
    resolviendo en segundo plano, uno a la vez (un solo Chrome), mientras
    el formulario queda libre de inmediato para seguir escribiendo.
"""

import io
import sys
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os

import descargar_coberturas as core


def _ocultar_consola():
    """En Windows, si el script se corrio con python.exe (no pythonw.exe)
    aparece una ventana de consola negra detras de la GUI. La ocultamos
    por comodidad; en cualquier otro sistema operativo esto no hace
    nada."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        pass


class ColaComoArchivo(io.TextIOBase):
    """Simula un archivo de escritura (como stdout) pero en vez de imprimir,
    guarda cada linea en una cola para que la interfaz grafica la muestre."""

    def __init__(self, cola):
        self.cola = cola

    def write(self, texto):
        if texto:
            self.cola.put(("log", texto))
        return len(texto)

    def flush(self):
        pass


class CampoFecha(ttk.Entry):
    """Entry que solo deja escribir numeros y va insertando los guiones
    de DD-MM-YYYY automaticamente a medida que se escribe (no hace falta
    escribir los guiones a mano).

    Ademas:
    - Al entrar al campo (con clic o con Tab) selecciona todo el texto,
      para que si ya tiene una fecha (por ejemplo la de hoy, precargada)
      escribir directamente la reemplace en vez de mezclarse con los
      numeros nuevos y desordenar la fecha.
    - Solo acepta digitos al escribir (letras y simbolos se rechazan,
      no se limitan a desaparecer despues).
    - Al borrar o corregir un numero en medio de la fecha, el cursor se
      recoloca en el lugar correcto en vez de saltar siempre al final.
    """

    def __init__(self, master, textvariable, **kw):
        self._var = textvariable
        vcmd = (master.register(self._validar_tecla), "%S", "%d")
        super().__init__(
            master, textvariable=textvariable,
            validate="key", validatecommand=vcmd, **kw
        )
        self._formateando = False
        self._var.trace_add("write", self._al_escribir)
        self.bind("<FocusIn>", self._al_enfocar)

    def _validar_tecla(self, texto_insertado, tipo_accion):
        # tipo_accion "1" = insercion de texto, "0" = borrado (siempre se permite borrar)
        if tipo_accion != "1":
            return True
        # solo se permiten digitos; el guion lo pone el propio campo
        return texto_insertado.isdigit()

    def _al_enfocar(self, _evento=None):
        # selecciona todo para que escribir de una vez reemplace la fecha
        # existente (por ejemplo la fecha de hoy que viene precargada)
        self.selection_range(0, tk.END)
        self.icursor(tk.END)

    def _al_escribir(self, *_args):
        if self._formateando:
            return
        texto = self._var.get()
        try:
            pos_cursor = self.index("insert")
        except tk.TclError:
            pos_cursor = len(texto)

        # cuantos digitos hay antes del cursor, para recolocarlo despues
        # de reformatear sin que salte siempre al final
        digitos_antes = sum(1 for ch in texto[:pos_cursor] if ch.isdigit())

        solo_digitos = "".join(ch for ch in texto if ch.isdigit())[:8]
        partes = []
        if solo_digitos[0:2]:
            partes.append(solo_digitos[0:2])
        if solo_digitos[2:4]:
            partes.append(solo_digitos[2:4])
        if solo_digitos[4:8]:
            partes.append(solo_digitos[4:8])
        nuevo = "-".join(partes)

        if nuevo == texto:
            return

        pos_nueva = len(nuevo)
        if digitos_antes <= 0:
            pos_nueva = 0
        else:
            digitos_vistos = 0
            for idx, ch in enumerate(nuevo):
                if ch.isdigit():
                    digitos_vistos += 1
                if digitos_vistos == digitos_antes:
                    pos_nueva = idx + 1
                    break

        self._formateando = True
        try:
            self._var.set(nuevo)
        finally:
            self._formateando = False

        # Tk repone su propio cursor al sincronizar el Entry con la variable
        # (y lo hace despues de que corre este mismo metodo), asi que hay que
        # recolocarlo una vez que esa sincronizacion ya termino, con after_idle,
        # o el cursor terminaria saltando a un lugar distinto al que queremos.
        self.after_idle(lambda: self._reponer_cursor(pos_nueva))

    def _reponer_cursor(self, pos_nueva):
        try:
            if self.winfo_exists():
                self.icursor(pos_nueva)
        except tk.TclError:
            pass


class DialogoCopiaExcel(tk.Toplevel):
    """Ventana emergente donde se elige QUE copiar (todo / solo hoy / un
    mes completo / un rango de fechas, segun la columna 'Fecha atencion'
    de la matriz) y en que carpeta guardar la copia, antes de generarla.
    Se abre tanto desde la ventana principal como desde la de ingreso
    manual."""

    def __init__(self, padre):
        super().__init__(padre)
        self.title("Generar copia en Excel")
        self.resizable(False, False)
        self.transient(padre)
        self.grab_set()

        cont = ttk.Frame(self, padding=14)
        cont.pack(fill="both", expand=True)

        ttk.Label(cont, text="¿Que quieres copiar?", font=("", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 6)
        )

        self.var_modo = tk.StringVar(value="todo")

        ttk.Radiobutton(
            cont, text="Todo", variable=self.var_modo, value="todo",
            command=self._actualizar_estado_campos
        ).grid(row=1, column=0, columnspan=2, sticky="w")

        ttk.Radiobutton(
            cont, text="Solo hoy", variable=self.var_modo, value="hoy",
            command=self._actualizar_estado_campos
        ).grid(row=2, column=0, columnspan=2, sticky="w")

        ttk.Radiobutton(
            cont, text="Un mes completo:", variable=self.var_modo, value="mes",
            command=self._actualizar_estado_campos
        ).grid(row=3, column=0, sticky="w")
        self.var_fecha_mes = tk.StringVar()
        self.campo_mes = CampoFecha(cont, textvariable=self.var_fecha_mes, width=12)
        self.campo_mes.grid(row=3, column=1, sticky="w", padx=(4, 0))
        ttk.Label(
            cont, text="(cualquier fecha de ese mes; en blanco = mes actual)", foreground="#777"
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 4))

        ttk.Radiobutton(
            cont, text="Rango de fechas:", variable=self.var_modo, value="rango",
            command=self._actualizar_estado_campos
        ).grid(row=5, column=0, sticky="w")
        marco_rango = ttk.Frame(cont)
        marco_rango.grid(row=5, column=1, sticky="w", padx=(4, 0))
        ttk.Label(marco_rango, text="Desde").pack(side="left")
        self.var_fecha_desde = tk.StringVar()
        self.campo_desde = CampoFecha(marco_rango, textvariable=self.var_fecha_desde, width=11)
        self.campo_desde.pack(side="left", padx=(4, 8))
        ttk.Label(marco_rango, text="Hasta").pack(side="left")
        self.var_fecha_hasta = tk.StringVar()
        self.campo_hasta = CampoFecha(marco_rango, textvariable=self.var_fecha_hasta, width=11)
        self.campo_hasta.pack(side="left", padx=(4, 0))

        ttk.Separator(cont, orient="horizontal").grid(row=6, column=0, columnspan=2, sticky="we", pady=10)

        ttk.Label(cont, text="Carpeta de destino:").grid(row=7, column=0, columnspan=2, sticky="w")
        marco_carpeta = ttk.Frame(cont)
        marco_carpeta.grid(row=8, column=0, columnspan=2, sticky="we", pady=(4, 0))
        marco_carpeta.columnconfigure(0, weight=1)
        self.var_carpeta = tk.StringVar(value=core._carpeta_descargas_windows())
        ttk.Entry(marco_carpeta, textvariable=self.var_carpeta).grid(row=0, column=0, sticky="we")
        ttk.Button(marco_carpeta, text="Examinar...", command=self._elegir_carpeta).grid(row=0, column=1, padx=(6, 0))

        marco_botones = ttk.Frame(cont)
        marco_botones.grid(row=9, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(marco_botones, text="Cancelar", command=self.destroy).pack(side="left", padx=(0, 6))
        ttk.Button(marco_botones, text="Generar copia", command=self._generar).pack(side="left")

        self._actualizar_estado_campos()
        self.bind("<Return>", lambda _e: self._generar())
        self.bind("<Escape>", lambda _e: self.destroy())

    def _actualizar_estado_campos(self):
        modo = self.var_modo.get()
        self.campo_mes.configure(state=("normal" if modo == "mes" else "disabled"))
        self.campo_desde.configure(state=("normal" if modo == "rango" else "disabled"))
        self.campo_hasta.configure(state=("normal" if modo == "rango" else "disabled"))

    def _elegir_carpeta(self):
        carpeta = filedialog.askdirectory(initialdir=self.var_carpeta.get() or os.path.expanduser("~"))
        if carpeta:
            self.var_carpeta.set(carpeta)

    def _generar(self):
        modo = self.var_modo.get()
        fecha_desde = fecha_hasta = None

        if modo == "mes":
            texto = self.var_fecha_mes.get().strip()
            if texto:
                fecha_desde = core._parsear_fecha_flexible(texto)
                if not fecha_desde:
                    messagebox.showerror(
                        "Fecha invalida", "Revisa la fecha del mes (formato DD-MM-YYYY).", parent=self
                    )
                    return
        elif modo == "rango":
            fecha_desde = core._parsear_fecha_flexible(self.var_fecha_desde.get().strip())
            fecha_hasta = core._parsear_fecha_flexible(self.var_fecha_hasta.get().strip())
            if not fecha_desde or not fecha_hasta:
                messagebox.showerror(
                    "Fechas invalidas", "Ingresa 'Desde' y 'Hasta' en formato DD-MM-YYYY.", parent=self
                )
                return
            if fecha_desde > fecha_hasta:
                messagebox.showerror(
                    "Rango invalido", "La fecha 'Desde' no puede ser posterior a 'Hasta'.", parent=self
                )
                return

        carpeta_destino = self.var_carpeta.get().strip() or core._carpeta_descargas_windows()

        try:
            destino = core.generar_copia_matriz(
                carpeta_destino=carpeta_destino,
                modo_fecha=modo,
                fecha_desde=fecha_desde,
                fecha_hasta=fecha_hasta,
            )
        except FileNotFoundError as e:
            messagebox.showwarning("Matriz no encontrada", str(e), parent=self)
            return
        except Exception as e:
            messagebox.showerror("Error al generar la copia", str(e), parent=self)
            return

        self.destroy()

        if messagebox.askyesno(
            "Copia generada",
            f"Se genero el archivo:\n{os.path.basename(destino)}\n\n¿Abrir la carpeta que lo contiene?"
        ):
            try:
                os.startfile(os.path.dirname(destino))  # Windows
            except AttributeError:
                messagebox.showinfo("Carpeta", os.path.dirname(destino))


def _generar_copia_excel_y_avisar(padre):
    """Abre el dialogo donde se elige que copiar (todo/hoy/mes/rango) y
    en que carpeta guardarlo. 'padre' es la ventana (Tk o Toplevel) sobre
    la que se debe centrar el dialogo."""
    DialogoCopiaExcel(padre)


SEXO_OPCIONES = ("Masculino", "Femenino")
SEXO_A_CODIGO = {"Masculino": "M", "Femenino": "F"}

ESTADOS_PENDIENTES = ("En cola", "Procesando...")
ESTADOS_OK = ("OK",)
# Estados que se pintan en rojo (algo que amerita revisar), pero ojo:
# no todos son reintentables (ver ESTADOS_REINTENTABLES mas abajo).
ESTADOS_ERROR = ("Error", "Sin matriz (revisar)", "Sin seguro (confirmado)")
# Unicamente estos se pueden volver a poner en la cola con "Reintentar":
# fallas reales (pagina caida, timeout guardando la matriz, etc). Un
# "Sin seguro (confirmado)" YA es un resultado valido -- no una falla --
# asi que nunca se debe reencolar, para no perder tiempo reintentando a
# alguien que de verdad no tiene seguro.
ESTADOS_REINTENTABLES = ("Error",)


class VentanaPrincipal:
    def __init__(self, root):
        self.root = root
        self.root.title("Descarga de Coberturas - CORESALUD")
        self.root.geometry("980x600")
        self.root.minsize(760, 460)

        self.cola = queue.Queue()
        self.filas_por_indice = {}
        self.corriendo = False
        self.ventana_manual = None

        self._armar_interfaz()
        self._cargar_pacientes_iniciales()
        self.root.after(100, self._procesar_cola)

    # ---------------------------------------------------------------
    def _armar_interfaz(self):
        estilo = ttk.Style()
        try:
            estilo.theme_use("vista")  # tema nativo de Windows si esta disponible
        except Exception:
            pass

        barra_superior = ttk.Frame(self.root, padding=10)
        barra_superior.pack(fill="x")

        self.boton_iniciar = ttk.Button(barra_superior, text="Iniciar descarga (por lotes)", command=self._iniciar)
        self.boton_iniciar.pack(side="left")

        self.boton_manual = ttk.Button(
            barra_superior, text="Ingresar paciente manual", command=self._abrir_ventana_manual
        )
        self.boton_manual.pack(side="left", padx=(8, 0))

        self.boton_abrir_carpeta = ttk.Button(
            barra_superior, text="Abrir carpeta de resultados", command=self._abrir_carpeta, state="disabled"
        )
        self.boton_abrir_carpeta.pack(side="left", padx=(8, 0))

        self.boton_generar_excel = ttk.Button(
            barra_superior, text="Generar copia en Excel", command=lambda: _generar_copia_excel_y_avisar(self.root)
        )
        self.boton_generar_excel.pack(side="left", padx=(8, 0))

        self.etiqueta_progreso = ttk.Label(barra_superior, text="Listo para iniciar")
        self.etiqueta_progreso.pack(side="right")

        self.barra_progreso = ttk.Progressbar(self.root, mode="determinate")
        self.barra_progreso.pack(fill="x", padx=10)

        # --- Tabla de pacientes ---
        marco_tabla = ttk.Frame(self.root, padding=(10, 10, 10, 5))
        marco_tabla.pack(fill="both", expand=True)

        columnas = ("num", "nombre", "cedula", "fechas", "estado")
        self.tabla = ttk.Treeview(marco_tabla, columns=columnas, show="headings", height=12)
        self.tabla.heading("num", text="#")
        self.tabla.heading("nombre", text="Paciente")
        self.tabla.heading("cedula", text="Cedula")
        self.tabla.heading("fechas", text="Fecha(s)")
        self.tabla.heading("estado", text="Estado")
        self.tabla.column("num", width=40, anchor="center")
        self.tabla.column("nombre", width=320)
        self.tabla.column("cedula", width=110, anchor="center")
        self.tabla.column("fechas", width=200)
        self.tabla.column("estado", width=140, anchor="center")

        self.tabla.tag_configure("ok", background="#e3f6e3")
        self.tabla.tag_configure("error", background="#fbe1e1")
        self.tabla.tag_configure("procesando", background="#fff6d9")

        scroll_tabla = ttk.Scrollbar(marco_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_tabla.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_tabla.pack(side="right", fill="y")

        # --- Log detallado ---
        marco_log = ttk.LabelFrame(self.root, text="Detalle", padding=6)
        marco_log.pack(fill="both", expand=False, padx=10, pady=(0, 10))

        self.texto_log = tk.Text(marco_log, height=10, state="disabled", wrap="word", bg="#111", fg="#ddd")
        scroll_log = ttk.Scrollbar(marco_log, orient="vertical", command=self.texto_log.yview)
        self.texto_log.configure(yscrollcommand=scroll_log.set)
        self.texto_log.pack(side="left", fill="both", expand=True)
        scroll_log.pack(side="right", fill="y")

    # ---------------------------------------------------------------
    def _cargar_pacientes_iniciales(self):
        try:
            registros = core.leer_excel(core.ARCHIVO_EXCEL)
        except Exception as e:
            messagebox.showerror("Error al leer el Excel", str(e))
            return

        if core.LIMITE_PRUEBA:
            registros = registros[: core.LIMITE_PRUEBA]

        self.total = len(registros)
        self.barra_progreso.configure(maximum=max(self.total, 1))

        for i, reg in enumerate(registros, start=1):
            fechas_txt = ", ".join(f.strftime("%d-%m-%Y") for f in reg["fechas"])
            item_id = self.tabla.insert(
                "", "end",
                values=(i, reg["nombre"], reg["cedula"], fechas_txt, "Pendiente")
            )
            self.filas_por_indice[i] = item_id

        self.etiqueta_progreso.configure(text=f"0 / {self.total} pacientes")

    # ---------------------------------------------------------------
    def _iniciar(self):
        if self.corriendo:
            return
        self.corriendo = True
        self.boton_iniciar.configure(state="disabled")
        self.boton_manual.configure(state="disabled")
        self._agregar_log("Iniciando... se abrira Chrome en segundo plano.\n")

        hilo = threading.Thread(target=self._correr_descarga, daemon=True)
        hilo.start()

    def _correr_descarga(self):
        salida_original = sys.stdout
        sys.stdout = ColaComoArchivo(self.cola)
        try:
            core.main(
                callback_fila=lambda i, estado: self.cola.put(("fila", i, estado)),
                callback_progreso=lambda hechos, total: self.cola.put(("progreso", hechos, total)),
            )
            self.cola.put(("fin_ok", None))
        except Exception as e:
            self.cola.put(("fin_error", str(e)))
        finally:
            sys.stdout = salida_original

    # ---------------------------------------------------------------
    def _procesar_cola(self):
        try:
            while True:
                item = self.cola.get_nowait()
                tipo = item[0]

                if tipo == "log":
                    self._agregar_log(item[1])

                elif tipo == "fila":
                    _, indice, estado = item
                    item_id = self.filas_por_indice.get(indice)
                    if item_id:
                        valores = list(self.tabla.item(item_id, "values"))
                        valores[4] = estado
                        etiqueta = "procesando"
                        if estado == "OK":
                            etiqueta = "ok"
                        elif estado == "Con errores":
                            etiqueta = "error"
                        self.tabla.item(item_id, values=valores, tags=(etiqueta,))
                        self.tabla.see(item_id)

                elif tipo == "progreso":
                    _, hechos, total = item
                    self.barra_progreso.configure(value=hechos)
                    self.etiqueta_progreso.configure(text=f"{hechos} / {total} pacientes")

                elif tipo == "fin_ok":
                    self.corriendo = False
                    self.boton_iniciar.configure(state="normal")
                    self.boton_manual.configure(state="normal")
                    self.boton_abrir_carpeta.configure(state="normal")
                    self.etiqueta_progreso.configure(text=f"Terminado - {self.total} / {self.total} pacientes")
                    messagebox.showinfo("Listo", "La descarga termino. Revisa la carpeta de resultados.")

                elif tipo == "fin_error":
                    self.corriendo = False
                    self.boton_iniciar.configure(state="normal")
                    self.boton_manual.configure(state="normal")
                    messagebox.showerror("Error inesperado", item[1])

        except queue.Empty:
            pass
        self.root.after(100, self._procesar_cola)

    # ---------------------------------------------------------------
    def _agregar_log(self, texto):
        self.texto_log.configure(state="normal")
        self.texto_log.insert("end", texto)
        self.texto_log.see("end")
        self.texto_log.configure(state="disabled")

    def _abrir_carpeta(self):
        ruta = os.path.abspath(core.CARPETA_SALIDA)
        try:
            os.startfile(ruta)  # Windows
        except AttributeError:
            messagebox.showinfo("Carpeta de resultados", ruta)

    # ---------------------------------------------------------------
    def _abrir_ventana_manual(self):
        if self.corriendo:
            messagebox.showwarning("Ocupado", "Espera a que termine la descarga por lotes.")
            return
        if self.ventana_manual is not None and self.ventana_manual.winfo_exists():
            self.ventana_manual.lift()
            self.ventana_manual.focus_force()
            return
        self.root.withdraw()  # se oculta esta pantalla mientras se usa el modo manual
        self.ventana_manual = VentanaManual(self)


class VentanaManual(tk.Toplevel):
    """Formulario para ingresar pacientes uno tras otro. Cada paciente que
    se agrega entra a una cola y se procesa en segundo plano (un solo
    Chrome, uno a la vez, en el orden en que se agregaron); el formulario
    queda libre de inmediato para seguir escribiendo el siguiente, sin
    esperar a que el anterior termine."""

    def __init__(self, padre):
        super().__init__(padre.root)
        self.padre = padre
        self.title("Ingresar pacientes manualmente")
        self.geometry("760x680")
        self.minsize(680, 560)
        self.protocol("WM_DELETE_WINDOW", self._al_cerrar)

        self.cola = queue.Queue()          # eventos del hilo de trabajo -> GUI
        self.cola_pendientes = queue.Queue()  # pacientes en espera -> hilo de trabajo

        self.driver_holder = None
        self.carpeta_descargas_temp = None
        self.carpeta_diagnostico = None
        self.errores = []
        self.sin_seguro = []
        self.filas_agregadas = 0
        self.contador_ingresados = 0
        self.filas_tabla = {}
        self.datos_por_num = {}  # num -> datos originales, para poder reintentar

        self.ruta_matriz = None
        self.dependencias_validas = []

        self._armar_interfaz()
        self._inicializar_matriz()
        self.after(100, self._procesar_cola)

        self.hilo_trabajo = threading.Thread(target=self._trabajador, daemon=True)
        self.hilo_trabajo.start()

    # ---------------------------------------------------------------
    def _armar_interfaz(self):
        cont = ttk.Frame(self, padding=12)
        cont.pack(fill="both", expand=True)
        cont.columnconfigure(1, weight=1)

        fila = 0
        ttk.Label(cont, text="Responsable (persona que ingresa la informacion):").grid(
            row=fila, column=0, columnspan=3, sticky="w", pady=(0, 2)
        )
        fila += 1
        self.var_responsable = tk.StringVar()
        ttk.Entry(cont, textvariable=self.var_responsable).grid(
            row=fila, column=0, columnspan=3, sticky="we", pady=(0, 10)
        )
        fila += 1

        ttk.Separator(cont, orient="horizontal").grid(row=fila, column=0, columnspan=3, sticky="we", pady=(0, 10))
        fila += 1

        ttk.Label(cont, text="Cedula del paciente:").grid(row=fila, column=0, sticky="w", pady=4)
        self.var_cedula = tk.StringVar()
        self.entry_cedula = ttk.Entry(cont, textvariable=self.var_cedula, width=22)
        self.entry_cedula.grid(row=fila, column=1, sticky="w", pady=4)
        fila += 1

        ttk.Label(cont, text="Fecha de atencion:").grid(row=fila, column=0, sticky="w", pady=4)
        self.var_fecha_atencion = tk.StringVar(value=core.date.today().strftime("%d-%m-%Y"))
        CampoFecha(cont, textvariable=self.var_fecha_atencion, width=22).grid(row=fila, column=1, sticky="w", pady=4)
        ttk.Label(cont, text="(solo escribe los numeros: dia, mes, año)", foreground="#777").grid(
            row=fila, column=2, sticky="w", padx=(6, 0)
        )
        fila += 1

        ttk.Label(cont, text="Dependencia (tipo de consulta):").grid(row=fila, column=0, sticky="w", pady=4)
        self.var_dependencia = tk.StringVar()
        self.combo_dependencia = ttk.Combobox(
            cont, textvariable=self.var_dependencia, state="readonly", width=28
        )
        self.combo_dependencia.grid(row=fila, column=1, columnspan=2, sticky="w", pady=4)
        fila += 1

        ttk.Label(cont, text="Fecha de nacimiento:").grid(row=fila, column=0, sticky="w", pady=4)
        self.var_fecha_nacimiento = tk.StringVar()
        CampoFecha(cont, textvariable=self.var_fecha_nacimiento, width=22).grid(row=fila, column=1, sticky="w", pady=4)
        ttk.Label(cont, text="(solo escribe los numeros: dia, mes, año)", foreground="#777").grid(
            row=fila, column=2, sticky="w", padx=(6, 0)
        )
        fila += 1

        ttk.Label(cont, text="Sexo:").grid(row=fila, column=0, sticky="w", pady=4)
        self.var_sexo = tk.StringVar()
        ttk.Combobox(
            cont, textvariable=self.var_sexo, state="readonly", width=14, values=SEXO_OPCIONES
        ).grid(row=fila, column=1, sticky="w", pady=4)
        fila += 1

        ttk.Label(cont, text="Observaciones (opcional):").grid(row=fila, column=0, sticky="w", pady=4)
        self.var_observaciones = tk.StringVar()
        ttk.Entry(cont, textvariable=self.var_observaciones).grid(row=fila, column=1, columnspan=2, sticky="we", pady=4)
        fila += 1

        self.boton_agregar = ttk.Button(
            cont, text="Agregar a la cola y seguir con el siguiente", command=self._agregar_a_cola
        )
        self.boton_agregar.grid(row=fila, column=0, columnspan=3, pady=(12, 6))
        fila += 1

        marco_estado = ttk.Frame(cont)
        marco_estado.grid(row=fila, column=0, columnspan=3, sticky="we")
        marco_estado.columnconfigure(0, weight=1)
        self.etiqueta_estado = ttk.Label(marco_estado, text="0 en cola, 0 completados", foreground="#555")
        self.etiqueta_estado.grid(row=0, column=0, sticky="w")
        ttk.Button(
            marco_estado, text="Reintentar fallidos", command=self._reintentar_fallidos
        ).grid(row=0, column=1, sticky="e", padx=(0, 6))
        ttk.Button(
            marco_estado, text="Generar copia en Excel", command=lambda: _generar_copia_excel_y_avisar(self)
        ).grid(row=0, column=2, sticky="e")
        fila += 1

        # --- Tabla de la cola ---
        marco_tabla = ttk.Frame(cont)
        marco_tabla.grid(row=fila, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        cont.rowconfigure(fila, weight=1)
        fila += 1

        columnas = ("num", "cedula", "fecha", "dependencia", "estado")
        self.tabla = ttk.Treeview(marco_tabla, columns=columnas, show="headings", height=6)
        self.tabla.heading("num", text="#")
        self.tabla.heading("cedula", text="Cedula")
        self.tabla.heading("fecha", text="Fecha atencion")
        self.tabla.heading("dependencia", text="Dependencia")
        self.tabla.heading("estado", text="Estado")
        self.tabla.column("num", width=35, anchor="center")
        self.tabla.column("cedula", width=100, anchor="center")
        self.tabla.column("fecha", width=110, anchor="center")
        self.tabla.column("dependencia", width=180)
        self.tabla.column("estado", width=170, anchor="center")
        self.tabla.tag_configure("ok", background="#e3f6e3")
        self.tabla.tag_configure("error", background="#fbe1e1")
        self.tabla.tag_configure("procesando", background="#fff6d9")
        scroll_tabla = ttk.Scrollbar(marco_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll_tabla.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll_tabla.pack(side="right", fill="y")
        self.tabla.bind("<Double-1>", self._al_doble_clic_fila)

        ttk.Label(
            cont, text="Doble clic en una fila con 'Error' para reintentarla.",
            foreground="#777"
        ).grid(row=fila, column=0, columnspan=3, sticky="w", pady=(2, 0))
        fila += 1
        marco_log = ttk.LabelFrame(cont, text="Detalle", padding=6)
        marco_log.grid(row=fila, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        cont.rowconfigure(fila, weight=1)

        self.texto_log = tk.Text(marco_log, height=8, state="disabled", wrap="word", bg="#111", fg="#ddd")
        scroll_log = ttk.Scrollbar(marco_log, orient="vertical", command=self.texto_log.yview)
        self.texto_log.configure(yscrollcommand=scroll_log.set)
        self.texto_log.pack(side="left", fill="both", expand=True)
        scroll_log.pack(side="right", fill="y")

    # ---------------------------------------------------------------
    def _inicializar_matriz(self):
        try:
            self.ruta_matriz = core.localizar_archivo_matriz()
        except FileNotFoundError as e:
            messagebox.showwarning("Matriz no encontrada", str(e))
            self._agregar_log(f"AVISO: {e}\nSe podran descargar los PDF, pero no se llenara la matriz.\n")
            return

        try:
            # Esta apertura es solo para leer la lista de dependencias
            # validas (MAESTRO) y confirmar que el archivo esta bien;
            # NO se guarda para escribir mas tarde. Si esta ventana se
            # quedara con esta copia abierta durante toda la sesion
            # mientras otra PC tambien agrega pacientes a la misma
            # matriz por red, la que guarde de ultimo terminaria
            # borrando lo que la otra ya escribio. Cada fila se agrega
            # releyendo el archivo tal como este en ese momento, con
            # candado (ver core.agregar_filas_matriz_con_bloqueo).
            wb_inicial = core.abrir_matriz(self.ruta_matriz)
            self.dependencias_validas = core.obtener_dependencias_validas(wb_inicial)
            self.combo_dependencia["values"] = self.dependencias_validas
            self._agregar_log(f"Matriz cargada: {os.path.basename(self.ruta_matriz)}\n")
        except Exception as e:
            messagebox.showerror("Error al abrir la matriz", str(e))

    # ---------------------------------------------------------------
    def _validar_y_obtener_datos(self):
        solo_digitos = "".join(ch for ch in self.var_cedula.get() if ch.isdigit())
        if not solo_digitos:
            messagebox.showerror("Falta la cedula", "Ingresa la cedula del paciente.")
            self.entry_cedula.focus_set()
            return None
        cedula = str(int(solo_digitos)).zfill(10)
        if len(cedula) != 10:
            messagebox.showerror(
                "Cedula invalida", "La cedula debe tener 10 digitos. Revisa lo que escribiste."
            )
            self.entry_cedula.focus_set()
            return None

        fecha_atencion = core._parsear_fecha_flexible(self.var_fecha_atencion.get())
        if not fecha_atencion:
            messagebox.showerror(
                "Fecha de atencion invalida",
                "La fecha de atencion no es una fecha valida (revisa el dia, mes y año)."
            )
            return None

        dependencia = self.var_dependencia.get().strip()
        if not dependencia:
            messagebox.showerror("Falta la dependencia", "Elige la dependencia (tipo de consulta) de la lista.")
            return None

        fecha_nacimiento = core._parsear_fecha_flexible(self.var_fecha_nacimiento.get())
        if not fecha_nacimiento:
            messagebox.showerror(
                "Fecha de nacimiento invalida",
                "La fecha de nacimiento no es una fecha valida (revisa el dia, mes y año)."
            )
            return None
        if fecha_nacimiento > core.date.today():
            messagebox.showerror(
                "Fecha de nacimiento invalida", "La fecha de nacimiento no puede ser una fecha futura."
            )
            return None

        sexo_texto = self.var_sexo.get().strip()
        if sexo_texto not in SEXO_A_CODIGO:
            messagebox.showerror("Falta el sexo", "Elige Masculino o Femenino.")
            return None

        responsable = self.var_responsable.get().strip().upper()
        if not responsable:
            messagebox.showerror("Falta el responsable", "Ingresa el nombre de la persona que ingresa la informacion.")
            return None

        observaciones = self.var_observaciones.get().strip()

        return {
            "cedula": cedula,
            "fecha_atencion": fecha_atencion,
            "dependencia": dependencia,
            "fecha_nacimiento": fecha_nacimiento,
            "sexo_codigo": SEXO_A_CODIGO[sexo_texto],
            "responsable": responsable,
            "observaciones": observaciones,
        }

    # ---------------------------------------------------------------
    def _agregar_a_cola(self):
        datos = self._validar_y_obtener_datos()
        if not datos:
            return  # ya se avisO cual dato esta mal; el formulario NO se limpia para que se corrija

        self.contador_ingresados += 1
        datos["_num"] = self.contador_ingresados

        item_id = self.tabla.insert(
            "", "end",
            values=(
                datos["_num"], datos["cedula"], datos["fecha_atencion"].strftime("%d-%m-%Y"),
                datos["dependencia"], "En cola"
            )
        )
        self.filas_tabla[datos["_num"]] = item_id
        self.datos_por_num[datos["_num"]] = datos
        self.tabla.see(item_id)

        self.cola_pendientes.put(datos)
        self._actualizar_contador()

        # Limpia solo los campos del paciente; responsable queda igual
        # (normalmente no cambia entre un paciente y otro) y puedes
        # escribir el siguiente de inmediato, sin esperar a que este
        # termine de procesarse.
        self.var_cedula.set("")
        self.var_fecha_nacimiento.set("")
        self.var_sexo.set("")
        self.var_observaciones.set("")
        self.var_fecha_atencion.set(core.date.today().strftime("%d-%m-%Y"))
        self.entry_cedula.focus_set()

    # ---------------------------------------------------------------
    def _reencolar(self, num):
        """Vuelve a poner en la cola de trabajo el paciente #num, usando
        los mismos datos con los que se agrego la primera vez. Como la
        descarga de PDFs revisa si el archivo ya existe antes de volver a
        pedirlo, esto no genera duplicados: solo reintenta lo que
        realmente fallo la vez anterior."""
        datos = self.datos_por_num.get(num)
        if not datos:
            return
        item_id = self.filas_tabla.get(num)
        if item_id and self.tabla.exists(item_id):
            valores = list(self.tabla.item(item_id, "values"))
            valores[4] = "En cola"
            self.tabla.item(item_id, values=valores, tags=())
        self.cola_pendientes.put(datos)
        self._actualizar_contador()

    def _al_doble_clic_fila(self, evento):
        item_id = self.tabla.identify_row(evento.y)
        if not item_id:
            return
        valores = self.tabla.item(item_id, "values")
        if not valores:
            return
        estado_actual = valores[4]
        if estado_actual not in ESTADOS_REINTENTABLES:
            return
        num = int(valores[0])
        self._reencolar(num)

    def _reintentar_fallidos(self):
        pendientes = [
            int(self.tabla.set(iid, "num"))
            for iid in self.tabla.get_children()
            if self.tabla.set(iid, "estado") in ESTADOS_REINTENTABLES
        ]
        if not pendientes:
            messagebox.showinfo("Reintentar fallidos", "No hay ninguna fila en Error para reintentar.")
            return
        for num in pendientes:
            self._reencolar(num)

    # ---------------------------------------------------------------
    # Hilo de trabajo: procesa la cola de pacientes uno a la vez, en el
    # orden en que se agregaron, reutilizando un solo Chrome.
    # ---------------------------------------------------------------
    def _trabajador(self):
        while True:
            datos = self.cola_pendientes.get()
            if datos is None:
                break

            self.cola.put(("estado", datos["_num"], "Procesando..."))

            salida_original = sys.stdout
            sys.stdout = ColaComoArchivo(self.cola)
            try:
                if self.driver_holder is None:
                    os.makedirs(core.CARPETA_SALIDA, exist_ok=True)
                    self.carpeta_descargas_temp = os.path.join(core.CARPETA_SALIDA, "_descargas_temp")
                    os.makedirs(self.carpeta_descargas_temp, exist_ok=True)
                    self.carpeta_diagnostico = os.path.join(core.CARPETA_SALIDA, "_diagnostico")
                    self.driver_holder = [core.crear_driver(self.carpeta_descargas_temp)]

                reg = {
                    "cedula": datos["cedula"],
                    "nombre": None,
                    "fechas": [datos["fecha_atencion"]],
                    "cedula_padre": None,
                    "cedula_titular": None,
                }

                errores_antes = len(self.errores)
                sin_seguro_antes = len(self.sin_seguro)

                filas_matriz = core.procesar_registro(
                    reg, datos["_num"], datos["_num"], self.driver_holder,
                    self.carpeta_descargas_temp, self.carpeta_diagnostico,
                    self.errores, self.sin_seguro, recolectar_matriz=True
                )

                # Se distingue el resultado DEL PACIENTE PRINCIPAL (no de
                # un familiar) para decidir el estado: un error suyo
                # propio no lleva el prefijo "(ROL)" que si llevan los
                # errores de padre/madre/jefe de familia; y su fila de
                # "sin seguro" siempre queda marcada con rol "PACIENTE".
                error_paciente = any(
                    not str(e[3]).startswith("(") for e in self.errores[errores_antes:]
                )
                sin_seguro_paciente = any(
                    e[3] == "PACIENTE" for e in self.sin_seguro[sin_seguro_antes:]
                )

                agregado = False
                fallo_guardado_matriz = False
                if self.ruta_matriz is not None and filas_matriz:
                    try:
                        filas_escritas = core.agregar_filas_matriz_con_bloqueo(
                            self.ruta_matriz, filas_matriz, datos["dependencia"], datos["fecha_nacimiento"],
                            datos["sexo_codigo"], datos["observaciones"], datos["responsable"]
                        )
                    except TimeoutError as e:
                        print(f"   ERROR guardando en la matriz: {e}")
                        fallo_guardado_matriz = True
                        self.errores.append([
                            datos["cedula"], datos["cedula"],
                            datos["fecha_atencion"].strftime("%d-%m-%Y"), f"(MATRIZ) {e}"
                        ])
                    else:
                        self.filas_agregadas += len(filas_escritas)
                        for fila_escrita in filas_escritas:
                            print(f"   -> Fila {fila_escrita} agregada a la matriz ({core.HOJA_MATRIZ}).")
                        agregado = True

                if agregado:
                    estado = "OK"
                elif fallo_guardado_matriz:
                    estado = "Error"
                elif sin_seguro_paciente:
                    # Ya se confirmo que NO tiene seguro (por coberturasalud
                    # o por el respaldo del IESS): esto NO es una falla, asi
                    # que no debe quedar disponible para "Reintentar".
                    estado = "Sin seguro (confirmado)"
                elif error_paciente:
                    estado = "Error"
                elif filas_matriz:
                    estado = "Sin matriz (revisar)"
                else:
                    # No se pudo determinar con claridad que paso (por
                    # ejemplo, algo fallo solo con un familiar). Se deja
                    # como Error para poder revisarlo/reintentarlo, nunca
                    # se descarta en silencio.
                    estado = "Error"

                self.cola.put(("estado", datos["_num"], estado))

            except Exception as e:
                self.cola.put(("log", f"ERROR procesando paciente #{datos['_num']}: {e}\n"))
                self.cola.put(("estado", datos["_num"], "Error"))
            finally:
                sys.stdout = salida_original

            self.cola.put(("resumen", None))

        # Se recibio la señal de cierre (self.cola_pendientes.put(None)):
        # se cierra Chrome y se guardan los reportes finales, aunque la
        # ventana ya se haya cerrado.
        if self.driver_holder is not None:
            try:
                self.driver_holder[0].quit()
            except Exception:
                pass
        try:
            core._guardar_reportes_finales(self.errores, self.sin_seguro)
        except Exception:
            pass

    # ---------------------------------------------------------------
    def _procesar_cola(self):
        try:
            while True:
                item = self.cola.get_nowait()
                tipo = item[0]

                if tipo == "log":
                    self._agregar_log(item[1])

                elif tipo == "estado":
                    _, num, estado = item
                    item_id = self.filas_tabla.get(num)
                    if item_id and self.tabla.exists(item_id):
                        valores = list(self.tabla.item(item_id, "values"))
                        valores[4] = estado
                        etiqueta = "procesando"
                        if estado in ESTADOS_OK:
                            etiqueta = "ok"
                        elif estado in ESTADOS_ERROR:
                            etiqueta = "error"
                        self.tabla.item(item_id, values=valores, tags=(etiqueta,))
                        self.tabla.see(item_id)

                elif tipo == "resumen":
                    self._actualizar_contador()

        except queue.Empty:
            pass

        if self.winfo_exists():
            self.after(100, self._procesar_cola)

    def _actualizar_contador(self):
        en_cola = procesando = completados = 0
        for iid in self.tabla.get_children():
            estado = self.tabla.set(iid, "estado")
            if estado == "En cola":
                en_cola += 1
            elif estado == "Procesando...":
                procesando += 1
            else:
                completados += 1
        self.etiqueta_estado.configure(
            text=f"{en_cola} en cola, {procesando} procesando ahora, "
                 f"{completados} completado(s) - {self.filas_agregadas} fila(s) en la matriz."
        )

    # ---------------------------------------------------------------
    def _agregar_log(self, texto):
        self.texto_log.configure(state="normal")
        self.texto_log.insert("end", texto)
        self.texto_log.see("end")
        self.texto_log.configure(state="disabled")

    def _al_cerrar(self):
        pendientes = sum(
            1 for iid in self.tabla.get_children()
            if self.tabla.set(iid, "estado") in ESTADOS_PENDIENTES
        )
        if pendientes:
            messagebox.showwarning(
                "Todavia hay pacientes pendientes",
                f"Quedan {pendientes} paciente(s) en cola o procesandose. "
                "Espera a que la tabla los marque como terminados antes de cerrar esta ventana."
            )
            return

        self.cola_pendientes.put(None)  # el hilo de trabajo se cierra solo (Chrome, reportes)
        self.padre.ventana_manual = None
        self.padre.root.deiconify()  # vuelve a mostrar la pantalla principal
        self.destroy()


if __name__ == "__main__":
    _ocultar_consola()
    root = tk.Tk()
    app = VentanaPrincipal(root)
    root.mainloop()
