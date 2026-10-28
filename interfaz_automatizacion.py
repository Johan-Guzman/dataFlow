"""
INTERFAZ DE CONFIGURACIÓN - Automatización de Excel

Permite elegir los archivos, editar encabezados, catálogos, equivalencias
y periodos, guardar/cargar la configuración y ejecutar el proceso.

Uso:  python interfaz_automatizacion.py
(requiere automatizacion_excel.py en la misma carpeta)
"""

import json
import os
import subprocess
import sys
import threading
import tkinter as tk
from dataclasses import asdict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import automatizacion_excel as ae

CONFIG_AUTO = Path(__file__).with_name("config_interfaz.json")

COLUMNAS_CATALOGO = [
    ("nivel", "Nivel educativo"),
    ("documento", "Tipo de documento"),
    ("programa", "Programa de interés"),
]
ENCABEZADOS = [
    ("col_nivel", "Nivel educativo"),
    ("col_documento", "Tipo de documento"),
    ("col_programa", "Programa de interés"),
    ("col_texto", "Texto a normalizar (Title Case)"),
    ("col_telefono", "Teléfono"),
    ("col_codigo_pais", "Código de país"),
    ("col_periodo", "Periodo"),
]


from utilidades_config import lineas, parsear_pares, pares_a_texto  # noqa: E402


# ---------------------------- interfaz ----------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automatización de Excel – Configuración")
        self.geometry("860x720")
        self.minsize(760, 620)

        self.v = {}          # variables simples
        self.txt = {}        # cuadros de texto
        self.ultimo_resultado = None

        self._crear_widgets()
        self._cargar_inicial()

    # ---------- construcción ----------
    def _crear_widgets(self):
        barra = ttk.Frame(self, padding=(10, 8))
        barra.pack(fill="x")
        ttk.Button(barra, text="Cargar configuración…", command=self.cargar_config).pack(side="left")
        ttk.Button(barra, text="Guardar configuración…", command=self.guardar_config).pack(side="left", padx=6)
        ttk.Button(barra, text="Restablecer valores por defecto", command=self.restablecer).pack(side="left")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10)
        self._tab_archivos(nb)
        self._tab_columnas(nb)
        self._tab_catalogos(nb)
        self._tab_periodos(nb)

        pie = ttk.Frame(self, padding=10)
        pie.pack(fill="both")
        fila = ttk.Frame(pie)
        fila.pack(fill="x")
        self.btn_ejecutar = ttk.Button(fila, text="▶ Ejecutar automatización", command=self.ejecutar)
        self.btn_ejecutar.pack(side="left")
        self.btn_abrir = ttk.Button(fila, text="Abrir archivo generado", command=self.abrir_salida, state="disabled")
        self.btn_abrir.pack(side="left", padx=6)
        self.log = tk.Text(pie, height=9, state="disabled", wrap="word", bg="#f7f7f7")
        self.log.pack(fill="x", pady=(8, 0))

    def _entrada(self, padre, fila, etiqueta, clave, ancho=45, ayuda=None):
        self.v[clave] = tk.StringVar()
        ttk.Label(padre, text=etiqueta).grid(row=fila, column=0, sticky="w", pady=4)
        e = ttk.Entry(padre, textvariable=self.v[clave], width=ancho)
        e.grid(row=fila, column=1, sticky="we", padx=8)
        if ayuda:
            ttk.Label(padre, text=ayuda, foreground="#666").grid(row=fila, column=2, sticky="w")
        return e

    def _tab_archivos(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="Archivos y opciones")
        f.columnconfigure(1, weight=1)
        self._entrada(f, 0, "Archivo original", "entrada")
        ttk.Button(f, text="Examinar…", command=lambda: self._elegir("entrada", False)).grid(row=0, column=2)
        self._entrada(f, 1, "Archivo de salida", "salida")
        ttk.Button(f, text="Examinar…", command=lambda: self._elegir("salida", True)).grid(row=1, column=2)
        ttk.Label(
            f, foreground="#666", wraplength=700,
            text="El original nunca se sobrescribe. Si es .xlsm, la salida se guarda como .xlsm para conservar las macros.",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 14))

        self.v["leer_hoja"] = tk.BooleanVar()
        self.v["reg_validos"] = tk.BooleanVar()
        ttk.Checkbutton(
            f, variable=self.v["leer_hoja"],
            text=f"Leer catálogos desde la hoja «{ae.HOJA_REGLAS}» (reemplaza los de esta ventana si hay coincidencia de encabezados)",
        ).grid(row=3, column=0, columnspan=3, sticky="w", pady=3)
        ttk.Checkbutton(
            f, variable=self.v["reg_validos"],
            text="Registrar también cada celda VALIDA en el reporte (por defecto solo se cuentan en el resumen)",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=3)

    def _tab_columnas(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="Columnas")
        f.columnconfigure(1, weight=1)
        ttk.Label(
            f, foreground="#666", wraplength=700,
            text="Escribe el encabezado (fila 1) de cada columna. Se ignoran mayúsculas, tildes, espacios y guiones bajos.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        for i, (clave, etiqueta) in enumerate(ENCABEZADOS, start=1):
            self._entrada(f, i, etiqueta, clave)

    def _tab_catalogos(self, nb):
        f = ttk.Frame(nb, padding=8)
        nb.add(f, text="Catálogos y equivalencias")
        f.columnconfigure(0, weight=1)
        for i, (clave, etiqueta) in enumerate(COLUMNAS_CATALOGO):
            f.rowconfigure(i, weight=1)
            marco = ttk.LabelFrame(f, text=etiqueta, padding=6)
            marco.grid(row=i, column=0, sticky="nsew", pady=4)
            marco.columnconfigure((0, 1), weight=1)
            marco.rowconfigure(1, weight=1)
            ttk.Label(marco, text="Valores oficiales (uno por línea)").grid(row=0, column=0, sticky="w")
            ttk.Label(marco, text="Equivalencias confirmadas:  clave = valor oficial").grid(row=0, column=1, sticky="w")
            self.txt[f"cat_{clave}"] = self._texto(marco, 1, 0)
            self.txt[f"eq_{clave}"] = self._texto(marco, 1, 1)

    def _tab_periodos(self, nb):
        f = ttk.Frame(nb, padding=14)
        nb.add(f, text="Periodos")
        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)
        ttk.Label(
            f, foreground="#666", wraplength=700,
            text="Código de 2 dígitos = descripción. Ej.: 10 = Primer Semestre. Los códigos no listados no reciben descripción.",
        ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self.txt["periodos"] = self._texto(f, 1, 0, columnspan=3, height=6)
        self._entrada(f, 2, "Año mínimo válido", "anio_min", ancho=10)
        self._entrada(f, 3, "Año máximo válido", "anio_max", ancho=10)

    def _texto(self, padre, fila, col, columnspan=1, height=5):
        marco = ttk.Frame(padre)
        marco.grid(row=fila, column=col, columnspan=columnspan, sticky="nsew", padx=3)
        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(0, weight=1)
        t = tk.Text(marco, height=height, width=30, wrap="none", undo=True)
        sb = ttk.Scrollbar(marco, command=t.yview)
        t.configure(yscrollcommand=sb.set)
        t.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        return t

    # ---------- configuración <-> interfaz ----------
    def _set_texto(self, clave, contenido):
        self.txt[clave].delete("1.0", "end")
        self.txt[clave].insert("1.0", contenido)

    def aplicar_cfg(self, cfg):
        for clave, _ in ENCABEZADOS:
            self.v[clave].set(getattr(cfg, clave))
        for clave, _ in COLUMNAS_CATALOGO:
            nombre = getattr(cfg, f"col_{clave}")
            self._set_texto(f"cat_{clave}", "\n".join(cfg.valores_validos.get(nombre, [])))
            self._set_texto(f"eq_{clave}", pares_a_texto(cfg.equivalencias.get(nombre, {})))
        self._set_texto("periodos", pares_a_texto(cfg.periodos))
        self.v["anio_min"].set(str(cfg.anio_min))
        self.v["anio_max"].set(str(cfg.anio_max))
        self.v["leer_hoja"].set(cfg.leer_catalogos_desde_hoja)
        self.v["reg_validos"].set(cfg.registrar_validos)

    def obtener_cfg(self):
        """Construye una Configuracion validada desde la interfaz."""
        cfg = ae.Configuracion()
        nombres = []
        for clave, etiqueta in ENCABEZADOS:
            valor = self.v[clave].get().strip()
            if not valor and clave != "col_codigo_pais":
                raise ValueError(f"El encabezado «{etiqueta}» no puede estar vacío.")
            setattr(cfg, clave, valor)
            nombres.append(ae.normalizar_encabezado(valor))
        if len(set(nombres)) != len(nombres):
            raise ValueError("Hay dos columnas con el mismo encabezado en la pestaña Columnas.")

        cfg.valores_validos, cfg.equivalencias = {}, {}
        for clave, etiqueta in COLUMNAS_CATALOGO:
            nombre = getattr(cfg, f"col_{clave}")
            cfg.valores_validos[nombre] = lineas(self.txt[f"cat_{clave}"].get("1.0", "end"))
            cfg.equivalencias[nombre] = parsear_pares(
                self.txt[f"eq_{clave}"].get("1.0", "end"), f"Equivalencias de {etiqueta}"
            )

        cfg.periodos = parsear_pares(self.txt["periodos"].get("1.0", "end"), "Periodos")
        for cod in cfg.periodos:
            if not (cod.isdigit() and len(cod) == 2):
                raise ValueError(f"Periodos: el código «{cod}» debe tener exactamente 2 dígitos.")
        try:
            cfg.anio_min = int(self.v["anio_min"].get())
            cfg.anio_max = int(self.v["anio_max"].get())
        except ValueError:
            raise ValueError("Los años mínimo y máximo deben ser números enteros.")
        if cfg.anio_min > cfg.anio_max:
            raise ValueError("El año mínimo no puede ser mayor que el máximo.")

        cfg.leer_catalogos_desde_hoja = self.v["leer_hoja"].get()
        cfg.registrar_validos = self.v["reg_validos"].get()
        return cfg

    def _cargar_inicial(self):
        cfg = ae.Configuracion()
        if CONFIG_AUTO.exists():
            try:
                cfg = ae.cargar_configuracion_json(CONFIG_AUTO)
                self._log(f"Configuración anterior cargada de {CONFIG_AUTO.name}")
            except Exception as e:
                self._log(f"No se pudo leer {CONFIG_AUTO.name}: {e}")
        self.aplicar_cfg(cfg)
        self.v["entrada"].set(str(ae.ARCHIVO_ENTRADA))
        self.v["salida"].set(str(ae.ARCHIVO_SALIDA))

    # ---------- acciones ----------
    def _elegir(self, clave, guardar):
        tipos = [("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")]
        if guardar:
            ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=tipos)
        else:
            ruta = filedialog.askopenfilename(filetypes=tipos)
        if ruta:
            self.v[clave].set(ruta)
            if clave == "entrada" and not self.v["salida"].get().strip():
                self.v["salida"].set(str(Path(ruta).with_name("datos_corregidos.xlsx")))

    def _guardar_json(self, ruta, cfg):
        datos = asdict(cfg)
        Path(ruta).write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")

    def guardar_config(self):
        try:
            cfg = self.obtener_cfg()
        except ValueError as e:
            return messagebox.showerror("Configuración inválida", str(e))
        ruta = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if ruta:
            self._guardar_json(ruta, cfg)
            self._log(f"Configuración guardada en {ruta}")

    def cargar_config(self):
        ruta = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not ruta:
            return
        try:
            self.aplicar_cfg(ae.cargar_configuracion_json(ruta))
            self._log(f"Configuración cargada de {ruta}")
        except Exception as e:
            messagebox.showerror("No se pudo cargar", str(e))

    def restablecer(self):
        if messagebox.askyesno("Restablecer", "¿Volver a los valores por defecto del script?"):
            self.aplicar_cfg(ae.Configuracion())

    def ejecutar(self):
        try:
            cfg = self.obtener_cfg()
        except ValueError as e:
            return messagebox.showerror("Configuración inválida", str(e))

        entrada, salida = self.v["entrada"].get().strip(), self.v["salida"].get().strip()
        if not entrada or not Path(entrada).exists():
            return messagebox.showerror("Archivo", "Selecciona un archivo original que exista.")
        if not salida:
            return messagebox.showerror("Archivo", "Indica el archivo de salida.")
        if Path(salida).exists() and not messagebox.askyesno(
            "Archivo existente", f"{Path(salida).name} ya existe. ¿Reemplazarlo?"
        ):
            return

        self._guardar_json(CONFIG_AUTO, cfg)  # recuerda la última configuración usada
        self.btn_ejecutar.state(["disabled"])
        self._log("Procesando…")
        threading.Thread(target=self._trabajo, args=(entrada, salida, cfg), daemon=True).start()

    def _trabajo(self, entrada, salida, cfg):
        try:
            r = ae.procesar_excel(entrada, salida, cfg)
            self.after(0, self._terminado, r, None)
        except Exception as e:  # se muestra en la interfaz, no se cierra
            self.after(0, self._terminado, None, e)

    def _terminado(self, r, error):
        self.btn_ejecutar.state(["!disabled"])
        if error:
            self._log(f"ERROR: {error}")
            return messagebox.showerror("Error al procesar", str(error))
        self.ultimo_resultado = r
        lineas_log = [f"Listo. Archivo generado: {r['salida']}", f"Hojas procesadas: {r['hojas_procesadas']}"]
        lineas_log += [f"  {e}: {n}" for e, n in sorted(r["resumen"].items())]
        lineas_log += [f"AVISO: {a}" for a in r["avisos"]]
        lineas_log.append("Revisa las hojas Reporte_validacion y Resumen_validacion.")
        self._log("\n".join(lineas_log))
        self.btn_abrir.state(["!disabled"])

    def abrir_salida(self):
        if not self.ultimo_resultado:
            return
        ruta = str(self.ultimo_resultado["salida"])
        if sys.platform.startswith("win"):
            os.startfile(ruta)
        elif sys.platform == "darwin":
            subprocess.run(["open", ruta])
        else:
            subprocess.run(["xdg-open", ruta])

    def restablecer_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _log(self, texto):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.insert("end", texto)
        self.log.configure(state="disabled")


if __name__ == "__main__":
    App().mainloop()
