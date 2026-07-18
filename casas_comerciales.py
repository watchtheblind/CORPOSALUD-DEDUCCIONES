import pdfplumber
import csv
import re
import tkinter as tk
from tkinter import ttk, messagebox
import os
import platform
import sys
import io
import json

# Configuración para evitar errores de caracteres en la terminal de Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ---------------------------------------------------------------------------
# SELECTOR DE ARCHIVOS PERSONALIZADO (reemplaza filedialog de tkinter)
# ---------------------------------------------------------------------------

class SelectorArchivos:
    """Selector de múltiples PDFs con panel de directorios y lista de archivos."""

    def __init__(self, titulo="Seleccionar archivos PDF", extension=".pdf"):
        self.titulo = titulo
        self.extension = extension.lower()
        self.resultado = []
        self._ruta_actual = os.path.expanduser("~")

    def abrir(self):
        raiz = tk.Tk()
        raiz.title(self.titulo)
        raiz.geometry("900x580")
        raiz.minsize(700, 420)
        raiz.configure(bg="#1e1e2e")
        raiz.resizable(True, True)

        # ── Estilos ──────────────────────────────────────────────────────────
        style = ttk.Style(raiz)
        style.theme_use("clam")
        style.configure("TFrame",       background="#1e1e2e")
        style.configure("TLabel",       background="#1e1e2e", foreground="#cdd6f4",
                         font=("Segoe UI", 10))
        style.configure("TButton",      background="#89b4fa", foreground="#1e1e2e",
                         font=("Segoe UI", 10, "bold"), relief="flat", padding=6)
        style.map("TButton",            background=[("active", "#74c7ec")])
        style.configure("Accent.TButton", background="#a6e3a1", foreground="#1e1e2e",
                         font=("Segoe UI", 11, "bold"), relief="flat", padding=8)
        style.map("Accent.TButton",     background=[("active", "#94e2d5")])
        style.configure("Treeview",     background="#313244", fieldbackground="#313244",
                         foreground="#cdd6f4", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background="#45475a", foreground="#89b4fa",
                         font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Treeview",           background=[("selected", "#585b70")])
        style.configure("TEntry",       fieldbackground="#313244", foreground="#cdd6f4",
                         insertcolor="#cdd6f4", font=("Segoe UI", 10), relief="flat")
        style.configure("TSeparator",   background="#45475a")

        # ── Layout principal ─────────────────────────────────────────────────
        marco_top = ttk.Frame(raiz, padding=(10, 8, 10, 4))
        marco_top.pack(fill="x")

        ttk.Label(marco_top, text="📁  Ruta actual:").pack(side="left")
        self._var_ruta = tk.StringVar(value=self._ruta_actual)
        entrada_ruta = ttk.Entry(marco_top, textvariable=self._var_ruta, width=60)
        entrada_ruta.pack(side="left", padx=(6, 4), fill="x", expand=True)
        ttk.Button(marco_top, text="Ir",
                   command=lambda: self._navegar(self._var_ruta.get(), tree_dirs, lista_arch, var_busq)).pack(side="left")

        # Panel separado en dos columnas
        panel = ttk.Frame(raiz, padding=(10, 0, 10, 0))
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(0, weight=1)
        panel.columnconfigure(1, weight=3)
        panel.rowconfigure(0, weight=1)

        # ── Árbol de directorios (izquierda) ─────────────────────────────────
        marco_dirs = ttk.Frame(panel)
        marco_dirs.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        ttk.Label(marco_dirs, text="Directorios",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(4, 2))
        sb_dirs = ttk.Scrollbar(marco_dirs)
        sb_dirs.pack(side="right", fill="y")
        tree_dirs = ttk.Treeview(marco_dirs, yscrollcommand=sb_dirs.set,
                                  show="tree", selectmode="browse")
        tree_dirs.pack(fill="both", expand=True)
        sb_dirs.config(command=tree_dirs.yview)

        # ── Lista de archivos (derecha) ───────────────────────────────────────
        marco_arch = ttk.Frame(panel)
        marco_arch.grid(row=0, column=1, sticky="nsew")

        marco_busq = ttk.Frame(marco_arch)
        marco_busq.pack(fill="x", pady=(4, 2))
        ttk.Label(marco_busq, text="🔍 Filtrar:").pack(side="left")
        var_busq = tk.StringVar()
        entrada_busq = ttk.Entry(marco_busq, textvariable=var_busq, width=30)
        entrada_busq.pack(side="left", padx=(4, 0))
        var_busq.trace_add("write",
            lambda *_: self._actualizar_archivos(lista_arch, var_busq.get()))

        sb_arch = ttk.Scrollbar(marco_arch)
        sb_arch.pack(side="right", fill="y")
        lista_arch = ttk.Treeview(marco_arch,
                                   columns=("nombre", "tamaño"),
                                   show="headings",
                                   yscrollcommand=sb_arch.set,
                                   selectmode="extended")
        lista_arch.heading("nombre", text="Archivo")
        lista_arch.heading("tamaño", text="Tamaño")
        lista_arch.column("nombre", width=380)
        lista_arch.column("tamaño", width=90, anchor="e")
        lista_arch.pack(fill="both", expand=True)
        sb_arch.config(command=lista_arch.yview)

        # ── Panel de seleccionados ────────────────────────────────────────────
        marco_sel = ttk.Frame(raiz, padding=(10, 4, 10, 4))
        marco_sel.pack(fill="x")
        ttk.Label(marco_sel, text="Seleccionados:",
                  font=("Segoe UI", 9, "bold")).pack(side="left")
        self._var_conteo = tk.StringVar(value="0 archivo(s)")
        ttk.Label(marco_sel, textvariable=self._var_conteo,
                  foreground="#a6e3a1").pack(side="left", padx=6)
        ttk.Button(marco_sel, text="✖ Limpiar selección",
                   command=lambda: [lista_arch.selection_remove(*lista_arch.selection()),
                                    self._var_conteo.set("0 archivo(s)")]).pack(side="right")

        # ── Botones inferiores ────────────────────────────────────────────────
        marco_bot = ttk.Frame(raiz, padding=(10, 6, 10, 10))
        marco_bot.pack(fill="x")
        ttk.Button(marco_bot, text="⬆  Subir directorio",
                   command=lambda: self._subir(tree_dirs, lista_arch, var_busq)
                   ).pack(side="left", padx=(0, 6))
        ttk.Button(marco_bot, text="Cancelar",
                   command=raiz.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(marco_bot, text="✔  Agregar seleccionados",
                   style="Accent.TButton",
                   command=lambda: self._confirmar(lista_arch, raiz)
                   ).pack(side="right")

        # ── Eventos ───────────────────────────────────────────────────────────
        tree_dirs.bind("<<TreeviewSelect>>",
            lambda e: self._seleccionar_dir(tree_dirs, lista_arch, var_busq))
        lista_arch.bind("<<TreeviewSelect>>",
            lambda e: self._var_conteo.set(
                f"{len(lista_arch.selection())} archivo(s) seleccionado(s)"))
        lista_arch.bind("<Double-1>",
            lambda e: self._confirmar(lista_arch, raiz))
        raiz.bind("<Return>",
            lambda e: self._confirmar(lista_arch, raiz))
        raiz.bind("<Escape>", lambda e: raiz.destroy())

        # ── Carga inicial ─────────────────────────────────────────────────────
        self._poblar_dirs(tree_dirs, lista_arch, var_busq)
        raiz.mainloop()
        return self.resultado

    # ── Métodos internos ──────────────────────────────────────────────────────

    def _poblar_dirs(self, tree, lista, var_busq):
        """Reconstruye el árbol de directorios desde la ruta actual."""
        tree.delete(*tree.get_children())
        ruta = self._ruta_actual
        # Niveles hacia arriba
        partes = []
        r = ruta
        while True:
            padre, hijo = os.path.split(r)
            if hijo:
                partes.append((r, hijo))
                r = padre
            else:
                partes.append((r, r))
                break
        partes.reverse()
        nodo_padre = ""
        ultimo = ""
        for full, nombre in partes:
            nodo = tree.insert(nodo_padre, "end", iid=full,
                               text=f"📂 {nombre}", open=True, values=[full])
            nodo_padre = nodo
            ultimo = full
        # Subdirectorios del nivel actual
        try:
            for entry in sorted(os.scandir(ruta), key=lambda e: e.name.lower()):
                if entry.is_dir() and not entry.name.startswith("."):
                    tree.insert(ultimo, "end", iid=entry.path,
                                text=f"📁 {entry.name}", values=[entry.path])
        except PermissionError:
            pass
        self._actualizar_archivos(lista, var_busq.get() if var_busq else "")

    def _actualizar_archivos(self, lista, filtro=""):
        """Rellena la lista de archivos del directorio actual con filtro opcional."""
        lista.delete(*lista.get_children())
        try:
            entries = sorted(os.scandir(self._ruta_actual), key=lambda e: e.name.lower())
            for entry in entries:
                if not entry.is_file(): continue
                if not entry.name.lower().endswith(self.extension): continue
                if filtro and filtro.lower() not in entry.name.lower(): continue
                tam = entry.stat().st_size
                tam_str = (f"{tam/1024:.1f} KB" if tam < 1_048_576
                           else f"{tam/1_048_576:.1f} MB")
                lista.insert("", "end", iid=entry.path,
                             values=(entry.name, tam_str))
        except PermissionError:
            pass

    def _navegar(self, ruta, tree, lista, var_busq):
        if os.path.isdir(ruta):
            self._ruta_actual = os.path.abspath(ruta)
            self._var_ruta.set(self._ruta_actual)
            self._poblar_dirs(tree, lista, var_busq)

    def _seleccionar_dir(self, tree, lista, var_busq):
        sel = tree.selection()
        if sel:
            ruta = sel[0]
            if os.path.isdir(ruta):
                self._ruta_actual = ruta
                self._var_ruta.set(ruta)
                self._poblar_dirs(tree, lista, var_busq)

    def _subir(self, tree, lista, var_busq):
        padre = os.path.dirname(self._ruta_actual)
        if padre != self._ruta_actual:
            self._ruta_actual = padre
            self._var_ruta.set(padre)
            self._poblar_dirs(tree, lista, var_busq)

    def _confirmar(self, lista, raiz):
        sel = lista.selection()
        if not sel:
            messagebox.showwarning("Atención", "Selecciona al menos un archivo.")
            return
        self.resultado = list(sel)   # iid == ruta completa
        raiz.destroy()


class SelectorCarpeta:
    """Selector de carpeta de destino con árbol de directorios."""

    def __init__(self, titulo="Seleccionar carpeta de salida"):
        self.titulo = titulo
        self.resultado = ""
        self._ruta_actual = os.path.expanduser("~")

    def abrir(self):
        raiz = tk.Tk()
        raiz.title(self.titulo)
        raiz.geometry("620x460")
        raiz.minsize(480, 340)
        raiz.configure(bg="#1e1e2e")

        style = ttk.Style(raiz)
        style.theme_use("clam")
        style.configure("TFrame",       background="#1e1e2e")
        style.configure("TLabel",       background="#1e1e2e", foreground="#cdd6f4",
                         font=("Segoe UI", 10))
        style.configure("TButton",      background="#89b4fa", foreground="#1e1e2e",
                         font=("Segoe UI", 10, "bold"), relief="flat", padding=6)
        style.map("TButton",            background=[("active", "#74c7ec")])
        style.configure("Accent.TButton", background="#a6e3a1", foreground="#1e1e2e",
                         font=("Segoe UI", 11, "bold"), relief="flat", padding=8)
        style.map("Accent.TButton",     background=[("active", "#94e2d5")])
        style.configure("Treeview",     background="#313244", fieldbackground="#313244",
                         foreground="#cdd6f4", rowheight=24, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background="#45475a", foreground="#89b4fa",
                         font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Treeview",           background=[("selected", "#585b70")])
        style.configure("TEntry",       fieldbackground="#313244", foreground="#cdd6f4",
                         insertcolor="#cdd6f4", font=("Segoe UI", 10), relief="flat")

        marco_top = ttk.Frame(raiz, padding=(10, 8, 10, 4))
        marco_top.pack(fill="x")
        ttk.Label(marco_top, text="📁  Ruta:").pack(side="left")
        self._var_ruta = tk.StringVar(value=self._ruta_actual)
        entrada = ttk.Entry(marco_top, textvariable=self._var_ruta, width=55)
        entrada.pack(side="left", padx=(6, 4), fill="x", expand=True)
        ttk.Button(marco_top, text="Ir",
                   command=lambda: self._navegar(self._var_ruta.get(), tree)
                   ).pack(side="left")

        marco_tree = ttk.Frame(raiz, padding=(10, 0, 10, 0))
        marco_tree.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(marco_tree)
        sb.pack(side="right", fill="y")
        tree = ttk.Treeview(marco_tree, yscrollcommand=sb.set,
                             show="tree", selectmode="browse")
        tree.pack(fill="both", expand=True)
        sb.config(command=tree.yview)

        self._var_sel = tk.StringVar(value="")
        marco_sel = ttk.Frame(raiz, padding=(10, 4))
        marco_sel.pack(fill="x")
        ttk.Label(marco_sel, text="Carpeta seleccionada:").pack(side="left")
        ttk.Label(marco_sel, textvariable=self._var_sel,
                  foreground="#a6e3a1").pack(side="left", padx=6)

        marco_bot = ttk.Frame(raiz, padding=(10, 6, 10, 10))
        marco_bot.pack(fill="x")
        ttk.Button(marco_bot, text="⬆  Subir",
                   command=lambda: self._subir(tree)).pack(side="left", padx=(0, 6))
        ttk.Button(marco_bot, text="Cancelar",
                   command=raiz.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(marco_bot, text="✔  Usar esta carpeta",
                   style="Accent.TButton",
                   command=lambda: self._confirmar(tree, raiz)).pack(side="right")

        tree.bind("<<TreeviewSelect>>",
                  lambda e: self._en_seleccion(tree))
        tree.bind("<Double-1>",
                  lambda e: self._expandir(tree))
        raiz.bind("<Return>", lambda e: self._confirmar(tree, raiz))
        raiz.bind("<Escape>", lambda e: raiz.destroy())

        self._poblar(tree)
        raiz.mainloop()
        return self.resultado

    def _poblar(self, tree, expandir=None):
        tree.delete(*tree.get_children())
        ruta = self._ruta_actual
        partes = []
        r = ruta
        while True:
            padre, hijo = os.path.split(r)
            if hijo:
                partes.append((r, hijo))
                r = padre
            else:
                partes.append((r, r))
                break
        partes.reverse()
        nodo_padre = ""
        ultimo = ""
        for full, nombre in partes:
            nodo = tree.insert(nodo_padre, "end", iid=full,
                               text=f"📂 {nombre}", open=True)
            nodo_padre = nodo
            ultimo = full
        try:
            for entry in sorted(os.scandir(ruta), key=lambda e: e.name.lower()):
                if entry.is_dir() and not entry.name.startswith("."):
                    tree.insert(ultimo, "end", iid=entry.path,
                                text=f"📁 {entry.name}")
        except PermissionError:
            pass
        self._var_sel.set(ruta)

    def _en_seleccion(self, tree):
        sel = tree.selection()
        if sel:
            self._var_sel.set(sel[0])

    def _expandir(self, tree):
        sel = tree.selection()
        if sel and os.path.isdir(sel[0]):
            self._ruta_actual = sel[0]
            self._var_ruta.set(sel[0])
            self._poblar(tree)

    def _navegar(self, ruta, tree):
        if os.path.isdir(ruta):
            self._ruta_actual = os.path.abspath(ruta)
            self._var_ruta.set(self._ruta_actual)
            self._poblar(tree)

    def _subir(self, tree):
        padre = os.path.dirname(self._ruta_actual)
        if padre != self._ruta_actual:
            self._ruta_actual = padre
            self._var_ruta.set(padre)
            self._poblar(tree)

    def _confirmar(self, tree, raiz):
        sel = tree.selection()
        ruta = sel[0] if sel else self._ruta_actual
        if not os.path.isdir(ruta):
            messagebox.showwarning("Atención", "Selecciona una carpeta válida.")
            return
        self.resultado = ruta
        raiz.destroy()

# ---------------------------------------------------------------------------

def cargar_configuracion(nombre_archivo="config.json"):
    ruta_base = os.path.dirname(os.path.abspath(__file__))
    ruta_config = os.path.join(ruta_base, nombre_archivo)
    try:
        with open(ruta_config, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error cargando config.json: {e}")
        return {}

def limpiar_monto(texto):
    """Detecta dinámicamente si el decimal es coma o punto y lo estandariza."""
    if not texto: return 0.0
    t = texto.strip().replace('$', '').replace('Bs', '').replace('"', '')
    pos_coma = t.rfind(',')
    pos_punto = t.rfind('.')
    if pos_coma > pos_punto:
        t = t.replace('.', '').replace(',', '.')
    elif pos_punto > pos_coma:
        t = t.replace(',', '')
    try:
        return float(t)
    except ValueError:
        return 0.0

def estandarizar_concepto(concepto):
    """Estandariza nombres de conceptos a códigos canónicos (case-insensitive)."""
    c = concepto.upper()
    if "FONDO PENSIONES" in c:
        return "FPJ"
    if "FONDO DE AHORRO" in c or "AHORRO" in c:
        return "FAOV"
    if "S.S.O." in c or "4%" in c:
        return "SSO"
    return concepto

def obtener_nombre_centro_desde_pdf(pdf, etiquetas):
    anclas = etiquetas.get("ancla_centro", [])
    if isinstance(anclas, str): anclas = [anclas]
    centro_detectado = "CENTRO DESCONOCIDO"
    ignorar = ["NRO CONTROL", "MES:", "AÑO:", "RECIBO:", "PÁGINA", "FECHA:", "PAGINA", "LISTADO DE"]
    if len(pdf.pages) > 0:
        pagina = pdf.pages[0]
        texto = pagina.extract_text()
        if texto:
            lineas = texto.split('\n')
            iterador = iter(lineas)
            for linea in iterador:
                l = linea.strip()
                for ancla in anclas:
                    if ancla in l:
                        extra = l.replace(ancla, "").strip(": ").strip()
                        if len(extra) > 3 and not any(x in extra.upper() for x in ignorar):
                            return extra
                        else:
                            try:
                                siguiente = next(iterador).strip()
                                if not any(x in siguiente.upper() for x in ignorar):
                                    return siguiente
                            except StopIteration: pass
            if len(lineas) > 3:
                for i in [2, 3, 4]:
                    if len(lineas) > i:
                        c = lineas[i].strip()
                        if len(c) > 3 and not any(x in c.upper() for x in ignorar):
                            if "CORPORACION" not in c.upper() and "GOBERNACION" not in c.upper():
                                return c
    return centro_detectado

def extraer_datos_formato_tabla(pdf, nombre_centro):
    datos = []
    concepto_actual = ""
    # Exclusiones estrictas
    exclusiones = ["PÁGINA", "PAGINA", "CONCEPTO:", "TRAB.", "EMPRESA", "RECIBO:", "MES:", "AÑO:", "FECHA:"]
    
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if not texto: continue
        lineas = texto.split('\n')
        
        for linea in lineas:
            linea_up = linea.upper().strip()
            
            # 1. Saltamos basura explícita
            if "LISTADO DE" in linea_up or any(ex in linea_up for ex in exclusiones):
                if "CONCEPTO:" in linea_up:
                    concepto_actual = estandarizar_concepto(linea_up.split("CONCEPTO:")[-1].strip())
                continue

            if not concepto_actual: continue

            # 2. Caso Especial: TOTAL DEDUCCIONES
            if "TOTAL DEDUCCIONES" in linea_up:
                # Buscamos números al FINAL de la línea (para evitar la fecha del encabezado)
                match_nums = re.findall(r'(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)', linea)
                if len(match_nums) >= 2:
                    # En totales, la cantidad suele ser el primer número tras el texto
                    # pero los montos son los últimos dos
                    cant = match_nums[0]
                    trab_raw = match_nums[-2] if len(match_nums) >= 2 else "0,00"
                    emp_raw = match_nums[-1]
                    
                    datos.append([
                        nombre_centro, "TOTAL GENERAL", 
                        f"DEDUCCION {concepto_actual} TOTAL {trab_raw} {emp_raw}",
                        concepto_actual, cant, limpiar_monto(trab_raw), limpiar_monto(emp_raw)
                    ])
                continue

            # 3. Caso Normal: Grupos de nómina
            match_nums = re.findall(r'(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2,3})|\b\d+\b)', linea)
            if len(match_nums) >= 2:
                nombre_grupo = linea.split(match_nums[0])[0].strip()
                
                # Si el grupo es basura o muy corto, saltar
                if len(nombre_grupo) < 4 or "TOTALES" in nombre_grupo.upper():
                    continue

                cant = match_nums[0]
                trab_raw = match_nums[1]
                emp_raw = match_nums[2] if len(match_nums) > 2 else "0,00"
                
                datos.append([
                    nombre_centro, nombre_grupo, 
                    f"DEDUCCION {concepto_actual} {nombre_grupo} {trab_raw} {emp_raw}", 
                    concepto_actual, cant, limpiar_monto(trab_raw), limpiar_monto(emp_raw)
                ])
    return datos

def extraer_datos_formato_listado(pdf, nombre_centro):
    datos = []
    grupo_actual = "SIN GRUPO"
    conteo_trabajadores = 0
    en_listado = False
    
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if not texto: continue
        lineas = texto.split('\n')
        for linea in lineas:
            linea_up = linea.upper().strip()
            
            if "GRUPO DE NÓMINA" in linea_up:
                grupo_actual = linea_up.split(':')[-1].strip()
                continue
            
            if "LISTADO DE" in linea_up:
                en_listado = True
                conteo_trabajadores = 0
                continue
            
            if en_listado and re.match(r'^\d{5,9}\s+', linea.strip()):
                conteo_trabajadores += 1
                
            if "TOTAL DEDUCCIONES" in linea_up:
                en_listado = False
                montos = re.findall(r'(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)', linea)
                
                if len(montos) >= 1:
                    trab_raw = montos[-2] if len(montos) >= 2 else montos[-1]
                    emp_raw = montos[-1] if len(montos) >= 2 else "0,00"
                    
                    concepto = estandarizar_concepto(linea_up.replace("TOTAL DEDUCCIONES", "").split('(')[0].strip())
                    datos.append([
                        nombre_centro, grupo_actual, linea.strip(), concepto, 
                        conteo_trabajadores, limpiar_monto(trab_raw), limpiar_monto(emp_raw)
                    ])
    return datos

def ejecutor_final():
    config = cargar_configuracion()
    etiquetas = config.get("etiquetas_pdf", {})
    encabezado = ["Centro", "Grupo", "Referencia", "Concepto", "Cant", "Aporte Trabajador", "Aporte Empresa"]

    pdfs = SelectorArchivos(titulo="Seleccionar PDFs de Casas Comerciales").abrir()
    if not pdfs: return
    destino = SelectorCarpeta(titulo="Seleccionar carpeta de salida").abrir()
    if not destino: return

    consolidado_total = []

    for ruta in pdfs:
        nombre_archivo_raw = os.path.basename(ruta).replace(".pdf", "")
        print(f"-> {nombre_archivo_raw}")
        try:
            with pdfplumber.open(ruta) as pdf:
                nombre_centro_real = obtener_nombre_centro_desde_pdf(pdf, etiquetas)
                test_text = pdf.pages[0].extract_text().upper()
                
                if "POR GRUPOS DE NÓMINA" in test_text:
                    datos_archivo = extraer_datos_formato_tabla(pdf, nombre_centro_real)
                else:
                    datos_archivo = extraer_datos_formato_listado(pdf, nombre_centro_real)

                if datos_archivo:
                    ruta_ind = os.path.join(destino, f"Resumen_{nombre_archivo_raw}.csv")
                    with open(ruta_ind, mode='w', encoding='utf-8-sig', newline='') as f:
                        writer = csv.writer(f, delimiter=';')
                        writer.writerow(encabezado)
                        writer.writerows(datos_archivo)
                    consolidado_total.extend(datos_archivo)
        except Exception as e:
            print(f"Error: {e}")

    if consolidado_total:
        with open(os.path.join(destino, "_CONSOLIDADO_TOTAL.csv"), mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(encabezado)
            writer.writerows(consolidado_total)
        if platform.system() == "Windows": os.startfile(destino)

if __name__ == "__main__":
    ejecutor_final()