import pdfplumber
import json
import csv
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import platform
import sys
from reporte_excel import ReporteExcel

def limpiar_monto(texto):
    """Detecta dinámicamente si el decimal es coma o punto y lo estandariza."""
    if not texto: return 0.0
    t = texto.strip().replace('$', '').replace('Bs', '')
    
    pos_coma = t.rfind(',')
    pos_punto = t.rfind('.')
    
    if pos_coma > pos_punto:
        # Formato 1.234,56 -> 1234.56
        t = t.replace('.', '').replace(',', '.')
    elif pos_punto > pos_coma:
        # Formato 1,234.56 -> 1234.56
        t = t.replace(',', '')
    
    try:
        return float(t)
    except ValueError:
        return 0.0

class NominaExtractor:
    def __init__(self, nombre_config="config.json"):
        # Detectar ruta automática del script
        if getattr(sys, 'frozen', False):
            ruta_base = os.path.dirname(sys.executable)
        else:
            ruta_base = os.path.dirname(os.path.abspath(__file__))
            
        ruta_config = os.path.join(ruta_base, nombre_config)
        
        if not os.path.exists(ruta_config):
            raise FileNotFoundError(f"No se encontró {nombre_config} en: {ruta_config}")

        with open(ruta_config, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.mapa_conceptos = self.config["mapa_conceptos"]
        self.re_datos = re.compile(self.config["regex_datos"])
        self.etiquetas = self.config["etiquetas_pdf"]

    def procesar_pdf(self, ruta_pdf):
        datos_finales = []
        centro_actual = "" 
        recibo_actual = ""
        concepto_actual = ""
        
        # Cargamos las anclas de respaldo del config
        anclas_guia = self.etiquetas.get("ancla_centro", [])
        if isinstance(anclas_guia, str): 
            anclas_guia = [anclas_guia]

        try:
            with pdfplumber.open(ruta_pdf) as pdf:
                total_pags = len(pdf.pages)
                nombre_fichero = os.path.basename(ruta_pdf)
                print(f"\n--- Procesando: {nombre_fichero} ---")
                
                for i, pagina in enumerate(pdf.pages):
                    print(f"   -> Leyendo página {i+1}/{total_pags}...", end='\r')
                    texto = pagina.extract_text()
                    if not texto: continue
                    
                    lineas = texto.split('\n')
                    
                    # 1. DETECCIÓN AUTOMÁTICA DE CENTRO (Por posición en el encabezado)
                    if not centro_actual and len(lineas) > 3:
                        # Buscamos en las líneas 3 y 4 del documento
                        linea_3 = lineas[2].strip()
                        linea_4 = lineas[3].strip() if len(lineas) > 3 else ""
                        
                        if "CORPORACION DE SALUD" in linea_3.upper() or "GOBERNACION" in linea_3.upper():
                            centro_actual = linea_4
                        else:
                            centro_actual = linea_3
                        
                        print(f"   [Centro detectado: {centro_actual}]")

                    iterador = iter(lineas)
                    for linea in iterador:
                        l = linea.strip()
                        if not l or "TOTALES" in l: continue

                        # 2. RESPALDO DE CENTRO (Si aparece una frase guía de RRHH)
                        if any(ancla in l for ancla in anclas_guia):
                            try:
                                for ancla in anclas_guia:
                                    if ancla in l:
                                        extra = l.replace(ancla, "").strip(": ").strip()
                                        if len(extra) > 3:
                                            centro_actual = extra
                                        else:
                                            centro_actual = next(iterador).strip()
                                        break
                            except StopIteration: pass
                            continue

                        # 3. CAPTURAR RECIBO
                        match_recibo = re.search(self.etiquetas["ancla_recibo"], l)
                        if match_recibo:
                            recibo_actual = match_recibo.group(1)
                            continue

                        # 4. CAPTURAR CONCEPTO
                        if self.etiquetas["ancla_concepto"] in l:
                            concepto_crudo = l.replace(self.etiquetas["ancla_concepto"], "").strip()
                            # El mapa ahora devuelve el nombre con espacios
                            concepto_actual = self.mapa_conceptos.get(concepto_crudo, concepto_crudo)
                            continue

                        # 5. EXTRAER DATOS Y APLICAR LÓGICA DE APORTES
                        match_datos = self.re_datos.match(l)
                        if match_datos and concepto_actual:
                            m_trab_raw = match_datos.group(3).strip()
                            m_empr_raw = match_datos.group(4).strip()

                            # Convertir a número usando el limpiador inteligente
                            val_trab = limpiar_monto(m_trab_raw)
                            val_empr = limpiar_monto(m_empr_raw)

                            # Lógica especial: Si es Caja de Ahorro, todo va a Empresa
                            claves_solo_empresa = ["CAHORMINSAS", "CAJA DE AHORRO", "CAEMINSA", "SERVICIOS FUNERARIOS CAHORMINSAS"]
                            if any(c in concepto_actual.upper() for c in claves_solo_empresa):
                                val_empr = val_trab + val_empr
                                val_trab = 0.0

                            datos_finales.append({
                                "CENTRO": centro_actual,
                                "GREMIO": match_datos.group(1).strip().replace('"', ''),
                                "TRABAJADOR": val_trab,
                                "EMPRESA": val_empr,
                                "CONCEPTO": concepto_actual,
                                "NRO_RECIBO": recibo_actual,
                                "TRAB_ORIGINAL": m_trab_raw, # Columna de reserva
                                "EMPR_ORIGINAL": m_empr_raw  # Columna de reserva
                            })
            return datos_finales
        except Exception as e:
            print(f"\nError en {ruta_pdf}: {e}")
            return []

def main():
    root = tk.Tk()
    root.withdraw()
    
    archivos_pdf = filedialog.askopenfilenames(title="Seleccionar PDFs de Nómina", filetypes=[("PDF files", "*.pdf")])
    if not archivos_pdf: return

    nombre_excel = filedialog.asksaveasfilename(
        title="Guardar reporte como...",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="Auditoria_Deducciones_Generado.xlsx"
    )
    if not nombre_excel: return

    try:
        extractor = NominaExtractor()
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return

    todos_los_datos = []
    for ruta in archivos_pdf:
        datos = extractor.procesar_pdf(ruta)
        todos_los_datos.extend(datos)

    if todos_los_datos:
        reporte = ReporteExcel(nombre_excel)
        if reporte.generar_reporte(todos_los_datos):
            try:
                if platform.system() == 'Windows': os.startfile(nombre_excel)
                else: os.system(f'{"open" if platform.system()=="Darwin" else "xdg-open"} "{nombre_excel}"')
            except: pass
            messagebox.showinfo("Éxito", f"Proceso completado. Registros: {len(todos_los_datos)}")
    else:
        messagebox.showwarning("Atención", "No se extrajeron datos.")

if __name__ == "__main__":
    main()