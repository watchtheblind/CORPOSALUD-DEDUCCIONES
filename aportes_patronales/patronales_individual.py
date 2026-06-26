import pdfplumber
import json
from reporte_excel import ReporteExcel
import csv
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import platform
class ArchivoManager:
    """Maneja la selección de archivos y la escritura final."""
    def seleccionar_pdf(self):
        root = tk.Tk()
        root.withdraw()
        ruta = filedialog.askopenfilename(
            title="Seleccione el archivo de nómina (PDF)",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        return ruta

    def guardar_csv(self, datos, salida="resultado_nomina.csv"):
        if not datos:
            messagebox.showwarning("Atención", "No se extrajeron datos. Revisa el formato del PDF.")
            return

        columnas = ["CENTRO", "GREMIO", "TRABAJADOR", "EMPRESA", "CONCEPTO", "NRO_RECIBO"]
        try:
            with open(salida, mode='w', encoding='utf-8-sig', newline='') as f:
                escritor = csv.DictWriter(f, fieldnames=columnas)
                escritor.writeheader()
                escritor.writerows(datos)
            messagebox.showinfo("Éxito", f"Archivo '{salida}' generado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el CSV: {e}")

class NominaExtractor:
    def __init__(self, nombre_config="config.json"):
        # Cargamos la configuración desde el JSON
        ruta_base = os.path.dirname(os.path.abspath(__file__))
        ruta_config = os.path.join(ruta_base, nombre_config)

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

        try:
            with pdfplumber.open(ruta_pdf) as pdf:
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if not texto: continue
                    
                    lineas = texto.split('\n')
                    iterador = iter(lineas)

                    for linea in iterador:
                        l = linea.strip()
                        if not l or "TOTALES" in l: continue

                        # 1. Capturar Centro usando el ancla del JSON
                        if self.etiquetas["ancla_centro"] in l:
                            try:
                                centro_actual = next(iterador).strip()
                            except StopIteration: pass
                            continue

                        # 2. Capturar Recibo con Regex del JSON
                        match_recibo = re.search(self.etiquetas["ancla_recibo"], l)
                        if match_recibo:
                            recibo_actual = match_recibo.group(1)
                            continue

                        # 3. Capturar Concepto
                        if self.etiquetas["ancla_concepto"] in l:
                            concepto_crudo = l.replace(self.etiquetas["ancla_concepto"], "").strip()
                            concepto_actual = self.mapa_conceptos.get(concepto_crudo, concepto_crudo)
                            continue

                        # 4. Extraer Datos de Fila
                        match_datos = self.re_datos.match(l)
                        if match_datos and concepto_actual:
                            # Aquí puedes mapear dinámicamente según la posición en la Regex
                            datos_finales.append({
                                "CENTRO": centro_actual,
                                "GREMIO": match_datos.group(1).strip().replace('"', ''),
                                "TRABAJADOR": match_datos.group(3).replace(',', ''),
                                "EMPRESA": match_datos.group(4).replace(',', ''),
                                "CONCEPTO": concepto_actual,
                                "NRO_RECIBO": recibo_actual
                            })
            return datos_finales
        except Exception as e:
            print(f"Error procesando PDF: {e}")
            return []
        

# ... (tus otros imports se mantienen igual)

def main():
    root = tk.Tk()
    root.withdraw()
    
    nombre_excel = "Auditoria_Deducciones_Generado.xlsx" # Lo definimos en una variable
    archivo_pdf = filedialog.askopenfilename(title="Seleccionar PDF de Nómina", filetypes=[("PDF files", "*.pdf")])
    if not archivo_pdf: return

    extractor = NominaExtractor()
    datos = extractor.procesar_pdf(archivo_pdf)

    if datos:
        reporte = ReporteExcel(nombre_excel)
        if reporte.generar_reporte(datos):
            # --- LÓGICA PARA ABRIR EL ARCHIVO ---
            try:
                if platform.system() == 'Windows':
                    os.startfile(nombre_excel)
                elif platform.system() == 'Darwin':  # macOS
                    os.system(f'open "{nombre_excel}"')
                else:  # Linux
                    os.system(f'xdg-open "{nombre_excel}"')
            except Exception as e:
                print(f"No se pudo abrir el archivo automáticamente: {e}")

            messagebox.showinfo("Proceso Exitoso", f"Se ha generado y abierto: {nombre_excel}")
    else:
        messagebox.showwarning("Sin datos", "No se pudo extraer información del PDF.")
if __name__ == "__main__":
    main()