import pdfplumber
import csv
import re
import tkinter as tk
from tkinter import filedialog
import os
import platform
import sys
import io
import json
import pandas as pd

# Configuración para evitar errores de caracteres en la terminal de Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

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

def obtener_nombre_centro_desde_pdf(pdf, etiquetas):
    anclas = etiquetas.get("ancla_centro", [])
    if isinstance(anclas, str): anclas = [anclas]
    
    ignorar = [
        "NRO CONTROL", "MES:", "AÑO:", "RECIBO:", "PÁGINA", "FECHA:", "PAGINA", 
        "LISTADO DE", "DIRECCIÓN REGIONAL DE RECURSOS HUMANOS", 
        "DIRECCION DE RECURSOS HUMANOS", "RECURSOS HUMANOS"
    ]

    if len(pdf.pages) > 0:
        pagina = pdf.pages[0]
        texto = pagina.extract_text()
        if texto:
            lineas = texto.split('\n')
            
            # 1. Por Anclas
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

            # 2. Por Posición (Líneas 3 a 6)
            for i in range(2, min(len(lineas), 7)):
                c = lineas[i].strip()
                if len(c) > 3 and not any(x in c.upper() for x in ignorar):
                    if "CORPORACION" not in c.upper() and "GOBERNACION" not in c.upper():
                        return c
    return "CENTRO DESCONOCIDO"

def estandarizar_concepto(concepto):
    """Unifica variantes de nombres de conceptos para el agrupamiento en Excel."""
    c = concepto.upper().strip()
    if "FUNERARIOS" in c and "CAHORMINSAS" in c:
        return "CAHORMINSAS"
    if "CAJA DE AHORRO CAEMINSA" in c or "CAEMINSA" in c:
        return "CAEMINSA"
    # Puedes añadir más reglas aquí si aparecen otros casos
    return c

def extraer_datos_formato_tabla(pdf, nombre_centro):
    datos = []
    concepto_actual = ""
    exclusiones = ["PÁGINA", "PAGINA", "CONCEPTO:", "TRAB.", "EMPRESA", "RECIBO:", "MES:", "AÑO:", "FECHA:"]
    
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if not texto: continue
        lineas = texto.split('\n')
        for linea in lineas:
            linea_up = linea.upper().strip()
            
            if "LISTADO DE" in linea_up or any(ex in linea_up for ex in exclusiones):
                if "CONCEPTO:" in linea_up:
                    concepto_raw = linea_up.split("CONCEPTO:")[-1].strip()
                    concepto_actual = estandarizar_concepto(concepto_raw)
                continue
            if not concepto_actual: continue

            if "TOTAL DEDUCCIONES" in linea_up:
                match_nums = re.findall(r'(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)', linea)
                if len(match_nums) >= 2:
                    cant = match_nums[0]
                    trab_raw = match_nums[-2]
                    emp_raw = match_nums[-1]
                    datos.append([nombre_centro, "TOTAL GENERAL", f"DEDUCCION {concepto_actual} TOTAL {trab_raw} {emp_raw}", concepto_actual, cant, limpiar_monto(trab_raw), limpiar_monto(emp_raw)])
                continue

            match_nums = re.findall(r'(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2,3})|\b\d+\b)', linea)
            if len(match_nums) >= 2:
                nombre_grupo = linea.split(match_nums[0])[0].strip()
                if len(nombre_grupo) < 4 or "TOTALES" in nombre_grupo.upper(): continue
                cant = match_nums[0]
                trab_raw = match_nums[1]
                emp_raw = match_nums[2] if len(match_nums) > 2 else "0,00"
                datos.append([nombre_centro, nombre_grupo, f"DEDUCCION {concepto_actual} {nombre_grupo} {trab_raw} {emp_raw}", concepto_actual, cant, limpiar_monto(trab_raw), limpiar_monto(emp_raw)])
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
                    concepto_raw = linea_up.replace("TOTAL DEDUCCIONES", "").split('(')[0].strip()
                    concepto_final = estandarizar_concepto(concepto_raw)
                    datos.append([nombre_centro, grupo_actual, linea.strip(), concepto_final, conteo_trabajadores, limpiar_monto(trab_raw), limpiar_monto(emp_raw)])
    return datos

def ejecutor_final():
    config = cargar_configuracion()
    etiquetas = config.get("etiquetas_pdf", {})
    encabezado = ["Centro", "Grupo", "Referencia", "Concepto", "Cant", "Aporte Trabajador", "Aporte Empresa"]

    root = tk.Tk()
    root.withdraw()
    pdfs = filedialog.askopenfilenames(title="Seleccionar PDFs")
    if not pdfs: return
    destino = filedialog.askdirectory(title="Carpeta de Salida")
    if not destino: return

    consolidado_total = []

    for ruta in pdfs:
        nombre_archivo_raw = os.path.basename(ruta).replace(".pdf", "")
        print(f"Procesando: {nombre_archivo_raw}")
        try:
            with pdfplumber.open(ruta) as pdf:
                nombre_centro_real = obtener_nombre_centro_desde_pdf(pdf, etiquetas)
                test_text = pdf.pages[0].extract_text().upper()
                
                datos_archivo = extraer_datos_formato_tabla(pdf, nombre_centro_real) if "POR GRUPOS DE NÓMINA" in test_text else extraer_datos_formato_listado(pdf, nombre_centro_real)

                if datos_archivo:
                    # Individual CSV
                    ruta_csv_ind = os.path.join(destino, f"Resumen_{nombre_archivo_raw}.csv")
                    with open(ruta_csv_ind, mode='w', encoding='utf-8-sig', newline='') as f:
                        writer = csv.writer(f, delimiter=';')
                        writer.writerow(encabezado)
                        writer.writerows(datos_archivo)
                    consolidado_total.extend(datos_archivo)
        except Exception as e:
            print(f"Error en {nombre_archivo_raw}: {e}")

    if consolidado_total:
        df = pd.DataFrame(consolidado_total, columns=encabezado)
        ruta_excel = os.path.join(destino, "_AUDITORIA_CASAS_COMERCIALES.xlsx")
        
        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='BASE', index=False)
            for concepto in sorted(df['Concepto'].unique()):
                nombre_hoja = str(concepto)[:31].replace(':', '').replace('/', '').strip()
                df[df['Concepto'] == concepto].to_excel(writer, sheet_name=nombre_hoja, index=False)

        print(f"\n¡Proceso Exitoso! Archivos guardados en: {destino}")
        if platform.system() == "Windows": os.startfile(destino)

if __name__ == "__main__":
    ejecutor_final()