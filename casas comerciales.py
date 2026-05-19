import pdfplumber
import csv
import re
import tkinter as tk
from tkinter import filedialog
import os
import platform
import sys
import io

# Configuración para evitar errores de caracteres en la terminal de Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def barra_progreso(actual, total, archivo=""):
    ancho = 40
    progreso = int(ancho * actual / total)
    barra = "#" * progreso + "-" * (ancho - progreso)
    porcentaje = (actual / total) * 100
    sys.stdout.write(f"\r|{barra}| {porcentaje:.1f}% Procesando: {archivo[:30]}...")
    sys.stdout.flush()

def extraer_datos_formato_tabla(pdf, nombre_centro):
    datos = []
    concepto_actual = "DESCONOCIDO"
    exclusiones = ["TOTALES", "PÁGINA", "CONCEPTO:", "TRAB.", "EMPRESA", "GRUPO DE NÓMINA", "RECIBO:"]
    
    for pagina in pdf.pages:
        texto = pagina.extract_text()
        if not texto: continue
        lineas = texto.split('\n')
        
        for linea in lineas:
            linea_up = linea.upper().strip()
            
            if "CONCEPTO:" in linea_up:
                concepto_actual = linea_up.split("CONCEPTO:")[-1].strip()
                continue

            if any(c.isdigit() for c in linea) and not any(ex in linea_up for ex in exclusiones):
                # NUEVO REGEX: Soporta tildes (áéíóú) y eñes (ñ) tanto en minúsculas como mayúsculas
                # También captura los números con sus separadores de miles y decimales
                partes = re.findall(r'([a-zA-ZáéíóúÁÉÍÓÚñÑ\s\(\)\-\/]+)|(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)', linea)
                
                elementos = [p[0].strip() if p[0] else p[1] for p in partes if any(p)]
                
                # A veces el regex deja espacios vacíos al final del nombre del gremio
                if len(elementos) >= 3:
                    grupo_visto = elementos[0] # Ahora captura "Médicos" sin romperse [cite: 7, 56]
                    cant = elementos[1]        # Cantidad de trabajadores [cite: 7]
                    trab = elementos[2]        # Aporte trabajador [cite: 7]
                    emp = elementos[3] if len(elementos) > 3 else "0.00"
                    
                    cadena_simulada = f"DEDUCCION {concepto_actual} {grupo_visto} {trab} {emp}"
                    
                    fila_excel = len(datos) + 2
                    f_trabajador = f'=ESPACIOS(IZQUIERDA(DERECHA(SUSTITUIR(C{fila_excel};" ";REPETIR(" ";100));200);100))'
                    f_empresa = f'=ESPACIOS(DERECHA(SUSTITUIR(C{fila_excel};" ";REPETIR(" ";100));100))'

                    datos.append([
                        nombre_centro, 
                        grupo_visto, 
                        cadena_simulada, 
                        concepto_actual, 
                        cant, 
                        f_trabajador, 
                        f_empresa
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
                partes = linea_up.split(':')
                grupo_actual = partes[-1].strip() if len(partes) > 1 else "VARIOS"
                continue
            if "LISTADO DE" in linea_up:
                en_listado = True
                conteo_trabajadores = 0
                continue
            if en_listado and re.match(r'^\d{5,9}\s+', linea.strip()):
                conteo_trabajadores += 1
            if "TOTAL DEDUCCIONES" in linea_up:
                en_listado = False
                cadena_original = linea.strip()
                concepto = linea_up.replace("TOTAL DEDUCCIONES", "").split('(')[0].strip()
                fila = len(datos) + 2
                f_trabajador = f'=ESPACIOS(IZQUIERDA(DERECHA(SUSTITUIR(C{fila};" ";REPETIR(" ";100));200);100))'
                f_empresa = f'=ESPACIOS(DERECHA(SUSTITUIR(C{fila};" ";REPETIR(" ";100));100))'
                datos.append([nombre_centro, grupo_actual, cadena_original, concepto, conteo_trabajadores, f_trabajador, f_empresa])
    return datos

def ejecutor_final():
    root = tk.Tk()
    root.withdraw()
    
    pdfs = filedialog.askopenfilenames(title="1. Elige los archivos PDF")
    if not pdfs: return
    
    destino = filedialog.askdirectory(title="2. Elige dónde guardar los CSV")
    if not destino: return

    print(f"\n Procesando {len(pdfs)} archivos...\n")

    for i, ruta in enumerate(pdfs, 1):
        nombre = os.path.basename(ruta).replace(".pdf", "")
        barra_progreso(i, len(pdfs), nombre)
        
        try:
            with pdfplumber.open(ruta) as pdf:
                # Detección de formato basada en el contenido de la primera página [cite: 4, 17]
                test_text = pdf.pages[0].extract_text().upper()
                if "POR GRUPOS DE NÓMINA" in test_text:
                    datos = extraer_datos_formato_tabla(pdf, nombre)
                else:
                    datos = extraer_datos_formato_listado(pdf, nombre)

            # Guardar CSV con codificación Excel (utf-8-sig) y delimitador punto y coma
            archivo_csv = os.path.join(destino, f"Resumen_{nombre}.csv")
            with open(archivo_csv, mode='w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(["Centro", "Grupo", "Referencia", "Concepto", "Cant", "Aporte Trabajador", "Aporte Empresa"])
                writer.writerows(datos)

        except Exception as e:
            print(f"\nError en {nombre}: {e}")

    print(f"\n\n ¡Listo! Carpeta abierta: {destino}")
    if platform.system() == "Windows":
        os.startfile(destino)

if __name__ == "__main__":
    ejecutor_final()