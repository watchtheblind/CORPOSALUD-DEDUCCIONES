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

# Diccionario de mapeo proporcionado
MAPEO_CENTROS = {
    "SAANA": ["S.A.A.N.A.", "SAANA", "S.A.A.N.A"],
    "SAGER": ["S.A.G.E.R.", "SAGER", "S.A.G.E.R"],
    "AMBULATORIO DE TURMERO": ["AMB. DE TURMERO", "AMBULATORIO DE TURMERO", "SALA DE PARTO DE TURMERO"],
    "SOCIEDAD CIVIL HOSPITAL DEL SUR": ["HOSPITAL DEL SUR", "SOCIEDAD CIVIL HOSPITAL DEL SUR", "SOCIEDAD CIVIL"],
    "D.M.S CAMATAGUA": ["D.M.S. CAMATAGUA", "D.M.S CAMATAGUA-URDANETA", "D.M.S CAMATAGUA-"],
    "HOSPITAL CENTRAL DE MARACAY": ["S.A. HOSPITAL CENTRAL DE MARACAY"],
    "HOSPITAL JOSE MARIA BENITEZ": ["S.A. HOSPITAL JOSE MARIA BENITEZ"],
    "HOSPITAL JOSE MARIA VARGAS": ["HOSPITAL JOSE MARIA VARGAS"],
    "HOSPITAL JOSE RANGEL": ["HOSPITAL JOSE RANGEL"],
    "HOSPITAL LAS TEJERIAS": ["HOSPITAL LAS TEJERIAS"],
    "CLINICA PSIQUIATRICA DE MARACAY": ["CLINICA PSIQUIATRICA DE MARACAY"],
    "D.M.S. GIRARDOT": ["D.M.S. GIRARDOT"],
    "D.M.S. MARIÑO": ["D.M.S. MARIÑO"],
    "D.M.S. RIBAS": ["D.M.S. RIBAS"],
    "D.M.S. SUCRE": ["D.M.S. SUCRE"],
    "D.M.S. ZAMORA": ["D.M.S. ZAMORA"],
    "D.M.S. LIBERTADOR": ["D.M.S. LIBERTADOR"],
    "D.M.S. SANTOS MICHELENA": ["D.M.S. SANTOS MICHELENA"],
    "D.M.S. TOVAR": ["D.M.S. TOVAR"],
    "D.M.S. REVENGA": ["D.M.S. REVENGA"],
    "D.M.S. SAN CASIMIRO": ["D.M.S. SAN CASIMIRO"],
    "D.M.S. BOLIVAR": ["D.M.S. BOLIVAR"],
    "D.M.S. FRANCISCO LINARES ALCANTARA": ["D.M.S. FRANCISCO LINARES ALCANTARA"],
    "D.M.S. MARIO BRICEÑO IRAGORRI": ["D.M.S. MARIO BRICEÑO IRAGORRI"],
    "AMB. PALO NEGRO": ["AMB. PALO NEGRO"],
    "AMB. LA CANDELARIA": ["AMB. LA CANDELARIA"],
    "SERVICIOS CENTRALES": ["SERVICIOS CENTRALES"],
    "CORPORACION DE SALUD DEL ESTADO ARAGUA": ["CORPORACION DE SALUD DEL ESTADO ARAGUA"]
}

def normalizar_nombre_centro(nombre_archivo):
    """
    Compara el nombre del archivo contra los sinónimos del diccionario.
    """
    # Convertimos el nombre del archivo a mayúsculas y quitamos espacios extra
    archivo_upper = nombre_archivo.upper().strip()
    
    for nombre_estandar, sinonimos in MAPEO_CENTROS.items():
        for s in sinonimos:
            # Si el sinónimo está contenido en el nombre del archivo
            if s.upper() in archivo_upper:
                return nombre_estandar
    
    # Si no hubo coincidencia, devolvemos el original
    return nombre_archivo

def barra_progreso(actual, total, archivo=""):
    ancho = 40
    progreso = int(ancho * actual / total)
    barra = "#" * progreso + "-" * (ancho - progreso)
    porcentaje = (actual / total) * 100
    sys.stdout.write(f"\r|{barra}| {porcentaje:.1f}% Procesando: {archivo[:30]}...")
    sys.stdout.flush()

def extraer_datos_formato_tabla(pdf, nombre_centro, offset_consolidado=0):
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
                partes = re.findall(r'([a-zA-ZáéíóúÁÉÍÓÚñÑ\s\(\)\-\/]+)|(\d{1,3}(?:[\.\,]\d{3})*(?:[\.\,]\d{2})|\b\d+\b)', linea)
                elementos = [p[0].strip() if p[0] else p[1] for p in partes if any(p)]
                
                if len(elementos) >= 3:
                    grupo_visto = elementos[0]
                    cant = elementos[1]
                    trab = elementos[2]
                    emp = elementos[3] if len(elementos) > 3 else "0.00"
                    
                    cadena_simulada = f"DEDUCCION {concepto_actual} {grupo_visto} {trab} {emp}"
                    
                    fila_excel = len(datos) + 2 + offset_consolidado
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

def extraer_datos_formato_listado(pdf, nombre_centro, offset_consolidado=0):
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
                
                fila = len(datos) + 2 + offset_consolidado
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

    consolidado_total = []
    encabezado = ["Centro", "Grupo", "Referencia", "Concepto", "Cant", "Aporte Trabajador", "Aporte Empresa"]

    for i, ruta in enumerate(pdfs, 1):
        # Obtenemos el nombre base del archivo sin extensión
        nombre_archivo_raw = os.path.basename(ruta).replace(".pdf", "")
        
        # NORMALIZACIÓN REAL: Aquí comparamos el nombre del archivo físico
        nombre_centro_mapeado = normalizar_nombre_centro(nombre_archivo_raw)
        
        barra_progreso(i, len(pdfs), nombre_archivo_raw)
        
        try:
            with pdfplumber.open(ruta) as pdf:
                test_text = pdf.pages[0].extract_text().upper()
                
                # Procesamiento individual (usando el nombre mapeado)
                if "POR GRUPOS DE NÓMINA" in test_text:
                    datos_ind = extraer_datos_formato_tabla(pdf, nombre_centro_mapeado)
                else:
                    datos_ind = extraer_datos_formato_listado(pdf, nombre_centro_mapeado)

                # Guardamos el CSV individual
                archivo_csv = os.path.join(destino, f"Resumen_{nombre_archivo_raw}.csv")
                with open(archivo_csv, mode='w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(encabezado)
                    writer.writerows(datos_ind)

                # Consolidación (usando el nombre mapeado y ajustando fórmulas)
                offset = len(consolidado_total)
                if "POR GRUPOS DE NÓMINA" in test_text:
                    datos_cons = extraer_datos_formato_tabla(pdf, nombre_centro_mapeado, offset_consolidado=offset)
                else:
                    datos_cons = extraer_datos_formato_listado(pdf, nombre_centro_mapeado, offset_consolidado=offset)
                
                consolidado_total.extend(datos_cons)

        except Exception as e:
            print(f"\nError en {nombre_archivo_raw}: {e}")

    # Guardado del gran consolidado
    if consolidado_total:
        ruta_consolidado = os.path.join(destino, "_CONSOLIDADO_TOTAL.csv")
        with open(ruta_consolidado, mode='w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(encabezado)
            writer.writerows(consolidado_total)
        print(f"\n\n Archivo consolidado generado exitosamente.")

    print(f"\n ¡Listo! Carpeta abierta: {destino}")
    if platform.system() == "Windows":
        os.startfile(destino)

if __name__ == "__main__":
    ejecutor_final()