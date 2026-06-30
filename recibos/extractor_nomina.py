import pdfplumber
import os
import re
import pandas as pd
import tkinter as tk
from datetime import datetime

def registrar_auditoria(mensaje):
    fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open("auditoria_recibos.txt", "a", encoding="utf-8") as f:
        f.write(f"[{fecha_hora}] {mensaje}\n")

def obtener_nombre_hoja_base():
    meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
             "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
    fecha = datetime.now()
    return f"BASE {meses[fecha.month - 1]} {fecha.year}"

def limpiar_monto_estandar(texto_monto):
    if not texto_monto: return 0.0
    texto = texto_monto.strip()
    u_coma = texto.rfind(',')
    u_punto = texto.rfind('.')
    if u_coma > u_punto:
        texto = texto.replace('.', '').replace(',', '.')
    elif u_punto > u_coma:
        texto = texto.replace(',', '')
    else:
        texto = texto.replace(',', '.')
    texto = "".join(c for c in texto if c.isdigit() or c == '.')
    try: return float(texto)
    except: return 0.0

def transformar_y_filtrar(linea):
    conceptos_permitidos = {
        #SSO
        "LEY S.S.O. (4%)": "210",
        "S.S.O. (DEUDA)": "210.1.1",
        #COLEGIOS
        "COLEGIO BIOANALISTA": "255",
        "COLEGIO NUTRIC. Y DIET. VZLA": "286",
        "COLEGIO ENFERMERAS": "224",           
        "COLEGIO DE ENFERMERAS ESTADO CARABOBO": "967",
        "COLEGIO DE ENFERMERA EDO. MIRANDA": "603",
        #DIAS DESCANSO
        "NO LEY DESC. DIA(S) NO LABORADO(S)": "179",
        "DESC. DIA(S) NO LABORADO(S)": "179",
        #PIE
        "LEY PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.": "244",
        "LEY PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.)": "244",
        "PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.": "244",
        "PERDIDA INVOLUNTARIA DEL EMPLEO (P.I.E.)": "244",
        "PERDIDA INVOLUNTARIA DEL EMPLEO (DEUDA)": "244.2",
        "LEY PERDIDA INVOLUNTARIA DEL EMPLEO (DEUDA)": "244.2",
        #FAOV
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIV": "245",
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVI": "245",
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA": "245",
        "FONDO DE AHORRO OBLIGATORIO PARA LA VIV": "245.3",
        #DELEGACIONES
        "DELEGACIÓN REGIONAL INPRENFERMERA ARAGU": "370",
        "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA": "370",
        #PENSIONES
        "FONDO PENSIONES (JUBILACION)": "301",
        "FONDO PENSIONES JUBILACION (DEUDA)": "301.2",
        #CAHORMINSAS
        "NO LEY CAHORMINSAS": "212",
        "CAHORMINSAS": "212",
        "SERVICIOS FUNERARIOS CAHORMINSAS": "269",
        #CAJA AHORRO
        "CAJA DE AHORRO CAEMINSA": "238",
        "PRESTAMO CAJA DE AHORRO": "223",
        #SINDICATOS, SOCIEDADES
        "SINDICATO UNICO DE TRAB.DE LA SALUD Y S": "233",
        "SINDICATO OSBESS ARAGUA": "978",
        "SOCIEDAD ANESTECIOLOGOS": "257",
        "SINBOPROENF": "321",
        "SUNEP-SAS": "213",
        "FENASISTRASALUD": "513",
        "SISTRASALUD": "583",
        "SAPTRASEZ": "964",
        "SITRASSS-MIRANDA": "581",
        "SISTRASSS MIRANDA": "581",
        "ASUNAJUPENSAPROSO": "799",
        #DELEGACIONES
        "DELEGACIÓN REGIONAL INPREENFERMERA ARAGU": "370",
        "DELEGACIÓN REGIONAL INPREENFERMERA ARAGUA": "370",
        #TRIBUNALES
        "TRIBUNAL DE PROTECCION DE NIÑOS, NIÑAS": "807",
        "TRIBUNAL (PERMANENTE)": "215",
    }

    linea_up = linea.upper()
    candidatos = []
    for nombre, codigo in conceptos_permitidos.items():
        if nombre.upper() in linea_up:
            candidatos.append((nombre, codigo))
    if candidatos:
        mejor = max(candidatos, key=lambda x: len(x[0]))
        return f"{mejor[0]} {mejor[1]}"
    return None
MAPEO_CENTROS = {
    "SAANA": ["S.A.A.N.A.", "SAANA", "S.A.A.N.A"],
    "SAGER": ["S.A.G.E.R.", "SAGER", "S.A.G.E.R"],
    "AMBULATORIO DE TURMERO": ["AMB. DE TURMERO", "AMBULATORIO DE TURMERO"],
    "SALA DE PARTO DE TURMERO": ["SALA DE PARTO DE TURMERO"],
    "SOCIEDAD CIVIL HOSPITAL DEL SUR": ["HOSPITAL DEL SUR", "SOCIEDAD CIVIL HOSPITAL DEL SUR", "SOCIEDAD CIVIL"],
    "D.M.S CAMATAGUA": ["D.M.S. CAMATAGUA", "D.M.S CAMATAGUA-URDANETA", "D.M.S CAMATAGUA-"],
    "HOSPITAL CENTRAL DE MARACAY": ["S.A. HOSPITAL CENTRAL DE MARACAY"],
    "HOSPITAL JOSE MARIA BENITEZ": ["S.A. HOSPITAL JOSE MARIA BENITEZ"],
    "HOSPITAL JOSE MARIA VARGAS": ["HOSPITAL JOSE MARIA VARGAS"],
    "HOSPITAL JOSE RANGEL": ["HOSPITAL JOSE RANGEL"],
    "HOSPITAL LAS TEJERIAS": ["HOSPITAL LAS TEJERIAS"],
    "HOSPITAL NUESTRA SEÑORA DE LA CARIDAD": ["HOSPITAL NUESTRA SEÑORA DE LA CARIDAD"],
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

def normalizar_entidad(nombre_crudo):
    if not nombre_crudo: return "CENTRO DESCONOCIDO"
    nombre_limpio = nombre_crudo.strip().upper()
    candidatos = []
    for nombre_estandar, variaciones in MAPEO_CENTROS.items():
        if any(v.upper() in nombre_limpio for v in variaciones):
            candidatos.append(nombre_estandar)
    if candidatos:
        return max(candidatos, key=len)
    return nombre_limpio

def extraer_datos_de_carpeta(ruta_carpeta, entidades_dict, entidades_encontradas, es_deuda=False):
    if not ruta_carpeta or not os.path.exists(ruta_carpeta):
        registrar_auditoria(f"AVISO: La carpeta {ruta_carpeta} no existe")
        return []
        
    archivos = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith('.pdf')]
    lista_dfs = []
    tipo = "DEUDAS" if es_deuda else "NOMINAS"
    registrar_auditoria(f"--- INICIO ESCANEO CARPETA {tipo}: {ruta_carpeta} ---")

    for archivo in archivos:
        ruta_pdf = os.path.join(ruta_carpeta, archivo)
        nro_recibo, entidad = "S/N", "DESCONOCIDA"
        try:
            with pdfplumber.open(ruta_pdf) as pdf:
                texto_p1 = pdf.pages[0].extract_text().replace("\n", " ")
                #entidad = next((ent for ent in entidades_dict if ent in texto_p1), "OTRA ENTIDAD")
                entidad = normalizar_entidad(texto_p1)
                
                if entidad != "OTRA ENTIDAD":
                    entidades_encontradas.add(entidad)
                
                match_recibo = re.search(r"Recibo\s*N[º°]?\s*(\d+)", texto_p1)
                nro_recibo = match_recibo.group(1) if match_recibo else "S/N"
                
                texto_completo = "\n".join([p.extract_text() for p in pdf.pages])
                lineas = texto_completo.split('\n')
                datos_pdf = []
                activado = True if es_deuda else False
                
                for linea in lineas:
                    if not es_deuda and "TOTAL DE DEDUCCIONES POR CONCEPTO" in linea:
                        activado = True
                        continue
                    if "Total Asignaciones" in linea or "Total Pago Único" in linea:
                        if activado: break 
                    if activado:
                        partes = linea.strip().rsplit(' ', 1)
                        if len(partes) == 2:
                            conc = transformar_y_filtrar(partes[0])
                            if conc:
                                datos_pdf.append({
                                    "Conceptos": conc, "Entidad": entidad, 
                                    "Recibo": nro_recibo, "Montos": limpiar_monto_estandar(partes[1])
                                })
                if datos_pdf: lista_dfs.append(pd.DataFrame(datos_pdf))
                else: registrar_auditoria(f"AVISO: Recibo {nro_recibo} ({entidad}) sin conceptos relevantes")
        except Exception as e:
            registrar_auditoria(f"ERROR: No se pudo leer {archivo}. Detalle: {e}")
    return lista_dfs

def ejecutar_sistema():
    with open("auditoria_recibos.txt", "w", encoding="utf-8") as f:
        f.write(f"=== AUDITORÍA DE PROCESAMIENTO - {datetime.now().strftime('%d/%m/%Y')} ===\n")
    
    ruta_base = os.path.dirname(os.path.abspath(__file__))
    dir_nom = os.path.join(ruta_base, "recibos")
    dir_deu = os.path.join(ruta_base, "deudas")
    
    entidades_detectadas = set()
    dfs_nom = extraer_datos_de_carpeta(dir_nom, MAPEO_CENTROS, entidades_detectadas, False)
    dfs_deu = extraer_datos_de_carpeta(dir_deu, MAPEO_CENTROS, entidades_detectadas, True)

    faltantes = [e for e in MAPEO_CENTROS if e not in entidades_detectadas]
    if faltantes:
        registrar_auditoria("--- ENTIDADES NO ENCONTRADAS ---")
        for f in faltantes: registrar_auditoria(f"AUSENTE: {f}")
    else:
        registrar_auditoria("PROCESO: Todas las entidades fueron detectadas")

    if not dfs_nom and not dfs_deu: return None

    nombre_hoja = obtener_nombre_hoja_base()
    ruta_excel = os.path.join(ruta_base, "Consolidado_Final.xlsx")
    
    with pd.ExcelWriter(ruta_excel, engine='xlsxwriter') as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet(nombre_hoja)
        fmt_num = workbook.add_format({'num_format': '0.00', 'border': 1})
        fmt_head = workbook.add_format({'bold': True, 'bg_color': '#CFE2F3', 'border': 1})
        fmt_sep = workbook.add_format({'bold': True, 'bg_color': '#FFEB3B', 'align': 'center', 'border': 2})

        fila = 0
        def escribir_bloque(dfs, titulo):
            nonlocal fila
            if not dfs: return
            worksheet.merge_range(fila, 0, fila, 3, titulo, fmt_sep)
            fila += 2
            for df in dfs:
                for c, v in enumerate(df.columns): worksheet.write(fila, c, v, fmt_head)
                for i, row in df.iterrows():
                    worksheet.write(fila+i+1, 0, row['Conceptos'])
                    worksheet.write(fila+i+1, 1, row['Entidad'])
                    worksheet.write(fila+i+1, 2, row['Recibo'])
                    worksheet.write(fila+i+1, 3, row['Montos'], fmt_num)
                fila += len(df) + 4
        
        escribir_bloque(dfs_nom, "--- SECCIÓN NÓMINAS ---")
        escribir_bloque(dfs_deu, "--- SECCIÓN DEUDAS ---")
    
    return ruta_excel, nombre_hoja