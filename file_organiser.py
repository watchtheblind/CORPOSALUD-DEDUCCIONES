import shutil
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from datetime import datetime
import unicodedata
import concurrent.futures
import os
import re
import fitz  # PyMuPDF
from functools import partial

# ==============================================================================
# MAPA DE CONCEPTOS PARA CASAS COMERCIALES
# ==============================================================================
MAPA_CONCEPTOS_CASAS = {
    "ASUNAJUPENSAPROSO": "ASUNAJUPENSAPROSO",
    "CAHORMINSAS": "CAHORMINSAS",
    "CAJA DE AHORRO": "CAJA DE AHORRO",
    "CAJA DE AHORRO CAEMINSA": "CAEMINSAS",
    "COLEGIO BIOANALISTA": "COLEGIO BIOANALISTA",
    "COLEGIO DE ENFERMERA EDO. MIRANDA": "COLEGIO DE ENFERMERA EDO. MIRANDA",
    "COLEGIO DE ENFERMERAS ESTADO CARABOBO": "COLEGIO DE ENFERMERAS ESTADO CARABOBO",
    "COLEGIO ENFERMERAS": "COLEGIO ENFERMERAS",
    "COLEGIO NUTRIC. Y DIET. VZLA": "COLEGIO NUTRIC. Y DIET. VZLA",
    "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA": "DELEGACIÓN REGIONAL INPRENFERMERA ARAGUA",
    "DESC. DIA(S) NO LABORADO(S)": "DESC. DIA(S) NO LABORADO(S)",
    "FENASISTRASALUD": "FENASISTRASALUD",
    "JUZGADO PRIMERO DE MENORES": "JUZGADO PRIMERO DE MENORES",
    "PENSIÓN DE ALIMENTACIÓN": "PENSIÓN DE ALIMENTACIÓN",
    "PRESTAMO CAJA DE AHORRO": "PRESTAMO CAJA DE AHORRO",
    "SAPTRASEZ": "SAPTRASEZ",
    "SERVICIOS FUNERARIOS CAHORMINSAS": "SERVICIOS FUNERARIOS CAHORMINSAS",
    "SINBOPROENF": "SINBOPROENF",
    "SINDICATO OSBESS ARAGUA": "SINDICATO OSBESS ARAGUA",
    "SINDICATO UNICO DE TRAB.DE LA SALUD Y SUS SIMILARES DEL EDO.ARAGUA": "SINDICATO UNICO DE TRAB.DE LA SALUD Y SUS SIMILARES DEL EDO.ARAGUA",
    "SISTRASALUD": "SISTRASALUD",
    "SITRASSS-MIRANDA": "SITRASSS-MIRANDA",
    "SOCIEDAD ANESTECIOLOGOS": "SOCIEDAD ANESTECIOLOGOS",
    "SUNEP-SAS": "SUNEP-SAS",
    "CUOTA EXTRAORDINARIA SUNEP-SAS": "CUOTA EXTRAORDINARIA SUNEP-SAS",
    "TRIBUNAL (PERMANENTE)": "TRIBUNAL (PERMANENTE)",
    "TRIBUNAL DE PROTECCION DE NIÑOS. NIÑAS Y ADOLESCENTES": "TRIBUNAL DE PROTECCION DE NIÑOS. NIÑAS Y ADOLESCENTES"
}

# ==============================================================================
# CONFIGURACIÓN DE CASOS DE USO
# ==============================================================================
CASOS_DE_USO = {
    "1": {
        "nombre": "Recibos",
        "textos_requeridos": ["TOTAL DE DEDUCCIONES POR CONCEPTO"],
        "unificar_pdf": False,
        "incluir_concepto_en_nombre": False,
        "activo": True
    },
    "2": {
        "nombre": "Aportes Patronales",
        "textos_requeridos": [
            "FONDO DE AHORRO OBLIGATORIO PARA LA VIVIENDA",
            "FONDO PENSIONES",
            "S.S.O.",
            "PERDIDA INVOLUNTARIA",
            "CAHORMINSAS",
            "CAEMINSA"
        ],
        "unificar_pdf": True,
        "incluir_concepto_en_nombre": True,
        "activo": True
    },
    "3": {
        "nombre": "Casas Comerciales",
        "textos_requeridos": list(MAPA_CONCEPTOS_CASAS.keys()),
        "mapa_conceptos": MAPA_CONCEPTOS_CASAS,
        "unificar_pdf": True,
        "incluir_concepto_en_nombre": True,
        "activo": True
    }
}

PALABRA_DEUDA = "DEUDA"


def quitar_acentos(texto: str) -> str:
    """Elimina tildes y caracteres especiales."""
    return "".join(
        c for c in unicodedata.normalize("NFD", texto)
        if unicodedata.category(c) != "Mn"
    )


def sanitizar_nombre_archivo(nombre: str, max_len: int = 90) -> str:
    """Sanea caracteres no válidos y recorta la longitud para evitar errores de ruta en Windows."""
    nombre_limpio = quitar_acentos(nombre)
    nombre_limpio = re.sub(r'[\\/*?:"<>|]', '_', nombre_limpio)
    nombre_limpio = nombre_limpio.strip(". ")
    
    if len(nombre_limpio) > max_len:
        nombre_limpio = nombre_limpio[:max_len].strip(". ")
        
    return nombre_limpio or "ARCHIVO_PROCESADO"


def seleccionar_caso() -> dict:
    """Muestra el menú interactivo en consola."""
    print("\n=========================================")
    print("      SELECCIÓN DE CASO DE USO          ")
    print("=========================================")
    for clave, caso in CASOS_DE_USO.items():
        estado = "" if caso["activo"] else " (Próximamente)"
        print(f" {clave}. {caso['nombre']}{estado}")
    print("=========================================")

    while True:
        opcion = input("Selecciona una opción (1-3): ").strip()
        if opcion in CASOS_DE_USO:
            caso_sel = CASOS_DE_USO[opcion]
            if not caso_sel["activo"]:
                print("⚠️  Esa opción aún no está activa. Elige otra.")
                continue
            return caso_sel
        print("❌ Opción inválida. Intenta de nuevo.")


def seleccionar_origen():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("\nSelecciona la carpeta BASE de origen en la ventana emergente...")
    base_str = filedialog.askdirectory(title="Selecciona la Carpeta BASE de Origen")
    return Path(base_str) if base_str else None


def evaluar_pdf_rapido(
    pdf_path: Path, 
    mapa_busqueda: dict[str, str], 
    deuda_limpia: str,
    evaluar_no_ley: bool = False
) -> tuple[Path, bool, str, list[str], str]:
    """
    Evaluación con PyMuPDF.
    Acepta el PDF si contiene AL MENOS UNO de los conceptos configurados O si aplica regla "NO LEY".
    Retorna: (pdf_path, es_valido, categoria, conceptos_hallados, mensaje_log)
    """
    try:
        # REGISTRO Y REGLA NO LEY
        nombre_norm = quitar_acentos(pdf_path.name.upper())
        tiene_no_ley_en_nombre = evaluar_no_ley and "NO LEY" in nombre_norm

        doc = fitz.open(pdf_path)
        num_paginas = len(doc)
        if num_paginas == 0:
            doc.close()
            return pdf_path, False, "NINGUNA", [], "PDF sin páginas o corrupto."

        tiene_requerido = tiene_no_ley_en_nombre
        tiene_deuda = False
        conceptos_hallados = ["NO LEY"] if tiene_no_ley_en_nombre else []

        # Estrategia de lectura: Pág 1 y última página primero
        indices_a_revisar = [0]
        if num_paginas > 1:
            indices_a_revisar.append(num_paginas - 1)
        
        for i in range(1, num_paginas - 1):
            indices_a_revisar.append(i)

        for index_pag in indices_a_revisar:
            pagina = doc[index_pag]
            texto_raw = pagina.get_text() or ""
            
            if not texto_raw.strip():
                continue

            texto_sin_guiones = texto_raw.replace("-\n", "").replace("- ", "")
            texto_limpio = quitar_acentos(" ".join(texto_sin_guiones.split()).upper())

            # Búsqueda de conceptos
            for texto_normalizado, concepto_salida in mapa_busqueda.items():
                if texto_normalizado in texto_limpio:
                    tiene_requerido = True
                    if concepto_salida not in conceptos_hallados:
                        conceptos_hallados.append(concepto_salida)

            if deuda_limpia in texto_limpio:
                tiene_deuda = True

            # Early Exit: Si encontramos coincidencias requeridas y deuda, terminamos temprano
            if tiene_requerido and tiene_deuda:
                break

        doc.close()

        txt_hallado = ", ".join(conceptos_hallados) if conceptos_hallados else "N/A"

        if tiene_requerido:
            categoria = "DEUDA" if tiene_deuda else "NORMAL"
            mensaje = f"Concepto(s): [{txt_hallado}]" + (f" Y '{PALABRA_DEUDA}'" if tiene_deuda else "")
            return pdf_path, True, categoria, conceptos_hallados, mensaje

        return pdf_path, False, "NINGUNA", [], f"No coincide en sus {num_paginas} páginas."

    except Exception as e:
        return pdf_path, False, "NINGUNA", [], f"Error de lectura: {str(e)}"


def buscar_archivos_en_subcarpeta_paralelo(
    subcarpeta: Path, 
    lineas_log: list[str], 
    mapa_busqueda: dict[str, str], 
    deuda_limpia: str,
    evaluar_no_ley: bool = False
) -> list[tuple[Path, str, list[str]]]:
    
    hijos_dir = [d for d in subcarpeta.iterdir() if d.is_dir()]

    dirs_ord = [d for d in hijos_dir if "ORD" in d.name.upper() and "ORDINARIA" not in d.name.upper()]
    dirs_ordinaria = [d for d in hijos_dir if "ORDINARIA" in d.name.upper()]
    dirs_sueldo = [d for d in hijos_dir if "SUELDO" in d.name.upper()]

    rutas_a_inspeccionar = dirs_ord + dirs_ordinaria + dirs_sueldo + [subcarpeta]

    rutas_unicas = []
    for ruta in rutas_a_inspeccionar:
        if ruta not in rutas_unicas:
            rutas_unicas.append(ruta)

    archivos_pdfs = []
    archivos_revisados = set()

    for ruta in rutas_unicas:
        for pdf in ruta.rglob("*.pdf"):
            if pdf not in archivos_revisados:
                archivos_revisados.add(pdf)
                archivos_pdfs.append(pdf)

    lineas_log.append("\n=========================================")
    lineas_log.append(f"SUBCARPETA EVALUADA: [{subcarpeta.name}] ({len(archivos_pdfs)} PDFs detectados)")
    lineas_log.append(f"Ruta completa: {subcarpeta}")
    lineas_log.append("=========================================")

    if not archivos_pdfs:
        lineas_log.append("  [VACÍO] No se encontraron archivos .pdf.\n")
        return []

    encontrados = []
    max_workers = min(32, (os.cpu_count() or 4) + 4)
    
    funcion_evaluar = partial(
        evaluar_pdf_rapido, 
        mapa_busqueda=mapa_busqueda, 
        deuda_limpia=deuda_limpia,
        evaluar_no_ley=evaluar_no_ley
    )

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        resultados = executor.map(funcion_evaluar, archivos_pdfs)

        for pdf_path, es_valido, categoria, conceptos, detalle in resultados:
            if es_valido:
                encontrados.append((pdf_path, categoria, conceptos))
                lineas_log.append(f"  [ACEPTADO - {categoria}] {pdf_path.name}")
                lineas_log.append(f"             Ruta: {pdf_path}")
                lineas_log.append(f"             Detalle: {detalle}\n")
            else:
                lineas_log.append(f"  [RECHAZADO] {pdf_path.name}")
                lineas_log.append(f"              Ruta: {pdf_path}")
                lineas_log.append(f"              Razón: {detalle}\n")

    return encontrados


def copiar_y_renombrar(pdf_origen: Path, sub_destino: Path, nombre_base: str) -> Path:
    """Copia y renombra un PDF asegurando que no exceda el límite de caracteres de Windows."""
    sub_destino.mkdir(parents=True, exist_ok=True)
    nombre_sano = sanitizar_nombre_archivo(nombre_base)
    destino_final = sub_destino / f"{nombre_sano}.pdf"
    
    contador = 1
    while destino_final.exists():
        contador += 1
        destino_final = sub_destino / f"{nombre_sano}_{contador}.pdf"

    shutil.copy(pdf_origen, destino_final)
    return destino_final


def unificar_pdfs(rutas_pdf: list[Path], ruta_salida: Path):
    """Junta múltiples PDFs en un único archivo consolidado usando PyMuPDF."""
    doc_unificado = fitz.open()
    for pdf_path in rutas_pdf:
        doc_temp = fitz.open(pdf_path)
        doc_unificado.insert_pdf(doc_temp)
        doc_temp.close()
    
    doc_unificado.save(ruta_salida)
    doc_unificado.close()


# --- INICIO DEL PROCESO ---
if __name__ == "__main__":
    caso_seleccionado = seleccionar_caso()
    base = seleccionar_origen()

    if base:
        start_time = datetime.now()
        now_str = start_time.strftime("%Y%m%d_%H%M%S")
        escritorio = Path.home() / "Desktop"

        # Evaluar si el caso seleccionado aplica la excepción "NO LEY"
        es_caso_especial_no_ley = caso_seleccionado["nombre"] in ["Aportes Patronales", "Casas Comerciales"]

        # Construir mapa de búsqueda interna: {TextoNormalizado: ConceptoSalida}
        mapa_busqueda = {}
        if "mapa_conceptos" in caso_seleccionado:
            for k, v in caso_seleccionado["mapa_conceptos"].items():
                if k.strip().lower() != "(en blanco)":
                    key_norm = quitar_acentos(" ".join(k.split()).upper())
                    mapa_busqueda[key_norm] = v
        else:
            for t in caso_seleccionado["textos_requeridos"]:
                key_norm = quitar_acentos(" ".join(t.split()).upper())
                mapa_busqueda[key_norm] = t

        deuda_limpia = quitar_acentos(PALABRA_DEUDA.upper())

        tag_caso = caso_seleccionado["nombre"].replace(" ", "_")
        destino_normal = escritorio / f"Procesados_{tag_caso}_{now_str}"
        destino_deuda = escritorio / f"DEUDAS_{tag_caso}_{now_str}"
        nombre_log = f"log_{tag_caso}_{now_str}.txt"

        lineas_log = [
            "=======================================================",
            f"LOG DE PROCESAMIENTO: {caso_seleccionado['nombre'].upper()}",
            f"Fecha/Hora: {now_str}",
            f"Origen: {base}",
            "=======================================================\n",
        ]

        subcarpetas_origen = [item for item in base.iterdir() if item.is_dir()]

        if not subcarpetas_origen:
            print("No se encontraron subcarpetas en la carpeta base elegida.")
        else:
            print(f"\n🚀 Iniciando procesamiento para: [{caso_seleccionado['nombre']}]")

            for sub_origen in subcarpetas_origen:
                print(f"\n⚡ Evaluando subcarpeta: [{sub_origen.name}]...")

                hallazgos = buscar_archivos_en_subcarpeta_paralelo(
                    sub_origen, 
                    lineas_log, 
                    mapa_busqueda, 
                    deuda_limpia,
                    evaluar_no_ley=es_caso_especial_no_ley
                )

                if hallazgos:
                    copiados_normal = []
                    copiados_deuda = []

                    for pdf, categoria, conceptos in hallazgos:
                        raiz_salida = destino_deuda if categoria == "DEUDA" else destino_normal
                        sub_salida = raiz_salida / sub_origen.name

                        # REGLA "NO LEY": Si aplica y el archivo trae "NO LEY" en el nombre, se mantiene tal cual
                        es_no_ley = es_caso_especial_no_ley and "NO LEY" in quitar_acentos(pdf.name.upper())

                        if es_no_ley:
                            nombre_base = pdf.stem  # Copia tal cual el nombre original sin modificarlo
                        elif caso_seleccionado["incluir_concepto_en_nombre"] and conceptos:
                            if len(conceptos) > 2:
                                concepto_str = f"{conceptos[0]}_Y_{len(conceptos)-1}_MAS"
                            else:
                                concepto_str = "_".join(conceptos)
                            nombre_base = f"{sub_origen.name}_{concepto_str}"
                        else:
                            nombre_base = sub_origen.name

                        final_path = copiar_y_renombrar(pdf, sub_salida, nombre_base)
                        print(f"  ├─ 🟢 [{categoria}] Guardado ➔ {final_path.name}")

                        if categoria == "DEUDA":
                            copiados_deuda.append(final_path)
                        else:
                            copiados_normal.append(final_path)

                    # MERGE / UNIFICACIÓN DE PDFS (Sigue su curso normal e incluye a los "NO LEY")
                    if caso_seleccionado["unificar_pdf"]:
                        if copiados_normal:
                            sub_salida_norm = destino_normal / sub_origen.name
                            nombre_unif = sanitizar_nombre_archivo(f"{sub_origen.name}_UNIFICADO", max_len=60)
                            ruta_unificada = sub_salida_norm / f"{nombre_unif}.pdf"
                            unificar_pdfs(copiados_normal, ruta_unificada)
                            print(f"  ├─ 📦 [MERGE NORMAL] Generado ➔ {ruta_unificada.name}")

                        if copiados_deuda:
                            sub_salida_deuda = destino_deuda / sub_origen.name
                            nombre_unif_d = sanitizar_nombre_archivo(f"{sub_origen.name}_DEUDAS_UNIFICADO", max_len=60)
                            ruta_unificada_d = sub_salida_deuda / f"{nombre_unif_d}.pdf"
                            unificar_pdfs(copiados_deuda, ruta_unificada_d)
                            print(f"  ├─ 📦 [MERGE DEUDA] Generado ➔ {ruta_unificada_d.name}")

                else:
                    print(f"  └─ 🔴 Sin coincidencias en [{sub_origen.name}]")

            # Guardar Logs
            log_contenido = "\n".join(lineas_log)
            if destino_normal.exists():
                with open(destino_normal / nombre_log, "w", encoding="utf-8") as f:
                    f.write(log_contenido)

            if destino_deuda.exists():
                with open(destino_deuda / nombre_log, "w", encoding="utf-8") as f:
                    f.write(log_contenido)

            tiempo_total = (datetime.now() - start_time).total_seconds()

            print("\n=======================================================")
            print(f"✅ ¡Proceso finalizado en {tiempo_total:.2f} segundos!")
            if destino_normal.exists():
                print(f"📁 Salida Normal: {destino_normal}")
            if destino_deuda.exists():
                print(f"💳 Salida Deudas: {destino_deuda}")
            print("=======================================================")