import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber


def seleccionar_directorio():
    """Abre un diálogo para seleccionar la carpeta contenedora de los PDFs."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    directorio = filedialog.askdirectory(
        title="Selecciona la carpeta con los archivos PDF"
    )
    root.destroy()
    return directorio


def seleccionar_archivo_salida():
    """Abre un diálogo para guardar el archivo TXT resultante."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    archivo_salida = filedialog.asksaveasfilename(
        title="Selecciona dónde guardar el archivo de resultados TXT",
        defaultextension=".txt",
        filetypes=[
            ("Archivos de texto", "*.txt"),
            ("Todos los archivos", "*.*"),
        ],
        initialfile="resultado_recibos.txt",
    )
    root.destroy()
    return archivo_salida


def procesar_recibos():
    print("Por favor, selecciona la carpeta con los PDFs...")
    directorio_pdfs = seleccionar_directorio()

    if not directorio_pdfs:
        print("Operación cancelada: No se seleccionó ninguna carpeta.")
        return

    print("Por favor, selecciona el destino del archivo TXT de salida...")
    archivo_salida = seleccionar_archivo_salida()

    if not archivo_salida:
        print("Operación cancelada: No se seleccionó la ruta de salida.")
        return

    # Expresión regular ajustada para capturar 'Recibo', tolerando símbolos de número/formato LaTeX/unicode y espacios
    patron_recibo = re.compile(
        r"Recibo\s*(?:[Nn][°ºo]|N\^o|\$N\^\{o\}\$)?\s*(\d+)", re.IGNORECASE
    )

    resultados = []
    archivos_procesados = 0

    # Recorrer todos los archivos en el directorio seleccionado
    for archivo in sorted(os.listdir(directorio_pdfs)):
        if archivo.lower().endswith(".pdf"):
            archivos_procesados += 1
            ruta_pdf = os.path.join(directorio_pdfs, archivo)

            try:
                with pdfplumber.open(ruta_pdf) as pdf:
                    if len(pdf.pages) > 0:
                        # Extraer texto exacto de la primera página
                        texto_pagina_1 = pdf.pages[0].extract_text() or ""

                        # Buscar la coincidencia
                        coincidencia = patron_recibo.search(texto_pagina_1)

                        nombre_sin_ext = os.path.splitext(archivo)[0]

                        if coincidencia:
                            num_recibo = coincidencia.group(1)
                            resultados.append(
                                f"{nombre_sin_ext} = {num_recibo}"
                            )
                        else:
                            # Búsqueda de respaldo si varían los símbolos entre la palabra 'Recibo' y los dígitos
                            coincidencia_fallback = re.search(
                                r"Recibo[^\n\r]*?(\d{5,8})",
                                texto_pagina_1,
                                re.IGNORECASE,
                            )
                            if coincidencia_fallback:
                                resultados.append(
                                    f"{nombre_sin_ext} = {coincidencia_fallback.group(1)}"
                                )
                            else:
                                resultados.append(
                                    f"{nombre_sin_ext} = NO ENCONTRADO"
                                )
            except Exception as e:
                print(f"Error leyendo el archivo {archivo}: {e}")

    if archivos_procesados == 0:
        print("No se encontraron archivos PDF en la carpeta seleccionada.")
        return

    # Guardar los resultados en el archivo TXT seleccionado
    with open(archivo_salida, "w", encoding="utf-8") as f:
        for linea in resultados:
            f.write(linea + "\n")

    mensaje = f"¡Proceso completado con éxito!\n\nSe procesaron {archivos_procesados} archivo(s) PDF.\nGuardado en: {archivo_salida}"
    print(mensaje)

    # Mostrar mensaje modal de confirmación
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("Proceso Finalizado", mensaje)
    root.destroy()


if __name__ == "__main__":
    procesar_recibos()