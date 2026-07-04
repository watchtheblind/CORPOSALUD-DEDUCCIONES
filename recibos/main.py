import extractor_nomina
import procesador_reportes
import os
from tkinter import messagebox, filedialog

if __name__ == "__main__":
    try:
        dir_nom = filedialog.askdirectory(title="Seleccione la carpeta de RECIBOS (nóminas)")
        if not dir_nom:
            exit()

        dir_deu = filedialog.askdirectory(title="Seleccione la carpeta de DEUDAS")
        if not dir_deu:
            exit()

        resultado = extractor_nomina.ejecutar_sistema(dir_nom, dir_deu)
        if resultado:
            ruta, hoja_base = resultado
            procesador_reportes.generar_reportes_consolidados(ruta, hoja_base)

            os.startfile(ruta)
            if extractor_nomina.RUTA_AUDITORIA and os.path.exists(extractor_nomina.RUTA_AUDITORIA):
                os.startfile(extractor_nomina.RUTA_AUDITORIA)

            messagebox.showinfo("Proceso Exitoso", "Archivos generados correctamente")
    except Exception as e:
        messagebox.showerror("Error", f"Fallo en ejecucion: {e}")