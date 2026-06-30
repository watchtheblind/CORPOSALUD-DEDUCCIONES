import extractor_nomina
import procesador_reportes
import os
from tkinter import messagebox

if __name__ == "__main__":
    try:
        resultado = extractor_nomina.ejecutar_sistema()
        if resultado:
            ruta, hoja_base = resultado
            procesador_reportes.generar_reportes_consolidados(ruta, hoja_base)
            
            os.startfile(ruta)
            if os.path.exists("auditoria_recibos.txt"):
                os.startfile("auditoria_recibos.txt")
                
            messagebox.showinfo("Proceso Exitoso", "Archivos generados correctamente")
    except Exception as e:
        messagebox.showerror("Error", f"Fallo en ejecucion: {e}")