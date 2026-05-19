import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

class ReporteExcel:
    def __init__(self, nombre_archivo="Auditoria_Nomina.xlsx", columnas=None):
        self.nombre_archivo = nombre_archivo
        self.columnas = columnas or ["CENTRO", "GREMIO", "TRABAJADOR", "EMPRESA", "CONCEPTO", "NRO_RECIBO"]
        self.wb = openpyxl.Workbook()
        
        # Pestaña 1: Base de Datos
        self.ws_base = self.wb.active
        self.ws_base.title = "Base de Datos"
        
        # Pestaña 2: Cuadre y Verificación
        self.ws_cuadre = self.wb.create_sheet("Cuadre y Verificación")

    def generar_reporte(self, datos, totales_pdf_dict=None):
        if not datos: return False
        
        self.ws_base.append(self.columnas)
        for fila in datos:
            # Formateamos el número para que Excel lo vea con coma
            row = [
                fila["CENTRO"], fila["GREMIO"], 
                fila["TRABAJADOR"], fila["EMPRESA"], 
                fila["CONCEPTO"], fila["NRO_RECIBO"]
            ]
            self.ws_base.append(row)
            
        self._configurar_cuadre(datos, totales_pdf_dict)
        self.wb.save(self.nombre_archivo)
        return True
    
    def _configurar_cuadre(self, datos, totales_pdf_dict):
        headers_cuadre = ["CONCEPTO", "TIPO", "TOTAL PDF", "TOTAL EXCEL", "DIFERENCIA", "ESTADO"]
        self.ws_cuadre.append(headers_cuadre)
        
        conceptos = sorted(list(set(d["CONCEPTO"] for d in datos)))
        row_idx = 2

        for concepto in conceptos:
            for tipo in ["TRABAJADOR", "EMPRESA"]:
                # Buscamos el valor real que venía en el PDF
                valor_pdf = 0
                if totales_pdf_dict and concepto in totales_pdf_dict:
                    valor_pdf = totales_pdf_dict[concepto][tipo]

                col_suma = "C" if tipo == "TRABAJADOR" else "D"
                f_excel = f"=SUMIF('Base de Datos'!E:E, A{row_idx}, 'Base de Datos'!{col_suma}:{col_suma})"
                f_dif = f"=C{row_idx}-D{row_idx}"
                f_estado = f'=IF(ABS(E{row_idx})<0.01, "CUADRA", "REVISAR")'

                self.ws_cuadre.append([concepto, tipo, valor_pdf, f_excel, f_dif, f_estado])
                
                # Aplicamos formato numérico con coma a las celdas C y D
                self.ws_cuadre[f"C{row_idx}"].number_format = '#,##0.00'
                self.ws_cuadre[f"D{row_idx}"].number_format = '#,##0.00'
                row_idx += 1