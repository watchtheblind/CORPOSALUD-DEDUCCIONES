import openpyxl

class ExcelWorker:
    def __init__(self, nombre_archivo="Auditoria_Nomina.xlsx"):
        self.nombre_archivo = nombre_archivo
        # Columnas extendidas para auditoría
        self.columnas = ["CENTRO", "GREMIO", "TRABAJADOR", "EMPRESA", "CONCEPTO", "NRO_RECIBO", "TRAB_REF_PDF", "EMPR_REF_PDF"]
        self.wb = openpyxl.Workbook()
        self.ws_base = self.wb.active
        self.ws_base.title = "BASE"

    def generar_reporte(self, datos):
        if not datos: return False
        self.ws_base.append(self.columnas)
        for fila in datos:
            row = [
                fila["CENTRO"], fila["GREMIO"], fila["TRABAJADOR"], fila["EMPRESA"], 
                fila["CONCEPTO"], fila["NRO_RECIBO"], fila["TRAB_ORIGINAL"], fila["EMPR_ORIGINAL"]
            ]
            self.ws_base.append(row)
            # Formato contable para Excel
            idx = self.ws_base.max_row
            self.ws_base[f"C{idx}"].number_format = '#,##0.00'
            self.ws_base[f"D{idx}"].number_format = '#,##0.00'
            
        self.wb.save(self.nombre_archivo)
        return True
    