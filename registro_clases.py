from openpyxl import Workbook
from datetime import datetime, date
import os

class RegistroClases:
    def __init__(self, archivo_excel='Registro_clases.xlsx'):
        self.archivo_excel = archivo_excel
        print(f'📁 Inicializando registro en: {self.archivo_excel}')

    def inicializar_archivo(self):
        '''Crear el archivo Excel con encabezados si no existe'''
        if not os.path.exists(self.archivo_excel):
            wb = Workbook()
            ws = wb.active
            ws.title = 'Registro de Clases'

            # Encabezados
            encabezados = ['Nombre de la Clase', 'Fecha', 'Duración (minutos)', 'Valor ($)', 'Total']

            for col, encabezados in enumerate(encabezados,1):
                ws.cell(row=1, column=col, value=encabezados)

            wb.save(self.archivo_excel)
            print('✅ Archivo creado con encabezados')
        else:
            print('✅ El archivo ya existe')

if __name__ == '__main__':
    registro = RegistroClases()
    registro.inicializar_archivo()

