from openpyxl import Workbook, load_workbook
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

    def agregar_clase(self, nombre_clase, fecha, duracion, valor):
        '''Agrega una nueva clase al registro'''
        try:
            # Cargar el archivo existente
            wb = load_workbook(self.archivo_excel)
            ws = wb.active

            # Encontrar la siguiente fila vacía
            siguiente_fila = ws.max_row + 1

            # Insertar datos
            ws.cell(row=siguiente_fila, column=1, value=nombre_clase)
            ws.cell(row=siguiente_fila, column=2, value=fecha)
            ws.cell(row=siguiente_fila, column=3, value=duracion)
            ws.cell(row=siguiente_fila, column=4, value=valor)
            ws.cell(row=siguiente_fila, column=5, value=f'=C{siguiente_fila}*D{siguiente_fila}')

            wb.save(self.archivo_excel)
            print(f"✅ Clase '{nombre_clase}' agregada en fila {siguiente_fila}")

        except Exception as e:
            print(f"❌ Error al agregar la clase: {e}")


if __name__ == '__main__':
    registro = RegistroClases()
    registro.inicializar_archivo()

