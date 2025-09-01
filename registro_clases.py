from openpyxl import Workbook, load_workbook
from datetime import datetime, date
import os

class RegistroClases:
    def __init__(self, archivo_excel='Registro_clases.xlsx'):
        self.archivo_excel = archivo_excel
        self.inicializar_archivo()

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

    def mostrar_registro(self):
        '''Muestra todas las clases registradas'''
        try:
            wb = load_workbook(self.archivo_excel)
            ws = wb.active

            print("\n📋 REGISTRO DE CLASES:")
            print("-" * 50)

            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]: # Si la celda de nombre no está vacía
                    print(f'🏫 : {row[0]}')
                    print(f'📅: {row[1]}, ⏰: {row[2]}Min | 💰: ${row[3]}/h')
                    print(f'💵 Total: ${row[4] if row[4] else 'Calculando...'}')
                    print('-' * 30)

        except Exception as e:
            print(f"❌ Error al leer el registro: {e}")


def main():
    registro = RegistroClases()

    while True:
        print("\n" + "="*50)
        print("       SISTEMA DE REGISTRO DE CLASES")
        print("="*50)
        print("1. Agregar nueva clase")
        print("2. Ver registro completo")
        print("3. Salir")

        opcion = input('\nSelecciona una opción (1-3): ')
        if opcion == '1':
            print("\n➕ NUEVA CLASE")
            nombre = input("Nombre de la clase: ")
            fecha = input("Fecha (DD--MM--YYYY): ")
            duracion = float(input("Duración en minutos: "))
            valor = float(input("Valor por hora: $"))

            if registro.agregar_clase(nombre, fecha, duracion, valor):
                print("✅ Clase agregada exitosamente!")
            else:
                print("❌ Error al agregar la clase")

        elif opcion == '2':
            registro.mostrar_registro()

        elif opcion == '3':
            print("👋 ¡Hasta luego!")
            break

        else:
            print("❌ Opción no válida")

if __name__ == '__main__':
    main()
    #registro.inicializar_archivo()

