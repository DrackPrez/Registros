import datetime
import openpyxl
from openpyxl import Workbook

# Intentar cargar el archivo Excel existente o crear uno nuevo si no existe
try:
    wb = openpyxl.load_workbook('registro_asistencia.xlsx')
    sheet = wb.active
except FileNotFoundError:
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'Registro'
    sheet.append(['Fecha', 'Hora', 'Tipo', 'Descripción'])  # Encabezados

# Menú principal en un bucle
while True:
    print("\nMenú:")
    print("1. Marcar entrada")
    print("2. Marcar salida")
    print("3. Salir")
    opcion = input("Elija una opción: ")

    if opcion == '1':
        # Registrar entrada
        ahora = datetime.datetime.now()
        fecha = ahora.strftime("%Y-%m-%d")  # Formato: AAAA-MM-DD
        hora = ahora.strftime("%H:%M:%S")   # Formato: HH:MM:SS
        descripcion = input("Ingrese una descripción (opcional): ")
        sheet.append([fecha, hora, 'Entrada', descripcion])
        wb.save('registro_asistencia.xlsx')  # Guardar inmediatamente
        print("Entrada registrada y guardada.")

    elif opcion == '2':
        # Registrar salida
        ahora = datetime.datetime.now()
        fecha = ahora.strftime("%Y-%m-%d")
        hora = ahora.strftime("%H:%M:%S")
        sheet.append([fecha, hora, 'Salida', ''])
        wb.save('registro_asistencia.xlsx')  # Guardar inmediatamente
        print("Salida registrada y guardada.")

    elif opcion == '3':
        # Salir del programa
        print("Saliendo del programa. Todos los registros ya están guardados.")
        break

    else:
        print("Opción no válida. Por favor, elija 1, 2 o 3.")