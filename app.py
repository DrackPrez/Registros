import datetime
import openpyxl
from openpyxl import Workbook
from flask import Flask, request, render_template, redirect, url_for

app = Flask(__name__)

# Intentar cargar el archivo Excel existente o crear uno nuevo si no existe
try:
    wb = openpyxl.load_workbook('registro_asistencia.xlsx')
    sheet = wb.active
except FileNotFoundError:
    wb = Workbook()
    sheet = wb.active
    sheet.title = 'Registro'
    sheet.append(['Fecha', 'Hora', 'Tipo', 'Descripción'])  # Encabezados

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/registrar', methods=['POST'])
def registrar():
    tipo = request.form['tipo']
    descripcion = request.form.get('descripcion', '')
    ahora = datetime.datetime.now()
    fecha = ahora.strftime("%Y-%m-%d")  # Formato: AAAA-MM-DD
    hora = ahora.strftime("%H:%M:%S")   # Formato: HH:MM:SS
    sheet.append([fecha, hora, tipo, descripcion])
    wb.save('registro_asistencia.xlsx')  # Guardar inmediatamente
    return redirect(url_for('index') + '?mensaje=Registro guardado correctamente.')

if __name__ == '__main__':
    app.run(debug=True)