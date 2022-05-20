from openpyxl import *
from pywinauto import *
from pywinauto.keyboard import send_keys

wb = load_workbook('productos.xlsx')
print(wb.sheetnames)
print(wb.active)
wb.active = wb['Familias']
ws = wb.active

cell_range = ws['A2':'A5']
print(cell_range)
cod_familia = []
nom_familia = []
for row in ws.iter_cols(min_col=1, max_col=1, max_row=5, min_row=2, values_only=True):
    for cell in row:
        cod_familia.append(cell)
        print(cell)

for row in ws.iter_cols(min_col=2, max_col=2, max_row=5, min_row=2, values_only=True):
    for cell in row:
        nom_familia.append(cell)
        print(cell)
#r
print(cod_familia)
print(nom_familia)

app = Application(backend="uia").start('Stock/Stock.exe')
# .connect(title='Control de deposito',timeout=10)
dialog = app['Stock']
# dlg = app.top_window()
var = app['dialog']['Control de deposito']
# app.ControlDeDeposito.print_control_identifiers()
menuMantenimiento = app.ControlDeDeposito.child_window(title="Mantenimiento", control_type="MenuItem").wrapper_object()
menuMantenimiento.click_input()
send_keys('{DOWN}{DOWN}{ENTER}')
app.ControlDeDeposito.print_control_identifiers()

codigo = app.ControlDeDeposito.child_window(auto_id="4", control_type="Edit").wrapper_object()
nombreFamilia = app.ControlDeDeposito.child_window(auto_id="5", control_type="Edit").wrapper_object()
guardar = app.ControlDeDeposito.child_window(title="Guardar", auto_id="3", control_type="Button").wrapper_object()

index = 0
while index < len(cod_familia):
    codigo.type_keys(cod_familia[index])
    nombreFamilia.type_keys(nom_familia[index], with_spaces=True)
    guardar.click_input()
    index += 1
    codigo.set_text('')
    nombreFamilia.set_text('')

salir = app.ControlDeDeposito.child_window(title="Salir", auto_id="2", control_type="Button").wrapper_object()
salir.click_input()
# familias.click_input()


# agregar articulos
wb = load_workbook('productos.xlsx')
wb.active = wb['Productos']
ws = wb.active
print(wb.active)
productos = []
for row in ws.iter_cols(min_row=2, values_only=True):
    for cell in row:
        productos.append(cell)
        print(cell)

print(productos, len(productos))

menuMantenimiento = app.ControlDeDeposito.child_window(title="Mantenimiento", control_type="MenuItem").wrapper_object()
menuMantenimiento.click_input()
# app.ControlDeDeposito.print_control_identifiers()
send_keys('{DOWN}{ENTER}')
app.ControlDeDeposito.print_control_identifiers()

numero = app.ControlDeDeposito.child_window(auto_id="7", control_type="Edit").wrapper_object()
nombre = app.ControlDeDeposito.child_window(auto_id="8", control_type="Edit").wrapper_object()
familia = app.ControlDeDeposito.child_window(auto_id="3", control_type="Edit").wrapper_object()
buscarButton = app.ControlDeDeposito.child_window(title="Buscar", auto_id="2", control_type="Button").wrapper_object()
detalle = app.ControlDeDeposito.child_window(auto_id="6", control_type="Edit").wrapper_object()
costo = app.ControlDeDeposito.child_window(auto_id="9", control_type="Edit").wrapper_object()
guardar = app.ControlDeDeposito.child_window(title="Guardar", auto_id="10", control_type="Button").wrapper_object()
# familia.type_keys(productos[20], with_spaces=True)
# buscarButton.click_input()
# appBuscar.BuscarAlgo.print_control_identifiers()
confirmarFamilia = app.ControlDeDeposito.child_window(class_name="ThunderRT6FormDC").wrapper_object()
index = 0
while index <= 9:
    numero.type_keys(productos[index])
    nombre.type_keys(productos[index + 10], with_spaces=True)
    detalle.type_keys(productos[index + 30], with_spaces=True)
    costo.type_keys(productos[index + 40])
    familia.type_keys(productos[index + 20], with_spaces=True)
    buscarButton.click_input()
    handle = findwindows.find_window(best_match='Buscar algo')
    appBuscar = Application(backend="win32").connect(handle=handle)
    appBuscar.BuscarAlgo.click_input()
    send_keys('%{F4}')
    guardar.click_input()
    index += 1
    numero.set_text('')
    nombre.set_text('')
    familia.set_text('')
    detalle.set_text('')
    costo.set_text('')

