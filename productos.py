from openpyxl import *
from pywinauto import *
from pywinauto.keyboard import send_keys

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
app = Application(backend="uia").start('Stock/Stock.exe')
menuMantenimiento = app.ControlDeDeposito.child_window(title="Mantenimiento", control_type="MenuItem").wrapper_object()
menuMantenimiento.click_input()
send_keys('{DOWN}{ENTER}')


numero = app.ControlDeDeposito.child_window(auto_id="7", control_type="Edit").wrapper_object()
nombre = app.ControlDeDeposito.child_window(auto_id="8", control_type="Edit").wrapper_object()
familia = app.ControlDeDeposito.child_window(auto_id="3", control_type="Edit").wrapper_object()
buscarButton = app.ControlDeDeposito.child_window(title="Buscar", auto_id="2", control_type="Button").wrapper_object()
detalle = app.ControlDeDeposito.child_window(auto_id="6", control_type="Edit").wrapper_object()
costo = app.ControlDeDeposito.child_window(auto_id="9", control_type="Edit").wrapper_object()
guardar = app.ControlDeDeposito.child_window(title="Guardar", auto_id="10", control_type="Button").wrapper_object()
familia.type_keys(productos[20], with_spaces=True)
buscarButton.click_input()
# tests *********************
# appBuscar.BuscarAlgo.print_control_identifiers()
# confirmarFamilia = app.ControlDeDeposito.child_window(class_name="MSFlexGridWndClass").wrapper_object()
# handle = findwindows.find_window(best_match='Buscar algo')
# appBuscar = Application(backend="win32").connect(handle=handle)
# appBuscar.BuscarAlgo.print_control_identifiers()
# appBuscar.BuscarAlgo.click_input()
# # confirmarFamilia.click_input(coords=(500, 100))
# send_keys('%{F4}')
# ****************************
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

