import tkinter
from tkinter import ttk
import openpyxl
from openpyxl.drawing.image import Image
import os, sys
from win32com import client
from datetime import datetime
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *

now = datetime.now()

#Asignar la fecha en una variable
fecha = now.strftime('%Y-%m-%d')

def resource_path(relative_path):
    #Get absolute path to resource, works for dev and for PyInstaller
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    #ruta relativa = solo una parte de la ruta, se tiene en cuenta el directorio actual desde el que se está trabajando. Donde se va guardar el archivo
    #base_path = directorio actual, es la raiz.
    return os.path.join(base_path, relative_path)

#Seleccioanr la plantilla excel (xlsx)
pathplantilla = resource_path('FORMATO - COTIZACION - PLANTILLA.xlsx')
#Cargar la plantilla
archivo_excel = openpyxl.load_workbook(pathplantilla)
#Seleccionar la hoja del libro excel
hoja_trabajo = archivo_excel['PLANTILLA']
#Seleccionar la imagen de plantilla para pegarlo en el libro excel
pathimg = resource_path('PlantillaCotizacionoriginal_final.png')
#Cargar la imagen y guardarla en una variable
imag = Image(pathimg)


#funcion para usar junto con add_item y new_invoice (Cada que ves que se agrega un item o se crea un nuevo invoice, la cantidad y precio se reestrablecen a 1 y 0,0)
def clear_item():
    qty_spinbox.delete(0, tkinter.END)
    qty_spinbox.insert(0, "1")
    descrip_entry.delete(0, tkinter.END)
    price_spinbox.delete(0, tkinter.END)
    price_spinbox.insert(0, "0.0")
    
invoice_list = []
#Se obtiene los valores de lo que ingresó el usuario y se guarda en invoice_item para insertarlo al Arbol (Tree), y en la invoice_list [].
def add_item():
    qty = int(qty_spinbox.get())
    descrip = descrip_entry.get()
    price = float(price_spinbox.get())
    line_total = qty*price
    invoice_item = [qty, descrip, price, line_total]
    tree.insert('',0, values=invoice_item)
    #clear_item()
    
    invoice_list.append(invoice_item)

#Se elimina todo los datos ingresados por el usuario y tambien los que se guardaron en el arbol (Tree), y el invoice_list [].  
def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    
    invoice_list.clear()

#Funcion para especificar una ruta relativa y guardar la pro-forma generado
def save_file():
    file = filedialog.asksaveasfilename(filetypes=[('text file','*.txt')],defaultextension='.txt')
    
#Funcion para generar una Pro-Forma    
def generate_invoice():
    #get () - conseguir el valor de las entradas de texto 
    name = first_name_entry.get()+last_name_entry.get()
    phone = phone_entry.get()

    hoja_trabajo['C14'] = name
    hoja_trabajo['C15'] = phone
    hoja_trabajo['H7'] = fecha
    # 17 productos como maximo

    #Para añadir cuantas veces se requiera, dependiendo cuantos elementos halla en invoice_list (productos) 
    for i in range(len(invoice_list)):
        num = 20 + i

        hoja_trabajo["C" + str(num)] = invoice_list[i][0]
        hoja_trabajo["D" + str(num)] = invoice_list[i][1]
        hoja_trabajo["G" + str(num)] = invoice_list[i][2]
        hoja_trabajo["H" + str(num)] = invoice_list[i][3]        
    
    #Para añadir la imagen plantilla en "A1" en la hoja excel
    hoja_trabajo.add_image(imag, 'A1')
    #Para asignar el nombre al documento que se va a generar, en el nombre va ir el nombre del cliente y la fecha exacta
    excel_name = resource_path("cotizacion-computexperu" + name + datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".xlsx")
    #Para guardar el archivo excel en una ruta relativa
    archivo_excel.save(excel_name)

    #Obtener la ruta de acceso al directorio actual 
    currentDir = os.getcwd()
    #Crea una nueva instancia para poder modificarlo
    xlApp = client.Dispatch("Excel.Application")
    #Abre la hoja de calculo de Excel
    books = xlApp.Workbooks.Open(os.path.join(currentDir,excel_name))
    #Obtiene la primera hoja de calculo
    ws = books.Worksheets[0]
    #Hace visible la hoja de calculo
    ws.Visible = 1
    #Exporta la hoja de calculo a un archivo PDF
    filename = 1
    #ASK DIRECTORY FALTA AGREGAR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    #ASK DIRECTORY FALTA AGREGAR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    #ASK DIRECTORY FALTA AGREGAR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    #ASK DIRECTORY FALTA AGREGAR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!    
    ws.ExportAsFixedFormat(0,resource_path(os.path.join(currentDir,"cotizacion-computexperu" + name + datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".pdf")))
    ws.ExportAsFixedFormat(0,os.path.join(currentDir, filename + "cotizacion-computexperu" + name + datetime.now().strftime("%Y-%m-%d-%H%M%S") + ".pdf"))
    #Cierra el libro de Excel
    books.Close()
    
    #Abre una ventana con un mensaje
    messagebox.showinfo("Pro-Forma ComputexPerú", "Pro-Forma Completado exitosamente")
    
    #Se ejecuta la funcion new_invoice para borrar todo lo ingresado anteriormente y poder ingresar nuevos productos, y generar otra pro-forma
    new_invoice()

#Abre una ventana
window = tkinter.Tk()
#Asigna un titulo para la ventana
window.title("Invoice Generator Form")

#Selecciona el logo que va llevar la ventana
path = resource_path("logo_computex.png")
icono = tkinter.PhotoImage(file=path)
window.iconphoto(True, icono)

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

first_name_label = tkinter.Label(frame, text="Nombres")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(frame, text="Apellidos")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

phone_label = tkinter.Label(frame, text="Celular")
phone_label.grid(row=0, column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1, column=2)

qty_label = tkinter.Label(frame, text="Cantidad de Productos")
qty_label.grid(row=2, column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

#----------------------------------

descrip_label = tkinter.Label(frame, text="Descripcion")
descrip_label.grid(row=2, column=1)
descrip_entry = tkinter.Entry(frame)
descrip_entry.grid(row=3, column=1)

price_label = tkinter.Label(frame, text="Precio Unitario")
price_label.grid(row=2, column=2)
price_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment=0.5)
price_spinbox.grid(row=3, column=2)

add_item_button = tkinter.Button(frame, text = "Añadir Item", command = add_item)
add_item_button.grid(row=4, column=2, pady=5)

columns = ('cant', 'descrip', 'precio', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('cant', text='Cantidad')
tree.heading('descrip', text='Descripcion')
tree.heading('precio', text='Precio Unitario')
tree.heading('total', text="Total")

    
tree.grid(row=5, column=0, columnspan=3, padx=20, pady=10)


save_invoice_button = tkinter.Button(frame, text="Generar Pro-Forma", command=generate_invoice)
save_invoice_button.grid(row=6, column=0, columnspan=3, sticky="news", padx=20, pady=5)
new_invoice_button = tkinter.Button(frame, text="Nuevo Pro-Forma", command=new_invoice)
new_invoice_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=20, pady=5)


window.mainloop()