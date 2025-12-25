######################################################################
##Created by Uriel Villavicencio
##Objective: Extract from pdf file:
###########  FOLIO FISCAL 
###########  FECHA DE EMISION
###########  TOTAL
###########  ESTADO DEL COMPROBANTE
###########  Paste this information on xlsx file  
######################################################################

import pdfquery
import xlsxwriter
import os
import tkinter as tk
from tkinter import filedialog

# Crear una ventana para seleccionar el archivo
def seleccionar_archivo():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    archivo_seleccionado = filedialog.askopenfilename(
        title="Seleccionar archivo PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    return archivo_seleccionado

# Seleccionar el archivo PDF
pathfile = seleccionar_archivo()
if not pathfile:
    print("No se seleccionó ningún archivo. Saliendo...")
    exit()

# Cargar el archivo PDF
pdffile = pdfquery.PDFQuery(pathfile)
pdffile.load()

# Convertir el archivo PDF a XML
pdffile.tree.write('processfile.xml', pretty_print=True)

# Extraer información del archivo XML
with open('processfile.xml', 'r') as xmlfile:
    # Listas para almacenar los datos extraídos
    foliofiscallist = []
    fechaemisionlist = []
    totalamountlist = []
    estadocomprobantelist = []
    filtereddata = []

    # Separar elementos de cada línea
    for i, line in enumerate(xmlfile):
        filtereddata.append(line.split(">"))
        #print(filtereddata) # debug

    # Procesar cada elemento de las filas filtradas
    for row in filtereddata:
        for element in row:
            # Extracción de 'Folio Fiscal'
            if "Folio Fiscal: " in element:
                foliofiscallist.append(element.split(":")[1].split("<")[0])

            # Extracción de 'Fecha de Emisión'
            if "Fecha Emisi&#243;n: " in element:
                fechaemisionlist.append(element.split(":")[1].split("T")[0].split("<")[0])
                print(element) #debug
                # print((element.split(":")[1]).split("<")[0]) #debug                


            # Extracción de 'Total'
            if "Total: " in element:
                totalamountlist.append(element.split(":")[1].split("<")[0])

            # Extracción de 'Estado del Comprobante'
            if "Estado del Comprobante: " in element:
                estadocomprobantelist.append((element.split(":")[1]).split(">")[0].split()[0])
                #print((element.split(":")[1]).split(">")[0].split()[0]) #debug


# Crear el archivo Excel
excelfilename = 'output.xlsx'
excelfile = xlsxwriter.Workbook(excelfilename)

# Agregar hoja al archivo Excel
sheetname = 'Resultados'
worksheet = excelfile.add_worksheet(sheetname)

# Escribir títulos
worksheet.write('A1', 'Folio Fiscal')
worksheet.write('B1', 'Fecha de Emisión')
worksheet.write('C1', 'Monto Total')
worksheet.write('D1', 'Estado del Comprobante')

# Transferir datos a la hoja de Excel
for idx, folio in enumerate(foliofiscallist, start=2):
    worksheet.write(f'A{idx}', folio)

for idx, fecha in enumerate(fechaemisionlist, start=2):
    worksheet.write(f'B{idx}', fecha)

for idx, monto in enumerate(totalamountlist, start=2):
    worksheet.write(f'C{idx}', monto)

for idx, estado in enumerate(estadocomprobantelist, start=2):
    worksheet.write(f'D{idx}', estado)

# Eliminar el archivo XML temporal
os.remove('processfile.xml')

# Cerrar el archivo Excel
excelfile.close()

print(f"Datos extraídos y guardados en {excelfilename}")