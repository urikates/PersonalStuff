######################################################################
##Created by Uriel Villavicencio
##Objective: Extract from pdf file:
###########  RFC EMISOR
###########  NOMBRE O RAZON SOCIAL (EMISOR)
###########  TOTAL
###########  EFECTO DEL COMPROBANTE
######################################################################


import tkinter as tk
from tkinter import filedialog
from unittest import skip
from xml.dom.xmlbuilder import Options
import pandas as pds
import pdfquery as pdfq



#tk.filedialog.SaveAs(master=None, )
def startprocess():
    ### PDF Convertion 
    pdf=pdfq.PDFQuery('CADL-JUN21-G.pdf')
    pdf.load()
    pdf.tree.write('CADL-JUN21-G.xml', pretty_print=True)
    pdf
    
    ###Extracting Information
    pages = pdf.doc.catalog['Pages'].resolve()['Count']
    pagenumber=pdf.doc.catalog['Pages'].resolve()['Count']
    while pagenumber < pages:
            skip ##Just to avoid errors 
           


### GUI
w=tk.Tk()
w.geometry('800x400')
w.title('CFDI to Excel')
w.config(bg='Snow')
pdflabel=tk.Label(w, text= 'Seleccionar pdf', font=('Georgia Bold',20), bg='White', fg='Black').pack()
close=tk.Button(w, text='Salir', command=w.destroy, bg='White', fg='Black', font='Georgia').place(x=25, y=350)
start=tk.Button(w, text='Iniciar', command=startprocess, bg='White', fg='Black', font='Georgia').place(x=725, y=350)

w.mainloop()



###############################Notes#############################################
#Missing to add PDF Selection from GUI 
#Testing how works PDF extraction, not the final workflow
#Missing to work with Excel
#Needs to delete XML after the process

