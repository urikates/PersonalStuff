######################################################################
##Created by Uriel Villavicencio
##Objective: Extract from pdf file:
###########  RFC EMISOR
###########  NOMBRE O RAZON SOCIAL (EMISOR)
###########  TOTAL
###########  EFECTO DEL COMPROBANTE
###########  Paste this information on xlsx file  
######################################################################


import tkinter as tk
import PyPDF2
           
openpdf = open('CADL-JUN21-G.pdf', 'rb')
pdffile = PyPDF2.PdfReader(openpdf)
# pdffile.getNumPages()
page = pdffile.pages[0]
# print(page.extract_text())

parts = []
def visitor_body(text, cm, tm, fontdict, fontsize):
    y=tm[5]
    if y>10:
        parts.append(text)

page.extract_text(visitor_text=visitor_body)
text ="".join(parts)

print(text)




### GUI
# w=tk.Tk()
# w.geometry('800x400')
# w.title('CFDI to Excel')
# w.config(bg='Snow')
# pdflabel=tk.Label(w, text= 'Seleccionar pdf', font=('Georgia Bold',20), bg='White', fg='Black').pack()
# close=tk.Button(w, text='Salir', command=w.destroy, bg='White', fg='Black', font='Georgia').place(x=25, y=350)
# start=tk.Button(w, text='Iniciar', command=startprocess, bg='White', fg='Black', font='Georgia').place(x=725, y=350)
# selectfile= tk.Button(w, text='Seleccionar archivo', command=savefile, bg='white', fg='black', font='Georgia').place(x=300,y=200)
# pdfsellabel= tk.Label(w, text='PDF no seleccionado', font=('Georgia Bold',15), bg='White', fg='Black'). place(x=25,y=100)

#w.mainloop()



###############################Notes#############################################
#Missing to add PDF Selection from GUI 
#Testing how works PDF extraction, not the final workflow
#Missing to work with Excel
#Needs to delete XML after the process

