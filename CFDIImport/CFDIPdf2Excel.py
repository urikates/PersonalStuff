######################################################################
##Created by Uriel Villavicencio
##Objective: Extract from pdf file:
###########  RFC EMISOR 
###########  NOMBRE O RAZON SOCIAL (EMISOR)
###########  TOTAL
###########  EFECTO DEL COMPROBANTE
###########  Paste this information on xlsx file  
######################################################################

import pandas
import pdfquery
import xlsxwriter #Documentation:https://xlsxwriter.readthedocs.io/index.html
import os
import tkinter as tk 

#output for debugging 
#debugoutput = open('splited info.txt', 'w')

#Load PDF File
pathfile = '/Users/urielvillavicencio/Documents/Repos/PersonalStuff/CFDIImport/gulk-ene25-gastos.pdf'
pdffile = pdfquery.PDFQuery(pathfile)
pdffile.load()

#Convert pdf file to XML 
pdffile.tree.write('processfile.xml', pretty_print = True)

#Extracting info from xml file 
with open('processfile.xml', 'r') as xmlfile:
    #List with rfc emisor data
    rfcemisorlist = []
    #List with razon social emisor data 
    razonsociallist = []
    #List with Total amount data
    totalamountlist = []
    #List with EfectoComprobante data
    efectocomprobantelist = []
    #List of lines after filter 
    filtereddata=[]
    
    #Separating elements of each line 
    # for line in enumerate(xmlfile):
    for i,line in enumerate(xmlfile):
        filtereddata.append(line.split(">")) 
    
    #Earning each element of the row after filter
    for row in filtereddata:
        for element in row:
            #Extraction of 'RFC Emisor' data
            if "RFC Emisor: " in element:
                rfcemisorlist.append(element.split()[2])
                #print(element.split()[2]) #debug
        
         #Extraction of "Nombre o Razon Social (Emisor) data
            #if "Nombre o Raz&#243;n Social:" in element and 'LETICIA CALDERON DIAZ' not in element and 'CALDERON DIAZ LETICIA' not in element and 'Leticia Calderon Diaz' not in element and 'Calderon Diaz Leticia' not in element:
            if "Nombre o Raz&#243;n Social: " in element and 'KARINA GUERRERO LUNA' not in element:
            #if "Nombre o Raz&#243;n Social: " in element and 'CARLOS SANCHEZ GONZALEZ' not in element: 
            # if  "Nombre o Raz&#243;n Social: " in element and 'JAIME SANCHEZ RAMOS' not in element:
                razonsociallist.append((element.split(":")[1]).split("<")[0])
                #print((element.split(":")[1]).split("<")[0]) #debug
        
         #Extraction of "Total" amount data
            if "Total: " in element:
                totalamountlist.append((element.split(":")[1]).split("<")[0])
                #print((element.split(":")[1]).split("<")[0]) #debug
         
         #Extraction of "Efecto del comprobante" data       
            if "Efecto del Comprobante: " in element:
                efectocomprobantelist.append((element.split(":")[-1]).split("<")[0])
                #print((element.split(":")[-1]).split("<")[0]) #debug


#Start working with excel file 
#Creation of excel file
excelfilename = 'gulk-ene25.xlsx'
excelfile = xlsxwriter.Workbook(excelfilename)

#Adding sheet based on document
#Name of the sheet 
sheetname = 'Enero 2025'
worksheet = excelfile.add_worksheet(sheetname)

#Writting titles
worksheet.write('A1', 'RFC Emisor')
worksheet.write('B1', 'Nombre o Razon Social')
worksheet.write('C1', 'Monto Total')
worksheet.write('D1', 'Efecto del Comprobante')

#Start transfer of lists 
#Indexes for each 
rfcidx = 2
nombreidx = 2
montoidx = 2
efectoidx = 2

#Transfer of RFC
for rfc in rfcemisorlist:
    cell = 'A' + str(rfcidx) 
    worksheet.write(cell, rfc)
    rfcidx = rfcidx + 1

#Transfer nombre o razon social 
for nombre in razonsociallist:
    if nombre != ' ' :
        cell = 'B' + str(nombreidx)
        worksheet.write(cell, nombre)
        nombreidx = nombreidx + 1

#Transfer monto total
for monto in totalamountlist:
    cell = 'C' + str(montoidx)
    worksheet.write(cell , monto)
    montoidx = montoidx + 1

#Transfer Efecto de comprobante 
for efecto in efectocomprobantelist:
    cell = 'D' +str(efectoidx)
    worksheet.write(cell , efecto)
    efectoidx = efectoidx + 1


os.remove('processfile.xml')
excelfile.close()
            
