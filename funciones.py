# ------------------------------------------------------------------------
# Autor: Fran Hervás Álvarez
# Descripción: Programa para unir automaticamente los checklist deseados.
# ------------------------------------------------------------------------
import os
from win32com import client
import win32com


def excel2pdf(file_location):

    excel = client.Dispatch("Excel.Application")
    excel.Interactive=False
    excel.Visible=False

    workbook= excel.Workbooks.open(file_location)
    output= os.path.splitext(file_location)[0]

    workbook.ExportAsFixedFormat(0,output)
    workbook.Close()

#excel2pdf('C:/Users/Usuario/Desktop/semen/2230002.xlsx')

# excel2pdf('C:/Users/Usuario/Desktop/PDFMERGER/semen/EPB2230002.xlsx') #FUNCIONA

directorioactual=os.getcwd()
print(directorioactual)
excel2pdf(f'{directorioactual}/semen/EPB2230002.xlsx')


