import sys, os
import win32com.client
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


#converter xls in xlsx
directory = 'C:\\GitHub\\ITC Upgrade Test\\Tables ITC (input)\\xls\\'
for file in os.listdir(directory):
    dot = file.find('.')
    end = file[dot:]
    name = file[:dot]
    OutFile ="C:\\GitHub\\ITC Upgrade Test\\Tables ITC (input)\\xlsx\\"+name+".xlsx"
    App = win32com.client.Dispatch("Excel.Application")
    App.Visible = True
    print('Converting: ',name)
    workbook= App.Workbooks.Open(directory+file)
    workbook.ActiveSheet.SaveAs(OutFile, 51)   #51 is for xlsx 
    workbook.Close(SaveChanges=True)
    App.Quit()
