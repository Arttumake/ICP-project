import openpyxl as xl
import pandas as pd
import os
import glob
from copy import copy
from win32com.client import Dispatch

path = os.getcwd()
xml_files = glob.glob(os.path.join(path, "*.xml"))
excel_name = "Testing"

def convert_xls_to_xlsx(oldName:str, newName:str):
    oldName = os.path.abspath(oldName)
    newName = os.path.abspath(newName)
    xlApp = Dispatch("Excel.Application")
    wb = xlApp.Workbooks.Open(oldName)
    wb.SaveAs(newName,51)
    wb.Close(True)   

num = 0
for file in xml_files:
    num += 1
    excel_file = f"{excel_name}_{num}.xlsx"
    convert_xls_to_xlsx(file, excel_file)
    wb = xl.load_workbook(excel_file)
    ws1 = wb.active
    ws2 = wb.create_sheet("Sorted")
    for row in ws1.iter_rows(min_row = 0, max_row = 2, min_col = 1, max_col = 13):
        for cell in row:
            print(cell.coordinate)
            ws2[f"{cell.coordinate}"] = cell.value
    for row in ws1.iter_rows(min_row = 3, min_col = 1, max_col = 13):
        for cell in row:
            print(cell.coordinate)
            ws2[f"{cell.coordinate}"] = cell.value
            if cell.has_style:
                    ws2[f"{cell.coordinate}"]._style = copy(cell._style) 
    wb.save(excel_file)



