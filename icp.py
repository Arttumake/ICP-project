import openpyxl as xl
import os
import glob
from copy import copy
from win32com.client import Dispatch

path = os.getcwd()
xml_files = glob.glob(os.path.join(path, "*.xml"))
excel_name = "Testing"
exclusion_list = ("Blank", "Kalib.blank", "Std1", "Std2", "Std3", "QC", "qc", "Vesi", "vesi") # List of labels to filter out
process_samples = ("AgPbR_", "ZnR_", "SR_", "M_", "J_", "ZnJ_", "AgPbJ_")
courier_samples = (" AgPbR", " ZnR", " SR", " M", " J", " ZnJ", " AgPbJ")
qc = ["Prep", "Pulp", "GBM", "Reag.blank", "GMR", "OREAS", "GSB", "Pyrref", "Pyr-ref", "SR-Ref"]
not_geological = []

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
    wb.create_sheet("Sorted")
    ws2 = wb["Sorted"]

    for row in ws1.iter_rows(min_row = 1, max_row = 2, min_col = 1, max_col = 13):
        for cell in row:
            ws2[f"{cell.coordinate}"] = cell.value

    for row in ws1.iter_rows(min_row = 3, min_col = 1, max_col = 13):   
        for cell in row:
            ws2[f"{cell.coordinate}"] = cell.value
            if cell.has_style:
                    ws2[f"{cell.coordinate}"]._style = copy(cell._style)

    # Remove rows containing a cell containing a string from exclusion list
    for i in reversed(range(2, ws2.max_row+1)):
        if ws2.cell(row= i, column = 2).value in exclusion_list:
            ws2.delete_rows(i)
    
    # Strip g and ml units from a few columns
    for row in ws2.iter_rows(min_row = 3, min_col= 5, max_col = 5):
        for cell in row:
            txt = cell.value
            stripped = txt.strip(" g")
            try:
                cell.value = float(stripped)
            except(ValueError):
                continue
    for row in ws2.iter_rows(min_row = 3, min_col= 6, max_col = 6):
        for cell in row:
            txt = cell.value
            stripped = txt.strip(" ml")
            try:
                cell.value = float(stripped)
            except(ValueError):
                continue

    wb.create_sheet("Process")
    wb.create_sheet("Courier")  
    wb.create_sheet("Sorter")
    wb.create_sheet("QC")  
    wb.create_sheet("Final Sorted")

    ws3 = wb["Process"]
    ws4 = wb["Courier"]
    ws5 = wb["Sorter"]
    ws6 = wb["QC"]
    
    def move_items(from_sheet, to_sheet, tuple):
        # Move data from a sheet to another if column contains string in list
        sheet1_row = 2
        sheet2_row = 0          
        for row in from_sheet.iter_rows(min_row = 3, min_col = 2, max_col = 2): 
            sheet1_row +=1
            for cell in row:
                if any(value in cell.value for value in tuple) and "Pulp".lower() not in cell.value.lower() and "Prep".lower() not in cell.value.lower():
                    not_geological.append(cell.value)
                    for row2 in from_sheet.iter_rows(min_row = sheet1_row, max_row = sheet1_row, min_col = 1):
                        sheet2_row +=1
                        for cell in row2:
                            to_sheet.cell(row=sheet2_row, column = cell.col_idx).value = cell.value
                            if cell.has_style:
                                    to_sheet.cell(row=sheet2_row, column = cell.col_idx)._style = copy(cell._style)   
                                   
    def move_sorters(from_sheet, to_sheet):
        sheet1_row = 2
        sheet2_row = 0
        for row in from_sheet.iter_rows(min_row = 3, min_col = 2, max_col = 2):
            sheet1_row +=1
            for cell in row:
                if "_SO_" in cell.value:
                    not_geological.append(cell.value)
                    for row2 in from_sheet.iter_rows(min_row = sheet1_row, max_row = sheet1_row, min_col = 1):
                        sheet2_row +=1
                        for cell in row2:
                            to_sheet.cell(row=sheet2_row, column = cell.col_idx).value = cell.value
                            if cell.has_style:
                                    to_sheet.cell(row=sheet2_row, column = cell.col_idx)._style = copy(cell._style)
    def move_qc(from_sheet, to_sheet):
        sheet1_row = 2
        sheet2_row = 0
        lower_qc = [value.lower() for value in qc]
        for row in from_sheet.iter_rows(min_row = 3, min_col = 2, max_col = 2):
            sheet1_row +=1
            for cell in row:
                """Checks if cell value has "prep" or "pulp" and if it does, loops through the column again and checks for the non-duplicate value
                and inserts that value before the prep/pulp value to the QC tab
                """
                if any(value in cell.value.lower() for value in lower_qc):
                    if "prep" in cell.value.lower() or "pulp" in cell.value.lower():
                        print(f"duplicate {cell.value}")
                        non_dup = cell.value.split("_", 1)
                        if "pulp" in cell.value.lower():
                            non_dup_split = non_dup[1].split(" pulp")
                            non_dup_final = "_" + non_dup_split[0]
                            print(non_dup_final)                           
                        else:
                            non_dup_split = non_dup[1].split(" prep")
                            non_dup_final = "_" + non_dup_split[0]
                            print(non_dup_final) 
                        this_row = 2    
                        for row in from_sheet.iter_rows(min_row = 3, min_col = 2, max_col = 2):
                            this_row += 1
                            for cell in row:
                                if non_dup_final in cell.value and "prep" not in cell.value.lower() and "pulp" not in cell.value.lower():
                                    print(cell.value)
                                    for row in from_sheet.iter_rows(min_row = this_row, max_row = this_row, min_col = 1):
                                        sheet2_row +=1
                                        for cell in row:
                                            to_sheet.cell(row=sheet2_row, column = cell.col_idx).value = cell.value
                                            if cell.has_style:
                                                    to_sheet.cell(row=sheet2_row, column = cell.col_idx)._style = copy(cell._style)                            
                                    break
                              
                    not_geological.append(cell.value)
                    for row2 in from_sheet.iter_rows(min_row = sheet1_row, max_row = sheet1_row, min_col = 1):
                        sheet2_row +=1
                        for cell in row2:
                            to_sheet.cell(row=sheet2_row, column = cell.col_idx).value = cell.value
                            if cell.has_style:
                                    to_sheet.cell(row=sheet2_row, column = cell.col_idx)._style = copy(cell._style)                        

    move_items(ws2,ws3, process_samples)
    move_items(ws2,ws4, courier_samples)
    move_sorters(ws2, ws5)
    move_qc(ws2, ws6)
   
    wb.save(excel_file)
    wb.close()
