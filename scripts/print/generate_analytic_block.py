import sys
import os
from datetime import datetime
import traceback
import time
import subprocess
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from excel.connect_excel import get_excel

def generate_analytic_block():
    
    #Sheet name in the excel 
    sheetname = "Reporte"
    
    
    #Excel route in files
    excel_route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12.xlsm"
    
    # Get excel connection
    try:
        
        excel, wb, ws = get_excel(sheetname, excel_route)
        
        cells_range = ws.Range("A46:AR50")
        
        cells_range.Copy()
        
        cell_destiny_range = ws.Range("A59")
        
        for cell in cells_range:
            
            if cell.MergeCells:
                # Obtener el área combinada de la celda
                merged_area = cell.MergeArea
                # Calcular la posición relativa dentro del rango de origen
                row_offset = cell.Row - cells_range.Row
                col_offset = cell.Column - cells_range.Column
                # Aplicar la misma combinación en el rango de destino
                dest_cell = cell_destiny_range.Offset(row_offset, col_offset)
                ws.Range(
                    dest_cell, 
                    dest_cell.Offset(merged_area.Rows.Count - 1, merged_area.Columns.Count - 1)
                ).Merge()
        
        cell_destiny_range.PasteSpecial(Paste=-4163)
        
        wb.Save()
        
        wb.Close()
        
    except Exception as e:
        
        print(f"An error occurred: {e}")
        
    finally:
        
        #Close the excel
        if 'wb' in locals():
            
            wb.Close(SaveChanges=False)
            
        if 'excel' in locals():
            
            excel.Quit()
        


#Function to close or kill the excel processes before we start printing data to it
def kill_excel_processes():
    
    """Finish all excel process"""
    
    try:
        
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
                       
                      stdout=subprocess.DEVNULL, 
                      stderr=subprocess.DEVNULL)
    except:
        
        print("Excel processes could not be completed or there were none.")
        
generate_analytic_block()
