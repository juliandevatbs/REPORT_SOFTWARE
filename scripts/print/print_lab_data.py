from datetime import datetime
import sys
import os
import traceback
import time
import subprocess
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from excel.connect_excel import get_excel
from get.get_lab_data import get_lab_data


"""
    This function print into the Reporte sheet of the excel the data extracted
    from the Chain of Custody 1 sheet 
"""

def print_lab_data():
    
    excel= None
    
    wb= None
    
    try:
        
        #First step, close excel
        kill_excel_processes()
        
        #Sheet name to print the data
        sheet_name= "Reporte"
        
        
        #Route of the excel
        route= r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (3).xlsm"
        
        #Get the excel connection
        excel, wb, ws = get_excel(sheet_name, route)
        
        #Verify if the connection excel returns the correct object
        if not (excel and wb and ws):
            
            print("Excel cannot be opened")
            return False
        
        #Assign the data
        row_data = get_lab_data()
        
        #Columns where to print 
        columns = ['B', 'G', 'K', 'Q', 'U', 'X', 'AC']
        
        # Get the real sample matrix by mapping
        
        matrix_codes = {
            
            "A": "Air",
            "GW": "Groundwater",
            "SE": "Sediment",
            "SO": "Soil",
            "SW": "Surface Water",
            "W": "Water",
            "HW": "Potencial Haz Wastw"
            
        }
        
        #From this row start the space for printing begins
        start_row = 13
            
        # Iterate the number of rows with data brought in
        for excel_row in range(len(row_data)):
                
            # Get one row data object from the all  row data list
            current_data_row = row_data[excel_row]
            #print(current_data_row)
           
           #Iterate through the columns 
            for index, column in enumerate(columns):
                  
                cell_to_write = ws.Range(f"{column}{start_row}")
                              
                data_to_write = current_data_row[index]
                
                #In the column x we need to map the true name of the matrix
                if column == 'X':

                    #Write the real matrix name
                    cell_to_write.Value = matrix_codes.get(data_to_write)
                    
                else:
                    # Write the data
                    cell_to_write.Value = data_to_write
                    
            # Next row
            start_row += 1
               
        #Save the workbook  
        wb.Save()
                    
        return row_data
    
    except Exception as e:
        
        traceback.print_exc()
        
        return False
        
        
#Function to close or kill the excel processes before we start printing data to it
def kill_excel_processes():
    
    """Finish all excel process"""
    
    try:
        
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
                       
                      stdout=subprocess.DEVNULL, 
                      stderr=subprocess.DEVNULL)
    except:
        
        print("Excel processes could not be completed or there were none.")
        

print_lab_data()
