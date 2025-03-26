import sys
import os
import traceback
import time
import subprocess
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from get.get_header_data import get_header_data
from excel.connect_excel import get_excel

"""
    This function print the header data in the excel cells
"""

def print_header_data():
    
    excel = None
    wb = None
    
    try:
        kill_excel_processes()
        
        sheet_name= "Reporte"
        
        route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12.xlsm"
        # Get the excel connection
        excel, wb, ws = get_excel(sheet_name, route)
        
        if not (excel and wb and ws):
            
            print("Excel cannot be opened")
            return False
        
        
        # Get the header data
        #print(get_header_data())
        client, address, city, company_phone, zip_code, state = get_header_data()
        
        # Make  sure that the data is a string 
        client_str = str(client) if client is not None else ""
        address_str = str(address) if address is not None else ""
        city_str = str(city) if city is not None else ""
        phone_str=  str(company_phone) if company_phone is not  None else ""
        zip_str=  str(zip_code) if zip_code is not None else ""
        state_str= str(state) if state is not None else ""
        
        #NA value
        na_value= 'NA'
                
        # Write data cells
        print("Writing data")
        ws.Range('K7').Value = client_str
        ws.Range('K40').Value = client_str
        ws.Range('K41').Value = client_str
        ws.Range('K8').Value = client_str
        ws.Range('K9').Value = address_str
        ws.Range('K42').Value = address_str
        ws.Range('K10').Value = city_str
        ws.Range('K43').Value = city_str
        ws.Range('AK9').Value = phone_str
        ws.Range('AK42').Value = phone_str
        ws.Range('M11').Value = zip_str
        ws.Range('M44').Value = zip_str
        ws.Range("I11").Value= state_str
        ws.Range("I44").Value= state_str
        
        # Facility  Id and Client´s  Project Number are NA
        ws.Range("AK7").Value= na_value
        ws.Range("AK40").Value= na_value    
        ws.Range("AK10").Value= na_value
        ws.Range("AK43").Value= na_value
        
        # Save the file
        wb.Save()
        
        # Wait a while for the saving process to complete
        time.sleep(1)
        
        print("File saved succesfully")
        
        return True
        
    except Exception as e:
        
        traceback.print_exc()
        
        return False
    
    finally:
        
        # Make sure that the excel is close
        try:
            if wb is not None:
                
                print("Closing the book")
                
                try:
                    wb.Close(SaveChanges=True) 
                     
                except:
                    pass
                    
            if excel is not None:
                
                print("Closing the excel")
                
                try:
                    
                    excel.Quit()
                    
                except:
                    
                    pass
                
            time.sleep(2)
            kill_excel_processes()
                
            
        except Exception as e:
            
            print("FAILED")
                
#Function to close or kill the excel processes before we start printing data to it            
def kill_excel_processes():
    
    """Finish all excel process"""
    
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], 
                      stdout=subprocess.DEVNULL, 
                      stderr=subprocess.DEVNULL)
    except:
        print("Excel processes could not be completed or there were none.")



# Function to  merge
def write_and_save ():
    print_header_data()
    kill_excel_processes()    
    

write_and_save()