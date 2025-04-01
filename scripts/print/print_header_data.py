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
    
    wb = None
    
    try:
        kill_excel_processes()
        
        sheet_name= "Reporte"
        
        route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx"

        wb, ws = get_excel(sheet_name, route)
        
        if not (wb and ws):
            
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
        ws['K7'].value = client_str
        ws['K40'].value = client_str
        ws['K41'].value = client_str
        ws['K8'].value = client_str
        ws['K9'].value = address_str
        ws['K42'].value = address_str
        ws['K10'].value = city_str
        ws['K43'].value = city_str
        ws['AK9'].value = phone_str
        ws['AK42'].value = phone_str
        ws['M11'].value = zip_str
        ws['M44'].value = zip_str
        ws["I11"].value= state_str
        ws["I44"].value= state_str
        
        # Facility  Id and Client´s  Project Number are NA
        ws["AK7"].value= na_value
        ws["AK40"].value= na_value
        ws["AK10"].value= na_value
        ws["AK43"].value= na_value
        
        # Save the file
        wb.save(route)
        
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
                    wb.close(SaveChanges=True)
                     
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