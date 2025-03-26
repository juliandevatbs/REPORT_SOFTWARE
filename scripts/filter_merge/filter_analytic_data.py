import sys
import os



# Agrega el directorio raíz al PATH
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from scripts.get.get_analytic_data import get_analytic_data
from scripts.get.get_lab_data import get_lab_data


#This function clean the data to print into analytic section
def filter_analytic_data():
    
    #Get the analytic data 
    row_data, constant_values = get_analytic_data()
    filter_data = []
    
    #print(row_data)
    
    #Iterate through data to clean '' or none values
    for row in row_data:
        
        #print(row)
        
        # Only verify lists
        if type(row) == list:
            
            # Only lists with SW value in the first position is correct
            if row[0] != None:
                
                #print(row[0])
                filter_data.append(row)
                
                
    #for row in filter_data:
        #print(row)
                
    return filter_data, constant_values
    
filter_analytic_data()



    

