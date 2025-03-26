import sys
import os




sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from scripts.filter_merge.merge_lab_analytic import merge_lab_analytic
from scripts.get.get_mdl import get_mdl_value


"""
    This function read the analytic data and return the only sets that results value > 
    MDL

"""

def filter_summary_data():
    
    mdl_value = get_mdl_value()
    
    analytic_data = merge_lab_analytic()
    filtered_data = []
    
    for row in analytic_data:
        
        if row[3] > mdl_value:
            
            filtered_data.append(row)
    
    #print(filtered_data)
    return filtered_data
    
filter_summary_data()