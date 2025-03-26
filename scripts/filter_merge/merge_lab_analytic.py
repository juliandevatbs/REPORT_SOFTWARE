import sys
import os
import datetime

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from scripts.filter_merge.filter_analytic_data import filter_analytic_data
from scripts.get.get_analytic_data import get_analytic_data
from scripts.get.get_lab_data import get_lab_data

def merge_lab_analytic():
    # Obtener los datos
    analytic_data, constant_values = filter_analytic_data()
    chain_data = get_lab_data()  # Esto reemplaza tu chain_data = []
    
    # Diccionario: SW-X → (id_cadena, fecha, tipo, método)
    chain_dict = {}
    for row in chain_data:
        if len(row) >= 7:  # Asegurarse que la fila tiene todos los elementos
            sw_code = row[1]  # 'SW-X' está en índice 1
            if sw_code:  # Ignorar si es None o vacío
                id_cadena = row[5]  # '2503014-XXX' en índice 5
                fecha_chain = row[2]  # Fecha en índice 2
                tipo = row[4]  # 'GW' en índice 4
                metodo = row[6]  # 'EPA 9212' en índice 6
                chain_dict[sw_code] = (id_cadena, fecha_chain, tipo, metodo)
    
    # Lista para el resultado final
    merged_data = []
    
    for analytic_row in analytic_data:
        if len(analytic_row) >= 4:  # Asegurarse que la fila tiene todos los elementos
            sw_code = analytic_row[0]  # 'SW-X' en analytic_data
            new_row = analytic_row.copy()
            
            if sw_code in chain_dict:
                id_cadena, fecha_chain, tipo, metodo = chain_dict[sw_code]
                # Agregar los datos de chain_data
                new_row.extend([id_cadena, fecha_chain, tipo, metodo])
            else:
                # Si no hay match, agregar campos vacíos
                new_row.extend([None, None, None, None])
            
            # Agregar constant_values al final de cada fila
            new_row.extend(constant_values)
            merged_data.append(new_row)
    
    #print(merged_data)
    return merged_data

merge_lab_analytic()