import traceback
import subprocess
from openpyxl import load_workbook
from datetime import datetime
import sys
import os
import time

# Configuración de rutas
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from scripts.filter_merge.filter_analytic_data import filter_analytic_data
from scripts.filter_merge.merge_lab_analytic import merge_lab_analytic
from scripts.get import get_analysis_requested

def safe_save_workbook(wb, route, max_attempts=3):
    """Intenta guardar el workbook con reintentos y manejo de errores"""
    for attempt in range(max_attempts):
        try:
            # Cerrar Excel antes de guardar
            kill_excel_processes()
            time.sleep(1)  # Esperar para asegurar cierre
            
            # Guardar con backup
            temp_route = route + ".temp"
            wb.save(temp_route)
            
            # Reemplazar archivo original
            if os.path.exists(route):
                os.remove(route)
            os.rename(temp_route, route)
            
            return True
        except Exception as e:
            print(f"Intento {attempt + 1} de guardar falló: {str(e)}")
            time.sleep(2)  # Try later
    return False

def print_analytical_data():
    """Escribe datos analíticos en la hoja de reporte"""
    try:
        # Configuración
        sheetname = "Reporte"
        route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (3).xlsx"
        start_row = 47
        row_spacing = 5

        # Obtener datos
        print("Obteniendo datos analíticos...")
        row_data, constant_values = filter_analytic_data()
        row_data = merge_lab_analytic()
        
        #print("ROWWW DAATAA")
        #print(row_data)
        if not row_data:
            print("Error: No hay datos para escribir")
            return False

        # Verificar si el archivo existe
        if not os.path.exists(route):
            print(f"Error: Archivo no encontrado en {route}")
            return False

        # Abrir el archivo Excel
        print("Abriendo workbook...")
        try:
            wb = load_workbook(filename=route)
            ws = wb[sheetname]
        except Exception as e:
            print(f"Error al abrir el archivo: {str(e)}")
            return False

        # Escribir datos
        print(f"Escribiendo {len(row_data)} bloques de datos...")
        for block_num, data_block in enumerate(row_data):
            if not isinstance(data_block, list):
                print(f"Bloque {block_num} no es una lista - omitiendo")
                continue

            first_line_row = start_row + (block_num * row_spacing)
            second_line_row = first_line_row + 2

            try:
                # Extraer datos con validación
                sw_code = str(data_block[0]) if len(data_block) > 0 else ""
                date_value = data_block[1] if len(data_block) > 1 else None
                by_value = str(data_block[2]) if len(data_block) > 2 else ""
                result_value = data_block[3] if len(data_block) > 3 else ""
                batch_id_value = data_block[4] if len(data_block) > 4 else ""
                matrix_id_value = data_block[6] if len(data_block) > 6 else ""
                
                

                # Formatear fechas
                date_value_str = date_value.strftime('%Y-%m-%d') if isinstance(date_value, datetime) else str(date_value)

                # Primera línea
                ws[f"B{first_line_row}"] = sw_code
                ws[f"J{first_line_row}"] = batch_id_value
                ws[f"R{first_line_row}"] = date_value_str
                ws[f"Z{first_line_row}"] = by_value
                ws[f"AJ{first_line_row}"] = matrix_id_value

                # Segunda línea
                ws[f"J{second_line_row}"] = result_value
                ws[f"AD{second_line_row}"] = date_value_str
                ws[f"AF{second_line_row}"] = by_value
                ws[f"AH{second_line_row}"] = batch_id_value
                
                print(f"Bloque {block_num} escrito en filas {first_line_row}-{second_line_row}")

            except Exception as block_error:
                print(f"Error en bloque {block_num}: {str(block_error)}")
                print(f"Datos problemáticos: {data_block}")
                continue

        # Guardar cambios
        print("Guardando workbook...")
        if not safe_save_workbook(wb, route):
            print("Error: No se pudo guardar el archivo después de varios intentos")
            return False
            
        print("Archivo guardado exitosamente")
        return True

    except Exception as e:
        print(f"Error crítico: {str(e)}")
        traceback.print_exc()
        return False
    finally:
        if 'wb' in locals():
            wb.close()
        kill_excel_processes()

def kill_excel_processes():
    """Cierra todos los procesos de Excel"""
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"], 
                      stdout=subprocess.DEVNULL, 
                      stderr=subprocess.DEVNULL)
        time.sleep(1)  
    except:
        pass

print_analytical_data()