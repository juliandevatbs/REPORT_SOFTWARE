import traceback
import subprocess
from openpyxl import load_workbook
from datetime import datetime
import sys
import os
import time

from scripts.get.get_quality_data import get_quality_data

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


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


def print_quality_data():
    """Escribe datos analíticos en la hoja de reporte"""
    try:

        sheetname = "Reporte"
        route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx"
        start_row = 182
        row_spacing = 7

        print("Obteniendo datos analíticos...")
        row_data, method = get_quality_data()

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
            second_line_row = first_line_row + 1
            third_line_row = first_line_row + 2
            fourth_line_row = first_line_row + 3
            fifth_line_row = first_line_row + 5


            try:
                # Extraer datos con validación
                client_sample_id = str(data_block[0])
                sampled = data_block[0]
                lab_sample_id = data_block[0]
                prep = data_block[0]
                analyzed_value = data_block[1]
                matrix_id_value = ''
                analyte_name = method
                by_value = data_block[4]


                # Primera línea
                ws[f"J{first_line_row}"] = client_sample_id
                ws[f"AC{first_line_row}"] = sampled


                # Segunda línea
                ws[f"J{second_line_row}"] = lab_sample_id
                ws[f"AC{second_line_row}"] = sampled
                ws[f"AE{second_line_row}"] = analyzed_value
                ws[f"AJ{second_line_row}"] = matrix_id_value

                ws[f"B{fifth_line_row}"] =method
                ws[f"W{fifth_line_row}"] = method
                ws[f"AD{fifth_line_row}"] = sampled
                ws[f"W{fifth_line_row}"] = by_value

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

    return True

def kill_excel_processes():
    """Cierra todos los procesos de Excel"""
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
                      stdout=subprocess.DEVNULL,
                      stderr=subprocess.DEVNULL)
        time.sleep(1)
    except:
        pass


print_quality_data()