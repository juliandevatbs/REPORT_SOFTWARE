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
from scripts.get.get_all_data import get_chain_data
from scripts.get.get_all_data import get_matrix_data



def safe_save_workbook(wb, route, max_attempts=3):
    """Intenta guardar el workbook con reintentos y manejo de errores"""
    for attempt in range(max_attempts):
        try:
            kill_excel_processes()
            time.sleep(1)

            temp_route = route + ".temp"
            wb.save(temp_route)

            if os.path.exists(route):
                os.remove(route)
            os.rename(temp_route, route)

            return True
        except Exception as e:
            print(f"Intento {attempt + 1} de guardar falló: {str(e)}")
            time.sleep(2)
    return False


def should_write_block(data_block):
    """Determina si el bloque debe escribirse según la condición data_block[18] > data_block[12]"""
    try:
        # Verificar que ambos valores existan y sean comparables
        if len(data_block) < 19:
            return False

        value_18 = float(data_block[18]) if data_block[18] is not None else None
        value_12 = float(data_block[12]) if data_block[12] is not None else None

        if value_18 is None or value_12 is None:
            return False

        return value_18 > value_12
    except (ValueError, TypeError):
        return False


def write_data_block(ws, data_block, first_line_row):
    """Escribe un bloque de datos válido en las filas especificadas"""
    try:
        # Extraer datos
        sw_code = str(data_block[1])
        lab_sample_id = str(data_block[7])
        date_value = data_block[2]
        by_value = str(data_block[20])
        result_value = data_block[18]
        batch_id_value = data_block[7]
        matrix_id_value = data_block[5]
        method_analyzed = data_block[15]
        units_value = data_block[14]
        df_value = data_block[11]
        mdl_value = data_block[12]
        pql_value = data_block[13]

        # Formatear fechas
        date_value_str = date_value.strftime('%Y-%m-%d') if isinstance(date_value, datetime) else str(date_value)

        # Definir posiciones de filas
        second_line_row = first_line_row + 2


        # Primera línea
        ws[f"B{first_line_row}"] = sw_code
        ws[f"R{first_line_row}"] = date_value_str
        ws[f"J{first_line_row}"] = batch_id_value
        ws[f"Z{first_line_row}"] = by_value
        ws[f"AJ{first_line_row}"] = matrix_id_value

        # Segunda línea
        ws[f"J{second_line_row}"] = result_value
        ws[f"R{second_line_row}"] = units_value
        ws[f"U{second_line_row}"] = df_value
        ws[f"V{second_line_row}"] = mdl_value
        ws[f"W{second_line_row}"] = pql_value
        ws[f"Z{second_line_row}"] = method_analyzed
        ws[f"AD{second_line_row}"] = date_value_str
        ws[f"AF{second_line_row}"] = by_value
        ws[f"AH{second_line_row}"] = lab_sample_id



        return True
    except Exception as e:
        print(f"Error al escribir bloque: {str(e)}")
        return False


def print_summary_data():
    """Escribe datos analíticos en la hoja de reporte"""
    try:
        # Configuración
        config = {
            "sheetname": "Reporte",
            "filepath": r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx",
            "start_row": 114,
            "row_spacing": 5
        }

        # Obtener datos
        print("Obteniendo datos analíticos...")
        row_data = get_matrix_data(get_chain_data())

        if not row_data:
            print("Error: No hay datos para escribir")
            return False

        # Verificar archivo
        if not os.path.exists(config["filepath"]):
            print(f"Error: Archivo no encontrado en {config['filepath']}")
            return False

        # Abrir archivo Excel
        print("Abriendo workbook...")
        try:
            wb = load_workbook(filename=config["filepath"])
            ws = wb[config["sheetname"]]
        except Exception as e:
            print(f"Error al abrir el archivo: {str(e)}")
            return False

        # Procesar datos
        print(f"Procesando {len(row_data)} bloques de datos...")
        valid_blocks = 0
        current_row = config["start_row"]

        for block_num, data_block in enumerate(row_data):
            if not isinstance(data_block, list):
                print(f"Bloque {block_num} no es una lista - omitiendo")
                continue

            # Verificar condición data_block[18] > data_block[12]
            if should_write_block(data_block):
                if write_data_block(ws, data_block, current_row):
                    valid_blocks += 1
                    print(f"Bloque {block_num} escrito en fila {current_row}")
                    current_row += config["row_spacing"]
                else:
                    print(f"Error al escribir bloque {block_num} válido")
            else:
                print(f"Bloque {block_num} no cumple condición (18 > 12) - omitiendo")

        print(f"{valid_blocks} bloques válidos escritos de {len(row_data)} totales")

        # Guardar cambios
        print("Guardando workbook...")
        if not safe_save_workbook(wb, config["filepath"]):
            print("Error: No se pudo guardar el archivo")
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

print_summary_data()