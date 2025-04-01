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


def validate_data_block(data_block):
    """Valida que el bloque de datos tenga el formato correcto"""
    if not isinstance(data_block, list) or len(data_block) < 18:
        print(f"Bloque no válido. Se esperaba lista de 18 elementos. Recibido: {data_block}")
        return False
    return True


def format_date(value):
    """Formatea valores de fecha consistentemente"""
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    return str(value) if value is not None else ""


def write_data_block(ws, data_block, first_line_row):
    """Escribe un bloque de datos en las filas especificadas"""
    try:
        # Validar datos primero
        if not validate_data_block(data_block):
            return False

        # Extraer datos con validación
        sw_code = str(data_block[1]) if data_block[1] is not None else ""
        date_value = format_date(data_block[2])
        by_value = str(data_block[20]) if data_block[14] is not None else ""
        result_value = data_block[18] if data_block[18] is not None else ""
        batch_id_value = data_block[7] if data_block[7] is not None else ""
        matrix_id_value = data_block[5] if data_block[5] is not None else ""
        results_value = data_block[18]
        df_value = data_block[11]
        mdl_value = data_block[12]
        pql_value = data_block[13]
        units_value = data_block[14]
        analyzed_method = data_block[15]

        second_line_row = first_line_row + 2


        # Mapeo de celdas a escribir
        cell_mapping = {
            f"B{first_line_row}": sw_code,
            f"J{first_line_row}": batch_id_value,
            f"R{first_line_row}": date_value,
            f"Z{first_line_row}": by_value,
            f"AJ{first_line_row}": matrix_id_value,
            f"J{second_line_row}": result_value,
            f"AD{second_line_row}": date_value,
            f"AF{second_line_row}": by_value,
            f"AH{second_line_row}": batch_id_value,
            f"R{second_line_row}": units_value,
            f"U{second_line_row}": df_value,
            f"V{second_line_row}": mdl_value,
            f"W{second_line_row}": pql_value,
            f"Z{second_line_row}": analyzed_method,
        }

        # Escribir todos los valores
        for cell, value in cell_mapping.items():
            ws[cell] = value

        return True

    except Exception as e:
        print(f"Error al escribir bloque: {str(e)}")
        return False


def print_analytical_data():
    """Escribe datos analíticos en la hoja de reporte"""
    try:
        # Configuración
        config = {
            "sheetname": "Reporte",
            "filepath": r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx",
            "start_row": 47,
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

        # Escribir datos
        print(f"Escribiendo {len(row_data)} bloques de datos...")
        success_count = 0

        current_row = config["start_row"]

        for block_num, data_block in enumerate(row_data):
            if validate_data_block(data_block):  # Si cumple condiciones
                if write_data_block(ws, data_block, current_row):  # Intenta escribir
                    success_count += 1
                    current_row += config["row_spacing"]  # Solo avanza si fue exitoso

        print(f"{success_count}/{len(row_data)} bloques escritos exitosamente")

        # Guardar cambios
        print("Guardando workbook...")
        if not safe_save_workbook(wb, config["filepath"]):
            print("Error: No se pudo guardar el archivo")
            return False

        print("Proceso completado exitosamente")
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
