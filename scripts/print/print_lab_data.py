from datetime import datetime, timedelta
import sys
import os
import traceback
import time
import subprocess
import re
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.get.get_all_data import get_chain_data
from scripts.get.get_all_data import get_matrix_data


def kill_excel_processes():
    """Cierra todos los procesos de Excel - Mismo que en clean_lab_data"""
    try:
        subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        time.sleep(1)
    except:
        pass


def safe_save_workbook(wb, route):
    """Versión simplificada idéntica a la de clean_lab_data"""
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
        print(f"Error al guardar: {str(e)}")
        return False


def escribir_en_celda(ws, celda_coord, valor):
    """
    Escribe un valor en una celda, incluso si es parte de un rango combinado.

    Args:
        ws: Hoja de trabajo
        celda_coord: Coordenada de la celda (ej: 'B13')
        valor: Valor a escribir

    Returns:
        bool: True si se pudo escribir, False en caso contrario
    """
    try:
        # Usar expresión regular para separar letras y números
        match = re.match(r'([A-Za-z]+)(\d+)', celda_coord)
        if not match:
            print(f"Formato de coordenada incorrecto: {celda_coord}")
            return False

        col_str, row_str = match.groups()
        row = int(row_str)
        col = column_index_from_string(col_str)

        # Intentar obtener la celda directamente
        celda = ws.cell(row=row, column=col)

        # Si no es una celda combinada, escribir directamente
        if not isinstance(celda, MergedCell):
            celda.value = valor
            return True

        # Si es una celda combinada, buscar la celda principal
        for rango in ws.merged_cells.ranges:
            min_row, min_col, max_row, max_col = rango.min_row, rango.min_col, rango.max_row, rango.max_col

            # Verificar si la celda está dentro del rango
            if min_row <= row <= max_row and min_col <= col <= max_col:
                # Escribir en la celda principal
                celda_principal = ws.cell(row=min_row, column=min_col)
                celda_principal.value = valor
                return True

        # No se encontró un rango que contuviera esta celda
        print(f"No se encontró un rango combinado para la celda {celda_coord}")
        return False

    except Exception as e:
        print(f"Error en escribir_en_celda con coordenada {celda_coord}: {str(e)}")
        return False

def format_date(value):
    """Formatea valores de fecha consistentemente"""
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    return str(value) if value is not None else ""


def print_lab_data():
    """Versión que maneja row_data como lista de listas y escribe correctamente en celdas combinadas"""
    sheet_name = "Reporte"
    route = r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx"

    try:
        # Verificación inicial
        if not os.path.exists(route):
            print(f"Error: Archivo no encontrado en {route}")
            return False

        # Abrir el archivo
        wb = load_workbook(filename=route)
        ws = wb[sheet_name]

        # Obtener datos
        all_rows_data = get_matrix_data(get_chain_data())
        if not all_rows_data:
            print("Error: No se recibieron datos del laboratorio")
            return False

        # Verificar que tenemos una lista de listas
        if not isinstance(all_rows_data, list) or not all(isinstance(row, list) for row in all_rows_data):
            print(f"Error: Datos recibidos en formato incorrecto. Se esperaba lista de listas.")
            return False

        start_row = 13

        for row_idx, row_data in enumerate(all_rows_data):
            current_row = start_row + row_idx

            # Verificar que la fila tiene suficientes datos
            # Ajustar según el índice máximo que necesites (15 en este caso)
            if len(row_data) < 16:
                print(f"Advertencia: Fila {row_idx} no tiene suficientes datos. Se omitirá.")
                continue

            try:
                # Item (B)
                escribir_en_celda(ws, f'B{current_row}', row_data[0])

                # Lab Sample ID (G)
                escribir_en_celda(ws, f'G{current_row}', row_data[7])

                # Client Sample ID (K)
                escribir_en_celda(ws, f'K{current_row}', row_data[1])


                date_value = format_date(ws[f'T{current_row}'].value)
                escribir_en_celda(ws, f'T{current_row}', date_value)



                # Tiempo (U)
                if isinstance(row_data[3], timedelta):
                    total_seconds = int(row_data[3].total_seconds())
                    hours = total_seconds // 3600
                    minutes = (total_seconds % 3600) // 60
                    tiempo_formateado = f"{hours:02d}:{minutes:02d}"
                    escribir_en_celda(ws, f'U{current_row}', tiempo_formateado)

                # Sample Matrix (X)
                escribir_en_celda(ws, f'X{current_row}', row_data[5])

                # Analysis Requested (AC) - ahora usando índice 15 según el error
                escribir_en_celda(ws, f'AC{current_row}', row_data[15])

            except Exception as e:
                print(f"Error al procesar fila {current_row}: {str(e)}")
                traceback.print_exc()
                continue

        if not safe_save_workbook(wb, route):
            return False

        return True

    except Exception as e:
        print(f"Error crítico: {str(e)}")
        traceback.print_exc()
        return False
    finally:
        if 'wb' in locals():
            try:
                wb.close()
            except:
                pass
        kill_excel_processes()


if __name__ == "__main__":
    result = print_lab_data()
    print(f"Resultado: {'Éxito' if result else 'Falló'}")
    sys.exit(0 if result else 1)