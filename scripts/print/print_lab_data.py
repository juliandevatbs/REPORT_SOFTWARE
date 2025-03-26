from datetime import datetime
import sys
import os
import traceback
import time
import subprocess
from openpyxl import load_workbook

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from get.get_lab_data import get_lab_data


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


def print_lab_data():
    """Versión simplificada siguiendo el mismo patrón que clean_lab_data"""
    sheet_name = "Reporte"
    route = r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (3).xlsx"

    try:
        # Verificación inicial (igual que en clean_lab_data)
        if not os.path.exists(route):
            print(f"Error: Archivo no encontrado en {route}")
            return False

        # Abrir el archivo (mismo método)
        wb = load_workbook(filename=route)
        ws = wb[sheet_name]

        # Obtener datos
        row_data = get_lab_data()
        if not row_data:
            print("Error: No se recibieron datos del laboratorio")
            return False

        # Configuración de columnas
        columns = ['B', 'G', 'K', 'Q', 'U', 'X', 'AC']
        matrix_codes = {
            "A": "Air",
            "GW": "Groundwater",
            "SE": "Sediment",
            "SO": "Soil",
            "SW": "Surface Water",
            "W": "Water",
            "HW": "Potencial Haz Waste"
        }
        start_row = 13

        # Escribir datos (mismo estilo que clean_lab_data)
        for excel_row in range(len(row_data)):
            current_data_row = row_data[excel_row]

            for index, column in enumerate(columns):
                cell = ws[f"{column}{start_row}"]
                if column == 'X':
                    cell.value = matrix_codes.get(current_data_row[index], current_data_row[index])
                else:
                    cell.value = current_data_row[index]

            start_row += 1

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