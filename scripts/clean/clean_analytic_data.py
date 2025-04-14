import sys
import os
import traceback

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.utils.write_cell import write_cell
from scripts.utils.safe_save import safe_save_workbook


def clean_analytical_data(wb, ws, row_data, n_records) -> bool:
    try:
        start_row = 51
        row_spacing = 5
        empty_value = ''

        if n_records is None and row_data:
            n_records = len(row_data)
        if not n_records:
            print("No hay datos para limpiar")
            return False

        print(f"Limpiando {n_records} bloques de datos analíticos...")
        success_count = 0

        for i in range(n_records):
            current_row = start_row + (i * row_spacing)
            second_line_row = current_row + 2

            # Primera línea
            write_cell(ws, f"B{current_row}", empty_value)
            write_cell(ws, f"J{current_row}", empty_value)
            write_cell(ws, f"R{current_row}", empty_value)
            write_cell(ws, f"Z{current_row}", empty_value)
            write_cell(ws, f"AJ{current_row}", empty_value)

            # Segunda línea
            write_cell(ws, f"B{second_line_row}", empty_value)
            write_cell(ws, f"J{second_line_row}", empty_value)
            write_cell(ws, f"AD{second_line_row}", empty_value)
            write_cell(ws, f"AF{second_line_row}", empty_value)
            write_cell(ws, f"AH{second_line_row}", empty_value)
            write_cell(ws, f"R{second_line_row}", empty_value)
            write_cell(ws, f"U{second_line_row}", empty_value)
            write_cell(ws, f"V{second_line_row}", empty_value)
            write_cell(ws, f"W{second_line_row}", empty_value)
            write_cell(ws, f"Z{second_line_row}", empty_value)
            write_cell(ws, f"AJ{second_line_row}", empty_value)

            success_count += 1

            # Opcional: imprimir progreso cada 10 bloques
            if i % 10 == 0:
                print(f"Limpiados {i} bloques...")

        print(f"Bloques analíticos limpiados exitosamente: {success_count}/{n_records}")
        return True

    except Exception as e:
        print(f"Error crítico al limpiar datos analíticos: {str(e)}")
        traceback.print_exc()
        return False