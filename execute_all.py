import sys
import os
from datetime import datetime
import time
import gc
from openpyxl import load_workbook

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.excel.connect_excel import get_excel
from scripts.get.get_all_data import get_matrix_data_flattened, get_chain_data
from scripts.get.get_header_data import get_header_data
from scripts.export.export_to_pdf import export_pdf_vertical
from scripts.utils.safe_save import safe_save_workbook
from scripts.print.print_header_space import header_space
from scripts.print.print_lab_space import lab_space
from scripts.print.print_lab_format import print_lab_format_row
from scripts.print.print_lab_data import print_lab_data
from scripts.print.print_footer import print_footer
from scripts.print.print_analytic_format import print_analytic_format
from scripts.print.print_summary_space import print_summary_space
from scripts.print.print_summary_format import print_summary_format
from scripts.print.print_quality_space import print_quality_space

# Definir constantes para evitar repetición
EXCEL_PATH = r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx"
SAVE_RETRIES = 3


def execute_all(status_callback=None):
    try:
        start_time = time.time()
        print(f"Iniciando procesamiento del archivo: {EXCEL_PATH}")

        # Abrir el libro de Excel una sola vez
        wb = get_excel(EXCEL_PATH)
        wsd = wb["Final"]

        # Ejecutar todas las operaciones en secuencia
        # Header section
        next_row = header_space(wb, wsd, EXCEL_PATH, 1)

        # Lab section
        lab_space(wb, wsd, EXCEL_PATH, 14)
        print_lab_format_row(wb, wsd, 20, EXCEL_PATH, 15)

        # Footer after lab
        footer_row = print_footer(wb, wsd, EXCEL_PATH, 37 + 5)

        # Analytic section
        next_row = header_space(wb, wsd, EXCEL_PATH, footer_row)
        print_analytic_format(wb, wsd, 56, EXCEL_PATH, 23)

        # Footer after analytic
        footer_row = print_footer(wb, wsd, EXCEL_PATH, 151)

        # Summary section
        next_row = header_space(wb, wsd, EXCEL_PATH, footer_row)
        print_summary_space(wb, wsd, 165, EXCEL_PATH)
        print_summary_format(wb, wsd, 167, EXCEL_PATH, 23)

        # Footer after summary
        footer_row = print_footer(wb, wsd, EXCEL_PATH, 263)

        # Quality section
        next_row = header_space(wb, wsd, EXCEL_PATH, footer_row)
        print_quality_space(wb, wsd, 277, EXCEL_PATH)

        # Print lab data at the end
        print_lab_data(wb, wsd, get_chain_data())

        # Guardar una sola vez al final
        print("Guardando archivo final...")
        safe_save_workbook(wb, EXCEL_PATH, SAVE_RETRIES)

        # Liberar memoria explícitamente
        wb.close()
        wb = None
        gc.collect()

        end_time = time.time()
        print(f"Proceso completado. Tiempo total: {end_time - start_time:.2f}s")
        return True

    except Exception as e:
        print(f"ERROR en execute_all: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    execute_all()