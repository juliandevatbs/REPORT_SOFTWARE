import sys
import os
from datetime import datetime
import time
import gc
from openpyxl import load_workbook

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
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
from scripts.print.print_analytic_space import print_analyitic_space
from scripts.print.print_header_data import print_header_data

# Definir constantes para evitar repetición
EXCEL_PATH = r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx"
SAVE_RETRIES = 3
WB = get_excel(EXCEL_PATH)
WSD = WB["Final"]


def prepare_format():

    try:

        print(f"Preparing report format...")

        # Header section
        next_row, cell_mapping = header_space(WB, WSD, 1)
        header_data = get_header_data(WB, CC)
        print_header_data(WB, WSD, header_data, cell_mapping)

        # Lab section
        lab_space(WB, WSD, next_row)
        next_row = print_lab_format_row(WB, WSD, 20, next_row +1)


        # Footer after lab
        footer_row = print_footer(WB, WSD, next_row + 2)

        # Analytic section
        next_row = header_space(WB, WSD, footer_row)
        next_row = print_analyitic_space(WB, WSD, next_row)
        next_row = print_analytic_format(WB, WSD, next_row, EXCEL_PATH, 23)

        # Footer after analytic
        footer_row = print_footer(WB, WSD, next_row)

        # Summary section
        next_row = header_space(WB, WSD, footer_row)
        next_row = print_summary_space(WB, WSD, next_row)
        next_row = print_summary_format(WB, WSD, next_row, 23)

        # Footer after summary
        footer_row = print_footer(WB, WSD, next_row)

        # Quality section
        next_row = header_space(WB, WSD, footer_row)
        print_quality_space(WB, WSD, next_row)

        safe_save_workbook(WB, EXCEL_PATH, SAVE_RETRIES)



    except Exception as e:

        print(e)


CC = WB["Chain of Custody 1"]




def execute_all(status_callback=None):

    try:
        start_time = time.time()
        prepare_format()
        WB = None
        gc.collect()

        end_time = time.time()
        print(f"Process completed. Total time: {end_time - start_time:.2f}s")
        return True

    except Exception as e:
        print(f"ERROR en execute_all: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    execute_all()