import sys
import os
from datetime import datetime
import time
import gc
from openpyxl import load_workbook

from scripts.get.get_quality_data import get_quality_data
from scripts.print.print_quality_format import print_quality_format
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.excel.connect_excel import get_excel
from scripts.print.print_analytical_data import print_analytical_data
from scripts.print.print_summary_data import print_summary_data
from scripts.get.get_all_data import get_matrix_data_flattened
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
from scripts.get.get_all_data import get_chain_data


# Definir constantes para evitar repetición
EXCEL_PATH = "./excel/Reporte 2025-03-12 (4) 1.xlsx"
SAVE_RETRIES = 3
WB = get_excel(EXCEL_PATH)
WSD = WB["Final"]
WSC = WB["Chain of Custody 1"]


def prepare_format():

    try:
        chain_data, mixed_data = get_matrix_data_flattened(WB, WSC, EXCEL_PATH)
        
        q_mixed = len(mixed_data)
        
        # Header section
        next_row, cell_mapping = header_space(WB, WSD, 1)
        header_data = get_header_data(WB, CC)
        print_header_data(WSD, header_data, cell_mapping)

        # Lab section
        lab_space(WB, WSD, next_row)
        next_row = print_lab_format_row(WB, WSD, 20, next_row +1)
        print_lab_data(WSD, chain_data, next_row - 20)
        
        # Footer after lab
        footer_row = print_footer(WB, WSD, next_row + 2)

        # Analytic section
        next_row, spacing_data = header_space(WB, WSD, footer_row)
        print_header_data(WSD, header_data, spacing_data)
        next_row_f_p = print_analyitic_space(WB, WSD, next_row)
        next_row = print_analytic_format(WB, WSD, next_row_f_p, EXCEL_PATH, q_mixed)
        print_analytical_data(WSD, mixed_data, next_row_f_p + 1 )
        
        # Footer after analytic
        footer_row = print_footer(WB, WSD, next_row)

        # Summary section
        next_row, spacing_data = header_space(WB, WSD, footer_row)
        print_header_data(WSD, header_data, spacing_data)
        next_row_s = print_summary_space(WB, WSD, next_row)
        next_row = print_summary_format(WB, WSD, next_row_s, q_mixed)
        
        try:
            
            print_summary_data(WSD, mixed_data, next_row_s+1)
            
        except Exception as e:
            
            print(e)
        
        # Footer after summary
        footer_row = print_footer(WB, WSD, next_row +2)

        # Quality section
        next_row, spacing_data = header_space(WB, WSD, footer_row)
        print_header_data(WSD, header_data, spacing_data)
        print_quality_space(WB, WSD, next_row)
        
        quality_data = get_quality_data(WB, EXCEL_PATH)
        print("-----------------------------QUALITY DATA-------------------------------")
        for row in quality_data:
            print(row)
        try: 
            print_quality_format(WB, WSD, next_row, quality_data)
        except Exception as e:
            print(e)

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