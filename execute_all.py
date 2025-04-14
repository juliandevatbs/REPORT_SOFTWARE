import sys
import os
from datetime import datetime
import traceback
import time
import subprocess
import gc
from openpyxl import load_workbook

from scripts.copy_blocks.copy_block import copy_range_with_styles

"""from scripts.clean.clean_analytic_data import clean_analytic_data
from scripts.clean.clean_header import clean_header_data
from scripts.clean.clean_lab_data import clean_lab_data
from scripts.clean.clean_summary_data import clean_summary_data"""
from scripts.excel.connect_excel import get_excel
from scripts.get.get_all_data import get_matrix_data_flattened
from scripts.get.get_header_data import get_header_data
from scripts.get.get_all_data import get_chain_data
from scripts.export.export_to_pdf import export_pdf_vertical
from scripts.utils.safe_save import safe_save_workbook

# Add parent directory to path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def get_workbooks(main_sheet_name: str, report_sheet_name: str, file_path: str):
    """Load both workbooks once and return them"""
    source_wb, source_ws = get_excel(main_sheet_name, file_path)
    report_wb, report_ws = get_excel(report_sheet_name, file_path)
    return source_wb, source_ws, report_wb, report_ws


def process_header(source_wb, source_ws, report_wb, report_ws):
    """Process header data"""
    print("Procesando datos de encabezado...")
    from scripts.print.print_header_data import print_header_data
    header_data = get_header_data(source_wb, source_ws)
    print_header_data(report_wb, report_ws, header_data)
    print("✓ Datos de encabezado procesados")


def process_lab(source_wb, source_ws, report_wb, report_ws):
    """Process lab data"""
    print("Procesando datos de laboratorio...")
    from scripts.print.print_lab_data import print_lab_data
    lab_data = get_chain_data()
    print_lab_data(report_wb, report_ws, lab_data)
    print("✓ Datos de laboratorio procesados")


def process_analytic(source_wb, source_ws, report_wb, report_ws, file_path: str):
    """Process analytical data"""
    print("Procesando datos analíticos...")
    from scripts.print.print_analytical_data import print_analytical_data
    row_data = get_matrix_data_flattened(file_path)
    print_analytical_data(report_wb, report_ws, row_data)
    print("✓ Datos analíticos procesados")
    return row_data


def process_summary(source_wb, source_ws, report_wb, report_ws, file_path: str):
    """Process summary data"""
    print("Procesando resumen de datos...")
    from scripts.print.print_summary_data import print_summary_data
    analytic_data = process_analytic(source_wb, source_ws, report_wb, report_ws, file_path)
    last_lab_sample_id = print_summary_data(report_wb, report_ws, analytic_data)
    return last_lab_sample_id


"""def clean_report(file_path: str, report_sheet_name: str):

    clean_header_data(file_path, report_sheet_name)
    clean_analytic_data(file_path)
    clean_lab_data(file_path)
    clean_summary_data(file_path)
    print("EXCEL LIMPIO")
"""

def execute_all(status_callback = None):
    """Execute all report generation steps with optimized workbook handling"""
    start_time = time.time()

    # Configuration
    config = {
        'source_sheet': "Chain of Custody 1",
        'report_sheet': "Final",
        'file_path': r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx",
        'output_pdf': r"C:\Users\Duban Serrano\Desktop\reporte.pdf"
    }

    # Función de actualización de estado
    def update_status(message, progress=None):
        if status_callback:
            status_callback(message, progress)
        print(message)



    try:
        update_status("Cargando archivos Excel...", 0.1)

        wb = load_workbook(r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx")
        copy_ws = wb["Reporte"]
        source_ws = wb["Chain of Custody 1"]
        report_ws = wb["Final"]
        print(source_ws)
        print(report_ws)

        copy_range_with_styles(copy_ws, report_ws, "B1:AQ13", "B1")

        # Process data
        update_status("Procesando datos de encabezado...", 0.25)
        process_header(wb, source_ws, wb, report_ws)
        update_status("Procesando datos de laboratorio...", 0.5)

        start_row = 14

        lab_items = 20

        for items in range(lab_items):

            copy_range_with_styles(copy_ws, report_ws, f"B{14}:AQ{14}", f"B{start_row}")
            start_row += 1


        update_status("Procesando datos analiticos", 0.75)
        process_lab(wb, source_ws, wb, report_ws)

        first_footer = lab_items + start_row
        init_footer = 36
        final_footer = 49
        copy_range_with_styles(copy_ws, report_ws, f"B{init_footer}:AQ{final_footer}", f"B{first_footer}")

        start_row_analytic = first_footer + (final_footer - init_footer)

        copy_range_with_styles(copy_ws, report_ws,  f"B49:AQ53", f"B{start_row_analytic}")




        """
        # Process data
        update_status("Procesando datos analiticos...", 0.50)
        process_analytic(source_wb, source_ws, report_wb, report_ws, config['file_path'])
        update_status("Procesando datos analiticos...", 0.50)
        
       
        

        process_summary(source_wb, source_ws, report_wb, report_ws, config['file_path'])
        update_status("Guardando cambios...", 0.85)
        safe_save_workbook(report_wb, config['file_path'])

        # Export to PDF
        update_status("Exportando a PDF...", 0.9)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        export_pdf_vertical(
            excel_path=config['file_path'],
            pdf_path=config['output_pdf'],
            sheet_name=config['report_sheet']
        )

        # Open the PDF
        update_status("Abriendo PDF generado...", 1.0)
        try:
            os.startfile(config['output_pdf'])  # Windows
        except AttributeError:
            if sys.platform.startswith('darwin'):  # macOS
                subprocess.call(['open', config['output_pdf']])
            else:  # Linux
                subprocess.call(['xdg-open', config['output_pdf']])
"""
    finally:
        safe_save_workbook(wb, r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx")


        # Clean up
        if 'source_wb' in locals():
            wb.close()
        if 'report_wb' in locals():
            wb.close()
        gc.collect()

    end_time = time.time()

    print(f"Tiempo total de procesamiento: {end_time - start_time:.2f} segundos")

    return True

if __name__ == "__main__":
    execute_all()
