import os
import sys

from scripts.clean.clean_analytic_data import clean_analytical_data
from scripts.clean.clean_header import clean_header_data
from scripts.clean.clean_lab_data import clean_lab_data
from scripts.clean.clean_summary_data import clean_summary_data
from scripts.excel.connect_excel import get_excel

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


sheet_name = "Reporte"
excel_route = r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx"

wb, ws = get_excel(sheet_name, excel_route)

def clean_all(wb, ws, route: str, sheet_name: str) -> bool:

    clean_h = clean_header_data(wb, ws, route, sheet_name)
    clean_l = clean_lab_data(wb, ws, route, sheet_name, 20)
    clean_a = clean_analytical_data(wb, ws, route, sheet_name)
    clean_s = clean_summary_data(wb, ws, route, sheet_name)

    if clean_h and clean_l and clean_a and clean_s:

        print("CLEAN SUCCESFULLY")
        return True

    else:

        print("CLEAN FAILED")
        return False

clean_all(wb, ws, excel_route, sheet_name)

