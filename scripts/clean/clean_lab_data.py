import sys
import os
import traceback
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.utils.format_date import format_date
from scripts.utils.kill_excel_processes import kill_excel_processes
from scripts.get.get_all_data import get_chain_data
from scripts.utils.write_cell import  write_cell


def clean_lab_data(wb, ws, route: str, sheet_name: str, n_records: int) -> bool:
    try:

        start_row = 14

        empty_value = ''

        for row_idx, row_data in range(n_records):
            current_row = start_row + row_idx
            write_cell(ws, f"B{start_row}", empty_value)
            write_cell(ws, f"H{start_row}", empty_value)
            write_cell(ws, f"K{start_row}", empty_value)
            write_cell(ws, f"U{start_row}", empty_value)
            write_cell(ws, f"W{start_row}", empty_value)
            write_cell(ws, f"AC{start_row}",empty_value)
            write_cell(ws, f"AF{start_row}",empty_value)
            start_row += 1

        print("¡Lab Data cleaned!")
        return True

    except Exception as e:
        print(f"Fatal error: {str(e)}")
        traceback.print_exc()
        return False


