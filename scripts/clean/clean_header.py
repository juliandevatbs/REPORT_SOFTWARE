import re
import sys
import os

from openpyxl.cell import MergedCell
from openpyxl.utils import column_index_from_string

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.excel.connect_excel import get_excel
from scripts.get.get_header_data import get_header_data
from scripts.error.show_error import show_info
from scripts.utils.write_cell import  write_cell

def clean_header_data(wb, ws, route: str, sheet_name: str):
    """
        Writes header data to specified Excel report template.

        This function:
        1. Opens the specified Excel file and worksheet
        2. Retrieves header data from a source Excel file
        3. Writes the header information to designated cells in the template
        4. Saves the modified workbook

        Parameters:
            excel_route (str): Path to the destination Excel template file
            sheet_name (str): Name of the worksheet to modify

        Returns:
            bool: True if operation completed successfully
            None: If operation failed

        Raises:
            Prints errors to console and shows user-friendly messages via show_info()
        """

    try:



        empty_value = ''
        # No Apllicaton string
        na_value = 'No Application'

        # Write data cells
        write_cell(ws, "H7", empty_value)
        write_cell(ws, "H42", empty_value)
        write_cell(ws, "H120", empty_value)
        write_cell(ws, "H200", empty_value)

        write_cell(ws, "H8", empty_value)
        write_cell(ws, "H43", empty_value)
        write_cell(ws, "H121", empty_value)
        write_cell(ws, "H201", empty_value)
        write_cell(ws, "H251", empty_value)

        write_cell(ws, "H9", empty_value)
        write_cell(ws, "H44", empty_value)
        write_cell(ws, "H122", empty_value)
        write_cell(ws, "H202", empty_value)
        write_cell(ws, "H252", empty_value)

        write_cell(ws, "H10", empty_value)
        write_cell(ws, "H45", empty_value)
        write_cell(ws, "H123", empty_value)
        write_cell(ws, "H203", empty_value)
        write_cell(ws, "H253", empty_value)

        write_cell(ws, "H11", empty_value)
        write_cell(ws, "H46", empty_value)
        write_cell(ws, "H124", empty_value)
        write_cell(ws, "H204", empty_value)
        write_cell(ws, "H254", empty_value)

        write_cell(ws, "M11", empty_value)
        write_cell(ws, "M46", empty_value)
        write_cell(ws, "M124", empty_value)
        write_cell(ws, "M204", empty_value)
        write_cell(ws, "M254", empty_value)

        write_cell(ws, "AG6", empty_value)
        write_cell(ws, "AG41", empty_value)
        write_cell(ws, "AG119", empty_value)
        write_cell(ws, "AG199", empty_value)
        write_cell(ws, "AG249", empty_value)

        write_cell(ws, "AG7", empty_value)
        write_cell(ws, "AG42", empty_value)
        write_cell(ws, "AG120", empty_value)
        write_cell(ws, "AG200", empty_value)
        write_cell(ws, "AG250", empty_value)

        write_cell(ws, "AG8", empty_value)
        write_cell(ws, "AG43", empty_value)
        write_cell(ws, "AG120", empty_value)
        write_cell(ws, "AG201", empty_value)
        write_cell(ws, "AG251", empty_value)

        write_cell(ws, "AG9", empty_value)
        write_cell(ws, "AG44", empty_value)
        write_cell(ws, "AG122", empty_value)
        write_cell(ws, "AG202", empty_value)
        write_cell(ws, "AG252", empty_value)

        write_cell(ws, "AG10", empty_value)
        write_cell(ws, "AG45", empty_value)
        write_cell(ws, "AG123", empty_value)
        write_cell(ws, "AG203", empty_value)
        write_cell(ws, "AG253", empty_value)

        write_cell(ws, "AG11", empty_value)
        write_cell(ws, "AG46", empty_value)
        write_cell(ws, "AG124", empty_value)
        write_cell(ws, "AG204", empty_value)
        write_cell(ws, "AG254", empty_value)

    except Exception as e:

        show_info(f"Failed to write data, error: {e}")
        print(f"Failed to write data, error: {e}")
        return False

    return True


