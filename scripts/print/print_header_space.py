import os
import sys

from scripts.excel.connect_excel import get_excel
from scripts.utils.safe_save import safe_save_workbook

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.print.print_footer import print_footer

def header_space(wb, wsd, route_excel: str, init_row: int):

    try:


        header_source = wb["header_all_pages"]
        header_destination = wsd

        src_range = "A1:AQ13"
        destination_range = f"A{init_row}"

        last_cell = 13

        copy_range_with_styles(header_source, header_destination, src_range, destination_range)

        safe_save_workbook(wb, route_excel, 3)

        return last_cell

    except Exception as e:

        print(f"ERROR: {e}")






