from scripts.clean.clean_constant_data import safe_save_workbook
from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.excel.connect_excel import get_excel


def lab_space(wb, wsd, route_excel: str, last_cell: int):


    try:


        header_lab_source = wb["header_lab"]
        header_lab_destination = wsd

        src_range = "A1:AP2"
        destination_range = f"A{last_cell}"

        copy_range_with_styles(header_lab_source, header_lab_destination, src_range, destination_range)


        last_cell = 14

        safe_save_workbook(wb, route_excel, 3)

        return last_cell

    except Exception as e:

        print(f"ERROR: {e}")




