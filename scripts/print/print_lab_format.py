from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.excel.connect_excel import get_excel
from scripts.utils.safe_save import safe_save_workbook


def print_lab_format_row(wb, wsd, rows_l: int, route_excel: str, last_cell: int):

    try:


        header_lab_source = wb["block_lab"]
        header_lab_destination = wsd


        src_range = "A1:AP1"


        for i in range(rows_l):

            print(i)

            to_print_row = i + last_cell

            print(to_print_row)

            destination_range = f"A{to_print_row}"

            copy_range_with_styles(header_lab_source, header_lab_destination, src_range, destination_range)


        last_cell = 14

        safe_save_workbook(wb, route_excel, 3)

        return last_cell

    except Exception as e:

        print(f"ERROR: {e}")

