from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.excel.connect_excel import get_excel
from scripts.utils.safe_save import safe_save_workbook


def print_summary_format(wb, wsd, start_row:int, route_excel: str, q_rows:int):
    try:


        header_lab_source = wb["block_analytic"]
        header_lab_destination = wsd

        src_range = "A1:AP4"

        for i in range(q_rows):

            to_print_row = (i*4) + start_row


            destination_range = f"A{to_print_row}"

            copy_range_with_styles(header_lab_source, header_lab_destination, src_range, destination_range)


        safe_save_workbook(wb, route_excel, 3)



    except Exception as e:

        print(f"ERROR: {e}")

