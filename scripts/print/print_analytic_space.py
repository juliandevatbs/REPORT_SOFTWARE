from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.excel.connect_excel import get_excel
from scripts.utils.safe_save import safe_save_workbook


def print_analyitic_space(wb, wsd, start_row: int, route_excel: str):

    try:

        header_lab_source = wb["header_block_analytic"]
        header_lab_destination = wsd


        src_range = "A1:AP2"
        destination_range = f"A{start_row}"


        copy_range_with_styles(header_lab_source, header_lab_destination, src_range, destination_range)


        last_cell = f"A{start_row + 2}"

        safe_save_workbook(wb, route_excel, 3)

        return last_cell

    except Exception as e:

        print(f"ERROR: {e}")

    return True
print_analyitic_space(54, r"C:\Users\Duban Serrano\Desktop\REPORTES PYTHON\excel\Reporte 2025-03-12 (4).xlsx")