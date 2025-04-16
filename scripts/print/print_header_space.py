import os
import sys
from scripts.excel.connect_excel import get_excel
from scripts.utils.safe_save import safe_save_workbook
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.print.print_footer import print_footer

def header_space(wb, wsd, init_row: int):

    try:
        header_source = wb["header_all_pages"]
        header_destination = wsd

        src_range = "A1:AQ13"
        destination_range = f"A{init_row}"

        last_cell = init_row + 13

        copy_range_with_styles(header_source, header_destination, src_range, destination_range)

        #safe_save_workbook(wb, route_excel, 3)
        company_spacing = init_row + 4 + 1
        clients_spacing  = company_spacing + 1
        adress_spacing = clients_spacing + 1
        city_spacing  = adress_spacing + 1
        state_spacing = city_spacing +1
        zip_spacing = state_spacing
        requested_data_spacing = init_row + 4
        facility_id_spacing = requested_data_spacing +1
        project_location_spacing = facility_id_spacing + 1
        client_phone_spacing = project_location_spacing + 1
        project_number_spacing = client_phone_spacing + 1
        lab_reporting_batch_spacing = project_number_spacing +1

        cell_mapping = {

            "company_name": [f"G{company_spacing}", f"G{company_spacing}", f"G{company_spacing}", f"G{company_spacing}"],
            "client_name": [f"G{clients_spacing}", f"G{clients_spacing}", f"G{clients_spacing}", f"G{clients_spacing}", f"G{clients_spacing}"],
            "client_address": [f"G{adress_spacing}", f"G{adress_spacing}", f"G{adress_spacing}", f"G{adress_spacing}", f"G{adress_spacing}"],
            "city": [f"G{city_spacing}", f"G{city_spacing}", f"G{city_spacing}", f"G{city_spacing}", f"G{city_spacing}"],
            "state": [f"G{state_spacing}", f"G{state_spacing}", f"G{state_spacing}", f"G{state_spacing}", f"G{state_spacing}"],
            "zip_code": [f"L{zip_spacing}", f"L{zip_spacing}", f"L{zip_spacing}", f"L{zip_spacing}", f"L{zip_spacing}"],
            "requested_data": [f"AF{requested_data_spacing}", f"AF{requested_data_spacing}", f"AF{requested_data_spacing}", f"AF{requested_data_spacing}", f"AF{requested_data_spacing}"],
            "facility_id": [f"AF{facility_id_spacing}", f"AF{facility_id_spacing}", f"AF{facility_id_spacing}", f"AF{facility_id_spacing}", f"AF{facility_id_spacing}"],
            "project_location": [f"AF{project_location_spacing}", f"AF{project_location_spacing}", f"AF{project_location_spacing}", f"AF{project_location_spacing}", f"AF{project_location_spacing}"],
            "client_phone": [f"AF{client_phone_spacing}", f"AF{client_phone_spacing}", f"AF{client_phone_spacing}", f"AF{client_phone_spacing}", f"AF{client_phone_spacing}"],
            "project_number": [f"AF{project_number_spacing}", f"AF{project_number_spacing}", f"AF{project_number_spacing}", f"AF{project_number_spacing}", f"AF{project_number_spacing}"],
            "lab_reporting_batch_id": [f"AF{lab_reporting_batch_spacing}", f"AF{lab_reporting_batch_spacing}", f"AF{lab_reporting_batch_spacing}", f"AF{lab_reporting_batch_spacing}", f"AF{lab_reporting_batch_spacing}"]

        }


        return last_cell

    except Exception as e:

        print(f"ERROR: {e}")






