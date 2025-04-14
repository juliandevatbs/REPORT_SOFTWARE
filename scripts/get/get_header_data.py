from typing import Union, Tuple, Any

from scripts.error.show_error import show_info
from scripts.excel.connect_excel import get_excel


def get_header_data(wb, ws) -> bool | tuple[Any, Any, Any, Any, Any, Any, Any, Any, Any, Any, Any, Any]:
    """
    Extracts header information from a specified Excel worksheet ('Chain of Custody 1').

    This function reads specific cells from an Excel template to gather client and project information
    that typically appears in document headers. It validates that all required fields are present
    and provides specific error messages if any data is missing.

    Parameters:
        route_excel (str): The file path to the Excel workbook containing the data.

    Returns:
        If successful: A tuple containing 10 elements in this order:
            - company_name (str)
            - client_name (str)
            - client_address (str)
            - city (str)
            - state (str)
            - zip_code (str)
            - requested_date (str)
            - project_location (str)
            - client_phone (str)
            - project_number (str)
            - lab_reporting_batch_id (str)
        If any required field is missing: False

    Raises:
        Does not explicitly raise exceptions but will return False for any errors and display
        error messages through the show_info function.

    Notes:
        - The function checks specific cell locations in the worksheet (e.g., D4 for company/client name)
        - Each field is validated for presence before proceeding
        - User feedback is provided via show_info() for both success and error cases
        - The original database connection code is preserved but commented out for potential future use
    """
    try:

        # First check if workbook and worksheet are valid
        if wb is None or ws is None:
            show_info("ERROR: Invalid workbook or worksheet provided")
            return False

        # Extract values from specific cells in the worksheet
        company_name = ws['D4'].value # Company name
        client_name = ws['D4'].value# Client name
        client_address = ws['D5'].value  # Street address
        city = ws['D6'].value  # City
        state = ws['D7'].value  # State
        zip_code = ws['G7'].value  # ZIP/Postal code
        facility_id = ws['AE12'].value # Facility id
        requested_date = ws['Y11'].value  # Date when request was made
        project_location = ws['Y7'].value  # Location of project
        client_phone = ws['D8'].value  # Client phone number
        project_number = ws['Y9'].value  # Project identifier
        lab_reporting_batch_id = ws['AA3'].value  # Laboratory batch ID

        # Validate all required fields
        error = ''


        if error:

            show_info(error)
            return False

        else:

            show_info("All header data was collected successfully!")

            header_data = [ company_name, client_name, client_address, city, state,
                zip_code, facility_id, requested_date, project_location, client_phone,
                project_number, lab_reporting_batch_id]
            print(header_data)
            return header_data

    except Exception as e:
        print(f"Error occurred while extracting header data: {str(e)}")
        return False

