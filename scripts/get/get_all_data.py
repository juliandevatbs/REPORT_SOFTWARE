import os
import sys

from openpyxl import load_workbook

from scripts.excel.get_sheet_names import get_sheet_names

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def get_chain_data():

    sheetname = 'Chain of Custody 1'
    route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx"

    all_data = []

    wb = None

    try:

        # Verify file exists
        if not os.path.exists(route):
            print("Error: File not found")
            return None

        # Open workbook in read-only mode with data only (no formulas)
        wb = load_workbook(filename=route, read_only=True, data_only=True)

        # Verify sheet exists
        if sheetname not in wb.sheetnames:
            print(f"Error: Worksheet '{sheetname}' not found")
            return None

        ws = wb[sheetname]

        # The important data starts at this row
        start_row = 15

        specific_sheet = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']

        columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'Y']

        matrix_codes = {

            'A': 'Air',
            'GW': 'Groundwater',
            'SE': 'Sediment',
            'SO': 'Soil',
            'SW': 'Surface Water',
            'W': 'Water (Blanks)',
            'HW': 'Potencial Haz Wastw',
            'O': 'Other'

        }

        is_data = True

        while True:

            current_cell  = ws[f'B{start_row}'].value

            if current_cell is None or current_cell == '' or current_cell == 'Shipment Method:':

                is_data = False

                break

            row = []

            for column in columns:

                cell_value = ws[f'{column}{start_row}'].value

                if column == 'G':

                    matrix_id = matrix_codes[cell_value]

                    row.append(matrix_id)
                else:

                    row.append(cell_value)

            for sheet in specific_sheet:

                cell = ws[f'{sheet}{start_row}'].value

                if cell == 1:

                    sheet_name = ws[f'{sheet}13'].value
                    number_sheet = ws[f'{sheet}12'].value

                    final_sheet = f'{sheet_name} ({number_sheet})'

                    row.append(final_sheet)

            all_data.append(row)

            start_row += 1

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

    finally:

        if wb is not None:

            try:

                wb.close()
            except Exception as e:

                print(f"Error closing workbook: {e}")

    #print(all_data)
    return all_data


def get_matrix_data(chain_data: list):
    route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx"

    if not os.path.exists(route):
        print("Error: File not found")
        return None

    try:
        wb = load_workbook(filename=route, read_only=True, data_only=True)
        sheets_in_excel = get_sheet_names(route)

        cells_constant_value = [
            'N19', 'N20', 'N21', 'N22', 'N23', 'N24'
        ]






        for row in chain_data:
            constant_values = []



            #print(row)
            if len(row) < 1:
                continue

            sheet_name = row[-1]

            ws = wb[sheet_name]
            analysis_requested = ws['M7'].value

            for constant in cells_constant_value:
                constant_values.append(ws[constant].value)
            row.extend(constant_values)


            if sheet_name in sheets_in_excel:
                try:
                    ws = wb[sheet_name]
                    start_row = 26
                    columns = ['B', 'C', 'D', 'E', 'F', 'H', 'I', 'J']
                    found = False

                    #print(analysis_requested)
                    row.append(analysis_requested)





                    while True:
                        if ws[f'B{start_row}'].value == 'APPROVED BY':
                            break

                        cell_value = ws[f'C{start_row}'].value
                        if cell_value == row[1]:
                            row_data = []
                            for column in columns:
                                row_data.append(ws[f'{column}{start_row}'].value)
                            row.extend(row_data)

                            found = True
                            break

                        start_row += 1
                        if start_row > 1000:
                            break


                except Exception as e:
                    print(f"Error processing sheet {sheet_name}: {e}")
                    continue

    except Exception as e:
        print(f"An error occurred: {e}")
        return None
    finally:
        if 'wb' in locals():
            try:
                wb.close()
            except Exception as e:
                print(f"Error closing workbook: {e}")

    for row in chain_data:
        print(row)

    return chain_data

get_matrix_data(get_chain_data())




