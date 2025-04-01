import os
import sys
from datetime import datetime
from openpyxl import load_workbook

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def get_quality_data():

    sheet_name = "Chlorides (16887006)"
    route = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx"
    all_data = []

    try:
        # Verify file exists
        if not os.path.exists(route):
            print("Error: File not found")
            return None

        # Open workbook in read-only mode for better performance
        wb = load_workbook(filename=route, read_only=False, data_only=True)

        # Verify sheet exists
        if sheet_name not in wb.sheetnames:
            print(f"Error: Worksheet '{sheet_name}' not found")
            return None

        ws = wb[sheet_name]


        columns = ['B', 'C', 'D', 'E', 'F', 'H', 'I', 'J']
        start_row = 21

        # Process rows until empty B column or 'Shipment Method:' is found
        for row in range(5):

            b_val = ws[f'B{start_row}'].value

            # Extract values from specified columns
            row_data = []
            for col in columns:
                cell = ws[f'{col}{start_row}']
                row_data.append(cell.value)

            all_data.append(row_data)
            start_row += 1
        #print("All Data:", all_data)
        method = ws['M7'].value

        print(all_data)
        return all_data, method

    except Exception as e:
        print(f"Error: {str(e)}")
        # Print traceback for more detailed error information
        import traceback
        traceback.print_exc()
        return None

    finally:
        # Clean up resources
        if 'wb' in locals():
            try:
                wb.close()
            except Exception as e:
                print(f"Error closing workbook: {str(e)}")

get_quality_data()