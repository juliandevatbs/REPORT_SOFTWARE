import os
import sys
import traceback
from openpyxl import load_workbook

def get_excel(sheet_name: str, route: str):
    """
    Opens an Excel workbook and retrieves a specific worksheet using openpyxl.
    
    Args:
        sheet_name (str): Name of the worksheet to retrieve
        route (str): Path to the Excel workbook file
        
    Returns:
        tuple: A tuple containing (workbook_object, worksheet_object) if successful,
               or (None, None) if an error occurs.
               
    Note:
        Unlike the win32com version, this doesn't return an Excel application object
        since openpyxl doesn't work with the Excel application directly.
    """
    wb = None
    ws = None

    try:
        # Check if the file exists
        if not os.path.exists(route):
            raise FileNotFoundError(f"The file {route} does not exist")
        
        # Open the workbook (read-only mode is optional but often recommended)
        wb = load_workbook(filename=route, read_only=False, keep_vba=False)
        
        # Get the worksheet by name
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Worksheet '{sheet_name}' not found in workbook")
            
        ws = wb[sheet_name]
        
        return wb, ws

    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        traceback.print_exc()
        
        # Clean up if an error occurred
        if wb is not None:
            try:
                wb.close()
            except:
                pass
        
        return None, None