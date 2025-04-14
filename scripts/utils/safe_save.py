import traceback
import subprocess
from openpyxl import load_workbook
from datetime import datetime
import sys
import os
import time

from scripts.utils.kill_excel_processes import kill_excel_processes

# Route configuration
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

"""
    This function implements a safe secure save 

"""


def safe_save_workbook(wb, route, max_attempts=3):
    for attempt in range(max_attempts):
        try:
            kill_excel_processes()
            time.sleep(1)

            temp_route = route + ".temp"
            wb.save(temp_route)

            if os.path.exists(route):
                os.remove(route)
            os.rename(temp_route, route)

            return True
        except Exception as e:
            time.sleep(2)
    return False