import traceback
import subprocess
from openpyxl import load_workbook
from datetime import datetime
import sys
import os
import time

"""Close all excel tasks"""

def kill_excel_processes():

    try:

        subprocess.run(["taskkill", "/f", "/im", "excel.exe"],
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)

        time.sleep(1)
    except:

        pass