import sys
import os
from datetime import datetime
import traceback
import time
import subprocess
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from scripts.print.print_header_data import print_header_data
from scripts.print.print_lab_data import print_lab_data
from scripts.print.print_analytical_data import print_analytical_data
from scripts.print.print_summary_data import print_summary_data

def execute_all():

    print_header_data()
    print_lab_data()
    print_analytical_data()
    print_summary_data()

    return True







execute_all()
