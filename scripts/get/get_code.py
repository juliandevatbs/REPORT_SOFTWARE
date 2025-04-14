"""
    This function gets the code from chain of custody

"""
from scripts.excel.connect_excel import get_excel


def get_code(route:str):

    wb, ws = get_excel("Chain of Custody 1", route)

    try:

        code_value = ws["AA3"].value

        #print(code_value)

        return code_value

    except Exception as e:

        print(f"Error{str(e)}")
        raise

get_code("C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx")



