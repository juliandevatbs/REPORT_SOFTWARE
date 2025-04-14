import os
from openpyxl import load_workbook

from scripts.get.get_all_data import get_chain_data

# Prefijos para nombres de hoja
SHEET_PREFIXES = [
    "Alkalinity", "Ammonia", "Chlorides", "Nitrates",
    "Nitrites", "Oil & Grease", "Ortho-phosphates",
    "Sulfate", "Total Dissolved Solids", "Turbidity"
]

# Valores requeridos en columna B
REQUIRED_VALUES_B = {
    "Method Blank (MB)",
    "Laboratory Control Standard (LCS)",
    "QC",
    "Matrix Spike (MS)",
    "Matrix Spike Dup (MSD)"
}


def normalize_string(s):
    """Normaliza strings para comparación"""
    return " ".join(str(s).strip().split())


def generate_sample_id(base_id, increment):
    """Genera ID de muestra secuencial"""
    prefix, number = base_id.split('-')
    return f"{prefix}-{int(number) + increment:03d}"  # Formato 3 dígitos


def get_quality_data(chain_data, base_sample_id):
    excel_path = r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx"

    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Archivo no encontrado: {excel_path}")

    # Extraer nombres de hoja
    sheet_names = {item for sublist in chain_data for item in sublist[-3:]
                   if isinstance(item, str) and any(item.startswith(p) for p in SHEET_PREFIXES)}

    print(f"\nHojas a procesar: {sheet_names}")

    try:
        wb = load_workbook(excel_path, read_only=False, data_only=True)
        all_data = []
        current_increment = 1  # Comenzamos a incrementar desde 1

        for sheet_name in sheet_names:
            print(f"\nProcesando hoja: {sheet_name}")
            if sheet_name not in wb.sheetnames:
                print(f"Hoja no encontrada: {sheet_name}")
                continue

            ws = wb[sheet_name]
            constants = [ws[f"N{row}"].value for row in range(19, 25)]
            columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
            row_num = 21
            sheet_rows = []

            while row_num <= 100:  # Límite razonable
                cell_b = ws[f"C{row_num}"].value
                if cell_b is None:
                    break

                if normalize_string(cell_b).upper() == "SW":
                    row_num += 1
                    continue

                if any(normalize_string(cell_b) == req for req in REQUIRED_VALUES_B):
                    sample_id = generate_sample_id(base_sample_id, current_increment)
                    current_increment += 1

                    row_data = [ws[f"{col}{row_num}"].value for col in columns]
                    complete_row = [sample_id, sheet_name] + row_data + constants
                    sheet_rows.append(complete_row)
                    print(f"Fila {row_num}: ID {sample_id} | {normalize_string(cell_b)}")

                row_num += 1

            all_data.extend(sheet_rows)
            print(f"→ {len(sheet_rows)} filas válidas")

        return all_data

    except Exception as e:
        print(f"Error: {str(e)}")
        raise
    finally:
        if 'wb' in locals():
            wb.close()


def get_q_data(base_sample_id="2503014-020"):
    chain_data = get_chain_data()
    if not chain_data:
        print("Error: chain_data vacío")
        return []

    result = get_quality_data(chain_data, base_sample_id)



    for row in result:
        print(row)
    return result


if __name__ == "__main__":
    get_q_data(base_sample_id="2503014-020")