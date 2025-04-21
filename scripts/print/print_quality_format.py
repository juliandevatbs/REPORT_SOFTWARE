from scripts.copy_blocks.copy_block import copy_range_with_styles
from scripts.excel.connect_excel import get_excel
from scripts.utils.safe_save import safe_save_workbook


def print_quality_format(wb, wsd, init_row: int, quality_data: list):
    try:
        # Mapeo estricto (sin normalizar a minúsculas para evitar falsos positivos)
        FORMAT_QC_MAPPING = {
            "Method Blank (MB)": "Method Blank (MB)",
            "Laboratory Control Standard (LCS)": "Laboratory Control Standard",
            "QC": "QC",
            "Matrix Spike (MS)": "Matrix Spike (MS)",
            "Matrix Spike Dup (MSD)": "Matrix Spike Dup (MSD)"
        }

        print(f"Total de controles: {len(quality_data)}")
        
        src_range = "A1:AP6"
        current_row = init_row + 2  # Fila inicial
        
        for control in quality_data:
            print(f"\n--- Procesando control en FILA: {current_row} ---")
            print(f"Datos del control: {control}")
            
            quality_name = control[2].strip()
            if quality_name is None:
                continue
            
            # Buscar coincidencia EXACTA en el mapeo
            sheet_name = FORMAT_QC_MAPPING.get(quality_name)
            
            if not sheet_name:
                continue
            
            if sheet_name not in wb.sheetnames:
                continue
            
            # Copiar el bloque
            ws = wb[sheet_name]
            destiny_ws = wb["Final"]
            dst_cell = f"A{current_row}"
            
            copy_range_with_styles(ws, destiny_ws, src_range, dst_cell)
            
            current_row += 7  # Avanzar 7 filas para el siguiente control
        
        final_row = current_row
        return final_row
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return init_row