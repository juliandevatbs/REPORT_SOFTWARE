import traceback
from copy import deepcopy
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.styles import Border


def copy_range_with_styles(src_ws, dst_ws, src_range, dst_cell):
    try:
        # Obtener límites del rango fuente
        src_start_col, src_start_row, src_end_col, src_end_row = range_boundaries(src_range)

        # Obtener coordenadas destino
        dst_col_letter, dst_row = coordinate_from_string(dst_cell)
        dst_col = column_index_from_string(dst_col_letter)

        # 1. Primero manejar las celdas combinadas
        for merged_range in src_ws.merged_cells.ranges:
            m_start_col, m_start_row, m_end_col, m_end_row = range_boundaries(str(merged_range))

            if not (m_end_row < src_start_row or m_start_row > src_end_row or
                    m_end_col < src_start_col or m_start_col > src_end_col):

                row_offset = dst_row - src_start_row
                col_offset = dst_col - src_start_col

                new_min_col = m_start_col + col_offset
                new_min_row = m_start_row + row_offset
                new_max_col = m_end_col + col_offset
                new_max_row = m_end_row + row_offset

                if (new_min_col > 0 and new_min_row > 0 and
                        new_max_col > 0 and new_max_row > 0):
                    new_range = f"{get_column_letter(new_min_col)}{new_min_row}:{get_column_letter(new_max_col)}{new_max_row}"
                    dst_ws.merge_cells(new_range)

        # 2. Copiar valores y estilos
        for row in range(src_start_row, src_end_row + 1):
            for col in range(src_start_col, src_end_col + 1):
                dest_row = row - src_start_row + dst_row
                dest_col = col - src_start_col + dst_col
                cell_coord = f"{get_column_letter(dest_col)}{dest_row}"

                src_cell = src_ws.cell(row=row, column=col)

                # Verificar si la celda destino está en un rango combinado
                is_merged = False
                main_cell = None
                for merged_range in dst_ws.merged_cells.ranges:
                    if cell_coord in merged_range:
                        is_merged = True
                        min_col, min_row, _, _ = range_boundaries(str(merged_range))
                        main_cell = dst_ws.cell(row=min_row, column=min_col)
                        break

                # Determinar la celda objetivo
                target_cell = main_cell if is_merged else dst_ws.cell(row=dest_row, column=dest_col)

                # Solo copiar si es la celda principal o no está combinada
                if not is_merged or target_cell.coordinate == cell_coord:
                    # Copiar valor
                    target_cell.value = src_cell.value

                    # Copiar estilos (usando deepcopy para los bordes)
                    if src_cell.has_style:
                        target_cell.font = deepcopy(src_cell.font)
                        target_cell.fill = deepcopy(src_cell.fill)

                        # Manejo especial para bordes
                        if src_cell.border:
                            target_cell.border = Border(
                                left=deepcopy(src_cell.border.left),
                                right=deepcopy(src_cell.border.right),
                                top=deepcopy(src_cell.border.top),
                                bottom=deepcopy(src_cell.border.bottom),
                                diagonal=deepcopy(src_cell.border.diagonal),
                                diagonal_direction=deepcopy(src_cell.border.diagonal_direction),
                                outline=deepcopy(src_cell.border.outline),
                                vertical=deepcopy(src_cell.border.vertical),
                                horizontal=deepcopy(src_cell.border.horizontal)
                            )

                        target_cell.alignment = deepcopy(src_cell.alignment)
                        target_cell.number_format = src_cell.number_format

        # 3. Copiar dimensiones
        for col in range(src_start_col, src_end_col + 1):
            src_dim = src_ws.column_dimensions.get(get_column_letter(col))
            if src_dim and src_dim.width:
                dest_col = col - src_start_col + dst_col
                dst_ws.column_dimensions[get_column_letter(dest_col)].width = src_dim.width

        for row in range(src_start_row, src_end_row + 1):
            src_dim = src_ws.row_dimensions.get(row)
            if src_dim and src_dim.height:
                dest_row = row - src_start_row + dst_row
                dst_ws.row_dimensions[dest_row].height = src_dim.height

        return True

    except Exception as e:
        print(f"Error durante la copia: {str(e)}")
        traceback.print_exc()
        return False