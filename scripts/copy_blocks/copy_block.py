import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import traceback


def copy_range_with_styles(src_ws, dst_ws, src_range, dst_cell):

    try:
        # Convertir coordenadas
        src_start_col, src_start_row, src_end_col, src_end_row = openpyxl.utils.range_boundaries(src_range)
        dst_col, dst_row, _, _ = openpyxl.utils.range_boundaries(dst_cell + ':' + dst_cell)

        # Primero copiar celdas combinadas
        for merged_range in src_ws.merged_cells.ranges:
            if (merged_range.min_row >= src_start_row and merged_range.max_row <= src_end_row and
                    merged_range.min_col >= src_start_col and merged_range.max_col <= src_end_col):
                # Calcular nuevo rango
                new_min_row = merged_range.min_row - src_start_row + dst_row
                new_max_row = merged_range.max_row - src_start_row + dst_row
                new_min_col = merged_range.min_col - src_start_col + dst_col
                new_max_col = merged_range.max_col - src_start_col + dst_col

                new_range = f"{get_column_letter(new_min_col)}{new_min_row}:{get_column_letter(new_max_col)}{new_max_row}"
                dst_ws.merge_cells(new_range)

        for row in range(src_start_row, src_end_row + 1):
            for col in range(src_start_col, src_end_col + 1):
                src_cell = src_ws.cell(row=row, column=col)

                if isinstance(src_cell, openpyxl.cell.cell.MergedCell):
                    continue

                dest_row = row - src_start_row + dst_row
                dest_col = col - src_start_col + dst_col
                dst_cell = dst_ws.cell(row=dest_row, column=dest_col)

                dst_cell.value = src_cell.value

                if src_cell.has_style:
                    dst_cell.font = Font(
                        name=src_cell.font.name,
                        size=src_cell.font.size,
                        bold=src_cell.font.bold,
                        italic=src_cell.font.italic,
                        color=src_cell.font.color
                    )
                    dst_cell.fill = PatternFill(
                        start_color=src_cell.fill.start_color,
                        end_color=src_cell.fill.end_color,
                        fill_type=src_cell.fill.fill_type
                    )
                    dst_cell.border = Border(
                        left=Side(style=src_cell.border.left.border_style, color=src_cell.border.left.color),
                        right=Side(style=src_cell.border.right.border_style, color=src_cell.border.right.color),
                        top=Side(style=src_cell.border.top.border_style, color=src_cell.border.top.color),
                        bottom=Side(style=src_cell.border.bottom.border_style, color=src_cell.border.bottom.color)
                    )
                    dst_cell.alignment = Alignment(
                        horizontal=src_cell.alignment.horizontal,
                        vertical=src_cell.alignment.vertical,
                        wrap_text=src_cell.alignment.wrap_text
                    )
                    dst_cell.number_format = src_cell.number_format

        for col in range(src_start_col, src_end_col + 1):
            col_letter = get_column_letter(col)
            if col_letter in src_ws.column_dimensions:
                new_col = get_column_letter(col - src_start_col + dst_col)
                dst_ws.column_dimensions[new_col].width = src_ws.column_dimensions[col_letter].width

        for row in range(src_start_row, src_end_row + 1):
            if row in src_ws.row_dimensions:
                new_row = row - src_start_row + dst_row
                dst_ws.row_dimensions[new_row].height = src_ws.row_dimensions[row].height

        return True

    except Exception as e:
        print(f"Error durante la copia: {str(e)}")
        traceback.print_exc()
        return False