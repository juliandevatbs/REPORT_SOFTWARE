import win32com.client
import os
from pathlib import Path


def export_pdf_vertical(excel_path, pdf_path, sheet_name):

    try:

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        ws = wb.Worksheets(sheet_name)

        # 1. Configurar área de impresión exacta
        ws.PageSetup.PrintArea = "$B$1:$AQ$276"

        # 2. Ajustes CLAVE para orientación vertical
        ws.PageSetup.Orientation = 2  # 2 = Vertical (1=Horizontal)
        ws.PageSetup.Zoom = 50  # Desactivar zoom
        ws.PageSetup.FitToPagesTall = False  # No ajustar verticalmente
        ws.PageSetup.FitToPagesWide = 1  # Ajustar a 1 página de ancho

        ws.PageSetup.CenterHorizontally = True


        # 3. Márgenes mínimos (en puntos)
        ws.PageSetup.LeftMargin = 5  # (~2mm)
        ws.PageSetup.RightMargin = 5
        ws.PageSetup.TopMargin = 5
        ws.PageSetup.BottomMargin = 5

        # 4. Eliminar espacios adicionales
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically = True
        ws.PageSetup.LeftHeader = ""
        ws.PageSetup.RightHeader = ""

        # 5. Autoajustar columnas al contenido
        used_range = ws.Range("B1:AQ276")
        used_range.Columns.AutoFit()

        # Exportar
        # 0 = PDF
        ws.ExportAsFixedFormat(0, str(Path(pdf_path).resolve()))

        print(f"✓ PDF generado en VERTICAL correctamente: {pdf_path}")
        return True

    except Exception as e:

        print(f"✗ Error: {str(e)}")
        return False

    finally:

        if 'wb' in locals():

            wb.Close(False)

        if 'excel' in locals():

            excel.Quit()

        os.system('taskkill /f /im excel.exe')

export_pdf_vertical(

    excel_path=r"C:/Users/Duban Serrano/Desktop/REPORTES PYTHON/excel/Reporte 2025-03-12 (4).xlsx",
    pdf_path=r"C:\Users\Duban Serrano\Desktop\reporte.pdf",
    sheet_name="Reporte"

)