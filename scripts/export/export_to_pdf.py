import os
import sys
import time
from pathlib import Path


def export_pdf_vertical(excel_path, pdf_path, sheet_name):
    try:
        # Verificar si el archivo Excel existe
        if not Path(excel_path).exists():
            raise FileNotFoundError(f"El archivo Excel {excel_path} no existe")

        # Intenta importar win32com solo cuando sea necesario
        try:
            import win32com.client
        except ImportError:
            raise ImportError("pywin32 no está instalado. Ejecuta: pip install pywin32")

        excel = None
        wb = None

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            # Convertir rutas a formato absoluto
            excel_path = str(Path(excel_path).resolve())
            pdf_path = str(Path(pdf_path).resolve())

            wb = excel.Workbooks.Open(excel_path)
            ws = wb.Worksheets(sheet_name)

            # Configuración dinámica del rango
            last_row = ws.Cells.Find("*", SearchOrder=1, SearchDirection=2).Row
            last_col = ws.Cells.Find("*", SearchOrder=2, SearchDirection=2).Column
            last_col_letter = chr(64 + last_col) if last_col <= 26 else 'AQ'

            ws.PageSetup.PrintArea = f"$B$1:${last_col_letter}${last_row}"
            ws.PageSetup.Orientation = 1 # Vertical
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesTall = False
            ws.PageSetup.FitToPagesWide = 1

            # Márgenes en puntos (1 cm = 28.35 puntos)
            ws.PageSetup.LeftMargin = 28.35
            ws.PageSetup.RightMargin = 28.35
            ws.PageSetup.TopMargin = 28.35
            ws.PageSetup.BottomMargin = 28.35

            # Exportar a PDF
            ws.ExportAsFixedFormat(0, pdf_path)

            print(f"✓ PDF generado correctamente: {pdf_path}")
            return True

        except Exception as e:
            print(f"✗ Error durante la generación del PDF: {str(e)}", file=sys.stderr)
            return False

        finally:
            # Cerrar todo correctamente
            if wb is not None:
                wb.Close(False)
            if excel is not None:
                excel.Quit()

            # Esperar un momento antes de matar el proceso
            time.sleep(1)

            # Método más confiable para cerrar Excel
            try:
                os.system('taskkill /IM EXCEL.EXE /F')
            except:
                pass

    except Exception as e:
        print(f"✗ Error inicial: {str(e)}", file=sys.stderr)
        return False