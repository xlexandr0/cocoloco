import win32com.client as win32
import os
from tkinter import Tk, filedialog

def convertir_excels_a_pdf():
    # Seleccionar carpeta
    Tk().withdraw()
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos Excel")

    if not carpeta:
        print("No se seleccionó carpeta")
        return

    excel = win32.Dispatch("Excel.Application")

    # OPTIMIZACIONES
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    excel.Calculation = -4135  # xlCalculationManual
    excel.PrintCommunication = False

    anchos = [9.22, 9.11, 13.33, 12.11, 15.44, 38.89, 11.33, 17, 13.33]

    archivos = [f for f in os.listdir(carpeta) if f.endswith(".xlsx") and not f.startswith("~")]

    total = len(archivos)
    contador = 0

    for archivo in archivos:
        ruta = os.path.join(carpeta, archivo)

        try:
            wb = excel.Workbooks.Open(
                ruta,
                UpdateLinks=False,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False
            )

            ws = wb.Sheets(1)

            # Última fila en columna J
            ultima_fila = ws.Cells(ws.Rows.Count, "J").End(-4162).Row  # xlUp

            if ultima_fila >= 11:
                # Eliminar tablas existentes
                for tabla in ws.ListObjects:
                    if not excel.Intersect(tabla.Range, ws.Range(f"B11:J{ultima_fila}")) is None:
                        tabla.Unlist()

                # Crear tabla
                rango = ws.Range(f"B11:J{ultima_fila}")
                tabla = ws.ListObjects.Add(1, rango, None, 1)  # xlSrcRange, xlYes
                tabla.TableStyle = ""

                # Anchos
                for i, ancho in enumerate(anchos):
                    tabla.ListColumns(i + 1).Range.ColumnWidth = ancho

                # WrapText columnas B y G
                tabla.ListColumns(1).Range.WrapText = True
                tabla.ListColumns(1).Range.VerticalAlignment = -4160  # xlTop

                tabla.ListColumns(6).Range.WrapText = True
                tabla.ListColumns(6).Range.VerticalAlignment = -4160

                excel.PrintCommunication = True

                # PageSetup
                ps = ws.PageSetup
                ps.PrintArea = f"A1:J{ultima_fila}"
                ps.PrintTitleRows = "$11:$11"
                ps.CenterFooter = "Página &P"
                ps.Orientation = 1  # xlPortrait
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False
                ps.Zoom = False
                ps.PaperSize = 9  # xlPaperA4
                ps.LeftMargin = excel.CentimetersToPoints(1)
                ps.RightMargin = excel.CentimetersToPoints(1)
                ps.TopMargin = excel.CentimetersToPoints(1.5)
                ps.BottomMargin = excel.CentimetersToPoints(1.5)

                excel.PrintCommunication = False

                # Exportar PDF
                pdf = os.path.join(carpeta, archivo.replace(".xlsx", ".pdf"))
                ws.ExportAsFixedFormat(0, pdf)  # xlTypePDF

                contador += 1
                print(f"{contador} de {total} → {archivo}")

            wb.Close(False)

        except Exception as e:
            print(f"Error con {archivo}: {e}")

    # Restaurar Excel
    excel.PrintCommunication = True
    excel.Calculation = -4105  # xlCalculationAutomatic
    excel.EnableEvents = True
    excel.DisplayAlerts = True
    excel.ScreenUpdating = True
    excel.Quit()

    print("Proceso finalizado")

convertir_excels_a_pdf()
