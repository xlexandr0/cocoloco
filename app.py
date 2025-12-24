import os
import win32com.client as win32
from tkinter import Tk, Button, Label, filedialog, messagebox


def convertir_excels_a_pdf():
    # Seleccionar carpeta
    carpeta = filedialog.askdirectory(title="Selecciona la carpeta con los archivos Excel")

    if not carpeta:
        return

    try:
        excel = win32.Dispatch("Excel.Application")
    except Exception:
        messagebox.showerror("Error", "No se pudo iniciar Excel. Verifica que esté instalado.")
        return

    # OPTIMIZACIONES
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    excel.Calculation = -4135  # xlCalculationManual
    excel.PrintCommunication = False

    anchos = [9.22, 9.11, 13.33, 12.11, 15.44, 38.89, 11.33, 17, 13.33]

    archivos = [
        f for f in os.listdir(carpeta)
        if f.endswith(".xlsx") and not f.startswith("~")
    ]

    if not archivos:
        messagebox.showinfo("Información", "No se encontraron archivos Excel.")
        return

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

            # Última fila con datos en columna J
            ultima_fila = ws.Cells(ws.Rows.Count, "J").End(-4162).Row  # xlUp

            if ultima_fila >= 11:
                # Eliminar tablas existentes
                for tabla in ws.ListObjects:
                    if excel.Intersect(tabla.Range, ws.Range(f"B11:J{ultima_fila}")):
                        tabla.Unlist()

                # Crear tabla
                rango = ws.Range(f"B11:J{ultima_fila}")
                tabla = ws.ListObjects.Add(1, rango, None, 1)  # xlSrcRange, xlYes
                tabla.TableStyle = ""

                # Aplicar anchos
                for i, ancho in enumerate(anchos):
                    tabla.ListColumns(i + 1).Range.ColumnWidth = ancho

                # WrapText columnas B y G
                tabla.ListColumns(1).Range.WrapText = True
                tabla.ListColumns(1).Range.VerticalAlignment = -4160  # xlTop

                tabla.ListColumns(6).Range.WrapText = True
                tabla.ListColumns(6).Range.VerticalAlignment = -4160

                excel.PrintCommunication = True

                # Configuración de página
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
                ruta_pdf = os.path.join(
                    carpeta, archivo.replace(".xlsx", ".pdf")
                )

                ws.ExportAsFixedFormat(0, ruta_pdf)  # xlTypePDF

                contador += 1

        except Exception as e:
            messagebox.showwarning(
                "Advertencia",
                f"Ocurrió un error con el archivo:\n{archivo}\n\n{e}"
            )

        finally:
            try:
                wb.Close(False)
            except Exception:
                pass

    # Restaurar Excel
    excel.PrintCommunication = True
    excel.Calculation = -4105  # xlCalculationAutomatic
    excel.EnableEvents = True
    excel.DisplayAlerts = True
    excel.ScreenUpdating = True
    excel.Quit()

    messagebox.showinfo(
        "Proceso finalizado",
        f"Se convirtieron {contador} de {total} archivos."
    )


# ================= INTERFAZ =================

root = Tk()
root.title("Excel a PDF")
root.geometry("350x160")
root.resizable(False, False)

Label(
    root,
    text="Conversor Excel a PDF",
    font=("Arial", 12, "bold")
).pack(pady=10)

Button(
    root,
    text="Seleccionar carpeta y convertir",
    width=30,
    height=2,
    command=convertir_excels_a_pdf
).pack(pady=20)

Label(
    root,
    text="Usa Microsoft Excel para mantener el formato",
    font=("Arial", 9)
).pack(pady=5)

root.mainloop()
