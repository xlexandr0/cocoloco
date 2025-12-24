📄 Excel a PDF (Python + Tkinter)

Aplicación de escritorio en Python que convierte archivos Excel (.xlsx) a PDF, usando Microsoft Excel para mantener el formato original de impresión.
Incluye interfaz gráfica con Tkinter.

✅ Requisitos

Windows

Microsoft Excel instalado

Python 3.x

📦 Instalación

Instalar la dependencia necesaria con pip:

pip install pywin32


⚠️ tkinter ya viene incluido con Python, no se instala con pip.

▶️ Uso

Ejecutar el programa:

python excel_a_pdf.py


Se abrirá una ventana.

Hacer clic en “Seleccionar carpeta y convertir”.

Elegir la carpeta con los archivos Excel.

Los PDFs se generarán en la misma carpeta.

🖥️ Interfaz

La aplicación muestra una ventana simple con:

Un botón para seleccionar la carpeta

Conversión automática de todos los Excel a PDF

📝 Qué hace el programa

Procesa todos los archivos .xlsx

Ignora archivos temporales (~$)

Usa la primera hoja

Detecta datos desde la fila 11

Crea una tabla B11:J

Ajusta anchos de columnas

Configura impresión A4

Exporta a PDF sin guardar cambios en Excel

📂 Resultado
archivo.xlsx → archivo.pdf

⚠️ Notas

Solo funciona en Windows

Excel se ejecuta en segundo plano

No modifica ni guarda los archivos originales

Coco malditasea
