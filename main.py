    # Vamos a crear un programa que convierta una "factura" en Excell a factura el PDF.

from fpdf import FPDF
import pandas as pd
import glob  # para importar los archivos excell a pyhton.
from pathlib import Path

C = 0  # color.

filepaths = glob.glob(r"Invoices\*.xlsx")  # importamos los archivos que acaben en .xlsx (con '*.xlsx) de la carpeta Invoices.

# por cada archivo excell lo leemos y guardamos en una variable (tipo data frame de pandas) que va a ser una lista de diccionarios.
for filepath in filepaths:


    pdf = FPDF(orientation="P", unit="mm", format="A4")  # generamos la intancia pdf con el módulo FPDF.
    pdf.add_page()                                       # añadiremos una página pdf para cada factura (3)

    filename = Path(filepath).stem                       # con esto extraemos en nombre de la factura, que viene reflejado en el nombre que tiene el archivo y forma parte de la ruta. Por tanto extraemos dicho nombre desde la ruta.
    invoice_nr, invoice_date = filename.split("-")       # al hacer el split a filename se convierte en una lista y nos quedamos con el ínice 0 (número de factura)

    # creamos las celdas donde aparecerá el número de factura y la fecha, en cada página pdf.
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_date}", ln=1)

    # Creamos las cabeceras de cada tabla con los datos, una por cada página pdf (factura).
    excel_df = pd.read_excel(filepath, sheet_name="Sheet 1")        # estraemos cada tabla de las facturas en excell, con data frame de pandas.

    columns = list(excel_df.columns)                                # para las cabeceras de las tablas, extraemos las columnas del data frame antes creado y lo convermito en una lista.
    columns = [item.replace("_", " ").title() for item in columns]  # le damos formato a los elementos de la lista de columnas para que se muestren bien.
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(C, C, C,)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(w=36.5, h=8, txt=columns[0], align="C", fill=True, border=1)  # mostramos en cada página pdf las cabeceras de las tablas, una ves solo.
    pdf.cell(w=45, h=8, txt=columns[1], align="C", fill=True, border=1)
    pdf.cell(w=36.5, h=8, txt=columns[2], align="C", fill=True, border=1)
    pdf.cell(w=36.5, h=8, txt=columns[3], align="C", fill=True, border=1)
    pdf.cell(w=36.5, h=8, txt=columns[4], align="C", fill=True, border=1, ln=1)

    # recorremos los índices y filas del data frame antes creado  para luego mostrar en las páginas pdf todos los datos, fila por fila de cada factura.
    for index, row in excel_df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(C, C, C)
        pdf.cell(w=36.5, h=8, txt=str(row['product_id']), align="C", border=1)
        pdf.cell(w=45, h=8, txt=str(row['product_name']), align="C", border=1)
        pdf.cell(w=36.5, h=8, txt=str(row['amount_purchased']), align="C", border=1)
        pdf.cell(w=36.5, h=8, txt=str(row['price_per_unit']), align="C", border=1)
        pdf.cell(w=36.5, h=8, txt=str(row['total_price']), align="C", border=1, ln=1)

    # Mostramos en una nueva fila de la tabla sólo la suma total. Se muestra una vez por cada página pdf (factura).
    total_sum = sum(excel_df['total_price'])
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(C, C, C, )
    pdf.cell(w=36.5, h=8, txt="", border=1)  # mostramos en cada página pdf las cabeceras de las tablas, una vez solo.
    pdf.cell(w=45, h=8, txt="", border=1)
    pdf.cell(w=36.5, h=8, txt="", border=1)
    pdf.cell(w=36.5, h=8, txt="", border=1)
    pdf.cell(w=36.5, h=8, txt=str(total_sum), align="C", border=1, ln=1)

    # Añadimos fila, una sola por cada página pdf (factura), donde diga el total a pagar.
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(C, C, C, )
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(w=0, h=8, txt=f"The total price to pay is: {total_sum} $.", align="C")

    pdf.output(rf"PDFs\{filename}.pdf")