# We are going to create a program that converts an Excel "invoice" to a PDF invoice.

from fpdf import FPDF
import pandas as pd
import glob  # to import excell files into python.
from pathlib import Path

C = 0  # color.

filepaths = glob.glob(
    r"Invoices\*.xlsx"
)  # import the files ending in .xlsx (with '*.xlsx') from the Invoices folder.

# for each excell file we read it and save it in a variable (pandas data frame type) that will be a list of dictionaries.
for filepath in filepaths:

    pdf = FPDF(
        orientation="P", unit="mm", format="A4"
    )  # generate the pdf instance with the FPDF module.
    pdf.add_page()  # we will add a pdf page for each invoice (3)

    filename = Path(
        filepath
    ).stem  # with this we extract the invoice name, which is reflected in the name of the file and is part of the path. Therefore, we extract that name from the path.
    invoice_nr, invoice_date = filename.split(
        "-"
    )  # when splitting filename it becomes a list, and we keep index 0 (invoice number)

    # create the cells where the invoice number and date will appear, on each pdf page.
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_date}", ln=1)

    # Create the headers of each table with the data, one for each pdf page (invoice).
    excel_df = pd.read_excel(
        filepath, sheet_name="Sheet 1"
    )  # extract each table from the invoices in excell, with pandas data frame.

    columns = list(
        excel_df.columns
    )  # for the table headers, we extract the columns from the data frame created above and convert it into a list.
    columns = [
        item.replace("_", " ").title() for item in columns
    ]  # we format the elements of the column list to display correctly.
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(
        C,
        C,
        C,
    )
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(
        w=36.5, h=8, txt=columns[0], align="C", fill=True, border=1
    )  # show the table headers on each pdf page, only once.
    pdf.cell(w=45, h=8, txt=columns[1], align="C", fill=True, border=1)
    pdf.cell(w=36.5, h=8, txt=columns[2], align="C", fill=True, border=1)
    pdf.cell(w=36.5, h=8, txt=columns[3], align="C", fill=True, border=1)
    pdf.cell(w=36.5, h=8, txt=columns[4], align="C", fill=True, border=1, ln=1)

    # we iterate over the indexes and rows of the data frame created above to then show all the data on the pdf pages, row by row of each invoice.
    for index, row in excel_df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(C, C, C)
        pdf.cell(w=36.5, h=8, txt=str(row["product_id"]), align="C", border=1)
        pdf.cell(w=45, h=8, txt=str(row["product_name"]), align="C", border=1)
        pdf.cell(
            w=36.5, h=8, txt=str(row["amount_purchased"]), align="C", border=1
        )
        pdf.cell(w=36.5, h=8, txt=str(row["price_per_unit"]), align="C", border=1)
        pdf.cell(
            w=36.5, h=8, txt=str(row["total_price"]), align="C", border=1, ln=1
        )

    # Show the total sum in a new row of the table. It is displayed once for each pdf page (invoice).
    total_sum = sum(excel_df["total_price"])
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(
        C,
        C,
        C,
    )
    pdf.cell(
        w=36.5, h=8, txt="", border=1
    )  # show the table headers on each pdf page, only once.
    pdf.cell(w=45, h=8, txt="", border=1)
    pdf.cell(w=36.5, h=8, txt="", border=1)
    pdf.cell(w=36.5, h=8, txt="", border=1)
    pdf.cell(w=36.5, h=8, txt=str(total_sum), align="C", border=1, ln=1)

    # Add a row, only one per pdf page (invoice), where it says the total to pay. And then we add another row where we put the name of the company that makes that invoice.
    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(
        C,
        C,
        C,
    )
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(
        w=0, h=8, txt=f"The total price to pay is: {total_sum} $.", align="C"
    )

    company_name = "Tech Solutions"
    pdf.set_font(family="Times", size=14, style="IB")
    pdf.set_text_color(
        C,
        C,
        C,
    )
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(w=0, h=8, ln=1)
    pdf.cell(w=0, h=8, txt=f"{company_name.upper()} Group.", align="L")

    pdf.output(rf"PDFs\{filename}.pdf")