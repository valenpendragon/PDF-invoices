import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")

    filename = Path(filepath).stem
    invoice_no, invoice_date = filename.split("-")

    pdf.add_page()
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_no}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Clean up the appearance of the headers.
    headers = df.columns
    headers = [item.replace("_", " ").title().replace("Id", "ID") for item in headers]

    # Add the headers to the pdf.
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.set_fill_color(200, 200, 200)
    pdf.cell(w=30, h=8, txt=headers[0], border=1, fill=True)
    pdf.cell(w=70, h=8, txt=headers[1], border=1, fill=True)
    pdf.cell(w=35, h=8, txt=headers[2], border=1, fill=True, align="R")
    pdf.cell(w=30, h=8, txt=headers[3], border=1, fill=True, align="R")
    pdf.cell(w=30, h=8, txt=headers[4], ln=1, border=1, fill=True, align="R")

    # Add the data rows to the pdf.
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1, align="R")
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1, align="R")
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), ln=1, border=1, align="R")

    # Add the grand total lines to each invoice.
    invoice_total = df["total_price"].sum()
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=70, h=8, border=1)
    pdf.cell(w=35, h=8, border=1)
    pdf.cell(w=30, h=8, border=1)
    pdf.cell(w=30, h=8, txt=str(invoice_total), ln=1, border=1, align="R")

    # Add amount due statement,  and company logo.
    pdf.ln(15)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=f"The total amount due is {invoice_total} Euros.", ln=1)
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=25, h=8, txt=f"ABC Company")
    pdf.image("python-neon.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")
