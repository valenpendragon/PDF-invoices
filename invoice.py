import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


def generate(invoices_path, pdfs_path,
             product_id="product_id",
             product_name="product_name",
             amount_purchased="amount_purchased",
             price_per_unit="price_per_unit",
             total_price="total_price",
             currency_type="USD"):
    """
    This function requires the directory path for the spreadsheet invoices and the destination
    folder into which the function will print the PDF invoices. The remaining arguments are
    optional values for the columns and content in the printed PDFs. product_id, product_name,
    amount_purchase, and price_per_unit must match the corresponding column headings in the
    spreadsheet invoices or the wrong data will be extracted. currency_type is used to set the
    type of currency that will appear in the final PDF. There are default values for all
    of these strings, as they are needed for headings and PDF content.

    Before column headers are used in the final PDF, underscores are converted to spaces, the
    str.title() method is applied, and 'id' is converted to 'ID' if it appears in a string.
    Arguments:
    :param invoices_path: str, directory path
    :param pdfs_path: str, directory path
    :param product_id: str, defaults to "product_id"
    :param product_name: str, defaults to "product_name"
    :param amount_purchased: str, defaults to "amount_purchased"
    :param price_per_unit: str, defaults to "price_per_unit"
    :param total_price: str, defaults to "total_price"
    :param currency_type: str, defaults to "USD" for US dollars
    :return: None
    """
    filepaths = glob.glob(f"{invoices_path}/*.xlsx")

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
            pdf.cell(w=30, h=8, txt=str(row[product_id]), border=1)
            pdf.cell(w=70, h=8, txt=str(row[product_name]), border=1)
            pdf.cell(w=35, h=8, txt=str(row[amount_purchased]), border=1, align="R")
            pdf.cell(w=30, h=8, txt=str(row[price_per_unit]), border=1, align="R")
            pdf.cell(w=30, h=8, txt=str(row[total_price]), ln=1, border=1, align="R")

        # Add the grand total lines to each invoice.
        invoice_total = df[total_price].sum()
        pdf.cell(w=30, h=8, border=1)
        pdf.cell(w=70, h=8, border=1)
        pdf.cell(w=35, h=8, border=1)
        pdf.cell(w=30, h=8, border=1)
        pdf.cell(w=30, h=8, txt=str(invoice_total), ln=1, border=1, align="R")

        # Add amount due statement,  and company logo.
        pdf.ln(15)
        pdf.set_font(family="Times", size=10, style="B")
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=f"The total amount due is {invoice_total} {currency_type}.", ln=1)
        pdf.set_font(family="Times", size=10, style="B")
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=25, h=8, txt=f"ABC Company")
        pdf.image("python-neon.png", w=10)

        pdf.output(f"{pdfs_path}/{filename}.pdf")
