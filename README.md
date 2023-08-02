# PDF-invoices Package

This package contains the function, generate, which extracts invoice data from
spreadsheet files and converts that data into professional PDF invoices. The
function can be imported using the command below:

    from PDF_invoices import generate

Usage:

    generate(invoices_path, pdfs_path, company_logo, optional **kwargs)

The generate function requires the directory path for the spreadsheet invoices, the destination folder into which
the function will print the PDF invoices, and the relative filepath to an icon that will be used as a company
logo in each PDF invoice. There are no default values for these three arguments.

The remaining arguments are optional values for the columns and content in the printed PDFs. product_id, product_name, 
amount_purchase, and price_per_unit must match the corresponding column headings in the spreadsheet invoices or the  
wrong data will be extracted. currency_type is used to set the type of currency that will appear in the 
final PDF. There are default values for all of these strings, as they are needed for headings and PDF content.

Before column headers are used in the final PDF, underscores are converted to spaces, the str.title() method is 
applied, and 'id' is converted to 'ID' if it appears in a string.

Arguments:
- invoices_path: str, directory path, REQUIRED
- pdfs_path: str, directory path, REQUIRED
- company_logo: str, filepath to company logo icon, REQUIRED
- product_id: str, defaults to "product_id"
- product_name: str, defaults to "product_name"
- amount_purchased: str, defaults to "amount_purchased"
- price_per_unit: str, defaults to "price_per_unit"
- total_price: str, defaults to "total_price"
- currency_type: str, defaults to "USD" for US dollars

This program returns None.