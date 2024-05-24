import os
import pandas as pd
from docxtpl import DocxTemplate
from docx2pdf import convert

def load_invoice_data(excel_path):
    # Load data from 'InvoiceLog' sheet
    df_invoice = pd.read_excel(excel_path, sheet_name='InvoiceLog')
 
    # Convert DataFrame to a list of dictionaries
    invoice_data = df_invoice.to_dict(orient='records')
 
    return invoice_data
       
def render_invoice_template(template_path, data, output_folder):
    # Load the template based on the file extension
    _, template_extension = os.path.splitext(template_path)
    
    if template_extension.lower() == ".docx":
        # Load the Word template
        doc = DocxTemplate(template_path)
 
    # Render the template with the provided data
    doc.render(data)
 
    # Extract the Senders_Name company name from the data (adjust the key according to your data structure)
    Company_Name = data.get('Senders_Company', 'Unknown')
    Date = data.get('Date', 'Unknown')
    invoice_no = data.get('InvoiceNO', 'Unknown')
    contact_details = data.get('ContactDetails', 'Unknown')
    po_numbers = data.get('PO_Numbers', 'Unknown')
    line_items = data.get('LineItems', 'Unknown')
    services = data.get('Services', 'Unknown')
    service_price = data.get('Service_Price', 'Unknown')
    hours = data.get('Hours', 'Unknown')
    quantity = data.get('Quantity', 'Unknown')
    unit_price = data.get('Unit_Price', 'Unknown')
    total = data.get('Total', 'Unknown')
    receivers_company = data.get('Receivers_Company', 'Unknown')
    tax = data.get('Tax', 'Unknown')
    vat = data.get('VAT', 'Unknown')
    bank_name = data.get('Bank_Name', 'Unknown')
    bank_account_number = data.get('Bank_Account_Number', 'Unknown')
    currency = data.get('Currency', 'Unknown')
    exchange_currency = data.get('Exchange_Currency', 'Unknown')
    change = data.get('Change', 'Unknown')
    subtotals = data.get('Price_Subtotals', 'Unknown')
    totals = data.get('Totals', 'Unknown')
    
    # Save the rendered document
    docx_output_path = os.path.join(output_folder, f"Invoice_{Company_Name}.docx")
    doc.save(docx_output_path)
    print(f"Invoice saved as DOCX: {docx_output_path}")

    # Convert to PDF
    pdf_output_path = os.path.join(output_folder, f"Invoice_{Company_Name}.pdf")
    convert(docx_output_path, pdf_output_path)
    print(f"Invoice saved as PDF: {pdf_output_path}")

if __name__ == "__main__":
    # Specify the paths
    excel_path = "InvoiceDataLog.xlsx"  # Replace with your actual Excel file path

    # Word doc templates
    template_paths = [
        #"InvoiceWordTemplateDesigns\InvoiceSample1.docx",
       # "InvoiceWordTemplateDesigns\InvoiceSample2.docx",
       # "InvoiceWordTemplateDesigns\InvoiceSample3.docx",
      #  "InvoiceWordTemplateDesigns\InvoiceSample4.docx",
       # "InvoiceWordTemplateDesigns\InvoiceSample5.docx",
       # "InvoiceWordTemplateDesigns\InvoiceSample6.docx",
       # "InvoiceWordTemplateDesigns\InvoiceSample7.docx",
         #"InvoiceWordTemplateDesigns\InvoiceSample8.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample9.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample10.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample11.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample12.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample13.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample14.docx",
       # "InvoiceWordTemplateDesigns/InvoiceSample15.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample17.docx",
       # "InvoiceWordTemplateDesigns/InvoiceSample18.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample19.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample20.docx",
       # "InvoiceWordTemplateDesigns/InvoiceSample21.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample22.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample23.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample24.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample25.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample26.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample27.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample28.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample29.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample30.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample31.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample32.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample33.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample34.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample35.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample36.docx",
       # "InvoiceWordTemplateDesigns/InvoiceSample37.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample38.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample39.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample40.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample41.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample42.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample43.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample44.docx"
        #"InvoiceWordTemplateDesigns/InvoiceSample45.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample46.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample47.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample48.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample49.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample50.docx",
        #"InvoiceWordTemplateDesigns/InvoiceSample51.docx",
        #"Word Invoice Template Designs\InvoiceSample52.docx",
        #"Word Invoice Template Designs\InvoiceSample53.docx",
        #"Word Invoice Template Designs\InvoiceSample54.docx",
        #"Word Invoice Template Designs\InvoiceSample55.docx",
        #"Word Invoice Template Designs\InvoiceSample56.docx",
        

    ]

    # Output folders for Word and PDF files
    output_folder = "Generated Invoices Folder"

    # Create the output folders if they don't exist
    os.makedirs(output_folder, exist_ok=True)

    # Load data from Excel
    invoices_data = load_invoice_data(excel_path)

    for invoice_data in invoices_data:
        for template_path in template_paths:
            # Create the output folder for the template
            template_output_folder = os.path.join(output_folder, os.path.basename(template_path))
            os.makedirs(template_output_folder, exist_ok=True)

            # Render the invoice template
            render_invoice_template(template_path, invoice_data, template_output_folder)

            # Convert to PDF and save in the output folder
            pdf_output_path = os.path.join(template_output_folder, f"Invoice_{invoice_data.get('Senders_Company', 'Unknown')}.pdf")
            convert(os.path.join(template_output_folder, f"Invoice_{invoice_data.get('Senders_Company', 'Unknown')}.docx"), pdf_output_path)
            print(f"Invoice saved as PDF: {pdf_output_path}")




  