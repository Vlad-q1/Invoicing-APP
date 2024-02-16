import pandas as pd
import os
import locale
import zipfile
from mailmerge import MailMerge
from docx2pdf import convert
from queue import Queue

error_queue = Queue()

class MissingFieldError(Exception):
    pass

def merge_word_tempate(template_path, excel_path):
    
    locale.setlocale(locale.LC_ALL, 'ro_RO.UTF-8')
    
    data = pd.read_excel(excel_path)
    
    data.columns = data.columns.str.replace(' ', '_').str.replace("=", '_')
    
    required_fields = ['invoice_number', 'Month', 'DCU', 'print_value_ron', 'print_value_eur', 'total_in_ron_de_printat', 'print_value_eur_total', 'print_exchange_rate']
    
    if not all(field in data.columns for field in required_fields):
        raise MissingFieldError("One or more required fields are missing in the excel file")
    
    columns_to_convert = ['print_value_ron', 'print_value_eur', 'print_value_ron_1', 'total_in_ron_de_printat', 'print_value_eur_total', 'print_exchange_rate']
    
    os.makedirs('pdf', exist_ok=True)
    
    for index, row in data.iterrows():
        for column in columns_to_convert:
            if column in row and pd.notnull(row[column]):
                if column == 'print_exchange_rate':
                    row[column] = "{:,.4f}".format(row[column]).replace(",", " ").replace(".", ",").replace(" ", ".")
                else:
                    row[column] = "{:,.2f}".format(row[column]).replace(",", " ").replace(".", ",").replace(" ", ".")
        
        document = MailMerge(template_path)
        
        row_dict = row.astype(str).to_dict()
        
        row_dict = {k: v.replace('_', ' ') for k, v in row_dict.items()}
        
        document.merge(**row_dict)
        
        locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
        invoice_number = row['invoice_number']
        month = row['Month']
        dcu = row['DCU']
        locale.setlocale(locale.LC_TIME, 'ro_RO.UTF-8')
        
        file_name = f"{invoice_number} SAP Services {month} 2024 {dcu}" 
        
        docx_extension = "docx"
        doc_output_path = f"pdf/{file_name}.{docx_extension}"
        document.write(doc_output_path)
        
        pdf_extension = "pdf"
        pdf_output_path = f"pdf/{file_name}.{pdf_extension}"
        convert(doc_output_path, pdf_output_path)
        os.remove(doc_output_path)
    
    zip_pdfs('pdf', 'pdf')

def zip_pdfs(input_directory, output_directory):
    with zipfile.ZipFile(f'{output_directory}/invoices.zip', 'w') as zipf:
        for file in os.listdir(input_directory):
            if file.endswith('.pdf'):
                full_path = os.path.join(input_directory, file)
                zipf.write(full_path, arcname=file)
                os.remove(full_path)

def generate_invoice(excel_path, error_queue, close_callback=None):
        template_path = "invoice_template.docx"
        merge_word_tempate(template_path, excel_path)
        if close_callback:
            close_callback()
