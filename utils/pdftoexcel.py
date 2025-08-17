import pandas as pd
import os
import pdfplumber
from datetime import datetime

def pdf_to_excel(pdf_path, excel_path=None):
    """
    Extracts tables from each page of a PDF document and writes them to an Excel file,
    with each table in its own sheet for proper columns and rows.

    Args:
        pdf_path (str): Path to the input PDF file.
        excel_path (str, optional): Path to the output Excel file. If None, a timestamped file is created.
    """
    if not os.path.exists(pdf_path):
        print(f"Error: Input PDF file '{pdf_path}' not found.")
        return

    if excel_path is None:
        excel_filename = f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        excel_path = os.path.join(os.path.dirname(pdf_path), excel_filename)

    try:
        with pdfplumber.open(pdf_path) as pdf, pd.ExcelWriter(excel_path) as writer:
            table_found = False
            for i, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables()
                for j, table in enumerate(tables):
                    if table:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        sheet_name = f'Page_{i}Table{j+1}'
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        table_found = True
            if table_found:
                print(f"Extracted tables and saved to {excel_path}")
            else:
                print("No tables found in the PDF document.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
