from pdfminer.high_level import extract_text
import pandas as pd
import os
import io
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams

# File paths
pdf_path = "/Users/vedagya.saraswat/Porter/Ved experiments/pranjali_test.pdf"
output_path = "/Users/vedagya.saraswat/Porter/Ved experiments/extracted_tables.xlsx"

# Check if file exists
if not os.path.exists(pdf_path):
    print(f"File not found: {pdf_path}")
    output_path = None
else:
    try:
        # Use a more robust approach to handle non-ASCII85 digit errors
        resource_manager = PDFResourceManager()
        fake_file_handle = io.StringIO()
        converter = TextConverter(resource_manager, fake_file_handle, laparams=LAParams())
        page_interpreter = PDFPageInterpreter(resource_manager, converter)
        
        with open(pdf_path, 'rb') as file:
            for page in PDFPage.get_pages(file, caching=True, check_extractable=True):
                try:
                    page_interpreter.process_page(page)
                except Exception as e:
                    print(f"Warning: Error processing page: {e}")
                    continue
                    
        extracted_text = fake_file_handle.getvalue()
        converter.close()
        fake_file_handle.close()

        # Split text into lines
        lines = extracted_text.split("\n")

        # Convert lines to structured table format
        table_data = [line.split() for line in lines if line.strip()]  # Splitting words intelligently

        # Convert to DataFrame
        df = pd.DataFrame(table_data)

        # Save extracted table to Excel
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Extracted_Table", index=False, header=False)

        print(f"Tables successfully extracted to {output_path}")

    except Exception as e:
        print(f"Error processing PDF: {e}")
        output_path = None

print(output_path)