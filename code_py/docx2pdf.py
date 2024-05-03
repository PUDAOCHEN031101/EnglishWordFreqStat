import os
import shutil
from docx import Document
from fpdf import FPDF

def convert_docx_to_pdf(docx_path, output_folder):
    """Converts a DOCX or DOC document to PDF format, skipping non-Word files."""
    try:
        # Attempt to open the document to check if it's a valid Word file
        doc = Document(docx_path)
    except Exception as e:
        print(f"Skipping file {os.path.basename(docx_path)}: Not a valid Word file. Error: {str(e)}")
        return False

    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
        pdf.set_font('DejaVu', size=12)
        for para in doc.paragraphs:
            pdf.multi_cell(0, 10, para.text)
        output_path = os.path.join(output_folder, os.path.basename(docx_path)[:-4] + '.pdf')
        pdf.output(output_path)
        print(f"Converted {os.path.basename(docx_path)} to PDF and saved to {output_path}")
        return True
    except Exception as e:
        print(f"Error converting {os.path.basename(docx_path)} to PDF: {str(e)}")
        return False

def copy_pdf_files(source_folder, destination_folder):
    """Copies PDF files from the source folder to the destination folder."""
    for filename in os.listdir(source_folder):
        if filename.lower().endswith('.pdf'):
            shutil.copy2(os.path.join(source_folder, filename), os.path.join(destination_folder, filename))
            print(f"Copied {filename} to {destination_folder}")

def process_files(source_folder, destination_folder):
    """Converts DOCX and DOC files to PDF and copies all PDFs to the destination folder."""
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    for filename in os.listdir(source_folder):
        file_path = os.path.join(source_folder, filename)
        if filename.lower().endswith('.docx') or filename.lower().endswith('.doc'):
            convert_docx_to_pdf(file_path, destination_folder)
        elif filename.lower().endswith('.pdf'):
            shutil.copy2(file_path, os.path.join(destination_folder, filename))
            print(f"Copied {filename} to {destination_folder}")

# Example usage
source_folder = 'source_folder_path'  # Replace with your source folder path
destination_folder = 'destination_folder_path'  # Replace with your destination folder path
process_files(source_folder, destination_folder)
