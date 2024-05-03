import os
import shutil
import fitz  # PyMuPDF, ensure PyMuPDF is installed
from docx import Document
import re

def clean_text(value):
    """Clean the input text to be XML-compatible: Unicode or ASCII, no NULL bytes or control characters."""
    if value is None:
        return ''
    # Remove NULL bytes
    value = value.replace('\x00', '')
    # Remove control characters except tab, new line, and carriage return
    value = re.sub(r'[\x01-\x08\x0B\x0C\x0E-\x1F\x7F]', '', value)
    return value

def pdf_to_docx(pdf_path, docx_path):
    """Converts a PDF file to a DOCX file, including Chinese text handling."""
    document = Document()
    pdf_document = fitz.open(pdf_path)
    for page in pdf_document:
        try:
            # Extract text from each page and clean it
            text = clean_text(page.get_text())
            # Add a new paragraph for each page's text in the DOCX
            document.add_paragraph(text)
        except Exception as e:
            print(f"Error processing page: {e}")
    # Save the DOCX file
    document.save(docx_path)
    print(f"Converted {pdf_path} to {docx_path}")

def convert_folder_to_docx_and_copy(folder_path, output_folder):
    """Converts all PDF files in a folder to DOCX format and copies existing DOCX files."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.docx')
        if filename.lower().endswith('.pdf'):
            pdf_to_docx(file_path, output_path)
        elif filename.lower().endswith('.docx'):
            shutil.copy2(file_path, output_folder)
            print(f"Copied {filename} to destination folder")
        else:
            print(f"Skipping {filename}: Unsupported file format")

# Example usage
source_folder = 'source_folder_path'  # Replace with your source folder path
destination_folder = 'destination_folder_path'  # Replace with your destination folder path
convert_folder_to_docx_and_copy(source_folder, destination_folder)
