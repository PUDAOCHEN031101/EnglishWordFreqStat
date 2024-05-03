import os
from PyPDF2 import PdfReader
from docx import Document
import win32com.client as win32

def convert_pdf_to_txt(pdf_path):
    """Converts PDF document to TXT format while preserving gaps between words."""
    try:
        reader = PdfReader(pdf_path)
        text = []
        for page in reader.pages:
            extracted_text = page.extract_text()
            if extracted_text:
                text.append(extracted_text)
            else:
                print(f"Warning: No text found on one of the pages in {pdf_path}")
        return " ".join(text)
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {str(e)}")
        return None

def convert_docx_to_txt(docx_path):
    """Converts DOCX document to TXT format while preserving gaps between words."""
    try:
        doc = Document(docx_path)
        text = []
        for para in doc.paragraphs:
            if para.text:
                text.append(para.text)
        return " ".join(text)
    except Exception as e:
        print(f"Error reading DOCX {docx_path}: {str(e)}")
        return None

def convert_doc_to_txt(doc_path):
    """Converts DOC document to TXT format using win32com."""
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(doc_path)
        doc.Activate()
        text = doc.Content.Text
        doc.Close(False)
        word.Quit()
        return text
    except Exception as e:
        print(f"Error reading DOC {doc_path}: {str(e)}")
        return None

def convert_files(folder_path, output_folder):
    """Converts all PDF and Word documents in a folder to TXT format and saves them in a new folder."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        output_path = ""
        txt_content = ""
        if filename.lower().endswith('.pdf'):
            txt_content = convert_pdf_to_txt(file_path)
            output_path = os.path.join(output_folder, filename[:-4] + '.txt')
        elif filename.lower().endswith('.docx'):
            txt_content = convert_docx_to_txt(file_path)
            output_path = os.path.join(output_folder, filename[:-5] + '.txt')
        elif filename.lower().endswith('.doc'):
            txt_content = convert_doc_to_txt(file_path)
            output_path = os.path.join(output_folder, filename[:-4] + '.txt')
        else:
            continue  # Skip non-PDF/DOC/DOCX files
        if txt_content:
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(txt_content)
            print(f"Converted {filename} to TXT and saved to {output_path}")
        else:
            print(f"Failed to convert {filename}")

# Example usage
source_folder = 'source_folder_path'  # Replace with your source folder path
destination_folder = 'destination_folder_path'  # Replace with your destination folder path
convert_files(source_folder, destination_folder)
