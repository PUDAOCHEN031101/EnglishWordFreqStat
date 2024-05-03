import os
import shutil
import win32com.client as win32
import time

def convert_doc_to_docx(doc_path, docx_path):
    """Converts a DOC document to DOCX format."""
    try:
        word = win32.DispatchEx('Word.Application')
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(docx_path, FileFormat=16)  # FileFormat=16 for DOCX format
        doc.Close()
        word.Quit()
        return True
    except Exception as e:
        print(f"Error converting DOC to DOCX: {str(e)}")
        return False
    finally:
        # Ensure Word closes properly in case of failure
        doc = None
        word = None

def convert_files_to_docx(folder_path, output_folder):
    """Converts DOC files in a folder to DOCX format and copies PDF and DOCX files."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        output_path = os.path.join(output_folder, os.path.splitext(filename)[0] + '.docx')
        if filename.lower().endswith('.doc'):
            # Adding a slight delay to help with COM interface stability
            time.sleep(1)
            if convert_doc_to_docx(file_path, output_path):
                print(f"Converted {filename} to DOCX and saved to {output_path}")
            else:
                print(f"Failed to convert {filename}")
        elif filename.lower().endswith('.pdf') or filename.lower().endswith('.docx'):
            shutil.copy2(file_path, output_folder)
            print(f"Copied {filename} to destination folder")
        else:
            print(f"Skipping {filename}: Unsupported file format")

# Example usage
source_folder = 'source_folder_path'  # Replace with your source folder path
destination_folder = 'destination_folder_path'  # Replace with your destination folder path
convert_files_to_docx(source_folder, destination_folder)
