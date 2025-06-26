import os
import sys
from pathlib import Path
import PyPDF2

def pdf_to_text(pdf_path, txt_path):
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ''
            for page in reader.pages:
                text += page.extract_text() or ''
        with open(txt_path, 'w', encoding='utf-8') as out:
            out.write(text)
        print(f"Converted: {pdf_path} -> {txt_path}")
    except Exception as e:
        print(f"Error occurred: {pdf_path}: {e}")

def main(pdf_folder):
    pdf_folder = Path(pdf_folder)
    if not pdf_folder.exists() or not pdf_folder.is_dir():
        print("Please enter a valid folder path.")
        return
    pdf_files = list(pdf_folder.glob('*.pdf'))
    if not pdf_files:
        print("No PDF files found in the folder.")
        return
    for pdf_file in pdf_files:
        txt_file = pdf_file.with_suffix('.txt')
        pdf_to_text(pdf_file, txt_file)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pdf_to_text.py <pdf_folder>")
    else:
        main(sys.argv[1]) 