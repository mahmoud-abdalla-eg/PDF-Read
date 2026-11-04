import sys
import os
import warnings
warnings.filterwarnings("ignore")

from pdf2docx import Converter

def convert_pdf_to_word(pdf_path: str, output_docx: str):
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"Input PDF not found: {pdf_path}")

    cv = Converter(pdf_path)
    try:
        cv.convert(output_docx)
    finally:
        cv.close()

    print(f"WORD_EXPORT_SUCCESS: {output_docx}")

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: convert_pdf_to_word.py <input_pdf> <output_docx>")
        sys.exit(1)
    convert_pdf_to_word(sys.argv[1], sys.argv[2])
