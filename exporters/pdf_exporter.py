import os
import pythoncom
from docx2pdf import convert


def convert_docx_to_pdf(docx_path: str) -> str:
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    pythoncom.CoInitialize()
    try:
        convert(docx_path, pdf_path)
    finally:
        pythoncom.CoUninitialize()

    return pdf_path