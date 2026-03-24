import os


def is_pdf_export_supported() -> bool:
    return os.name == "nt"


def convert_docx_to_pdf(docx_path: str) -> str:
    if not is_pdf_export_supported():
        raise RuntimeError("PDF-Export ist nur unter Windows mit Microsoft Word verfuegbar.")

    import pythoncom
    from docx2pdf import convert

    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    pythoncom.CoInitialize()
    try:
        convert(docx_path, pdf_path)
    finally:
        pythoncom.CoUninitialize()

    return pdf_path
