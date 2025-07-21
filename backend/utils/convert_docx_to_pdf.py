import subprocess
import os

def convert_docx_to_pdf(docx_path: str, output_dir: str) -> str:
    """
    Converts a DOCX file to PDF using LibreOffice in headless mode.
    Returns the path to the generated PDF.
    """
    try:
        subprocess.run([
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            docx_path,
            "--outdir", output_dir
        ], check=True)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Failed to convert DOCX to PDF: {e}")

    pdf_path = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    return pdf_path
