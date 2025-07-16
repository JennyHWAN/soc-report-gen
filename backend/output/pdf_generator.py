# LaTeX-based PDF
import os
import tempfile
import shutil
from docx import Document
import streamlit as st

from backend.extract.ma_ar_parser import extract_ma_ar_sections
from backend.utils.latex_utils import (
    format_paragraphs_to_latex,
    latex_signature_block,
    latex_document_wrapper,
    render_latex_to_pdf
)


def generate_ma_ar_pdf(tmp_word_path, word_file, output_base_path: str) -> str:
    """
    Generates a PDF report (Part I & II) from an uploaded Word file.

    Args:
        tmp_word_path: The word uploaded.
        word_file: A file-like object containing the .docx Word input.
        output_base_path: Path without extension, e.g., 'generated_reports/Part_I_II'

    Returns:
        The full path to the generated PDF file.
    """
    doc = Document(tmp_word_path)

    # Extract and format text
    ma_text, ar_text = extract_ma_ar_sections(doc)
    ma_latex = format_paragraphs_to_latex(ma_text[:-1])
    ar_latex = format_paragraphs_to_latex(ar_text[:-3])

    body = rf"""
    \section*{{第一部分 – 管理层认定}}
    
    {ma_latex}
    
    {latex_signature_block(ma_text[-1], lines_before=4)}
    
    \newpage
    
    \section*{{第二部分 – 独立服务审计师报告}}
    
    {ar_latex}
    
    {latex_signature_block(r"\\".join(ar_text[-3:]), lines_before=8)}
    """

    # Wrap into full LaTeX document
    tex_content = latex_document_wrapper(body)

    # Render to PDF
    with tempfile.TemporaryDirectory() as tmp_dir:
        tex_path = os.path.join(tmp_dir, "part1_2.tex")
        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(tex_content)

            with st.spinner("Processing Part I & II..."):
                pdf_path = render_latex_to_pdf(tex_path)

        # === Copy only final outputs to persistent directory ===
        final_pdf_path = output_base_path + ".pdf"
        shutil.copy(pdf_path, final_pdf_path)
    # Temp files automatically cleaned
    return final_pdf_path