import os
import tempfile
import shutil
from typing import final

from docx import Document
import pandas as pd
import subprocess
import streamlit as st
from pathlib import Path
from backend.extract.ma_ar_parser import extract_ma_ar_sections
from pylatexenc.latexencode import unicode_to_latex
from backend.utils.latex_utils import (
    format_paragraphs_to_latex,
    latex_signature_block,
    latex_document_wrapper,
)

# === Utility Functions ===
def save_uploaded_file(uploaded_file, suffix):
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_file.write(uploaded_file.read())
    tmp_file.close()
    return tmp_file.name

def render_latex_to_pdf(tex_path):
    output_dir = os.path.dirname(tex_path)
    os.makedirs(output_dir, exist_ok=True)
    try:
        subprocess.run(["xelatex", "-output-directory", output_dir, tex_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError as e:
        st.error("❌ XeLaTeX compilation failed.")
        st.text(e.stderr.decode("utf-8"))
        raise
    except subprocess.TimeoutExpired:
        st.error("❌ XeLaTeX timed out.")
        raise
    return tex_path.replace(".tex", ".pdf")

def convert_tex_to_docx(tex_path):
    docx_path = tex_path.replace(".tex", ".docx")
    result = subprocess.run(
        ["pandoc", tex_path, "-o", docx_path],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    if result.returncode != 0:
        raise RuntimeError(f"Pandoc failed: {result.stderr.decode()}")
    return docx_path

def make_final_output_base(prefix):
    output_dir = os.path.join(os.getcwd(), "generated_reports")
    os.makedirs(output_dir, exist_ok=True)
    # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(output_dir, f"{prefix}")

# === Part I & II: MA & AR Word Input ===
def generate_part_i_ii(word_file):
    # Load document
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word:
        tmp_word.write(word_file.read())
        tmp_word_path = tmp_word.name

    doc = Document(tmp_word_path)

    try:
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

        tex_content = latex_document_wrapper(body)

        with tempfile.TemporaryDirectory() as tmp_dir:
            tex_path = os.path.join(tmp_dir, "part1_2.tex")
            with open(tex_path, "w", encoding="utf-8") as f:
                f.write(tex_content)

                with st.spinner("Processing Part I & II..."):
                    # st.write("Rendering LaTeX to PDF...")
                    pdf_path = render_latex_to_pdf(tex_path)
                    # st.write("Converting TeX to DOCX...")
                    docx_path = convert_tex_to_docx(tex_path)

                # === Copy only final outputs to persistent directory ===
                final_base = make_final_output_base("Part_I_II")
                shutil.copy(pdf_path, final_base + ".pdf")
                shutil.copy(docx_path, final_base + ".docx")

        # Temp files automatically cleaned
        return final_base

    finally:
        # ✅ This ensures the file is always cleaned up, even on error
        os.remove(tmp_word_path)

# === Part III & IV: Excel Input ===
def generate_part_iii_iv(excel_file):
    excel_path = save_uploaded_file(excel_file, ".xlsx")
    df = pd.read_excel(excel_path)

    # TODO: Convert each control/test into LaTeX entries
    tex_content = r"""
    \documentclass{article}
    \usepackage[margin=1in]{geometry}
    \begin{document}
    \section*{Description of Controls and Testing Procedures}
    Sample content from Excel goes here.
    \end{document}
    """

    tex_path = excel_path.replace(".xlsx", ".tex")
    with open(tex_path, "w") as f:
        f.write(tex_content)

    with st.spinner("Generating Part III & IV..."):
        pdf_path = render_latex_to_pdf(tex_path)
        docx_path = convert_tex_to_docx(tex_path)

    return tex_path.replace(".tex", "")

# === Final Report Assembly ===
def generate_final_report(files):
    combined_tex = r"""
    \documentclass{article}
    \usepackage[margin=1in]{geometry}
    \begin{document}
    \section*{SOC Report}
    Placeholder for merged content.
    \end{document}
    """
    tmp_dir = tempfile.mkdtemp()
    tex_path = os.path.join(tmp_dir, "soc_final.tex")
    with open(tex_path, "w") as f:
        f.write(combined_tex)

    with st.spinner("Generating final report..."):
        pdf_path = render_latex_to_pdf(tex_path)
        docx_path = convert_tex_to_docx(tex_path)

    return tex_path.replace(".tex", "")
