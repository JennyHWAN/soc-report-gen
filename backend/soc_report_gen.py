import os
import tempfile
import shutil
from typing import final

from docx import Document
import pandas as pd
import subprocess
import streamlit as st
from backend.output.pdf_generator import generate_ma_ar_pdf
from backend.output.word_generator import generate_ma_ar_docx
from backend.utils.convert_docx_to_pdf import convert_docx_to_pdf

# === Utility Functions ===
def save_uploaded_file(uploaded_file, suffix):
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_file.write(uploaded_file.read())
    tmp_file.close()
    return tmp_file.name

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

output_dir = os.path.join(os.getcwd(), "generated_reports")
os.makedirs(output_dir, exist_ok=True)

def make_final_output_base(prefix):
    # timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(output_dir, f"{prefix}")

# === Part I & II: MA & AR Word Input ===
def generate_part_i_ii(word_file, base_name: str = "Part_I_II"):
    output_base = os.path.join("generated_reports", base_name)
    output_dir = os.path.dirname(output_base)

    os.makedirs(output_dir, exist_ok=True)

    # Save uploaded file to temporary disk
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word:
        tmp_word.write(word_file.read())
        tmp_word_path = tmp_word.name

    try:
        docx_path = generate_ma_ar_docx(tmp_word_path, word_file, output_base)
        pdf_path = convert_docx_to_pdf(docx_path, output_dir)
    finally:
        # ✅ This ensures the file is always cleaned up, even on error
        os.remove(tmp_word_path)

    return pdf_path, docx_path

# === Part III & IV: Excel Input ===
def generate_part_iii_iv(excel_file):
    excel_path = save_uploaded_file(excel_file, ".xlsx")
    df = pd.read_excel(excel_path)

    # TODO: Convert each control/test into LaTeX entries, check 'latex word generation'
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
