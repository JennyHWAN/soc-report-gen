import os
import tempfile
from docx import Document
import pandas as pd
import subprocess
import streamlit as st
from pathlib import Path

# === Utility Functions ===
def save_uploaded_file(uploaded_file, suffix):
    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_file.write(uploaded_file.read())
    tmp_file.close()
    return tmp_file.name

def render_latex_to_pdf(tex_path):
    output_dir = os.path.dirname(tex_path)
    subprocess.run(["pdflatex", "-output-directory", output_dir, tex_path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
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

# === Part I & II: MA & AR Word Input ===
def generate_part_i_ii(word_file):
    word_path = save_uploaded_file(word_file, ".docx")
    doc = Document(word_path)

    # TODO: parse and extract the text from Word and convert to LaTeX
    tex_content = r"""
    \documentclass{article}
    \usepackage[margin=1in]{geometry}
    \begin{document}
    \section*{Management's Assertion}
    Sample content from Word goes here.
    \section*{Auditor's Report}
    Sample content from Word goes here.
    \end{document}
    """

    tex_path = word_path.replace(".docx", ".tex")
    with open(tex_path, "w") as f:
        f.write(tex_content)

    with st.spinner("Generating Part I & II..."):
        pdf_path = render_latex_to_pdf(tex_path)
        docx_path = convert_tex_to_docx(tex_path)

    return tex_path.replace(".tex", "")

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
