import os
import tempfile
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

def format_as_latex(text_list):
    # Escape LaTeX special characters and join paragraphs
    # Only escape actual LaTeX special characters — leave Chinese characters untouched
    def escape_latex(text):
        return (text.replace('\\', r'\textbackslash{}')
                .replace('&', r'\&')
                .replace('%', r'\%')
                .replace('$', r'\$')
                .replace('#', r'\#')
                .replace('_', r'\_')
                .replace('{', r'\{')
                .replace('}', r'\}')
                .replace('~', r'\textasciitilde{}')
                .replace('^', r'\textasciicircum{}'))

    return "\n\n".join(escape_latex(p) for p in text_list)

# === Part I & II: MA & AR Word Input ===
def generate_part_i_ii(word_file):
    word_path = save_uploaded_file(word_file, ".docx")
    doc = Document(word_path)

    ma_text, ar_text = extract_ma_ar_sections(doc)
    # ma_latex = format_paragraphs_to_latex(ma_text)
    # ar_latex = format_paragraphs_to_latex(ar_text)

    tex_content = rf"""
    \documentclass[12pt]{{article}}
    \usepackage[margin=1in]{{geometry}}
    \usepackage{{xeCJK}}
    \usepackage{{fontspec}}
    \setmainfont{{TeX Gyre Termes}}
    \setCJKmainfont{{Noto Sans CJK SC}}

    \clubpenalty=10000
    \widowpenalty=10000

    \begin{{document}}

    \section*{{第一部分 – 管理层认定}}

    {format_as_latex(ma_text)}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    上海外服（集团）有限公司

    \newpage

    \section*{{第二部分 – 独立服务审计师报告}}

    {format_as_latex(ar_text)}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    \vspace*{{3em}}

    安永华明会计师事务所（特殊普通合伙）上海分所

    \end{{document}}
    """

    tex_path = word_path.replace(".docx", ".tex")
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex_content)

    with st.spinner("Processing Part I & II..."):
        # st.write("Rendering LaTeX to PDF...")
        pdf_path = render_latex_to_pdf(tex_path)
        # st.write("Converting TeX to DOCX...")
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
