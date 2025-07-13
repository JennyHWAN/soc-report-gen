from pylatexenc.latexencode import unicode_to_latex

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

def format_paragraphs_to_latex(paragraphs, indent=False):
    """
    Safely include paragraphs in LaTeX using raw Unicode (for XeLaTeX).
    """
    return "\n\n".join(
        (r"\par\noindent " if indent else "") + p for p in paragraphs
    )

def latex_vspace_lines(lines=4):
    """
    Returns a vertical space block for a given number of lines (approx. 1em per line).
    """
    return "\n".join([r"\vspace*{3em}"] * lines)

def latex_signature_block(entity_name, lines_before=4):
    """
    Creates a signature block with spacing and organization name.
    """
    spacing = latex_vspace_lines(lines_before)
    return f"{spacing}\n{entity_name}"

def latex_document_wrapper(body: str, font_main='TeX Gyre Termes', font_cjk='Noto Sans CJK SC') -> str:
    """
    Wraps LaTeX body content in a complete XeLaTeX document structure.
    """
    return rf"""
    \documentclass[12pt]{{article}}
    \usepackage[margin=1in]{{geometry}}
    \usepackage{{xeCJK}}
    \usepackage{{fontspec}}
    \setmainfont{{{font_main}}}
    \setCJKmainfont{{{font_cjk}}}
    \clubpenalty=10000
    \widowpenalty=10000
    
    \begin{{document}}
    
    {body}
    
    \end{{document}}
    """
