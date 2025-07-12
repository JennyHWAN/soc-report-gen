from docx import Document
import re

def extract_ma_ar_sections(doc: Document):
    """Split paragraphs into MA and AR sections based on headings."""
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    ma_text = []
    ar_text = []
    section = None

    for para in paragraphs:
        if re.search(r"管理层认定", para):
            section = "ma"
            continue
        elif re.search(r"审计师报告", para):
            section = "ar"
            continue

        if section == "ma":
            ma_text.append(para)
        elif section == "ar":
            ar_text.append(para)

    return ma_text, ar_text