from docx import Document
import re

def extract_ma_ar_sections(doc: Document):
    """
        Split paragraphs into MA and AR sections based on keywords:
        - Start collecting MA section after detecting '管理层认定' or '第一部分'
        - Start collecting AR section after detecting '审计师报告' or '第二部分'
    """
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    ma_text = []
    ar_text = []
    section = None

    for para in paragraphs:
        # Detect the beginning of MA section
        if re.search(r"(管理层认定|第一部分)", para) and not ma_text:
            section = "ma"
            ma_text.append(para)
            continue

        # Detect the beginning of AR section
        elif re.search(r"(审计师报告|第二部分)", para) and not ar_text:
            section = "ar"
            ar_text.append(para)
            continue

        # Append to the correct section
        if section == "ma" and not ar_text:
            ma_text.append(para)
        elif section == "ar":
            ar_text.append(para)

    return ma_text, ar_text