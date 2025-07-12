from docx import Document
import re

def extract_ma_ar_sections(doc):
    """Split paragraphs into MA and AR sections based on headings."""
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    ma_text = []
    ar_text = []
    section = None

    for para in paragraphs:
        if re.search(r"^第?一部分|管理层认定", para):
            section = "ma"
            continue
        elif re.search(r"^第?二部分|审计师报告", para):
            section = "ar"
            continue

        if section == "ma":
            ma_text.append(para)
        elif section == "ar":
            ar_text.append(para)

    return ma_text, ar_text

if __name__ == '__main__':
    with open('/Users/jennyhuang/Documents/EY-project/SOC-report/SectionI&II.docx', 'rb') as f:
        extract_ma_ar_sections(f)