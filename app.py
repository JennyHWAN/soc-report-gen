from os import WCONTINUED

import streamlit as st
import tempfile
import os

from backend.soc_report_gen import generate_part_i_ii, generate_part_iii_iv
from backend.soc_report_gen import generate_part_i_ii, generate_part_iii_iv, generate_final_report

st.set_page_config(page_title="SOC Report Generator", layout="centered")
st.title("SOC Report Generator")

st.markdown("""
This application helps you generate a formatted SOC report (Parts I–V).
- Upload MA & AR (Part I & II) in Word format.
- Upload Description & Testing Procedures (Part III & IV) in Excel format.
- Generate final PDF and DOCX files.
""")

# Initialize state
if 'function_1' not in st.session_state:
    st.session_state.function_1 = ""
if 'function_2' not in st.session_state:
    st.session_state.function_2 = ""
if 'function_3' not in st.session_state:
    st.session_state.function_3 = ""
if 'generate_clicked_1' not in st.session_state:
    st.session_state.generate_clicked_1 = False
if 'generate_clicked_2' not in st.session_state:
    st.session_state.generate_clicked_2 = False
if 'generate_clicked_3' not in st.session_state:
    st.session_state.generate_clicked_3 = False
if 'output_path_1' not in st.session_state:
    st.session_state.output_path_1 = ""
if 'output_path_2' not in st.session_state:
    st.session_state.output_path_2 = ""
if 'output_path_3' not in st.session_state:
    st.session_state.output_path_3 = ""

def handle_generation_1():
    st.session_state.output_path_1 = generate_part_i_ii(word_file)
    st.session_state.generate_clicked_1 = True

def handle_generation_2():
    st.session_state.output_path_2 = generate_part_iii_iv(excel_file)
    st.session_state.generate_clicked_2 = True

def handle_generation_3():
    st.session_state.output_path_3 = generate_final_report(files)
    st.session_state.generate_clicked_3 = True

with st.expander("1. Upload and Generate Part I & II (MA & AR)"):
    word_file = st.file_uploader("Upload MA & AR Word File", type=["docx"])

    if word_file:
        # Detect new upload
        if st.session_state.function_1 != word_file.file_id:
            st.session_state.generate_clicked_1 = False
            st.session_state.function_1 = word_file.file_id

        # === Show Button Conditionally ===
        if not st.session_state.generate_clicked_1:
            st.button("Generate Part I & II", on_click=handle_generation_1())
        else:
            st.success("Formatted Part I & II generated!")

            # st.button("Formatted Part I & II generated, click here to download.")
            st.download_button("Download PDF", open(st.session_state.output_path_1 + ".pdf", "rb"),
                               file_name="Part_I_II.pdf")
            st.download_button("Download Word", open(st.session_state.output_path_1 + ".docx", "rb"),
                               file_name="Part_I_II.docx")
        # else:
        #     st.download_button("Download PDF", open(st.session_state.output_path + ".pdf", "rb"), file_name="Part_I_II.pdf")
        #     st.download_button("Download Word", open(st.session_state.output_path + ".docx", "rb"), file_name="Part_I_II.docx")

with (st.expander("2. Upload and Generate Part III & IV (Control + Test)")):
    excel_file = st.file_uploader("Upload Description & Testing Excel File", type=["xlsx"])
    if excel_file:
        # Detect new upload
        if st.session_state.function_2 != excel_file.file_id:
            st.session_state.generate_clicked_2 = False
            st.session_state.function_2 = excel_file.file_id

        # === Show Button Conditionally ===
        if not st.session_state.generate_clicked_2:
            st.button("Generate Part III & IV", on_click = handle_generation_2())
        else:
            st.success("Formatted Part III & IV generated!")
            st.download_button("Download PDF", open(st.session_state.output_path_2 + ".pdf", "rb"), file_name="Part_III_IV.pdf")
            st.download_button("Download Word", open(st.session_state.output_path_2 + ".docx", "rb"), file_name="Part_III_IV.docx")

with st.expander("3. Generate Final Report"):
    files = st.file_uploader("Upload All Parts (Word)", type=["docx"], accept_multiple_files=True)
    files_id = []
    for file in files:
        files_id.append(file.file_id)
    if files:
        if st.session_state.function_3 != files_id:
            st.session_state.generate_clicked_3 = False
            st.session_state.function_3 = files_id
        if not st.session_state.generate_clicked_3:
            st.button("Generate Final Report", on_click = handle_generation_3())
        else:
            st.success("Final Report generated!")
            st.download_button("Download Final PDF", open(st.session_state.output_path_3 + ".pdf", "rb"), file_name="SOC_Final_Report.pdf")
            st.download_button("Download Final Word", open(st.session_state.output_path_3 + ".docx", "rb"), file_name="SOC_Final_Report.docx")

st.markdown("---")
st.markdown("### 💬 Need Help?")
st.markdown("""
    <a href="mailto:jenny.yj.huang@cn.ey.com">
        <button class="contact-me-button">
            📧 Contact Me
        </button>
    </a>
    
    <style>
        .contact-me-button {
            background-color: #4CAF50;
            color: white;
            padding: 8px 16px;
            border: none;
            border-radius: 8px;
            transition: background-color 0.3s ease, color 0.3s ease;
            font-weight: 500;
        }
        .contact-me-button:hover {
            background-color: #007BFF;
            color: #ffffff;
            cursor: pointer;
        }
    </style>
""", unsafe_allow_html=True)