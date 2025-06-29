import streamlit as st
import tempfile
import os
# from backend.soc_report_generator import generate_part_i_ii, generate_part_iii_iv, generate_final_report

st.set_page_config(page_title="SOC Report Generator", layout="centered")
st.title("SOC Report Generator")

st.markdown("""
<style>
    /* Style expanders (background & text) */ 
    details summary:hover {
        color: #007BFF !important;
        cursor: pointer;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
This application helps you generate a formatted SOC report (Parts I–V).
- Upload MA & AR (Part I & II) in Word format.
- Upload Description & Testing Procedures (Part III & IV) in Excel format.
- Generate final PDF and DOCX files.
""")

with st.expander("1. Upload and Generate Part I & II (MA & AR)"):
    word_file = st.file_uploader("Upload MA & AR Word File", type=["docx"])
    if word_file and st.button("Generate Part I & II"):
        output_path = generate_part_i_ii(word_file)
        st.success("Formatted Part I & II generated!")
        st.download_button("Download PDF", open(output_path + ".pdf", "rb"), file_name="Part_I_II.pdf")
        st.download_button("Download Word", open(output_path + ".docx", "rb"), file_name="Part_I_II.docx")

with st.expander("2. Upload and Generate Part III & IV (Control + Test)"):
    excel_file = st.file_uploader("Upload Description & Testing Excel File", type=["xlsx"])
    if excel_file and st.button("Generate Part III & IV"):
        output_path = generate_part_iii_iv(excel_file)
        st.success("Formatted Part III & IV generated!")
        st.download_button("Download PDF", open(output_path + ".pdf", "rb"), file_name="Part_III_IV.pdf")
        st.download_button("Download Word", open(output_path + ".docx", "rb"), file_name="Part_III_IV.docx")

with st.expander("3. Generate Final Report"):
    files = st.file_uploader("Upload All Parts (Word)", type=["docx"], accept_multiple_files=True)
    if files and st.button("Generate Final Report"):
        final_path = generate_final_report(files)
        st.success("Final Report generated!")
        st.download_button("Download Final PDF", open(final_path + ".pdf", "rb"), file_name="SOC_Final_Report.pdf")
        st.download_button("Download Final Word", open(final_path + ".docx", "rb"), file_name="SOC_Final_Report.docx")

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