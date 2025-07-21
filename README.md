## Structure
```
soc-report-gen/
├── app.py                                  ← Streamlit frontend
├── backend/
│   ├── __init__.py
│   ├── soc_report_gen.py                   ← Orchestrator (calls helpers)
│   ├── extract/
│   │   ├── __init__.py\
│   │   └── ma_ar_parser.py                 ← MA & AR parsing logic
│   ├── output/
│   │   ├── __init__.py
│   │   ├── word_generator.py               ← Word (.docx) generation
│   │   └── pdf_generator.py                ← PDF (.tex/.pdf) generation
│   └── utils/
│       └── latex_utils.py                  ← Unicode/LaTeX encoding, spacing logic, etc.
├── generated_reports/                      ← Will generate automatically to save final .docx and .pdf file
├── tests/
│   ├── __init__.py
│   └── test_extract_ma_ar_sections.py      ← 🧪 tests for ma_ar_parser
```

# Run the app
```
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

# Prerequisite for system
```
sudo apt install texlive-xetex fonts-arphic-ukai fonts-arphic-uming fonts-noto-cjk pandoc
sudo apt install texlive-fonts-recommended texlive-lang-chinese
sudo apt install fonts-noto-cjk fonts-texgyre
sudo apt install libreoffice
```

# Test
Using `pytest`
