import pdfplumber
from balance_app import group_lines, find_columns

for name in ["demo_statement.pdf", "sample_A_us_style.pdf",
             "sample_B_uk_style.pdf", "sample_C_single_amount.pdf"]:
    with pdfplumber.open(name) as pdf:
        lines = group_lines(pdf.pages[0].extract_words())
        print(name, find_columns(lines))