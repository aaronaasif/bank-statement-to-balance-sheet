import pdfplumber
from balance_app import group_lines, find_columns, read_date

for name in ["demo_statement.pdf", "sample_A_us_style.pdf",
             "sample_B_uk_style.pdf", "sample_C_single_amount.pdf"]:
    with pdfplumber.open(name) as pdf:
        lines = group_lines(pdf.pages[0].extract_words())
        print(name, find_columns(lines))



from balance_app import read_date

print()
for name in ["demo_statement.pdf", "sample_A_us_style.pdf",
             "sample_B_uk_style.pdf", "sample_C_single_amount.pdf"]:
    with pdfplumber.open(name) as pdf:
        count = 0
        for page in pdf.pages:
            for ws in group_lines(page.extract_words()):
                if read_date(ws):
                    count += 1
        print(name, count, "rows with a date")