import pdfplumber
from balance_app import group_lines, find_columns, read_date

# for name in ["demo_statement.pdf", "sample_A_us_style.pdf",
#              "sample_B_uk_style.pdf", "sample_C_single_amount.pdf"]:
#     with pdfplumber.open(name) as pdf:
#         lines = group_lines(pdf.pages[0].extract_words())
#         print(name, find_columns(lines))



from balance_app import read_date

# print()
# for name in ["demo_statement.pdf", "sample_A_us_style.pdf",
#              "sample_B_uk_style.pdf", "sample_C_single_amount.pdf"]:
#     with pdfplumber.open(name) as pdf:
#         count = 0
#         for page in pdf.pages:
#             for ws in group_lines(page.extract_words()):
#                 if read_date(ws):
#                     count += 1
#         print(name, count, "rows with a date")


# from balance_app import read_row

# print()
# for name in ["demo_statement.pdf", "sample_A_us_style.pdf",
#              "sample_B_uk_style.pdf", "sample_C_single_amount.pdf"]:
#     rows, cols = [], None
#     with pdfplumber.open(name) as pdf:
#         for page in pdf.pages:
#             lines = group_lines(page.extract_words())
#             cols = find_columns(lines) or cols
#             for ws in lines:
#                 row = read_row(ws, cols)
#                 if row:
#                     rows.append(row)
#     print(name, len(rows), "rows")
#     print("   ", rows[0])

from balance_app import read_row, slash_format, newest_first, DATE_FORMATS

print()
for name in ["demo_statement.pdf", "sample_A_us_style.pdf", "sample_B_uk_style.pdf",
             "sample_C_single_amount.pdf", "sample_D_pk_dmy_with_tax.pdf"]:
    rows, cols = [], None
    with pdfplumber.open(name) as pdf:
        pages_words = [page.extract_words() for page in pdf.pages]
    formats = DATE_FORMATS + [slash_format(pages_words)]
    for words in pages_words:
        lines = group_lines(words)
        cols = find_columns(lines) or cols
        for ws in lines:
            row = read_row(ws, cols, formats)
            if row:
                rows.append(row)
    rows = newest_first(rows)
    print(name, len(rows), "rows", rows[0]["date"], "to", rows[-1]["date"])