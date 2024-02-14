# Produce by Joshua

# University of Aberdeen

# Development time: 2024/2/14 3:30

import PyPDF2
from pdfminer.high_level import extract_text, extract_pages
from pdfminer.layout import LTTextContainer, LTChar, LTRect, LTFigure
import pdfplumber
import pytesseract
import os


def extract_table(pdf_path, page_num, table_num):
    pdf = pdfplumber.open(pdf_path)
    table_page = pdf.pages[page_num]
    table = table_page.extract_tables()[table_num]
    return table


def table_converter(table):
    table_string = ''
    for row_num in range(len(table)):
        row = table[row_num]
        cleaned_row = [item.replace('\n', '') if item is not None and '\n' in item else 'None' if item is None else item for item in row ]
        table_string += ('|'+'|'.join(cleaned_row)+'|'+'\n')

    table_string = table_string[:-1]
    return table_string



pdf_path = '../resources/Centurion.pdf'
pdf_Obj = open(pdf_path, 'rb')
pdf_reader = PyPDF2.PdfReader(pdf_Obj)

text_per_page = {}
combined_text = ""

for pagenum, page in enumerate(extract_pages(pdf_path)):
    pageobj = pdf_reader.pages[pagenum]
    page_text = []
    line_format = []
    text_from_table = []
    page_content = []
    table_num = 0
    first_element = True
    table_extraction_flag = False
    pdf = pdfplumber.open(pdf_path)
    page_tables = pdf.pages[pagenum]
    tables = page_tables.find_tables()
    page_elements = [(element.y1, element) for element in page._objs]
    page_elements.sort(key=lambda a: a[0], reverse=True)

    for i, component in enumerate(page_elements):
        pos = component[0]
        element = component[1]
        if isinstance(element, LTRect):
            if first_element == True and (table_num + 1) <= len(tables):
                lower_side = page.bbox[3] - tables[table_num].bbox[3]
                upper_side = element.y1
                table = extract_table(pdf_path, pagenum, table_num)
                table_string = table_converter(table)
                text_from_table.append(table_string)
                page_content.append(table_string)
                table_extraction_flag = True
                first_element = False
                page_text.append('table')
                line_format.append('table')

                if element.y0 >= lower_side and element.y1 <= upper_side:
                    pass
                elif not isinstance(page_elements[i + 1][1], LTRect):
                    table_extraction_flag = False
                    first_element = True
                    table_num += 1

    dctkey = 'Page_' + str(pagenum)
    text_per_page[dctkey] = [page_text, line_format, text_from_table, page_content]

    # Append page text to the combined_text variable
    combined_text += ''.join(text_per_page[dctkey][3])

pdf_Obj.close()

# Save combined text to a single file
output_file = '../output/Centurion_output.txt'
with open(output_file, 'w', encoding='utf-8') as output_file:
    output_file.write(combined_text)

print(f'The Output for all pages has been saved to:', output_file)