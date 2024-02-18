import re
import pdfplumber
from openpyxl import load_workbook

import excel_management


def get_manufacture_model(description: str):
    workbook = load_workbook("../database/Full_list_of_Manufacturers_and_Models.xlsx")
    manufacturer_sheet = workbook['Manufacture']
    model_sheet = workbook['Model']
    description_keywords = description.lower().split()
    manufacture, model = "",""
    for row in manufacturer_sheet.iter_rows(min_row=2, values_only=True):
        keyword = row[0]
        if keyword and str(keyword).lower() in description_keywords:
            manufacture = keyword
            break
    for row in model_sheet.iter_rows(min_row=2, values_only=True):
        keyword = row[0]
        if keyword and str(keyword).lower() in description_keywords:
            model = keyword
            break
    return manufacture, model


def get_identification_parts_list(input_string: str, quantity: int):
    numeric_part = ''.join(filter(str.isdigit, input_string))
    part_list = list()
    if input_string[0].isdigit():
        alpha_part = input_string[len(numeric_part):]
        for i in range(0, quantity):
            part_list.append(f"{int(numeric_part) + i}{alpha_part}")
    else:
        alpha_part = input_string[:len(input_string)-len(numeric_part)]
        for i in range(0, quantity):
            part_list.append(f"{alpha_part}{int(numeric_part) + i}")
    return part_list


def get_identification_number_list(identification_numbers: str, quantity: int):
    #Take this as example (D971-1 to 6) or (MGL1 to MGL36)
    if "to" in identification_numbers.lower():
        # identification_number_first_part = D971-1 or MGL1
        identification_number_first_part = identification_numbers.split("to")[0].strip()
        # print(identification_number_first_part)
        # example: D971-1
        if "-" in identification_number_first_part:
            # first_part = D971, second_part = 1
            first_part, second_part = identification_number_first_part.split('-')
            # print(first_part, second_part)
            second_part_list = get_identification_parts_list(second_part, quantity)
            # print(second_part_list)
            identification_number_list = list()
            for second_part in second_part_list:
                identification_number_list.append(f"{first_part}-{second_part}")
        else:
            # example: MGL1
            identification_number_list = get_identification_parts_list(identification_number_first_part, quantity)
    elif re.search(r'x(\d+)', identification_numbers):
        # print(identification_numbers)
        identification_number_list = list()
        for num in identification_numbers.split(','):
            id_number = num.split('x')[0].strip()
            for i in range(len(id_number)-1, -1, -1):
                if id_number[i].isdigit():
                    id_number= id_number[:i+1]
                    break
            count = int(''.join(filter(str.isdigit, num.split('x')[1].strip())))
            for i in range(1,count+1):
                identification_number_list.append(f"{id_number}-{i}")
    elif "," in identification_numbers:
        identification_number_list = identification_numbers.split(',')

    # print(identification_number_list)
    return identification_number_list


def extract_sparrow_pdf(pdf_path):
    print("<------------extracting sparrow pdf------------>")
    pdf_doc = pdfplumber.open(pdf_path)
    extraction_info = dict()
    for i in range(0, len(pdf_doc.pages)):
        page = pdf_doc.pages[i]
        if page.extract_tables():
            print("page number:", i)
            page_tables = page.extract_tables()[0]
            table_data1 = page_tables[0][0].split('\n')
            table_data3 = page_tables[3]
            table_data4 = page_tables[4]
            identification_numbers, description, swl, quantity = None, None, None, None

            for index in range(0, len(table_data3)):
                if table_data3[index] is None:
                    continue
                text_to_compare = table_data3[index].lower()
                if not identification_numbers and  "identification" in text_to_compare:
                    identification_numbers = table_data4[index].strip()
                elif not description and "description" in text_to_compare:
                    description = table_data4[index].replace('\n', ' ')
                elif not swl and "swl" in text_to_compare:
                    swl = table_data4[index].strip()
                elif not quantity and "quantity" in text_to_compare:
                    quantity = int(float(table_data4[index]))
            if identification_numbers:
                page_info = dict()
                # page_info["Id Number"] = table_data4[0].strip()
                if description:
                    page_info["Item Description"] = description.split(':')[0]
                    manufacturer, model = get_manufacture_model(description)
                    page_info["Manufacturer"] = manufacturer
                    page_info["Model"] = model
                page_info["SWL"] = swl
                # report_number, date_of_examination, job_number, next_date_of__examination = None, None, None, None
                table_data1_mapping = dict()
                for data in table_data1:
                    data_list = data.split(':')
                    key = data_list[0].lower().replace(" ", "").strip()
                    value = data_list[-1].strip()
                    table_data1_mapping[key] = value
                page_info["Certificate No"] = table_data1_mapping["reportnumber"]
                page_info["Previous Inspection"] = table_data1_mapping["dateofthoroughexamination"]
                page_info["Provider Identification"] = "LOFT-" + table_data1_mapping["jobnumber"]
                page_info["Next Inspection Due Date"] = table_data1_mapping["duedateofnextthoroughexamination"]

                if quantity > 1:
                    identification_number_list = get_identification_number_list(identification_numbers, quantity)
                else:
                    identification_number_list = list()
                    identification_number_list.append(identification_numbers)

                for identification_number in identification_number_list:
                    extraction_info[identification_number] = page_info

                # print(identification_numbers, page_info)
            else:
                print("No identification error")
    print(len(extraction_info.keys()))
    excel_management.update_excel(extraction_info, "Sparrows")


if __name__ == "__main__":
    extract_sparrow_pdf("../resources/sparrows.pdf")
    # get_identification_number_list("A1 to A6", 8)
