import re
import pdfplumber
from openpyxl import load_workbook
from datetime import datetime

import excel_management

def get_manufacture(description: str):
    workbook = load_workbook("../database/Full_list_of_Manufacturers_and_Models.xlsx")
    manufacturer_sheet = workbook['Manufacture']
    model_sheet =workbook['Model']
    description_keyword =description.lower().split()
    manufacture, model = "",""
    for row in manufacturer_sheet.iter_rows(min_row=2, values_only=True):
        keyword = row[0]
        if keyword and str(keyword).lower() in description_keyword:
            manufacture = keyword
            break
    for row in model_sheet.iter_rows(min_row=2, values_only=True):
        keyword = row[0]
        if keyword and str(keyword).lower() in description_keyword:
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
    identification_number_list = []
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


def extract_quantity(text):
    match = re.match(r'^\s*(\d+)\s+\d*', text)
    if match:
        return int(match.group(1))
    else:
        return None


def extraction_centurion_pdf(pdf_path):
    print("<------------extracting centurion pdf------------>")
    pdf = pdfplumber.open(pdf_path)
    extraction_info = dict()
    for i in range(0, len(pdf.pages)):
        if i > 47:
            break

        page = pdf.pages[i]
        if page.extract_tables():
            print("page number:", i)
            page_tables = page.extract_tables()
            first_table = page_tables[0]
            # data1:Report Number / Date of Examination / Ref No
            table_data1 = first_table[0][13]
            # data2: Identify Company keywords
            table_data2 = first_table[1]
            # data3: Identify Text keywords
            table_data3 = first_table[2]
            # data4: Provide related Value
            table_data4 = first_table[4]
            # data5 ID number Value
            table_data5 = first_table[5]
            identification_number_list = list()
            if i < 27:

                id_numbers, description, wwl, next_thorough = None, None, None, None
                quantity = 1

                for index in range(0, len(table_data3)):
                    if table_data3[index] is None:
                        continue
                    text_to_compare = table_data3[index].lower()
                    if not description and "description" in text_to_compare:
                        description = table_data4[0].replace('\n', ' ')
                        serial_numbers = table_data5[index].split(":")[-1].strip()
                    elif not wwl and "wwl" in text_to_compare:
                        wwl = table_data4[index].strip()
                    elif not next_thorough and "next" in text_to_compare:
                        date_string = table_data4[index].strip()
                        date_obj = datetime.strptime(date_string, "%d/%m/%Y")
                        next_thorough = date_obj.strftime("%d/%m/%Y")
                    elif not id_numbers and "certificate" in text_to_compare:
                        id_numbers = table_data4[index].strip()


                if id_numbers:
                    page_info = dict()
                    if description:
                        page_info["Item Description"] = description.split(':')[0]
                        manufacturer, model = get_manufacture(description)
                        page_info["Manufacturer"] = manufacturer
                        page_info["Model"] = model
                    page_info["SWL"] = wwl
                    page_info["Next Inspection Due Date"] = next_thorough
                    # report_number, date_of_examination, job_number, next_date_of__examination = None, None, None, None
                    table_data1_mapping = dict()
                    table_data1 = table_data1.splitlines()

                    for data in table_data1:
                        data_list = data.split(':', 1)
                        if len(data_list) ==2:
                            key, value = data_list
                            formattted_key = key.lower().replace(" ", "").replace("/", "").replace(".", "")
                            table_data1_mapping[formattted_key] =value.strip()

                    page_info["Provider Identification"] = table_data1_mapping["custrefpono"]
                    page_info["Certificate No"] = page_info["Provider Identification"]
                    page_info["Previous Inspection"] = table_data1_mapping["dateofexamination"]

                    id_numbers = serial_numbers

                    if quantity == 1:
                        identification_number_list.append(id_numbers)

                    for identification_number in identification_number_list:
                        extraction_info[identification_number] = page_info

                else:
                    print("No identification error")
            else:
                quantity, id_numbers, description, wwl, next_thorough = None, None, None, None, None
                for index in range(0, len(table_data3)):
                    if table_data3[index] is None:
                        continue
                    text_to_compare = table_data3[index].lower()
                    if not description and "description" in text_to_compare:
                        description = table_data4[index].replace('\n', ' ')
                        serial_numbers = table_data5[index].split(":")[-1].strip()
                        quantity = extract_quantity(table_data4[index])
                    elif not wwl and "wwl" in text_to_compare:
                        wwl = table_data4[index].strip()
                    elif not next_thorough and "next" in text_to_compare:
                        date_string = table_data4[index].strip()
                        date_obj = datetime.strptime(date_string, "%d/%m/%Y")
                        next_thorough = date_obj.strftime("%d/%m/%Y")
                    elif not id_numbers and "certificate" in text_to_compare:
                        id_numbers = table_data4[index].strip()

                if id_numbers:
                    page_info = dict()
                    if description:
                        page_info["Item Description"] = description.split(':')[0]
                        manufacturer, model = get_manufacture(description)
                        page_info["Manufacturer"] = manufacturer
                        page_info["Model"] = model
                    page_info["SWL"] = wwl
                    page_info["Next Inspection Due Date"] = next_thorough
                    # report_number, date_of_examination, job_number, next_date_of__examination = None, None, None, None
                    table_data1_mapping = dict()
                    table_data1 = table_data1.splitlines()

                    for data in table_data1:
                        data_list = data.split(':', 1)
                        if len(data_list) ==2:
                            key, value = data_list
                            formattted_key = key.lower().replace(" ", "").replace("/", "").replace(".", "")
                            table_data1_mapping[formattted_key] =value.strip()

                    page_info["Provider Identification"] = table_data1_mapping["custrefpono"]
                    page_info["Certificate No"] = page_info["Provider Identification"]
                    page_info["Previous Inspection"] = table_data1_mapping["dateofexamination"]

                    id_numbers = serial_numbers

                    if quantity > 1:
                        identification_number_list = get_identification_number_list(id_numbers, quantity)

                    for identification_number in identification_number_list:
                        extraction_info[identification_number] = page_info

                    # print(identification_numbers, page_info)
                else:
                    print("No identification error")



    print(len(extraction_info.keys()))
    excel_management.update_excel(extraction_info, "Centurion")


if __name__ == "__main__":
    extraction_centurion_pdf("../resources/centurion.pdf")









