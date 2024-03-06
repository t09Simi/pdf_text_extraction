import re
import pdfplumber
from openpyxl import load_workbook
import excel_management

def split_id_numbers_with_range(id_numbers):
    new_id_numbers = []
    for id_number in id_numbers:

        match = re.match(r'([A-Za-z]+)(\d+)-(\d+)', id_number)
        if match:
            prefix_alpha = match.group(1)
            range_start = int(match.group(2))
            range_end = int(match.group(3))

        else:
            # Check if the second pattern matches
            match = re.match(r'([A-Za-z]+\d+/?\d*)/(\d+)-(\d+)', id_number)
            if match:
                prefix_alpha = match.group(1)
                prefix_numeric = match.group(2)
                range_start = int(prefix_numeric)
                range_end = int(match.group(3))
            else:
                # If neither pattern matches, append the id_number as is
                new_id_numbers.append(id_number)
                continue

        # Generate new id numbers based on the matched pattern
        for i in range(range_start, range_end + 1):
            new_id_numbers.append(f"{prefix_alpha}{i}")

    return new_id_numbers

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

def contains_keyword(first_row, keyword):
    #for row in table:
        if first_row is not None:
            row_text = ' '.join(cell if cell is not None else '' for cell in first_row)
            #print("Checking row:", first_row)
            if row_text.lower().startswith(keyword.lower()):
                #print("Keyword found in row:", row_text)
                return True
        return False

# Call to the First Integrated PDF
def extract_first_integrated_pdf(pdf_path):
    print("<------------extracting first_integrated pdf------------>")
    pdf_doc = pdfplumber.open(pdf_path)
    extraction_info = dict()
    for page_number, page in enumerate(pdf_doc.pages):
        page_tables = page.extract_tables()
        print("page number:", page_number)
        for table in page_tables:
            first_row = table[0]
            if contains_keyword(first_row, "Name &AddressofManufacturer"):
                process_table3(table, extraction_info)
            elif contains_keyword(first_row, "Date of Thorough Examination"):
                process_table2(table, extraction_info)
            else:
                process_table1(table, extraction_info)

    excel_management.update_excel(extraction_info, "First Integrated")


def process_table2(table, extraction_info):
    page_info = {}
    id_number = None
    keyword_date = "Date of Thorough Examination"
    for row_outer in table:
        #print("row_outer", row_outer)
        for cell in row_outer:
            if cell:
                # Extract information based on conditions
                if "identification of the equipment" in cell:
                    parts = cell.split('\n')
                    if len(parts) >= 2:
                        description = parts[1].strip()
                        id_number = parts[2].strip()
                        #page_info["Id Number"] = id_number
                        page_info["Item Description"] = description
                        id_number_list = []
                        if "-" in id_number:
                            split_id_numbers = split_id_numbers_with_range([id_number])
                            print("Split ID Numbers:", split_id_numbers)
                            id_number_list.extend(split_id_numbers)
                        else:
                            id_number_list.extend(id_number)
                        for id_number in id_number_list:
                            extraction_info[id_number] = page_info

                        #Get manufacture and model from description
                        manufacture, model = get_manufacture_model(description)
                        page_info["Manufacturer"] = manufacture
                        page_info["Model"] = model
                elif "WLL" in cell:
                    parts = cell.split('\n')
                    if len(parts) >= 2:
                        wll = parts[3].strip()
                        page_info["SWL"] = wll
                elif "ReportNumber" in cell:
                    certificate_number = cell.replace("ReportNumber:", "").strip()
                    page_info["Certificate No"] = certificate_number
                elif "Date of Thorough" in cell:
                    previous_inspection = cell.replace("Date of Thorough Examination:", "").strip()
                    page_info["Previous Inspection"] = previous_inspection
                elif "Latest date by which next" in cell:
                    next_inspection = cell.replace(
                        "Latest date by which next thorough\nexamination must be carried out:", "").strip()
                    page_info["Next Inspection Due Date"] = next_inspection


def process_table1(table, extraction_info):
    page_info = {}
    id_number = None
    id_pattern = re.compile(r'(\(\w+\)\s*)?[A-Z]{3}\d+')
    swl_pattern = re.compile(r'(\d+\.\d+)\s+(?:Tonnes|Kilos)\s', re.IGNORECASE)

    for row in table:
        row_text = ' '.join(cell if cell is not None else '' for cell in row)

        # Extract ID Number
        id_match = id_pattern.search(row_text)
        if id_match:
            id_number = id_match.group()
            #page_info["Id Number"] = id_number

        # Extract SWL
        swl_match = swl_pattern.search(row_text)
        if swl_match:
            page_info["SWL"] = swl_match.group().strip()
           # print("SWL:", swl_match)

        # Extract Description between ID Number and SWL
        if id_match:
            start_index = id_match.end()
            if swl_match:
                end_index = swl_match.start()
            else:
                end_index = -1
            description = row_text[start_index:end_index].strip()
            description = re.sub(r'\([^)]*\)\s*', '', description)
            page_info["Item Description"] = description

            #Get manufacture and model fro  description
            manufacture, model = get_manufacture_model(description)
            page_info["Manufacturer"] = manufacture
            page_info["Model"] = model

        # Extract Date of Next Inspection
        next_inspection_pattern = re.compile(r'(\d{2}/\d{2}/\d{4})$', re.IGNORECASE)
        next_inspection_match = next_inspection_pattern.search(row_text)
        if next_inspection_match:
            page_info["Next Inspection Due Date"] = next_inspection_match.group(1).strip()
            #print("Next Inspection Date:", next_inspection_match)

        # Extract Date of previous Inspection
        keyword_date = "Date of thorough examination"
        if keyword_date.lower() in row_text.lower():
            #print("Keyword found in row text:", row_text)
            date_pattern = re.compile(r'\b(\d{2}/\d{2}/\d{4})\b')
            previous_inspection_match = date_pattern.search(row_text)
            if previous_inspection_match:
                page_info["Previous Inspection"] = previous_inspection_match.group().strip()
                #print("Previous Inspection Date:", previous_inspection_match)


        #Extract Report No
        keyword_report = "Report Ref No"
        if keyword_report.lower() in row_text.lower():
            report_pattern = re.compile(r'[A-Z]{3}/\d{6}/\d{5}')
            report_ref_match = report_pattern.search(row_text)
            if report_ref_match:
                page_info["Certificate No"] = report_ref_match.group().strip()
                #print("Certificate Number", certificate_number)

        if id_number:
            extraction_info[id_number] = page_info

def process_table3(table,extraction_info):
    page_info = dict()
    id_number = None
    for i, row in enumerate(table):
        if row:
            # Check for labels
            for j, cell in enumerate(row):
                if cell:
                    # Check for Id Number label
                    if "Id Number" in cell:
                        # Get the value from the next row
                        id_number_row = table[i + 1] if i + 1 < len(table) else None
                        if id_number_row:
                            id_number = id_number_row[j].strip() if j < len(id_number_row) else None
                            id_number_list = []
                            if "-" in id_number:
                                split_id_numbers = split_id_numbers_with_range([id_number])
                                print("Split ID Numbers:", split_id_numbers)
                                id_number_list.extend(split_id_numbers)
                            else:
                                id_number_list.extend(id_number)
                            for id_number in id_number_list:
                                extraction_info[id_number] = page_info
                    # Check for Description label
                    elif "Description" in cell:
                        parts = cell.split('\n')
                        if len(parts) >= 2:
                            description = parts[1].strip()
                            page_info["Item Description"] = description
                    # Get manufacture and model from description
                            manufacture, model = get_manufacture_model(description)
                            page_info["Manufacturer"] = manufacture
                            page_info["Model"] = model
                    # Check for WLL label
                    elif "WLL" in cell:
                        # Get the value from the next row
                        wll_row = table[i + 1] if i + 1 < len(table) else None
                        if wll_row:
                            wll = wll_row[j].strip() if j < len(wll_row) else None
                            page_info["SWL"] = wll
                    elif "Certificate Number" in cell:
                        certificate_row = table[i + 1] if i + 1 < len(table) else None
                        if certificate_row:
                            certificate_number = certificate_row[j].strip() if j < len(
                                certificate_row) else None
                            page_info["Certificate No"] = certificate_number


if __name__ == "__main__":
    extract_first_integrated_pdf("../resources/First_Integrated.pdf")
    # get_identification_number_list("A1 to A6", 8)