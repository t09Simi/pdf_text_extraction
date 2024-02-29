import json
import boto3
from io import BytesIO
from openpyxl import load_workbook

s3 = boto3.client('s3')


def find_last_row(sheet):
    # last_row = sheet.max_row
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row):
        if row[0].value is None:
            return row[0].row - 1  # Return the row before the first empty cell
    return sheet.max_row


def save_workbook_to_s3(workbook, bucket, key):
    # Save the modified workbook to bytes
    print("<--------------Updating excel------------------>")
    buffer = BytesIO()
    workbook.save(buffer)
    # Upload the modified Excel file back to S3, overwriting the original file
    s3.put_object(Body=buffer.getvalue(), Bucket=bucket, Key=key)


def update_excel(extracted_data: dict, sheet_name: str, column_mapping: dict):
    # Specify the S3 bucket and file name
    excel_bucket = 'resources-and-extraction-data'
    excel_file_key = 'Extraction_data.xlsx'

    # Download the Excel file from S3
    excel_file_obj = s3.get_object(Bucket=excel_bucket, Key=excel_file_key)
    excel_file_content = excel_file_obj['Body'].read()

    # Load the workbook from the downloaded content
    workbook = load_workbook(filename=BytesIO(excel_file_content))
    # Select the specified sheet
    sheet = workbook[sheet_name]
    # Find the last row with data in column A
    last_row = find_last_row(sheet) + 1
    for key, data in extracted_data.items():
        sheet[f"A{last_row}"] = key
        for cell_name, value in data.items():
            # Create cell address based on column name and current row
            column_name = column_mapping[cell_name]
            cell_address = f"{column_name}{last_row}"
            sheet[cell_address] = value
        last_row += 1
    save_workbook_to_s3(workbook, excel_bucket, excel_file_key)
    # Close the workbook
    workbook.close()
    print("<-------------- excel updated successfully------------------>")


def lambda_handler(event, context):
    # Parse the payload from the event
    extracted_data = event['extracted_data']
    sheet_name = event['sheet_name']
    print(f"Received payload from other Lambda function. extracted_data: {extracted_data}, sheet_name: {sheet_name}")
    column_mapping = dict()
    column_mapping["Id Number"] = "A"
    column_mapping["Item Description"] = "D"
    column_mapping["SWL"] = "F"
    column_mapping["Certificate No"] = "H"
    column_mapping["Previous Inspection"] = "K"
    column_mapping["Provider Identification"] = "O"
    column_mapping["Next Inspection Due Date"] = "L"
    column_mapping["Manufacturer"] = "G"
    column_mapping["Model"] = "E"
    update_excel(extracted_data, sheet_name, column_mapping)
