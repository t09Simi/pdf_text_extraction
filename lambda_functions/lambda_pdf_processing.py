import json
import boto3
import urllib.parse
import pdfplumber
from io import BytesIO
from datetime import datetime


s3 = boto3.client('s3')
lambda_client = boto3.client('lambda')
retries = 3

def move_pdf_to_out_bucket(source_bucket, object_key, file_content):
    target_bucket = 'pdf-out-bucket'
    # Get the current date in the format YYYY-MM-DD
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Check if the folder with the current date exists in the target S3 bucket
    target_folder_key = f'Failure/{current_date}/{object_key}'

    # Check if the folder exists
    existing_objects = s3.list_objects(Bucket=target_bucket, Prefix=f'Failure/{current_date}/')

    if 'Contents' not in existing_objects:
        # Folder doesn't exist, create it
        s3.put_object(Body='', Bucket=target_bucket, Key=f'Failure/{current_date}/')

    # Upload the file to the target S3 bucket and folder
    s3.put_object(Body=file_content, Bucket=target_bucket, Key=target_folder_key)
    print(f"File moved to the target bucket: {target_bucket}, Folder: Failure/{current_date}")

    # Delete the original file from the source bucket (uncomment if needed)
    s3.delete_object(Bucket=source_bucket, Key=object_key)
    print(f"File deleted from source bucket: {source_bucket}, file: {object_key}")


def pdf_to_text(file_content):
    doc = pdfplumber.open(BytesIO(file_content))
    text = ""
    page = doc.pages[0]
    text += page.extract_text()
    return text


def search_keyword(text, keywords):
    for keyword in keywords:
        if keyword.lower() in text.lower():
            return keyword
    return None  # Return None if no keyword is found


def is_empty(text):
    return not bool(text.strip())


def invoke_pdf_extraction_lambda(source_bucket, object_key, lambda_function, file_content):
    # Payload to pass to the second Lambda function
    global retries
    payload = {
        'source_bucket': source_bucket,
        'object_key': object_key
    }

    # Invoke the second Lambda function asynchronously
    response = lambda_client.invoke(
        FunctionName=lambda_function,
        InvocationType='Event',
        Payload=json.dumps(payload)
    )

    # Optionally, you can check the response for success or handle errors
    status_code = response['StatusCode']
    if status_code == 202:
        print(f" Lambda function {lambda_function} invoked successfully.")
    else:
        print(f"Error invoking Lambda function {lambda_function}. Status code: {status_code}")
        if retries:
            retries -= 1
            print(f"Retrying to invoking Lambda function {lambda_function}")
            invoke_pdf_extraction_lambda(source_bucket, object_key, lambda_function, file_content)
        else:
            move_pdf_to_out_bucket(source_bucket, object_key, file_content)


def lambda_handler(event, context):
    try:
        # Extracting bucket and object key from the S3 event
        source_bucket = event['Records'][0]['s3']['bucket']['name']
        object_key = urllib.parse.unquote_plus(event['Records'][0]['s3']['object']['key'])
        # Download the file from S3
        pdf_file = s3.get_object(Bucket=source_bucket, Key=object_key)
        file_content = pdf_file['Body'].read()
        text_content = pdf_to_text(file_content)
        keywords = ["Sparrows", "Centurion"]

        if is_empty(text_content):
            print(f"No matching keyword found with existing clients, No procesing it, pushing this file to out failure folder")
            move_pdf_to_out_bucket(source_bucket, object_key, file_content)
        else:
            found_keyword = search_keyword(text_content, keywords)
            print(f"keyword found: {found_keyword}")
            if found_keyword and found_keyword == "Sparrows":
                invoke_pdf_extraction_lambda(source_bucket, object_key, 'sparrow_extraction', file_content)
            elif found_keyword and found_keyword == "Centurion":
                invoke_pdf_extraction_lambda(source_bucket, object_key, 'centurion_extraction', file_content)
            else:
                print(
                    f"No matching keyword found with existing clients, No procesing it, pushing this file to out failure folder")
                move_pdf_to_out_bucket(source_bucket, object_key, file_content)
    except Exception as e:
        print("Error in processing the PDF file:", e)
