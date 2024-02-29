import json
import boto3
import urllib.parse
import pdfplumber
from io import BytesIO

s3 = boto3.client('s3')
lambda_client = boto3.client('lambda')


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


def invoke_pdf_extraction_lambda(source_bucket, object_key, lambda_function):
    # Payload to pass to the second Lambda function
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


def lambda_handler(event, context):
    # Extracting bucket and object key from the S3 event
    source_bucket = event['Records'][0]['s3']['bucket']['name']
    object_key = urllib.parse.unquote_plus(event['Records'][0]['s3']['object']['key'])
    # Download the file from S3
    pdf_file = s3.get_object(Bucket=source_bucket, Key=object_key)
    file_content = pdf_file['Body'].read()
    text_content = pdf_to_text(file_content)
    keywords = ["Sparrows", "Centurion"]

    if is_empty(text_content):
        print(f"No text found")
    else:
        found_keyword = search_keyword(text_content, keywords)
        print(f"keyword found : {found_keyword}")
        if found_keyword and found_keyword == "Sparrows":
            invoke_pdf_extraction_lambda(source_bucket, object_key, 'sparrow_extraction')
        elif found_keyword and found_keyword == "Centurion":
            invoke_pdf_extraction_lambda(source_bucket, object_key, 'centurion_extraction')
        else:
            print(f"No matching keywords found")
