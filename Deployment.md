# Data extraction from PDF documents
## Overview
This guide outlines the steps required to deploy the PDF extraction project on AWS. Before proceeding, ensure that you have the necessary AWS credentials and permissions to create and manage resources.

## Prerequisites
- AWS account with appropriate permissions.
- Basic Understanding of AWS Services. 
- AWS Management Console Knowledge.

## Set Up AWS Environment
### 1. Create S3 Buckets
Navigate to the Amazon S3 service in AWS Management Console and create the following S3 bucket to store PDF files, resources and extracted data.
- pdf-in-bucket
- pdf-out-bucket
- excel-extraction-data
- layers-lib
- resources-and-extraction-data


### 2.  Create Lambda layers
Navigate to the AWS Lambda service in AWS Management Console and create the following Lambda layers. Select and upload respective library zip file present in lambda_layers directory. Set python 3.12.0. as runtime environment and rest of them set it to default options.
- openpyxl_layer
- pdfplumber_layer

### 3. Create IAM Role
Navigate to the AWS IAM service in AWS Management Console and create an IAM role with following permissions.

- AmazonS3FullAccess
- AmazonS3ReadOnlyAccess
- AWSLambdaBasicExecutionRole
- AWSLambdaRole

### 4.  Prepare Lambda Function
Navigate to the AWS Lambda service in AWS Management Console and create the following Lambda functions. Select python 3.12.0. as runtime environment and rest of them set it to default options.
- pdf_processing.py
- sparrow_extraction.py
- centurion_extraction.py
- excel_management.py
- first_integrated.py

Plese follow the being steps for each lambda function.
1. __Code Upload:__
Copy the respective code from lambda_functions directory and paste it in the code part of the created lambda function.

2. __Code Deploy:__
Click on deploy button to deploy the code.

3. __Add Lambda Layers:__
Next click add layer button and add previously created two lambda layers to the functions.

4. __Configuration:__
Now go to configuration tab and select general configuration. Set the memory, ephemeral storage and timeout to maximum provided value.

5. __Permissions:__
Next go to permissions section in configuration tab and add the previously created IAM role to the functions.

### 5. Integrate SNS
Navigate to the AWS SNS service in AWS Management Console.
- Create a topic with all the default options.
- Add subscriber to this topic using mobile or email-id or both.

Copy the ARN of the topic and paste it in the excel_management.py lambda function code. Look for the comments in the code for directions to add.


### 6. Configure Trigger
Navigate to the AWS Lambda service in AWS Management Console
- Select pdf_processing lambda function
- Next go to triggers section in configuration tab
- Click on Add trigger button
- Select source to S3 and bucket as pdf-in-bucket

## Test Deployment

### 1. Upload PDF File
Navigate to the AWS S3 service in AWS Management Console
- Select pdf-in-bucket from the buckets list
- click on upload button
- Then click on add files button
- select the PDF files you want to upload
- Finally, click the upload button to start uploading


### 2. Logging and Monitoring
Navigate to the AWS Cloudwatch service in AWS Management Console
- Select log group under logs section
- Select the log group you want to check among the list
- Go to log stream tab and select the log stream file you want to open


### 3. Download generated excel file
Navigate to the AWS S3 service in AWS Management Console
- Select excel-extraction-data from the buckets list
- Navigate to the pdf uploaded date directory
- select the excel file that was created using the same name as the pdf filename.
- Click on download button to start downloading the excel file containing extracted data

## Team Members
- Sairaj Naik Guguloth

- Aiswarya Jayasree Sasidharan
  
- Bangqi Liu
  
- Pengcheng Xu
  
- Qingyang Zeng

- Philip Chuka Ukwamedua

- Simi Mathai Simon
