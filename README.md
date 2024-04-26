# Data extraction from PDF documents
## Overview
This project aims to extract text data from PDF documents and store it in an Excel format for further analysis and processing. The extraction process involves identifying relevant information such as equipment details, descriptions, and manufacturers from PDF files and structuring it into an Excel spreadsheet.


## Features
- PDF Processing: Utilizes the Pdfplumber library to extract text, images, and tables from PDF documents.
- Excel File Creation: Generates Excel files containing extracted data for easy access and analysis.
- AWS Integration: Utilizes Amazon Web Services (AWS) Lambda functions and Simple Storage Service (S3) buckets for efficient processing and storage of PDF files and extracted data.
- Security: Implements encryption mechanisms for data at rest and access controls to ensure data confidentiality and integrity.
- Scalability: Utilizes serverless architecture with AWS Lambda for scalable and cost-effective PDF processing.
## Vision
Once complete, our application aims to revolutionize the way PDF files are processed and managed. With advanced text extraction capabilities, seamless integration with Amazon S3, and robust Excel file generation, our vision is to provide users with a comprehensive solution for extracting and managing data from PDF documents. 

In the case of Intebloc clients, the application has significantly reduced onboarding time by 90%, streamlining the process of integrating new clients into the system and facilitating quicker
access to information
## Requirements
All the packages and libraries required for this application to run can be found in requirements.txt file.
## Building the application
You need to clone/download the application  

```bash
git clone "repo_link"
```
Run the below command to install pyenv

```bash
pyenv update
pyenv install 3.12 # to install the pyenv on your server.
```

You need to create a virtual environment. Set present location in termial to root directory of the project and then run the following commands to create and start the virtunal environment.  
```bash
pyenv local 3.12.0 # this sets the local version of python to 3.10.7
python3 -m venv .venv # this creates the virtual environment for you
source .venv/bin/activate # this activates the virtual environment
pip install --upgrade pip [ this is optional]  # this installs pip, and upgrades it if required.
```

To install the dependencies run the following command.
```bash
pip install -r requirements.txt
```
## Testing the build
To test the build please run the following commands. Each commands runs the implemented unittest related to respective pdf type.
```bash
python3 -m unittest ./src/test/sparrow_test.py 
python3 -m unittest ./src/test/centurion_test.py 
python3 -m unittest ./src/test/firstintegrated_test.py 

```
 
## Running the application
- Copy you target PDF file type (sparrow/centurion/firstintegrated) to resources folder.
- Next configure the path for this file in pdf_processing.py file present in src folder.
- Now the below command to run the application which starts the pdf extraction.
```bash
  cd src  
  python3 pdf_processing.py
```
- The final output file is generated in database folder with filename same as the pdf filename.
  
## Deployment
This application supports AWS deployment by leveraging AWS lambda service's serverless architecture. Please refer to the deployment.md file to know more about AWS deployment.

## Team Members
- Sairaj Naik Guguloth

- Aiswarya Jayasree Sasidharan
  
- Bangqi Liu
  
- Pengcheng Xu
  
- Qingyang Zeng

- Philip Chuka Ukwamedua

- Simi Mathai Simon
