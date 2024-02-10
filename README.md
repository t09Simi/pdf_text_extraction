# Data extraction from PDF documents
A template repository for student software development teams to use in for coursework


## Vision
 Add something about what the application will do when more complete
  
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
pyenv install 3.10.7 # to install the pyenv on your server.
```

You need to create a virtual environment. Set present location in termial to root directory of the project and then run the following commands to create and start the virtunal environment.  
```bash
pyenv local 3.10.7 # this sets the local version of python to 3.10.7
python3 -m venv .venv # this creates the virtual environment for you
source .venv/bin/activate # this activates the virtual environment
pip install --upgrade pip [ this is optional]  # this installs pip, and upgrades it if required.
```

To install the dependencies run the following command.
```bash
pip install -r requirements.txt
```
## Testing the build
How do I test the code to ensure the build is correct?
  
## Running the application
 What do I do to deploy and/or run this?
  
## Team Members
 Sairaj Naik Guguloth

 Aiswarya Jayasree Sasidharan
  
 Bangqi Liu
  
 Pengcheng Xu
  
 Qingyang Zeng

 Philip Chuka Ukwamedua

 Simi Mathai Simon
